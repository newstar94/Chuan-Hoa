using System.Security.Cryptography;
using System.Text;
using Npgsql;
using ChuanHoa.Contracts;

namespace ChuanHoa.Infrastructure.Admin;

public sealed class PostgresAccountAccessStore(NpgsqlDataSource dataSource, TimeProvider clock) : IAccountAccessStore, IActivationKeyAdminStore
{
    private const int PasswordIterations = 600_000;

    public async Task<AccountSessionResponse> RegisterAsync(RegisterAccountRequest request, CancellationToken cancellationToken)
    {
        var email = NormalizeEmail(request.Email);
        ValidatePassword(request.Password);
        ValidateDevice(request.DeviceThumbprint);
        var now = clock.GetUtcNow();
        var userId = Guid.NewGuid();
        await using var connection = await dataSource.OpenConnectionAsync(cancellationToken);
        await using var transaction = await connection.BeginTransactionAsync(cancellationToken);
        await using (var insert = new NpgsqlCommand("""
            INSERT INTO users (id, normalized_email, display_name, status, password_hash, created_at_utc, updated_at_utc)
            VALUES ($1, $2, $3, 'ACTIVE', $4, $5, $5);
            """, connection, transaction))
        {
            insert.Parameters.AddWithValue(userId);
            insert.Parameters.AddWithValue(email);
            insert.Parameters.AddWithValue(string.IsNullOrWhiteSpace(request.DisplayName) ? email : request.DisplayName.Trim());
            insert.Parameters.AddWithValue(PasswordHash(request.Password));
            insert.Parameters.AddWithValue(now);
            await insert.ExecuteNonQueryAsync(cancellationToken);
        }
        var session = await BindAndCreateSessionAsync(connection, transaction, userId, email, request.DisplayName, request.DeviceThumbprint, now, cancellationToken);
        await transaction.CommitAsync(cancellationToken);
        return session;
    }

    public async Task<AccountSessionResponse> LoginAsync(LoginAccountRequest request, CancellationToken cancellationToken)
    {
        var email = NormalizeEmail(request.Email);
        ValidateDevice(request.DeviceThumbprint);
        await using var connection = await dataSource.OpenConnectionAsync(cancellationToken);
        await using var transaction = await connection.BeginTransactionAsync(cancellationToken);
        await using var lookup = new NpgsqlCommand("SELECT id, display_name, password_hash, status FROM users WHERE normalized_email = $1 FOR UPDATE", connection, transaction);
        lookup.Parameters.AddWithValue(email);
        await using var reader = await lookup.ExecuteReaderAsync(cancellationToken);
        if (!await reader.ReadAsync(cancellationToken) || !string.Equals(reader.GetString(3), "ACTIVE", StringComparison.Ordinal))
            throw new UnauthorizedAccessException("EMAIL_OR_PASSWORD_INVALID");
        var userId = reader.GetGuid(0);
        var displayName = reader.GetString(1);
        var storedHash = reader.IsDBNull(2) ? string.Empty : reader.GetString(2);
        if (!VerifyPassword(request.Password, storedHash)) throw new UnauthorizedAccessException("EMAIL_OR_PASSWORD_INVALID");
        await reader.DisposeAsync();
        var session = await BindAndCreateSessionAsync(connection, transaction, userId, email, displayName, request.DeviceThumbprint, clock.GetUtcNow(), cancellationToken);
        await transaction.CommitAsync(cancellationToken);
        return session;
    }

    public async Task LogoutAsync(string token, CancellationToken cancellationToken)
    {
        var hash = HashToken(token);
        await using var command = dataSource.CreateCommand("""
            WITH current_session AS (
                SELECT user_id, device_thumbprint FROM account_sessions WHERE token_hash = $1 FOR UPDATE
            ), revoked AS (
                UPDATE account_sessions s SET revoked_at_utc = now()
                FROM current_session c
                WHERE s.user_id = c.user_id AND s.device_thumbprint = c.device_thumbprint AND s.revoked_at_utc IS NULL
                RETURNING s.user_id, s.device_thumbprint
            )
            UPDATE account_device_bindings b SET released_at_utc = now(), release_reason = 'logout'
            FROM current_session c
            WHERE b.user_id = c.user_id AND b.device_thumbprint = c.device_thumbprint AND b.released_at_utc IS NULL
            """);
        command.Parameters.AddWithValue(hash);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task<AccountSessionResponse?> GetSessionAsync(string token, CancellationToken cancellationToken)
    {
        var hash = HashToken(token);
        await using var command = dataSource.CreateCommand("""
            SELECT s.user_id, u.normalized_email, u.display_name, s.issued_at_utc, s.expires_at_utc, s.device_thumbprint
            FROM account_sessions s JOIN users u ON u.id = s.user_id
            WHERE s.token_hash = $1 AND s.revoked_at_utc IS NULL AND s.expires_at_utc > now() AND u.status = 'ACTIVE'
            """);
        command.Parameters.AddWithValue(hash);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        if (!await reader.ReadAsync(cancellationToken)) return null;
        return new AccountSessionResponse(reader.GetGuid(0), reader.GetString(1), reader.GetString(2), reader.GetFieldValue<DateTimeOffset>(3), reader.GetFieldValue<DateTimeOffset>(4), reader.GetString(5), token);
    }

    public async Task<VipActivationResponse> ActivateVipKeyAsync(ActivateVipKeyRequest request, CancellationToken cancellationToken)
    {
        ValidateDevice(request.DeviceThumbprint);
        var normalized = NormalizeKey(request.ActivationKey);
        var hash = SHA256.HashData(Encoding.UTF8.GetBytes(normalized));
        var now = clock.GetUtcNow();
        await using var connection = await dataSource.OpenConnectionAsync(cancellationToken);
        await using var transaction = await connection.BeginTransactionAsync(cancellationToken);
        await using var command = new NpgsqlCommand("SELECT id, key_prefix, feature_codes, max_devices, expires_at_utc, status FROM activation_keys WHERE key_hash = $1 FOR UPDATE", connection, transaction);
        command.Parameters.AddWithValue(hash);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        if (!await reader.ReadAsync(cancellationToken)) throw new UnauthorizedAccessException("ACTIVATION_KEY_INVALID");
        var keyId = reader.GetGuid(0); var prefix = reader.GetString(1); var features = reader.GetFieldValue<string[]>(2);
        var max = reader.GetInt32(3); var expiry = reader.IsDBNull(4) ? (DateTimeOffset?)null : reader.GetFieldValue<DateTimeOffset>(4);
        var status = reader.GetString(5); await reader.DisposeAsync();
        if (!string.Equals(status, "ACTIVE", StringComparison.Ordinal) || expiry <= now) throw new UnauthorizedAccessException(expiry <= now ? "ACTIVATION_KEY_EXPIRED" : "ACTIVATION_KEY_REVOKED");
        await using var existing = new NpgsqlCommand("SELECT released_at_utc IS NULL FROM activation_key_devices WHERE activation_key_id = $1 AND device_thumbprint = $2", connection, transaction);
        existing.Parameters.AddWithValue(keyId); existing.Parameters.AddWithValue(request.DeviceThumbprint);
        var sameDevice = await existing.ExecuteScalarAsync(cancellationToken);
        if (sameDevice is bool active && active)
        {
            await transaction.CommitAsync(cancellationToken);
            return await ActivationSummaryAsync(connection, keyId, prefix, features, max, expiry, cancellationToken);
        }
        await using var count = new NpgsqlCommand("SELECT count(*) FROM activation_key_devices WHERE activation_key_id = $1 AND released_at_utc IS NULL", connection, transaction);
        count.Parameters.AddWithValue(keyId);
        if ((long)(await count.ExecuteScalarAsync(cancellationToken) ?? 0L) >= max) throw new InvalidOperationException("ACTIVATION_KEY_DEVICE_LIMIT");
        await using var insert = new NpgsqlCommand("""
            INSERT INTO activation_key_devices (id, activation_key_id, device_thumbprint, activated_at_utc)
            VALUES ($1, $2, $3, $4)
            ON CONFLICT (activation_key_id, device_thumbprint) DO UPDATE SET released_at_utc = NULL, release_reason = NULL
            """, connection, transaction);
        insert.Parameters.AddWithValue(Guid.NewGuid()); insert.Parameters.AddWithValue(keyId); insert.Parameters.AddWithValue(request.DeviceThumbprint); insert.Parameters.AddWithValue(now);
        await insert.ExecuteNonQueryAsync(cancellationToken);
        var result = await ActivationSummaryAsync(connection, keyId, prefix, features, max, expiry, cancellationToken);
        await transaction.CommitAsync(cancellationToken);
        return result;
    }

    public async Task<CreatedVipKeyResponse> CreateVipKeyAsync(CreateVipKeyRequest request, CancellationToken cancellationToken)
    {
        if (request.MaxDevices is < 1 or > 100_000 || request.FeatureCodes is null || request.FeatureCodes.Length == 0)
            throw new ArgumentException("VIP_KEY_CONFIGURATION_INVALID");
        if (request.ExpiresAtUtc is not null && request.ExpiresAtUtc <= clock.GetUtcNow()) throw new ArgumentException("VIP_KEY_EXPIRY_INVALID");
        var raw = GenerateVipKey();
        var normalized = NormalizeKey(raw);
        var hash = SHA256.HashData(Encoding.UTF8.GetBytes(normalized));
        var id = Guid.NewGuid();
        await using var connection = await dataSource.OpenConnectionAsync(cancellationToken);
        await using var transaction = await connection.BeginTransactionAsync(cancellationToken);
        await using (var insert = new NpgsqlCommand("""
            INSERT INTO activation_keys (id, key_prefix, key_hash, status, product_id, feature_codes, max_devices, created_at_utc, created_by, created_by_external_actor, expires_at_utc, note)
            VALUES ($1, $2, $3, 'ACTIVE', $4, $5, $6, $7, $8, $9, $10, $11)
            """, connection, transaction))
        {
            insert.Parameters.AddWithValue(id); insert.Parameters.AddWithValue(raw[..8]); insert.Parameters.AddWithValue(hash);
            insert.Parameters.AddWithValue(request.ProductId); insert.Parameters.AddWithValue(request.FeatureCodes); insert.Parameters.AddWithValue(request.MaxDevices);
            insert.Parameters.AddWithValue(clock.GetUtcNow());
            // Bidding and Chuẩn Hóa accounts are intentionally independent;
            // the external actor is recorded in integration audit, not as a
            // local users FK.
            insert.Parameters.AddWithValue(DBNull.Value);
            insert.Parameters.AddWithValue(request.ActorId);
            insert.Parameters.AddWithValue((object?)request.ExpiresAtUtc ?? DBNull.Value); insert.Parameters.AddWithValue(request.Note ?? string.Empty);
            await insert.ExecuteNonQueryAsync(cancellationToken);
        }
        await using (var audit = new NpgsqlCommand("""
            INSERT INTO admin_integration_audit (event_id, occurred_at_utc, source_application, external_actor_id, target_type, target_id, action_code, result_code, correlation_id, metadata)
            VALUES ($1, $2, 'biddingflow', $3, 'activation_key', $4, 'admin.activation_key.create', 'success', $5, $6::jsonb)
            """, connection, transaction))
        {
            audit.Parameters.AddWithValue(Guid.NewGuid()); audit.Parameters.AddWithValue(clock.GetUtcNow()); audit.Parameters.AddWithValue(request.ActorId);
            audit.Parameters.AddWithValue(id.ToString("D")); audit.Parameters.AddWithValue(request.CorrelationId);
            audit.Parameters.AddWithValue("{\"keyPrefix\":\"" + raw[..8] + "\",\"maxDevices\":" + request.MaxDevices + "}");
            await audit.ExecuteNonQueryAsync(cancellationToken);
        }
        await transaction.CommitAsync(cancellationToken);
        return new CreatedVipKeyResponse(id, raw, raw[..8], request.MaxDevices, request.ExpiresAtUtc);
    }

    public Task ReleaseKeyDeviceAsync(AdminDeviceMutationRequest request, CancellationToken cancellationToken)
        => ExecuteAdminMutationAsync("UPDATE activation_key_devices SET released_at_utc = $3, release_reason = 'admin_release' WHERE activation_key_id = $1 AND device_thumbprint = $2 AND released_at_utc IS NULL", request.TargetId, request.DeviceThumbprint, request.ActorId, request.CorrelationId, "activation_device", cancellationToken);

    public Task RevokeKeyAsync(AdminKeyMutationRequest request, CancellationToken cancellationToken)
        => ExecuteAdminMutationAsync("UPDATE activation_keys SET status = 'REVOKED', revoked_at_utc = $2 WHERE id = $1 AND status = 'ACTIVE'", request.KeyId, string.Empty, request.ActorId, request.CorrelationId, "activation_key", cancellationToken);

    public Task ResetAccountDeviceAsync(AdminDeviceMutationRequest request, CancellationToken cancellationToken)
        => ExecuteAdminMutationAsync("UPDATE account_device_bindings SET released_at_utc = $3, release_reason = 'admin_reset' WHERE user_id = $1 AND device_thumbprint = $2 AND released_at_utc IS NULL", request.TargetId, request.DeviceThumbprint, request.ActorId, request.CorrelationId, "account_device", cancellationToken);

    public async Task<ActivationKeyAdminPage> ListAsync(string? search, int page, int pageSize, CancellationToken cancellationToken)
    {
        page = Math.Max(1, page); pageSize = Math.Clamp(pageSize, 1, 100); search = string.IsNullOrWhiteSpace(search) ? null : search.Trim();
        await using var connection = await dataSource.OpenConnectionAsync(cancellationToken);
        await using var count = new NpgsqlCommand("SELECT count(*) FROM activation_keys WHERE ($1 IS NULL OR key_prefix ILIKE '%' || $1 || '%' OR status ILIKE '%' || $1 || '%' OR created_by_external_actor ILIKE '%' || $1 || '%')", connection);
        count.Parameters.AddWithValue((object?)search ?? DBNull.Value);
        var total = Convert.ToInt64(await count.ExecuteScalarAsync(cancellationToken));
        await using var command = new NpgsqlCommand("""
            SELECT k.id, k.key_prefix, k.status, k.product_id, k.max_devices,
                   count(d.id) FILTER (WHERE d.released_at_utc IS NULL), k.created_at_utc,
                   k.expires_at_utc, k.created_by_external_actor, k.note
            FROM activation_keys k LEFT JOIN activation_key_devices d ON d.activation_key_id = k.id
            WHERE ($1 IS NULL OR k.key_prefix ILIKE '%' || $1 || '%' OR k.status ILIKE '%' || $1 || '%' OR k.created_by_external_actor ILIKE '%' || $1 || '%')
            GROUP BY k.id ORDER BY k.created_at_utc DESC OFFSET $2 LIMIT $3
            """, connection);
        command.Parameters.AddWithValue((object?)search ?? DBNull.Value); command.Parameters.AddWithValue((page - 1) * pageSize); command.Parameters.AddWithValue(pageSize);
        var items = new List<ActivationKeyAdminView>();
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            var used = reader.GetInt64(5); var max = reader.GetInt32(4);
            items.Add(new ActivationKeyAdminView(reader.GetGuid(0), reader.GetString(1), reader.GetString(2), reader.GetGuid(3), max, (int)used, max - (int)used, reader.GetFieldValue<DateTimeOffset>(6), reader.IsDBNull(7) ? null : reader.GetFieldValue<DateTimeOffset>(7), reader.IsDBNull(8) ? null : reader.GetString(8), reader.IsDBNull(9) ? null : reader.GetString(9)));
        }
        return new ActivationKeyAdminPage(items, page, pageSize, total);
    }

    private async Task ExecuteAdminMutationAsync(string sql, Guid targetId, string deviceOrActor, string actorId, Guid correlationId, string targetType, CancellationToken cancellationToken)
    {
        await using var connection = await dataSource.OpenConnectionAsync(cancellationToken);
        await using var transaction = await connection.BeginTransactionAsync(cancellationToken);
        await using var command = new NpgsqlCommand(sql, connection, transaction);
        command.Parameters.AddWithValue(targetId);
        if (sql.Contains("device_thumbprint", StringComparison.Ordinal)) { command.Parameters.AddWithValue(deviceOrActor); command.Parameters.AddWithValue(clock.GetUtcNow()); }
        else { command.Parameters.AddWithValue(clock.GetUtcNow()); }
        await command.ExecuteNonQueryAsync(cancellationToken);
        await using var audit = new NpgsqlCommand("INSERT INTO admin_integration_audit (event_id, occurred_at_utc, source_application, external_actor_id, target_type, target_id, action_code, result_code, correlation_id, metadata) VALUES ($1, $2, 'biddingflow', $3, $4, $5, $6, 'success', $7, '{}'::jsonb)", connection, transaction);
        audit.Parameters.AddWithValue(Guid.NewGuid()); audit.Parameters.AddWithValue(clock.GetUtcNow()); audit.Parameters.AddWithValue(actorId); audit.Parameters.AddWithValue(targetType); audit.Parameters.AddWithValue(targetId.ToString("D")); audit.Parameters.AddWithValue(targetType == "activation_key" ? "admin.activation_key.revoke" : targetType == "activation_device" ? "admin.activation_device.release" : "admin.account_device.reset"); audit.Parameters.AddWithValue(correlationId);
        await audit.ExecuteNonQueryAsync(cancellationToken);
        await transaction.CommitAsync(cancellationToken);
    }

    private static async Task<VipActivationResponse> ActivationSummaryAsync(NpgsqlConnection connection, Guid keyId, string prefix, string[] features, int max, DateTimeOffset? expiry, CancellationToken cancellationToken)
    {
        await using var count = new NpgsqlCommand("SELECT count(*) FROM activation_key_devices WHERE activation_key_id = $1 AND released_at_utc IS NULL", connection);
        count.Parameters.AddWithValue(keyId);
        var used = Convert.ToInt32(await count.ExecuteScalarAsync(cancellationToken));
        return new VipActivationResponse(keyId, prefix, new HashSet<string>(features, StringComparer.Ordinal), max, used, max - used, expiry);
    }

    private async Task<AccountSessionResponse> BindAndCreateSessionAsync(NpgsqlConnection connection, NpgsqlTransaction transaction, Guid userId, string email, string displayName, string device, DateTimeOffset now, CancellationToken cancellationToken)
    {
        await using var binding = new NpgsqlCommand("SELECT device_thumbprint, released_at_utc FROM account_device_bindings WHERE user_id = $1 FOR UPDATE", connection, transaction);
        binding.Parameters.AddWithValue(userId);
        await using var reader = await binding.ExecuteReaderAsync(cancellationToken);
        if (await reader.ReadAsync(cancellationToken))
        {
            var current = reader.GetString(0); var released = reader.IsDBNull(1) ? (DateTimeOffset?)null : reader.GetFieldValue<DateTimeOffset>(1);
            if (!string.Equals(current, device, StringComparison.Ordinal) && released is null)
            {
                await reader.DisposeAsync();
                await using var active = new NpgsqlCommand("SELECT EXISTS (SELECT 1 FROM account_sessions WHERE user_id = $1 AND device_thumbprint = $2 AND revoked_at_utc IS NULL AND expires_at_utc > now())", connection, transaction);
                active.Parameters.AddWithValue(userId); active.Parameters.AddWithValue(current);
                if (await active.ExecuteScalarAsync(cancellationToken) is true) throw new InvalidOperationException("ACCOUNT_DEVICE_IN_USE");
                await using var release = new NpgsqlCommand("UPDATE account_device_bindings SET released_at_utc = $2, release_reason = 'session_expired' WHERE user_id = $1 AND released_at_utc IS NULL", connection, transaction);
                release.Parameters.AddWithValue(userId); release.Parameters.AddWithValue(now); await release.ExecuteNonQueryAsync(cancellationToken);
            }
        }
        if (!reader.IsClosed) await reader.DisposeAsync();
        await using var upsert = new NpgsqlCommand("""
            INSERT INTO account_device_bindings (user_id, device_thumbprint, bound_at_utc, last_seen_at_utc, released_at_utc)
            VALUES ($1, $2, $3, $3, NULL)
            ON CONFLICT (user_id) DO UPDATE SET device_thumbprint = EXCLUDED.device_thumbprint, bound_at_utc = EXCLUDED.bound_at_utc, last_seen_at_utc = EXCLUDED.last_seen_at_utc, released_at_utc = NULL
            """, connection, transaction);
        upsert.Parameters.AddWithValue(userId); upsert.Parameters.AddWithValue(device); upsert.Parameters.AddWithValue(now); await upsert.ExecuteNonQueryAsync(cancellationToken);
        var rawToken = Convert.ToBase64String(RandomNumberGenerator.GetBytes(32));
        await using var insert = new NpgsqlCommand("INSERT INTO account_sessions (id, user_id, device_thumbprint, token_hash, issued_at_utc, expires_at_utc, last_seen_at_utc) VALUES ($1, $2, $3, $4, $5, $5 + interval '72 hours', $5)", connection, transaction);
        insert.Parameters.AddWithValue(Guid.NewGuid()); insert.Parameters.AddWithValue(userId); insert.Parameters.AddWithValue(device); insert.Parameters.AddWithValue(HashToken(rawToken)); insert.Parameters.AddWithValue(now); await insert.ExecuteNonQueryAsync(cancellationToken);
        return new AccountSessionResponse(userId, email, displayName, now, now.AddHours(72), device, rawToken);
    }

    private static string NormalizeEmail(string email) => string.IsNullOrWhiteSpace(email) ? throw new ArgumentException("EMAIL_REQUIRED") : email.Trim().ToUpperInvariant();
    private static string NormalizeKey(string key) => string.IsNullOrWhiteSpace(key) ? throw new ArgumentException("ACTIVATION_KEY_REQUIRED") : key.Replace("-", string.Empty, StringComparison.Ordinal).Replace(" ", string.Empty, StringComparison.Ordinal).ToUpperInvariant();
    private static string GenerateVipKey()
    {
        var bytes = RandomNumberGenerator.GetBytes(20);
        var alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
        var chars = new char[20];
        for (var i = 0; i < chars.Length; i++) chars[i] = alphabet[bytes[i] % alphabet.Length];
        return $"CHV-{new string(chars, 0, 4)}-{new string(chars, 4, 4)}-{new string(chars, 8, 4)}-{new string(chars, 12, 4)}-{new string(chars, 16, 4)}";
    }
    private static void ValidateDevice(string device) { if (string.IsNullOrWhiteSpace(device) || device.Length > 256) throw new ArgumentException("DEVICE_INVALID"); }
    private static void ValidatePassword(string password) { if (string.IsNullOrEmpty(password) || password.Length < 12) throw new ArgumentException("PASSWORD_POLICY_FAILED"); }
    private static byte[] HashToken(string token) => SHA256.HashData(Encoding.UTF8.GetBytes(token));
    private static string PasswordHash(string password)
    {
        var salt = RandomNumberGenerator.GetBytes(16); var hash = Rfc2898DeriveBytes.Pbkdf2(password, salt, PasswordIterations, HashAlgorithmName.SHA256, 32);
        return $"pbkdf2-sha256${PasswordIterations}${Convert.ToBase64String(salt)}${Convert.ToBase64String(hash)}";
    }
    private static bool VerifyPassword(string password, string encoded)
    {
        var parts = encoded.Split('$');
        if (parts.Length != 4 || !int.TryParse(parts[1], out var iterations) || !Convert.TryFromBase64String(parts[2], new byte[32], out _)) return false;
        try { var salt = Convert.FromBase64String(parts[2]); var expected = Convert.FromBase64String(parts[3]); var actual = Rfc2898DeriveBytes.Pbkdf2(password, salt, iterations, HashAlgorithmName.SHA256, expected.Length); return CryptographicOperations.FixedTimeEquals(actual, expected); } catch { return false; }
    }
}
