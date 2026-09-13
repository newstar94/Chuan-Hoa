using Npgsql;

namespace ChuanHoa.Infrastructure.Admin;

public sealed record AdminPage<T>(IReadOnlyList<T> Items, int Page, int PageSize, long Total);
public sealed record AdminAccount(Guid Id, string Email, string DisplayName, string Status, DateTimeOffset CreatedAtUtc);
public sealed record AdminOffer(Guid Id, Guid ProductId, string ProductCode, int Version, string Currency, long AmountMinor, string Status, DateTimeOffset EffectiveFromUtc, DateTimeOffset? EffectiveUntilUtc, IReadOnlyList<string> Features);
public sealed record AdminOrder(Guid Id, string StableOrderCode, Guid UserId, string Currency, long AmountMinor, string Status, DateTimeOffset CreatedAtUtc, DateTimeOffset UpdatedAtUtc);
public sealed record AdminSubscription(Guid Id, Guid UserId, Guid? OrganizationId, string ProductCode, string Source, string Status, DateTimeOffset EffectiveFromUtc, DateTimeOffset EffectiveUntilUtc, IReadOnlyList<string> Features);
public sealed record AdminPayment(Guid Id, Guid OrderId, string StableOrderCode, string Provider, string ProviderEventId, string EventType, bool SignatureValid, long? AmountMinor, string? Currency, DateTimeOffset ReceivedAtUtc, string? ProcessingResult);
public sealed record AdminMutation(Guid GrantId, Guid UserId, Guid ProductId, string Status, DateTimeOffset EffectiveFromUtc, DateTimeOffset EffectiveUntilUtc, IReadOnlyList<string> Features);
public sealed record AdminAudit(Guid EventId, DateTimeOffset OccurredAtUtc, string SourceApplication, string ExternalActorId, string TargetType, string TargetId, string ActionCode, string ResultCode, Guid CorrelationId, string MetadataJson);

public interface IIntegrationAdminStore
{
    Task<AdminPage<AdminAccount>> AccountsAsync(string? search, int page, int pageSize, CancellationToken cancellationToken);
    Task<AdminPage<AdminOffer>> OffersAsync(string? search, int page, int pageSize, CancellationToken cancellationToken);
    Task<AdminPage<AdminOrder>> OrdersAsync(string? search, int page, int pageSize, CancellationToken cancellationToken);
    Task<AdminPage<AdminSubscription>> SubscriptionsAsync(string? search, int page, int pageSize, CancellationToken cancellationToken);
    Task<AdminPage<AdminPayment>> PaymentsAsync(string? search, int page, int pageSize, CancellationToken cancellationToken);
    Task<AdminPage<AdminAudit>> AuditAsync(string? search, int page, int pageSize, CancellationToken cancellationToken);
    Task<AdminMutation> ExtendEntitlementAsync(Guid userId, Guid productId, IReadOnlyList<string> featureCodes, TimeSpan duration, string reason, string clientId, string idempotencyKey, string actorId, Guid correlationId, CancellationToken cancellationToken);
    Task RecordFailedEntitlementAsync(Guid userId, Guid productId, string clientId, string actorId, Guid correlationId, string resultCode, CancellationToken cancellationToken);
}

public sealed class PostgresIntegrationAdminStore(NpgsqlDataSource dataSource) : IIntegrationAdminStore
{
    public async Task RecordFailedEntitlementAsync(Guid userId, Guid productId, string clientId, string actorId, Guid correlationId, string resultCode, CancellationToken cancellationToken)
    {
        await using var command = dataSource.CreateCommand("INSERT INTO admin_integration_audit (event_id, occurred_at_utc, source_application, external_actor_id, target_type, target_id, action_code, result_code, correlation_id, metadata) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10::jsonb)");
        command.Parameters.AddWithValue(Guid.NewGuid());
        command.Parameters.AddWithValue(DateTimeOffset.UtcNow);
        command.Parameters.AddWithValue("biddingflow");
        command.Parameters.AddWithValue(actorId);
        command.Parameters.AddWithValue("entitlement");
        command.Parameters.AddWithValue($"{userId:D}:{productId:D}");
        command.Parameters.AddWithValue("admin.entitlement.extend");
        command.Parameters.AddWithValue(resultCode);
        command.Parameters.AddWithValue(correlationId);
        command.Parameters.AddWithValue($"{{\"clientId\":{System.Text.Json.JsonSerializer.Serialize(clientId)}}}");
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public Task<AdminPage<AdminAccount>> AccountsAsync(string? search, int page, int pageSize, CancellationToken cancellationToken) =>
        QueryAsync("""
            SELECT id, normalized_email, display_name, status::text, created_at_utc,
                   count(*) OVER() AS total
            FROM users
            WHERE ($1 = '' OR normalized_email ILIKE '%' || $1 || '%' OR display_name ILIKE '%' || $1 || '%')
            ORDER BY created_at_utc DESC, id DESC OFFSET $2 LIMIT $3;
            """, search, page, pageSize,
            r => new AdminAccount(r.GetGuid(0), r.GetString(1), r.GetString(2), r.GetString(3), r.GetFieldValue<DateTimeOffset>(4)), cancellationToken);

    public Task<AdminPage<AdminOffer>> OffersAsync(string? search, int page, int pageSize, CancellationToken cancellationToken) =>
        QueryAsync("""
            SELECT o.id, p.id, p.code, o.version, o.currency, o.amount_minor, o.status::text,
                   o.effective_from_utc, o.effective_until_utc, o.feature_codes, count(*) OVER() AS total
            FROM offers o JOIN products p ON p.id = o.product_id
            WHERE ($1 = '' OR p.code ILIKE '%' || $1 || '%' OR o.status::text ILIKE '%' || $1 || '%')
            ORDER BY o.effective_from_utc DESC, o.id DESC OFFSET $2 LIMIT $3;
            """, search, page, pageSize,
            r => new AdminOffer(r.GetGuid(0), r.GetGuid(1), r.GetString(2), r.GetInt32(3), r.GetString(4), r.GetInt64(5), r.GetString(6), r.GetFieldValue<DateTimeOffset>(7), r.IsDBNull(8) ? null : r.GetFieldValue<DateTimeOffset>(8), r.GetFieldValue<string[]>(9)), cancellationToken);

    public Task<AdminPage<AdminOrder>> OrdersAsync(string? search, int page, int pageSize, CancellationToken cancellationToken) =>
        QueryAsync("""
            SELECT id, stable_order_code, user_id, currency, amount_minor, status::text,
                   created_at_utc, updated_at_utc, count(*) OVER() AS total
            FROM orders
            WHERE ($1 = '' OR stable_order_code ILIKE '%' || $1 || '%' OR status::text ILIKE '%' || $1 || '%')
            ORDER BY created_at_utc DESC, id DESC OFFSET $2 LIMIT $3;
            """, search, page, pageSize,
            r => new AdminOrder(r.GetGuid(0), r.GetString(1), r.GetGuid(2), r.GetString(3), r.GetInt64(4), r.GetString(5), r.GetFieldValue<DateTimeOffset>(6), r.GetFieldValue<DateTimeOffset>(7)), cancellationToken);

    public Task<AdminPage<AdminSubscription>> SubscriptionsAsync(string? search, int page, int pageSize, CancellationToken cancellationToken) =>
        QueryAsync("""
            SELECT e.id, e.user_id, e.organization_id, p.code, e.source::text, e.status::text,
                   e.effective_from_utc, e.effective_until_utc, e.feature_codes, count(*) OVER() AS total
            FROM entitlement_grants e JOIN products p ON p.id = e.product_id
            WHERE ($1 = '' OR p.code ILIKE '%' || $1 || '%' OR e.status::text ILIKE '%' || $1 || '%'
                   OR e.source::text ILIKE '%' || $1 || '%')
            ORDER BY e.effective_until_utc DESC, e.id DESC OFFSET $2 LIMIT $3;
            """, search, page, pageSize,
            r => new AdminSubscription(r.GetGuid(0), r.GetGuid(1), r.IsDBNull(2) ? null : r.GetGuid(2), r.GetString(3), r.GetString(4), r.GetString(5), r.GetFieldValue<DateTimeOffset>(6), r.GetFieldValue<DateTimeOffset>(7), r.GetFieldValue<string[]>(8)), cancellationToken);

    public Task<AdminPage<AdminPayment>> PaymentsAsync(string? search, int page, int pageSize, CancellationToken cancellationToken) =>
        QueryAsync("""
            SELECT pe.id, pe.order_id, o.stable_order_code, pe.provider, pe.provider_event_id,
                   pe.event_type, pe.signature_valid, pe.amount_minor, pe.currency,
                   pe.received_at_utc, pe.processing_result, count(*) OVER() AS total
            FROM payment_events pe JOIN orders o ON o.id = pe.order_id
            WHERE ($1 = '' OR o.stable_order_code ILIKE '%' || $1 || '%' OR pe.provider_event_id ILIKE '%' || $1 || '%'
                   OR pe.event_type ILIKE '%' || $1 || '%')
            ORDER BY pe.received_at_utc DESC, pe.id DESC OFFSET $2 LIMIT $3;
            """, search, page, pageSize,
            r => new AdminPayment(r.GetGuid(0), r.GetGuid(1), r.GetString(2), r.GetString(3), r.GetString(4), r.GetString(5), r.GetBoolean(6), r.IsDBNull(7) ? null : r.GetInt64(7), r.IsDBNull(8) ? null : r.GetString(8), r.GetFieldValue<DateTimeOffset>(9), r.IsDBNull(10) ? null : r.GetString(10)), cancellationToken);

    public Task<AdminPage<AdminAudit>> AuditAsync(string? search, int page, int pageSize, CancellationToken cancellationToken) =>
        QueryAsync("""
            SELECT event_id, occurred_at_utc, source_application, external_actor_id,
                   target_type, target_id, action_code, result_code, correlation_id,
                   metadata::text, count(*) OVER() AS total
            FROM admin_integration_audit
            WHERE ($1 = '' OR external_actor_id ILIKE '%' || $1 || '%'
                   OR target_id ILIKE '%' || $1 || '%' OR action_code ILIKE '%' || $1 || '%'
                   OR result_code ILIKE '%' || $1 || '%')
            ORDER BY occurred_at_utc DESC, sequence_id DESC OFFSET $2 LIMIT $3;
            """, search, page, pageSize,
            r => new AdminAudit(r.GetGuid(0), r.GetFieldValue<DateTimeOffset>(1), r.GetString(2), r.GetString(3), r.GetString(4), r.GetString(5), r.GetString(6), r.GetString(7), r.GetGuid(8), r.GetString(9)), cancellationToken);

    public async Task<AdminMutation> ExtendEntitlementAsync(Guid userId, Guid productId, IReadOnlyList<string> featureCodes, TimeSpan duration, string reason, string clientId, string idempotencyKey, string actorId, Guid correlationId, CancellationToken cancellationToken)
    {
        if (userId == Guid.Empty || productId == Guid.Empty || featureCodes.Count is < 1 or > 100 || duration <= TimeSpan.Zero || duration > TimeSpan.FromDays(3660) || string.IsNullOrWhiteSpace(reason))
            throw new ArgumentException("ADMIN_ENTITLEMENT_REQUEST_INVALID");
        var features = featureCodes.Select(x => x.Trim()).Where(x => x.Length is > 0 and <= 128).Distinct(StringComparer.Ordinal).Order(StringComparer.Ordinal).ToArray();
        if (features.Length == 0) throw new ArgumentException("ADMIN_ENTITLEMENT_FEATURES_INVALID");
        var canonical = $"{userId:D}|{productId:D}|{string.Join(',', features)}|{duration.TotalSeconds:0}|{reason.Trim()}";
        var requestHash = System.Security.Cryptography.SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(canonical));
        await using var connection = await dataSource.OpenConnectionAsync(cancellationToken);
        // The transaction-scoped advisory lock serializes the idempotency key;
        // ReadCommitted avoids SSI aborts while retaining atomic target/audit/idempotency writes.
        await using var transaction = await connection.BeginTransactionAsync(System.Data.IsolationLevel.ReadCommitted, cancellationToken);
        await using (var lockCommand = new NpgsqlCommand("SELECT pg_advisory_xact_lock(hashtextextended($1, 0));", connection, transaction))
        {
            lockCommand.Parameters.AddWithValue($"{clientId}:{idempotencyKey}");
            await lockCommand.ExecuteNonQueryAsync(cancellationToken);
        }
        await using (var existing = new NpgsqlCommand("SELECT request_hash, response_body::text FROM admin_integration_idempotency WHERE client_id = $1 AND idempotency_key = $2 FOR UPDATE", connection, transaction))
        {
            existing.Parameters.AddWithValue(clientId); existing.Parameters.AddWithValue(idempotencyKey);
            await using var reader = await existing.ExecuteReaderAsync(cancellationToken);
            if (await reader.ReadAsync(cancellationToken))
            {
                var previousHash = (byte[])reader[0];
                var replayJson = reader.GetString(1);
                await reader.DisposeAsync();
                if (!previousHash.AsSpan().SequenceEqual(requestHash)) throw new InvalidOperationException("IDEMPOTENCY_KEY_CONFLICT");
                return System.Text.Json.JsonSerializer.Deserialize<AdminMutation>(replayJson)!;
            }
        }
        var now = DateTimeOffset.UtcNow;
        var result = new AdminMutation(Guid.NewGuid(), userId, productId, "ACTIVE", now, now.Add(duration), features);
        await using (var command = new NpgsqlCommand("""
            INSERT INTO entitlement_grants (id, user_id, product_id, source, status, effective_from_utc, effective_until_utc, feature_codes, source_reference, reason, created_at_utc)
            SELECT $1, u.id, p.id, 'ADMIN_EXTENSION', 'ACTIVE', $3, $4, $5, $6, $7, $3
            FROM users u CROSS JOIN products p WHERE u.id = $2 AND u.status = 'ACTIVE' AND p.id = $8 AND p.status = 'ACTIVE';
            """, connection, transaction))
        {
            command.Parameters.AddWithValue(result.GrantId); command.Parameters.AddWithValue(userId); command.Parameters.AddWithValue(now); command.Parameters.AddWithValue(result.EffectiveUntilUtc); command.Parameters.AddWithValue(features); command.Parameters.AddWithValue($"admin-integration:{idempotencyKey}"); command.Parameters.AddWithValue(reason.Trim()); command.Parameters.AddWithValue(productId);
            if (await command.ExecuteNonQueryAsync(cancellationToken) != 1) throw new InvalidOperationException("ADMIN_ENTITLEMENT_TARGET_INVALID");
        }
        var responseJson = System.Text.Json.JsonSerializer.Serialize(result);
        await using (var audit = new NpgsqlCommand("INSERT INTO admin_integration_audit (event_id, occurred_at_utc, source_application, external_actor_id, target_type, target_id, action_code, result_code, correlation_id, metadata) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10::jsonb)", connection, transaction))
        {
            audit.Parameters.AddWithValue(Guid.NewGuid()); audit.Parameters.AddWithValue(now); audit.Parameters.AddWithValue("biddingflow"); audit.Parameters.AddWithValue(actorId); audit.Parameters.AddWithValue("entitlement"); audit.Parameters.AddWithValue(result.GrantId.ToString("D")); audit.Parameters.AddWithValue("admin.entitlement.extend"); audit.Parameters.AddWithValue("SUCCEEDED"); audit.Parameters.AddWithValue(correlationId); audit.Parameters.AddWithValue($"{{\"reason\":{System.Text.Json.JsonSerializer.Serialize(reason.Trim())},\"userId\":\"{userId:D}\"}}");
            await audit.ExecuteNonQueryAsync(cancellationToken);
        }
        await using (var idem = new NpgsqlCommand("INSERT INTO admin_integration_idempotency (id, client_id, idempotency_key, request_hash, response_status, response_body, created_at_utc) VALUES ($1,$2,$3,$4,$5,$6::jsonb,$7)", connection, transaction))
        {
            idem.Parameters.AddWithValue(Guid.NewGuid()); idem.Parameters.AddWithValue(clientId); idem.Parameters.AddWithValue(idempotencyKey); idem.Parameters.AddWithValue(requestHash); idem.Parameters.AddWithValue(200); idem.Parameters.AddWithValue(responseJson); idem.Parameters.AddWithValue(now);
            await idem.ExecuteNonQueryAsync(cancellationToken);
        }
        await transaction.CommitAsync(cancellationToken);
        return result;
    }

    private async Task<AdminPage<T>> QueryAsync<T>(string sql, string? search, int page, int pageSize, Func<NpgsqlDataReader, T> map, CancellationToken cancellationToken)
    {
        page = Math.Clamp(page, 1, 100000);
        pageSize = Math.Clamp(pageSize, 1, 100);
        await using var command = dataSource.CreateCommand(sql);
        command.Parameters.AddWithValue(search?.Trim() ?? string.Empty);
        command.Parameters.AddWithValue((page - 1) * pageSize);
        command.Parameters.AddWithValue(pageSize);
        var items = new List<T>(pageSize);
        long total = 0;
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            items.Add(map(reader));
            total = reader.GetInt64(reader.FieldCount - 1);
        }
        return new AdminPage<T>(items, page, pageSize, total);
    }
}
