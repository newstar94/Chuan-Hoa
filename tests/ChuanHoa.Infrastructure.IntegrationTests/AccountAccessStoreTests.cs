using System.Security.Cryptography;
using System.Text;
using ChuanHoa.Contracts;
using ChuanHoa.Infrastructure.Admin;
using Npgsql;

namespace ChuanHoa.Infrastructure.IntegrationTests;

public sealed class AccountAccessStoreTests : IAsyncLifetime
{
    private NpgsqlDataSource _dataSource = null!;
    private PostgresAccountAccessStore _store = null!;
    private readonly FixedTimeProvider _clock = new(DateTimeOffset.Parse("2026-09-13T10:00:00Z"));

    public async Task InitializeAsync()
    {
        var connectionString = Environment.GetEnvironmentVariable("CHUANHOA_TEST_CONNECTION_STRING");
        if (string.IsNullOrWhiteSpace(connectionString)) throw new InvalidOperationException("CHUANHOA_TEST_CONNECTION_STRING is required.");
        _dataSource = NpgsqlDataSource.Create(connectionString);
        _store = new PostgresAccountAccessStore(_dataSource, _clock);
    }

    public async Task DisposeAsync() => await _dataSource.DisposeAsync();

    [Fact]
    public async Task Registration_issues_absolute_72_hour_session_without_plaintext_password()
    {
        var email = $"user-{Guid.NewGuid():N}@example.invalid";
        var session = await _store.RegisterAsync(new RegisterAccountRequest(email, "A-strong-password-123", "Test User", "device-a"), CancellationToken.None);

        Assert.Equal(_clock.GetUtcNow().AddHours(72), session.ExpiresAtUtc);
        Assert.NotEqual("A-strong-password-123", await ScalarAsync("SELECT password_hash FROM users WHERE id = $1", session.UserId));
        Assert.Equal(session, await _store.GetSessionAsync(session.SessionToken, CancellationToken.None));
    }

    [Fact]
    public async Task Second_device_cannot_take_an_active_account_binding()
    {
        var email = $"user-{Guid.NewGuid():N}@example.invalid";
        await _store.RegisterAsync(new RegisterAccountRequest(email, "A-strong-password-123", "Test User", "device-a"), CancellationToken.None);
        var error = await Assert.ThrowsAsync<InvalidOperationException>(() => _store.LoginAsync(new LoginAccountRequest(email, "A-strong-password-123", "device-b"), CancellationToken.None));
        Assert.Equal("ACCOUNT_DEVICE_IN_USE", error.Message);
    }

    [Fact]
    public async Task Logout_releases_binding_so_another_device_can_login()
    {
        var email = $"user-{Guid.NewGuid():N}@example.invalid";
        var first = await _store.RegisterAsync(new RegisterAccountRequest(email, "A-strong-password-123", "Test User", "device-a"), CancellationToken.None);
        await _store.LogoutAsync(first.SessionToken, CancellationToken.None);

        var second = await _store.LoginAsync(new LoginAccountRequest(email, "A-strong-password-123", "device-b"), CancellationToken.None);

        Assert.Equal("device-b", second.DeviceThumbprint);
        Assert.Null(await _store.GetSessionAsync(first.SessionToken, CancellationToken.None));
    }

    [Fact]
    public async Task Expired_session_does_not_lock_account_to_old_device_forever()
    {
        var email = $"user-{Guid.NewGuid():N}@example.invalid";
        var first = await _store.RegisterAsync(new RegisterAccountRequest(email, "A-strong-password-123", "Test User", "device-a"), CancellationToken.None);
        await using (var expire = _dataSource.CreateCommand("UPDATE account_sessions SET expires_at_utc = now() - interval '1 second', issued_at_utc = now() - interval '72 hours 1 second' WHERE user_id = $1"))
        {
            expire.Parameters.AddWithValue(first.UserId); await expire.ExecuteNonQueryAsync();
        }

        var second = await _store.LoginAsync(new LoginAccountRequest(email, "A-strong-password-123", "device-b"), CancellationToken.None);

        Assert.Equal("device-b", second.DeviceThumbprint);
    }

    [Fact]
    public async Task Same_device_activation_is_idempotent_and_does_not_consume_a_second_slot()
    {
        var productId = Guid.NewGuid();
        var keyId = Guid.NewGuid();
        var rawKey = "CHV-AAAA-BBBB-CCCC-DDDD";
        var hash = SHA256.HashData(Encoding.UTF8.GetBytes(rawKey.Replace("-", string.Empty, StringComparison.Ordinal).ToUpperInvariant()));
        var now = _clock.GetUtcNow();
        await using (var product = _dataSource.CreateCommand("INSERT INTO products (id, code, display_name, status, created_at_utc) VALUES ($1, $2, $3, 'ACTIVE', $4)"))
        {
            product.Parameters.AddWithValue(productId); product.Parameters.AddWithValue($"product-{productId:N}"); product.Parameters.AddWithValue("Product"); product.Parameters.AddWithValue(now); await product.ExecuteNonQueryAsync();
        }
        await using (var key = _dataSource.CreateCommand("INSERT INTO activation_keys (id, key_prefix, key_hash, status, product_id, feature_codes, max_devices, created_at_utc) VALUES ($1, 'CHV-AAAA', $2, 'ACTIVE', $3, $4, 2, $5)"))
        {
            key.Parameters.AddWithValue(keyId); key.Parameters.AddWithValue(hash); key.Parameters.AddWithValue(productId); key.Parameters.AddWithValue(new[] { "FORMAT_SCAN" }); key.Parameters.AddWithValue(now); await key.ExecuteNonQueryAsync();
        }
        var first = await _store.ActivateVipKeyAsync(new ActivateVipKeyRequest(rawKey, "device-a"), CancellationToken.None);
        var second = await _store.ActivateVipKeyAsync(new ActivateVipKeyRequest(rawKey, "device-a"), CancellationToken.None);
        Assert.Equal(1, first.UsedDevices); Assert.Equal(1, second.UsedDevices); Assert.Equal(1, second.RemainingDevices);
    }

    [Fact]
    public async Task Concurrent_final_slot_activation_allows_only_one_device()
    {
        var productId = Guid.NewGuid();
        var keyId = Guid.NewGuid();
        var rawKey = "CHV-FFFF-GGGG-HHHH-JJJJ";
        var hash = SHA256.HashData(Encoding.UTF8.GetBytes(rawKey.Replace("-", string.Empty, StringComparison.Ordinal).ToUpperInvariant()));
        var now = _clock.GetUtcNow();
        await using (var product = _dataSource.CreateCommand("INSERT INTO products (id, code, display_name, status, created_at_utc) VALUES ($1, $2, $3, 'ACTIVE', $4)"))
        {
            product.Parameters.AddWithValue(productId); product.Parameters.AddWithValue($"product-{productId:N}"); product.Parameters.AddWithValue("Product"); product.Parameters.AddWithValue(now); await product.ExecuteNonQueryAsync();
        }
        await using (var key = _dataSource.CreateCommand("INSERT INTO activation_keys (id, key_prefix, key_hash, status, product_id, feature_codes, max_devices, created_at_utc) VALUES ($1, 'CHV-FFFF', $2, 'ACTIVE', $3, $4, 1, $5)"))
        {
            key.Parameters.AddWithValue(keyId); key.Parameters.AddWithValue(hash); key.Parameters.AddWithValue(productId); key.Parameters.AddWithValue(new[] { "FORMAT_SCAN" }); key.Parameters.AddWithValue(now); await key.ExecuteNonQueryAsync();
        }
        var tasks = new[] { "device-race-a", "device-race-b" }.Select(device => Task.Run(() => _store.ActivateVipKeyAsync(new ActivateVipKeyRequest(rawKey, device), CancellationToken.None))).ToArray();
        var results = await Task.WhenAll(tasks.Select(async task => { try { await task; return (success: true, error: (string?)null); } catch (InvalidOperationException error) { return (success: false, error: error.Message); } }));
        Assert.Equal(1, results.Count(result => result.success));
        Assert.Equal(1, results.Count(result => result.error == "ACTIVATION_KEY_DEVICE_LIMIT"));
        Assert.Equal(1L, Convert.ToInt64(await ScalarAsync("SELECT count(*) FROM activation_key_devices WHERE activation_key_id = $1 AND released_at_utc IS NULL", keyId)));
    }

    private async Task<object?> ScalarAsync(string sql, Guid id)
    {
        await using var command = _dataSource.CreateCommand(sql); command.Parameters.AddWithValue(id); return await command.ExecuteScalarAsync();
    }

    private sealed class FixedTimeProvider(DateTimeOffset now) : TimeProvider
    {
        public override DateTimeOffset GetUtcNow() => now;
    }
}
