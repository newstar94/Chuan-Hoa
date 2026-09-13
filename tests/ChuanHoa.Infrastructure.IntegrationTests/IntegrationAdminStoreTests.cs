using ChuanHoa.Infrastructure.Admin;
using Npgsql;

namespace ChuanHoa.Infrastructure.IntegrationTests;

public sealed class IntegrationAdminStoreTests : IAsyncLifetime
{
    private NpgsqlDataSource _dataSource = null!;
    private PostgresIntegrationAdminStore _store = null!;
    private readonly Guid _userId = Guid.NewGuid();
    private readonly Guid _productId = Guid.NewGuid();

    public async Task InitializeAsync()
    {
        var connectionString = Environment.GetEnvironmentVariable("CHUANHOA_TEST_CONNECTION_STRING");
        if (string.IsNullOrWhiteSpace(connectionString)) throw new InvalidOperationException("CHUANHOA_TEST_CONNECTION_STRING is required.");
        _dataSource = NpgsqlDataSource.Create(connectionString);
        _store = new PostgresIntegrationAdminStore(_dataSource);
        var now = DateTimeOffset.UtcNow;
        await using var userCommand = _dataSource.CreateCommand("INSERT INTO users (id, normalized_email, display_name, status, created_at_utc, updated_at_utc) VALUES ($1, $2, $3, 'ACTIVE', $4, $4);");
        userCommand.Parameters.AddWithValue(_userId); userCommand.Parameters.AddWithValue($"admin-store-{_userId:N}@example.invalid"); userCommand.Parameters.AddWithValue("Integration Admin User"); userCommand.Parameters.AddWithValue(now);
        await userCommand.ExecuteNonQueryAsync();
        await using var productCommand = _dataSource.CreateCommand("INSERT INTO products (id, code, display_name, status, created_at_utc) VALUES ($1, $2, $3, 'ACTIVE', $4);");
        productCommand.Parameters.AddWithValue(_productId); productCommand.Parameters.AddWithValue($"product-{_productId:N}"); productCommand.Parameters.AddWithValue("Integration Product"); productCommand.Parameters.AddWithValue(now);
        await productCommand.ExecuteNonQueryAsync();
    }

    public async Task DisposeAsync() => await _dataSource.DisposeAsync();

    [Fact]
    public async Task Same_key_concurrent_requests_create_one_grant_and_replay_same_result()
    {
        var key = $"admin-extend-{Guid.NewGuid():N}";
        var tasks = Enumerable.Range(0, 2).Select(_ => _store.ExtendEntitlementAsync(
            _userId, _productId, ["feature.a", "feature.b"], TimeSpan.FromDays(30), "support case",
            "bidding-admin", key, "actor-1", Guid.NewGuid(), CancellationToken.None));
        var results = await Task.WhenAll(tasks);

        Assert.Equal(results[0].GrantId, results[1].GrantId);
        await using var count = _dataSource.CreateCommand("SELECT count(*) FROM entitlement_grants WHERE id = $1;");
        count.Parameters.AddWithValue(results[0].GrantId);
        Assert.Equal(1L, Convert.ToInt64(await count.ExecuteScalarAsync()));
    }

    [Fact]
    public async Task Same_key_with_different_payload_is_rejected()
    {
        var key = $"admin-conflict-{Guid.NewGuid():N}";
        await _store.ExtendEntitlementAsync(_userId, _productId, ["feature.a"], TimeSpan.FromDays(30), "first", "bidding-admin", key, "actor-1", Guid.NewGuid(), CancellationToken.None);
        var error = await Assert.ThrowsAsync<InvalidOperationException>(() => _store.ExtendEntitlementAsync(_userId, _productId, ["feature.b"], TimeSpan.FromDays(30), "second", "bidding-admin", key, "actor-1", Guid.NewGuid(), CancellationToken.None));
        Assert.Equal("IDEMPOTENCY_KEY_CONFLICT", error.Message);
    }

    [Fact]
    public async Task Failed_command_attempt_is_recorded_in_target_audit()
    {
        var correlation = Guid.NewGuid();
        await _store.RecordFailedEntitlementAsync(_userId, _productId, "bidding-admin", "actor-1", correlation, "ADMIN_ENTITLEMENT_TARGET_INVALID", CancellationToken.None);
        await using var command = _dataSource.CreateCommand("SELECT source_application, external_actor_id, target_type, target_id, action_code, result_code, correlation_id FROM admin_integration_audit WHERE correlation_id = $1;");
        command.Parameters.AddWithValue(correlation);
        await using var reader = await command.ExecuteReaderAsync();
        Assert.True(await reader.ReadAsync());
        Assert.Equal("biddingflow", reader.GetString(0));
        Assert.Equal("actor-1", reader.GetString(1));
        Assert.Equal("entitlement", reader.GetString(2));
        Assert.Equal($"{_userId:D}:{_productId:D}", reader.GetString(3));
        Assert.Equal("admin.entitlement.extend", reader.GetString(4));
        Assert.Equal("ADMIN_ENTITLEMENT_TARGET_INVALID", reader.GetString(5));
        Assert.Equal(correlation, reader.GetGuid(6));
        await reader.DisposeAsync();

        var page = await _store.AuditAsync("actor-1", 1, 25, CancellationToken.None);
        Assert.Contains(page.Items, item => item.CorrelationId == correlation
            && item.SourceApplication == "biddingflow"
            && item.ResultCode == "ADMIN_ENTITLEMENT_TARGET_INVALID");
    }
}
