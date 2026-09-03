using System.Data;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using ChuanHoa.Application.Persistence;
using ChuanHoa.Contracts;
using ChuanHoa.Infrastructure.Persistence;
using Npgsql;

namespace ChuanHoa.Infrastructure.IntegrationTests;

public sealed class PostgresPersistenceTests : IAsyncLifetime
{
    private NpgsqlDataSource _dataSource = null!;
    private PostgresTransactionRunner _transactionRunner = null!;
    private PostgresIdempotencyStore _idempotencyStore = null!;
    private PostgresOutboxWriter _outboxWriter = null!;

    public Task InitializeAsync()
    {
        var connectionString = Environment.GetEnvironmentVariable("CHUANHOA_TEST_CONNECTION_STRING");
        if (string.IsNullOrWhiteSpace(connectionString))
        {
            throw new InvalidOperationException(
                "CHUANHOA_TEST_CONNECTION_STRING is required. Run this project through tools/database/verify_persistence.ps1.");
        }

        _dataSource = NpgsqlDataSource.Create(connectionString);
        _transactionRunner = new PostgresTransactionRunner(_dataSource);
        _idempotencyStore = new PostgresIdempotencyStore(_dataSource, _transactionRunner);
        _outboxWriter = new PostgresOutboxWriter();
        return Task.CompletedTask;
    }

    public async Task DisposeAsync()
    {
        await _dataSource.DisposeAsync();
    }

    [Fact]
    public async Task Begin_distinguishes_acquired_in_progress_and_conflicting_payload()
    {
        var now = DateTimeOffset.UtcNow;
        var scope = $"subject:{Guid.NewGuid():D}";
        var keyHash = Sha256($"key:{Guid.NewGuid():D}");
        var request = CreateRequest(scope, keyHash, Sha256("request-a"), now);

        var acquired = await _idempotencyStore.BeginAsync(request, CancellationToken.None);
        var inProgress = await _idempotencyStore.BeginAsync(request, CancellationToken.None);
        var conflict = await _idempotencyStore.BeginAsync(
            request with { RequestHash = Sha256("request-b") },
            CancellationToken.None);

        Assert.Equal(IdempotencyDisposition.Acquired, acquired.Disposition);
        Assert.NotNull(acquired.OwnerToken);
        Assert.Equal(IdempotencyDisposition.InProgress, inProgress.Disposition);
        Assert.Null(inProgress.OwnerToken);
        Assert.Equal(IdempotencyDisposition.Conflict, conflict.Disposition);
        Assert.Null(conflict.OwnerToken);
        Assert.Equal(acquired.RecordId, inProgress.RecordId);
        Assert.Equal(acquired.RecordId, conflict.RecordId);
    }

    [Fact]
    public async Task Completed_response_is_replayed_byte_for_byte()
    {
        var now = DateTimeOffset.UtcNow;
        var request = CreateRequest(
            $"subject:{Guid.NewGuid():D}",
            Sha256($"key:{Guid.NewGuid():D}"),
            Sha256("request-complete"),
            now);
        var acquired = await _idempotencyStore.BeginAsync(request, CancellationToken.None);
        var storedResponse = new StoredHttpResponse(
            201,
            "application/json",
            new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["ETag"] = "\"version-7\"",
                ["Location"] = "/v1/orders/00000000-0000-0000-0000-000000000001"
            },
            Encoding.UTF8.GetBytes("{\"status\":\"created\"}"));

        await _idempotencyStore.CompleteAsync(
            acquired.RecordId,
            acquired.OwnerToken!.Value,
            storedResponse,
            now.AddSeconds(1),
            CancellationToken.None);
        var replay = await _idempotencyStore.BeginAsync(request, CancellationToken.None);

        Assert.Equal(IdempotencyDisposition.Replay, replay.Disposition);
        Assert.NotNull(replay.ReplayResponse);
        Assert.Equal(storedResponse.StatusCode, replay.ReplayResponse.StatusCode);
        Assert.Equal(storedResponse.ContentType, replay.ReplayResponse.ContentType);
        Assert.Equal(storedResponse.Headers, replay.ReplayResponse.Headers);
        Assert.Equal(storedResponse.Body, replay.ReplayResponse.Body);
    }

    [Fact]
    public async Task Retryable_failure_can_be_reacquired_with_a_new_owner()
    {
        var now = DateTimeOffset.UtcNow;
        var request = CreateRequest(
            $"subject:{Guid.NewGuid():D}",
            Sha256($"key:{Guid.NewGuid():D}"),
            Sha256("request-retryable"),
            now);
        var first = await _idempotencyStore.BeginAsync(request, CancellationToken.None);

        await _idempotencyStore.MarkRetryableAsync(
            first.RecordId,
            first.OwnerToken!.Value,
            "UPSTREAM_TIMEOUT",
            now.AddSeconds(1),
            CancellationToken.None);
        var reacquired = await _idempotencyStore.BeginAsync(request, CancellationToken.None);

        Assert.Equal(IdempotencyDisposition.Acquired, reacquired.Disposition);
        Assert.Equal(first.RecordId, reacquired.RecordId);
        Assert.NotNull(reacquired.OwnerToken);
        Assert.NotEqual(first.OwnerToken, reacquired.OwnerToken);
    }

    [Fact]
    public async Task Outbox_write_commits_and_rolls_back_with_the_enclosing_transaction()
    {
        var rollbackKey = $"rollback:{Guid.NewGuid():D}";
        var rollbackMessage = CreateOutboxMessage(rollbackKey);

        await Assert.ThrowsAsync<ExpectedRollbackException>(async () =>
            await _transactionRunner.ExecuteAsync<int>(
                async (connection, transaction, cancellationToken) =>
                {
                    await _outboxWriter.EnqueueAsync(
                        connection,
                        transaction,
                        rollbackMessage,
                        cancellationToken);
                    throw new ExpectedRollbackException();
                },
                IsolationLevel.ReadCommitted,
                CancellationToken.None));
        Assert.Equal(0, await CountOutboxAsync(rollbackKey));

        var commitKey = $"commit:{Guid.NewGuid():D}";
        var commitMessage = CreateOutboxMessage(commitKey);
        await _transactionRunner.ExecuteAsync(
            async (connection, transaction, cancellationToken) =>
            {
                await _outboxWriter.EnqueueAsync(
                    connection,
                    transaction,
                    commitMessage,
                    cancellationToken);
                return 1;
            },
            IsolationLevel.ReadCommitted,
            CancellationToken.None);

        Assert.Equal(1, await CountOutboxAsync(commitKey));
    }

    private static IdempotencyRequest CreateRequest(
        string scope,
        byte[] keyHash,
        byte[] requestHash,
        DateTimeOffset now)
    {
        return new IdempotencyRequest(
            scope,
            keyHash,
            requestHash,
            "POST",
            "/v1/test-mutation",
            now,
            now.AddHours(24));
    }

    private static OutboxMessage CreateOutboxMessage(string idempotencyKey)
    {
        return new OutboxMessage(
            Guid.NewGuid(),
            "Order",
            Guid.NewGuid().ToString("D"),
            "OrderCreated",
            JsonDocument.Parse("{\"status\":\"CREATED\"}"),
            JsonDocument.Parse("{\"schema\":\"chuanhoa.outbox.headers.v1\"}"),
            DateTimeOffset.UtcNow,
            DateTimeOffset.UtcNow,
            idempotencyKey);
    }

    private async Task<long> CountOutboxAsync(string idempotencyKey)
    {
        await using var command = _dataSource.CreateCommand(
            "SELECT count(*) FROM outbox_messages WHERE idempotency_key = $1;");
        command.Parameters.AddWithValue(idempotencyKey);
        var result = await command.ExecuteScalarAsync(CancellationToken.None);
        return Convert.ToInt64(result);
    }

    private static byte[] Sha256(string value)
    {
        return SHA256.HashData(Encoding.UTF8.GetBytes(value));
    }

    private sealed class ExpectedRollbackException : Exception;
}
