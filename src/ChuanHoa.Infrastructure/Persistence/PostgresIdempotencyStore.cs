using System.Data;
using System.Text.Json;
using ChuanHoa.Application.Persistence;
using ChuanHoa.Contracts;
using Npgsql;
using NpgsqlTypes;

namespace ChuanHoa.Infrastructure.Persistence;

public sealed class PostgresIdempotencyStore(
    NpgsqlDataSource dataSource,
    PostgresTransactionRunner transactionRunner) : IIdempotencyStore
{
    public Task<IdempotencyBeginResult> BeginAsync(
        IdempotencyRequest request,
        CancellationToken cancellationToken)
    {
        ValidateRequest(request);
        return transactionRunner.ExecuteAsync(
            (connection, transaction, token) => BeginCoreAsync(connection, transaction, request, token),
            IsolationLevel.ReadCommitted,
            cancellationToken);
    }

    public async Task CompleteAsync(
        Guid recordId,
        Guid ownerToken,
        StoredHttpResponse response,
        DateTimeOffset completedAtUtc,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(response);
        if (response.StatusCode is < 100 or > 599)
        {
            throw new ArgumentOutOfRangeException(nameof(response), "Stored response status code must be a valid HTTP status code.");
        }

        const string sql = """
            UPDATE idempotency_records
               SET status = 'COMPLETED',
                   response_status = @response_status,
                   response_content_type = @response_content_type,
                   response_headers = @response_headers,
                   response_body = @response_body,
                   completed_at_utc = @completed_at_utc,
                   last_error_code = NULL
             WHERE id = @id
               AND owner_token = @owner_token
               AND status = 'IN_PROGRESS';
            """;

        await using var connection = await dataSource.OpenConnectionAsync(cancellationToken);
        await using var command = new NpgsqlCommand(sql, connection);
        command.Parameters.AddWithValue("response_status", response.StatusCode);
        command.Parameters.AddWithValue("response_content_type", response.ContentType);
        command.Parameters.AddWithValue("response_headers", NpgsqlDbType.Jsonb, JsonSerializer.Serialize(response.Headers));
        command.Parameters.AddWithValue("response_body", response.Body);
        command.Parameters.AddWithValue("completed_at_utc", completedAtUtc);
        command.Parameters.AddWithValue("id", recordId);
        command.Parameters.AddWithValue("owner_token", ownerToken);

        if (await command.ExecuteNonQueryAsync(cancellationToken) != 1)
        {
            throw new IdempotencyOwnershipException(recordId);
        }
    }

    public async Task MarkRetryableAsync(
        Guid recordId,
        Guid ownerToken,
        string errorCode,
        DateTimeOffset failedAtUtc,
        CancellationToken cancellationToken)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(errorCode);

        const string sql = """
            UPDATE idempotency_records
               SET status = 'FAILED_RETRYABLE',
                   last_error_code = @error_code,
                   completed_at_utc = @failed_at_utc
             WHERE id = @id
               AND owner_token = @owner_token
               AND status = 'IN_PROGRESS';
            """;

        await using var connection = await dataSource.OpenConnectionAsync(cancellationToken);
        await using var command = new NpgsqlCommand(sql, connection);
        command.Parameters.AddWithValue("error_code", errorCode);
        command.Parameters.AddWithValue("failed_at_utc", failedAtUtc);
        command.Parameters.AddWithValue("id", recordId);
        command.Parameters.AddWithValue("owner_token", ownerToken);

        if (await command.ExecuteNonQueryAsync(cancellationToken) != 1)
        {
            throw new IdempotencyOwnershipException(recordId);
        }
    }

    private static async Task<IdempotencyBeginResult> BeginCoreAsync(
        NpgsqlConnection connection,
        NpgsqlTransaction transaction,
        IdempotencyRequest request,
        CancellationToken cancellationToken)
    {
        var proposedRecordId = Guid.NewGuid();
        var proposedOwnerToken = Guid.NewGuid();

        const string insertSql = """
            INSERT INTO idempotency_records (
                id,
                scope_key,
                idempotency_key_hash,
                request_hash,
                method,
                path,
                status,
                owner_token,
                expires_at_utc,
                created_at_utc
            ) VALUES (
                @id,
                @scope_key,
                @idempotency_key_hash,
                @request_hash,
                @method,
                @path,
                'IN_PROGRESS',
                @owner_token,
                @expires_at_utc,
                @created_at_utc
            )
            ON CONFLICT (scope_key, idempotency_key_hash) DO NOTHING;
            """;

        await using (var insert = new NpgsqlCommand(insertSql, connection, transaction))
        {
            insert.Parameters.AddWithValue("id", proposedRecordId);
            insert.Parameters.AddWithValue("scope_key", request.ScopeKey);
            insert.Parameters.AddWithValue("idempotency_key_hash", request.IdempotencyKeyHash);
            insert.Parameters.AddWithValue("request_hash", request.RequestHash);
            insert.Parameters.AddWithValue("method", request.Method);
            insert.Parameters.AddWithValue("path", request.Path);
            insert.Parameters.AddWithValue("owner_token", proposedOwnerToken);
            insert.Parameters.AddWithValue("expires_at_utc", request.ExpiresAtUtc);
            insert.Parameters.AddWithValue("created_at_utc", request.NowUtc);

            if (await insert.ExecuteNonQueryAsync(cancellationToken) == 1)
            {
                return new IdempotencyBeginResult(
                    IdempotencyDisposition.Acquired,
                    proposedRecordId,
                    proposedOwnerToken,
                    null);
            }
        }

        const string selectSql = """
            SELECT id,
                   request_hash,
                   status,
                   response_status,
                   response_content_type,
                   response_headers,
                   response_body,
                   expires_at_utc
              FROM idempotency_records
             WHERE scope_key = @scope_key
               AND idempotency_key_hash = @idempotency_key_hash
             FOR UPDATE;
            """;

        await using var select = new NpgsqlCommand(selectSql, connection, transaction);
        select.Parameters.AddWithValue("scope_key", request.ScopeKey);
        select.Parameters.AddWithValue("idempotency_key_hash", request.IdempotencyKeyHash);
        await using var reader = await select.ExecuteReaderAsync(cancellationToken);
        if (!await reader.ReadAsync(cancellationToken))
        {
            throw new InvalidOperationException("The idempotency record disappeared during conflict resolution.");
        }

        var recordId = reader.GetGuid(0);
        var storedRequestHash = reader.GetFieldValue<byte[]>(1);
        var status = reader.GetString(2);
        int? responseStatus = reader.IsDBNull(3) ? null : reader.GetInt32(3);
        var responseContentType = reader.IsDBNull(4) ? null : reader.GetString(4);
        var responseHeadersJson = reader.IsDBNull(5) ? null : reader.GetString(5);
        var responseBody = reader.IsDBNull(6) ? null : reader.GetFieldValue<byte[]>(6);
        var expiresAtUtc = reader.GetFieldValue<DateTimeOffset>(7);
        await reader.DisposeAsync();

        if (!storedRequestHash.AsSpan().SequenceEqual(request.RequestHash))
        {
            return new IdempotencyBeginResult(IdempotencyDisposition.Conflict, recordId, null, null);
        }

        if (status == "COMPLETED" && expiresAtUtc > request.NowUtc)
        {
            if (responseStatus is null || responseContentType is null || responseHeadersJson is null || responseBody is null)
            {
                throw new InvalidOperationException("A completed idempotency record is missing its stored response.");
            }

            var headers = JsonSerializer.Deserialize<Dictionary<string, string>>(responseHeadersJson)
                ?? throw new InvalidOperationException("Stored idempotency response headers are invalid.");
            return new IdempotencyBeginResult(
                IdempotencyDisposition.Replay,
                recordId,
                null,
                new StoredHttpResponse(responseStatus.Value, responseContentType, headers, responseBody));
        }

        if (status == "IN_PROGRESS" && expiresAtUtc > request.NowUtc)
        {
            return new IdempotencyBeginResult(IdempotencyDisposition.InProgress, recordId, null, null);
        }

        const string reacquireSql = """
            UPDATE idempotency_records
               SET request_hash = @request_hash,
                   method = @method,
                   path = @path,
                   status = 'IN_PROGRESS',
                   owner_token = @owner_token,
                   response_status = NULL,
                   response_content_type = NULL,
                   response_headers = NULL,
                   response_body = NULL,
                   last_error_code = NULL,
                   expires_at_utc = @expires_at_utc,
                   completed_at_utc = NULL,
                   created_at_utc = @created_at_utc
             WHERE id = @id;
            """;

        await using var reacquire = new NpgsqlCommand(reacquireSql, connection, transaction);
        reacquire.Parameters.AddWithValue("request_hash", request.RequestHash);
        reacquire.Parameters.AddWithValue("method", request.Method);
        reacquire.Parameters.AddWithValue("path", request.Path);
        reacquire.Parameters.AddWithValue("owner_token", proposedOwnerToken);
        reacquire.Parameters.AddWithValue("expires_at_utc", request.ExpiresAtUtc);
        reacquire.Parameters.AddWithValue("created_at_utc", request.NowUtc);
        reacquire.Parameters.AddWithValue("id", recordId);
        await reacquire.ExecuteNonQueryAsync(cancellationToken);

        return new IdempotencyBeginResult(
            IdempotencyDisposition.Acquired,
            recordId,
            proposedOwnerToken,
            null);
    }

    private static void ValidateRequest(IdempotencyRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        ArgumentException.ThrowIfNullOrWhiteSpace(request.ScopeKey);
        ArgumentException.ThrowIfNullOrWhiteSpace(request.Method);
        ArgumentException.ThrowIfNullOrWhiteSpace(request.Path);

        if (request.IdempotencyKeyHash.Length != 32)
        {
            throw new ArgumentException("Idempotency key hash must be SHA-256 (32 bytes).", nameof(request));
        }

        if (request.RequestHash.Length != 32)
        {
            throw new ArgumentException("Request hash must be SHA-256 (32 bytes).", nameof(request));
        }

        if (request.ExpiresAtUtc <= request.NowUtc)
        {
            throw new ArgumentException("Idempotency record expiry must be after the current server time.", nameof(request));
        }
    }
}
