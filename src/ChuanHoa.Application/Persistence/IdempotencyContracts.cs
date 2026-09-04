using ChuanHoa.Contracts;

namespace ChuanHoa.Application.Persistence;

public sealed record IdempotencyRequest(
    string ScopeKey,
    byte[] IdempotencyKeyHash,
    byte[] RequestHash,
    string Method,
    string Path,
    DateTimeOffset NowUtc,
    DateTimeOffset ExpiresAtUtc);

public sealed record StoredHttpResponse(
    int StatusCode,
    string ContentType,
    IReadOnlyDictionary<string, string> Headers,
    byte[] Body);

public sealed record IdempotencyBeginResult(
    IdempotencyDisposition Disposition,
    Guid RecordId,
    Guid? OwnerToken,
    StoredHttpResponse? ReplayResponse);

public interface IIdempotencyStore
{
    Task<IdempotencyBeginResult> BeginAsync(
        IdempotencyRequest request,
        CancellationToken cancellationToken);

    Task CompleteAsync(
        Guid recordId,
        Guid ownerToken,
        StoredHttpResponse response,
        DateTimeOffset completedAtUtc,
        CancellationToken cancellationToken);

    Task MarkRetryableAsync(
        Guid recordId,
        Guid ownerToken,
        string errorCode,
        DateTimeOffset failedAtUtc,
        CancellationToken cancellationToken);
}

public sealed class IdempotencyOwnershipException : Exception
{
    public IdempotencyOwnershipException(Guid recordId)
        : base($"Idempotency record {recordId} is no longer owned by this request.")
    {
        RecordId = recordId;
    }

    public Guid RecordId { get; }
}
