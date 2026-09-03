using System.Text.Json.Serialization;

namespace ChuanHoa.Contracts;

public static class ApiContractVersions
{
    public const string ErrorV1 = "chuanhoa.error.v1";
    public const string HealthV1 = "chuanhoa.health.v1";
}

public sealed record ApiErrorResponse(
    string Schema,
    string Code,
    string Title,
    string Detail,
    string Recovery,
    int Status,
    string CorrelationId,
    string Instance,
    IReadOnlyDictionary<string, IReadOnlyList<string>>? ValidationErrors = null);

public sealed record HealthResponse(
    string Schema,
    string Status,
    string Service,
    DateTimeOffset ServerTimeUtc);

public static class ApiHeaders
{
    public const string CorrelationId = "X-Correlation-ID";
    public const string IdempotencyKey = "Idempotency-Key";
    public const string EntityTag = "ETag";
    public const string IfMatch = "If-Match";
}

[JsonConverter(typeof(JsonStringEnumConverter<IdempotencyDisposition>))]
public enum IdempotencyDisposition
{
    Acquired,
    Replay,
    InProgress,
    Conflict
}
