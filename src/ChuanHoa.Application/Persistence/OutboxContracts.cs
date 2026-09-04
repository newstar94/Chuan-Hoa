using System.Text.Json;

namespace ChuanHoa.Application.Persistence;

public sealed record OutboxMessage(
    Guid MessageId,
    string AggregateType,
    string AggregateId,
    string EventType,
    JsonDocument Payload,
    JsonDocument Headers,
    DateTimeOffset OccurredAtUtc,
    DateTimeOffset AvailableAtUtc,
    string IdempotencyKey);
