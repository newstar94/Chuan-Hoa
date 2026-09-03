using ChuanHoa.Application.Persistence;
using Npgsql;
using NpgsqlTypes;

namespace ChuanHoa.Infrastructure.Persistence;

public sealed class PostgresOutboxWriter
{
    public async Task EnqueueAsync(
        NpgsqlConnection connection,
        NpgsqlTransaction transaction,
        OutboxMessage message,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(connection);
        ArgumentNullException.ThrowIfNull(transaction);
        ArgumentNullException.ThrowIfNull(message);

        const string sql = """
            INSERT INTO outbox_messages (
                message_id,
                aggregate_type,
                aggregate_id,
                event_type,
                payload,
                headers,
                occurred_at_utc,
                available_at_utc,
                idempotency_key
            ) VALUES (
                @message_id,
                @aggregate_type,
                @aggregate_id,
                @event_type,
                @payload,
                @headers,
                @occurred_at_utc,
                @available_at_utc,
                @idempotency_key
            );
            """;

        await using var command = new NpgsqlCommand(sql, connection, transaction);
        command.Parameters.AddWithValue("message_id", message.MessageId);
        command.Parameters.AddWithValue("aggregate_type", message.AggregateType);
        command.Parameters.AddWithValue("aggregate_id", message.AggregateId);
        command.Parameters.AddWithValue("event_type", message.EventType);
        command.Parameters.AddWithValue("payload", NpgsqlDbType.Jsonb, message.Payload.RootElement.GetRawText());
        command.Parameters.AddWithValue("headers", NpgsqlDbType.Jsonb, message.Headers.RootElement.GetRawText());
        command.Parameters.AddWithValue("occurred_at_utc", message.OccurredAtUtc);
        command.Parameters.AddWithValue("available_at_utc", message.AvailableAtUtc);
        command.Parameters.AddWithValue("idempotency_key", message.IdempotencyKey);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }
}
