using ChuanHoa.Contracts.Integration;
using Npgsql;

namespace ChuanHoa.Infrastructure.Admin;

public sealed class PostgresIntegrationReplayStore(NpgsqlDataSource dataSource) : IIntegrationReplayStore
{
    public async Task<bool> TryClaimAsync(string clientId, string nonce, DateTimeOffset expiresAtUtc, CancellationToken cancellationToken)
    {
        await using var cleanup = dataSource.CreateCommand("DELETE FROM admin_integration_replay_nonces WHERE expires_at_utc <= now();");
        await cleanup.ExecuteNonQueryAsync(cancellationToken);
        await using var command = dataSource.CreateCommand("""
            INSERT INTO admin_integration_replay_nonces (client_id, nonce, expires_at_utc)
            VALUES ($1, $2, $3)
            ON CONFLICT (client_id, nonce) DO NOTHING
            RETURNING nonce;
            """);
        command.Parameters.AddWithValue(clientId);
        command.Parameters.AddWithValue(nonce);
        command.Parameters.AddWithValue(expiresAtUtc);
        return await command.ExecuteScalarAsync(cancellationToken) is not null;
    }
}
