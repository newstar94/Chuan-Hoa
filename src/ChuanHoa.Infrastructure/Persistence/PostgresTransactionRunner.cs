using System.Data;
using Npgsql;

namespace ChuanHoa.Infrastructure.Persistence;

public sealed class PostgresTransactionRunner(NpgsqlDataSource dataSource)
{
    public async Task<T> ExecuteAsync<T>(
        Func<NpgsqlConnection, NpgsqlTransaction, CancellationToken, Task<T>> command,
        IsolationLevel isolationLevel,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(command);

        await using var connection = await dataSource.OpenConnectionAsync(cancellationToken);
        await using var transaction = await connection.BeginTransactionAsync(isolationLevel, cancellationToken);

        try
        {
            var result = await command(connection, transaction, cancellationToken);
            await transaction.CommitAsync(cancellationToken);
            return result;
        }
        catch
        {
            await transaction.RollbackAsync(CancellationToken.None);
            throw;
        }
    }
}
