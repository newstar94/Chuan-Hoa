using ChuanHoa.Application.Persistence;
using ChuanHoa.Infrastructure.Persistence;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Npgsql;

namespace ChuanHoa.Infrastructure;

public static class ServiceCollectionExtensions
{
    public static IServiceCollection AddChuanHoaInfrastructure(
        this IServiceCollection services,
        IConfiguration configuration)
    {
        var connectionString = configuration.GetConnectionString("ChuanHoa");
        if (string.IsNullOrWhiteSpace(connectionString))
        {
            throw new InvalidOperationException(
                "ConnectionStrings:ChuanHoa is required. Supply it through the deployment secret store or environment; do not commit database credentials.");
        }

        services.AddSingleton(_ =>
        {
            var dataSourceBuilder = new NpgsqlDataSourceBuilder(connectionString);
            return dataSourceBuilder.Build();
        });
        services.AddSingleton<PostgresTransactionRunner>();
        services.AddSingleton<PostgresOutboxWriter>();
        services.AddSingleton<IIdempotencyStore, PostgresIdempotencyStore>();
        return services;
    }
}
