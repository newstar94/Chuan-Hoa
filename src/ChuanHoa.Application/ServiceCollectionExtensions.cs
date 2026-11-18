using ChuanHoa.Application.Scanning;
using Microsoft.Extensions.DependencyInjection;

namespace ChuanHoa.Application;

public static class ServiceCollectionExtensions
{
    public static IServiceCollection AddChuanHoaApplication(this IServiceCollection services)
    {
        services.AddSingleton<IClock, SystemClock>();
        services.AddSingleton<TechnicalDocumentScanner>();
        return services;
    }
}
