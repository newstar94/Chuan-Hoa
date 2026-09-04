using ChuanHoa.Domain.Commercial;
using ChuanHoa.Domain.Security;
using ChuanHoa.Domain.Trials;
using ChuanHoa.Application.Scanning;
using Microsoft.Extensions.DependencyInjection;

namespace ChuanHoa.Application;

public static class ServiceCollectionExtensions
{
    public static IServiceCollection AddChuanHoaApplication(this IServiceCollection services)
    {
        services.AddSingleton<IClock, SystemClock>();
        services.AddSingleton<TrialEligibilityResolver>();
        services.AddSingleton<OfferSelector>();
        services.AddSingleton<ExecutionGrantValidator>();
        services.AddSingleton<TechnicalDocumentScanner>();
        return services;
    }
}
