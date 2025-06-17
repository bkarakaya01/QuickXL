using Microsoft.Extensions.DependencyInjection;
using QuickXL.Core.Services;

namespace QuickXL;

public static class QuickXLServiceCollectionExtensions
{
    public static IServiceCollection AddQuickXL(
        this IServiceCollection services,
        Action<QuickXLOptions>? configure = null)
    {
        QuickXLOptions options = new();
        configure?.Invoke(options);

        services.AddSingleton(options);
        services.AddSingleton<IXLExporter, QuickXLExporter>();
        
        // TODO:
        // services.AddSingleton<IXLImporter, QuickXLImporter>();

        return services;
    }
}
