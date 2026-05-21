using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using PP.Application.Services.SharePoint;
using SharePoint.OnPrem.Core;

namespace SharePoint.OnPrem.PlanProvozuAdapter;

/// <summary>
/// Registers the legacy PlanProvozu-compatible adapter on top of the standalone SharePoint.OnPrem services.
/// Call this after <c>AddSharePointOnPrem(...)</c>.
/// </summary>
public static class ServiceCollectionExtensions
{
    public static IServiceCollection AddPlanProvozuSharePointCompatibilityAdapter(this IServiceCollection services)
    {
        ArgumentNullException.ThrowIfNull(services);

        services.AddHttpClient<PlanProvozuSharePointServiceAdapter>(ConfigureHttpClient);
        services.AddScoped<ISharePointService>(sp => sp.GetRequiredService<PlanProvozuSharePointServiceAdapter>());

        return services;
    }

    private static void ConfigureHttpClient(IServiceProvider serviceProvider, HttpClient client)
    {
        var options = serviceProvider.GetRequiredService<IOptions<SharePointOnPremOptions>>().Value;
        options.Validate();

        var siteBaseUrl = options.SiteBaseUrl.EndsWith("/", StringComparison.Ordinal)
            ? options.SiteBaseUrl
            : options.SiteBaseUrl + "/";

        client.BaseAddress = new Uri(siteBaseUrl, UriKind.Absolute);
        client.Timeout = TimeSpan.FromMinutes((double)options.HttpTimeoutMinutes);

        if (!client.DefaultRequestHeaders.Contains("Accept"))
            client.DefaultRequestHeaders.Add("Accept", "application/json;odata=nometadata");
    }
}

