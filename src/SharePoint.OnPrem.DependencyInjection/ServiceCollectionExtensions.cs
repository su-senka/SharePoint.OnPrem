using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using SharePoint.OnPrem.Abstractions;
using SharePoint.OnPrem.Core;
using SharePoint.OnPrem.Files;
using SharePoint.OnPrem.Security;

namespace SharePoint.OnPrem.DependencyInjection;

public static class ServiceCollectionExtensions
{
    public static IServiceCollection AddSharePointOnPrem(
        this IServiceCollection services,
        Action<SharePointOnPremOptions>? configureCore = null,
        Action<SharePointIdentityOptions>? configureIdentity = null,
        Action<SharePointStorageScopeOptions>? configureStorageScope = null)
    {
        ArgumentNullException.ThrowIfNull(services);

        services.AddOptions<SharePointOnPremOptions>();
        services.AddOptions<SharePointIdentityOptions>();
        services.AddOptions<SharePointStorageScopeOptions>();

        if (configureCore is not null)
            services.Configure(configureCore);

        if (configureIdentity is not null)
            services.Configure(configureIdentity);

        if (configureStorageScope is not null)
            services.Configure(configureStorageScope);

        services.AddSingleton(sp =>
        {
            var options = sp.GetRequiredService<IOptions<SharePointOnPremOptions>>().Value;
            options.Validate();
            return options;
        });

        services.AddSingleton(sp =>
        {
            var options = sp.GetRequiredService<IOptions<SharePointIdentityOptions>>().Value;
            options.Validate();
            return options;
        });

        services.AddSingleton(sp =>
        {
            var options = sp.GetRequiredService<IOptions<SharePointStorageScopeOptions>>().Value;
            options.Validate();
            return options;
        });

        services.AddSingleton<ILoginNameNormalizer, DefaultLoginNameNormalizer>();
        services.AddSingleton<BaseFolderPathScope>();
        services.AddSingleton<ISharePointServerRelativePathScope>(sp => sp.GetRequiredService<BaseFolderPathScope>());
#pragma warning disable CS0618
        services.AddSingleton<ISharePointPathScope>(sp => sp.GetRequiredService<BaseFolderPathScope>());
#pragma warning restore CS0618
        services.AddSingleton<IFormDigestProvider, SharePointFormDigestProvider>();
        services.AddSingleton<ISharePointRequestExecutor, SharePointRequestExecutor>();

        services.AddHttpClient<ISharePointFileClient, SharePointFileClient>(ConfigureHttpClient);
        services.AddHttpClient<ISharePointFolderClient, SharePointFolderClient>(ConfigureHttpClient);
        services.AddHttpClient<ISharePointSecurityClient, SharePointSecurityClient>(ConfigureHttpClient);

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
        client.Timeout = TimeSpan.FromMinutes(options.HttpTimeoutMinutes);

        if (!client.DefaultRequestHeaders.Contains("Accept"))
            client.DefaultRequestHeaders.Add("Accept", "application/json;odata=nometadata");
    }
}

