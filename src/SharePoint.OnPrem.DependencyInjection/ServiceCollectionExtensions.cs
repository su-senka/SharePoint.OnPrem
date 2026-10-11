using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Microsoft.Extensions.Options;
using SharePoint.OnPrem.Abstractions;
using SharePoint.OnPrem.Core;
using SharePoint.OnPrem.Files;
using SharePoint.OnPrem.Security;

namespace SharePoint.OnPrem.DependencyInjection;

/// <summary>
/// Registers the SharePoint on-prem clients with an <see cref="IServiceCollection"/>.
/// </summary>
public static class ServiceCollectionExtensions
{
    /// <summary>
    /// Name of the <see cref="HttpClient"/> shared by all SharePoint clients. Use it with
    /// <c>services.AddHttpClient(HttpClientName)</c> to add handlers, or pass <c>configureHttpClient</c> instead.
    /// </summary>
    public const string HttpClientName = "SharePoint.OnPrem";

    /// <summary>
    /// Registers <see cref="ISharePointFileClient"/>, <see cref="ISharePointFolderClient"/>, <see cref="ISharePointSecurityClient"/>,
    /// <see cref="ISharePointSiteClient"/> and their dependencies.
    /// </summary>
    /// <param name="services">The service collection.</param>
    /// <param name="configureCore">Connection settings; <see cref="SharePointOnPremOptions.SiteBaseUrl"/> is required.</param>
    /// <param name="configureIdentity">Login normalization settings used by the security client.</param>
    /// <param name="configureStorageScope">Base folder for <see cref="ISharePointServerRelativePathScope"/>.</param>
    /// <param name="configureHttpClient">
    /// Customizes the shared HTTP client, typically to set credentials:
    /// <c>b => b.ConfigurePrimaryHttpMessageHandler(() => new HttpClientHandler { UseDefaultCredentials = true })</c>.
    /// </param>
    public static IServiceCollection AddSharePointOnPrem(
        this IServiceCollection services,
        Action<SharePointOnPremOptions>? configureCore = null,
        Action<SharePointIdentityOptions>? configureIdentity = null,
        Action<SharePointStorageScopeOptions>? configureStorageScope = null,
        Action<IHttpClientBuilder>? configureHttpClient = null)
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

        services.TryAddSingleton(sp =>
        {
            var options = sp.GetRequiredService<IOptions<SharePointOnPremOptions>>().Value;
            options.Validate();
            return options;
        });

        services.TryAddSingleton(sp =>
        {
            var options = sp.GetRequiredService<IOptions<SharePointIdentityOptions>>().Value;
            options.Validate();
            return options;
        });

        services.TryAddSingleton(sp =>
        {
            var options = sp.GetRequiredService<IOptions<SharePointStorageScopeOptions>>().Value;
            options.Validate();
            return options;
        });

        services.TryAddSingleton<ILoginNameNormalizer, DefaultLoginNameNormalizer>();
        services.TryAddSingleton<BaseFolderPathScope>();
        services.TryAddSingleton<ISharePointServerRelativePathScope>(sp => sp.GetRequiredService<BaseFolderPathScope>());
#pragma warning disable CS0618
        services.TryAddSingleton<ISharePointPathScope>(sp => sp.GetRequiredService<BaseFolderPathScope>());
#pragma warning restore CS0618
        services.TryAddSingleton<IFormDigestProvider, SharePointFormDigestProvider>();
        services.TryAddSingleton<ISharePointRequestExecutor, SharePointRequestExecutor>();

        // One named client for all three typed clients: a single handler pipeline to configure (credentials, proxies, resilience).
        var httpClientBuilder = services.AddHttpClient(HttpClientName, ConfigureHttpClient);
        httpClientBuilder.AddTypedClient<ISharePointFileClient, SharePointFileClient>();
        httpClientBuilder.AddTypedClient<ISharePointFolderClient, SharePointFolderClient>();
        httpClientBuilder.AddTypedClient<ISharePointSecurityClient, SharePointSecurityClient>();
        httpClientBuilder.AddTypedClient<ISharePointSiteClient, SharePointSiteClient>();
        configureHttpClient?.Invoke(httpClientBuilder);

        return services;
    }

    /// <summary>
    /// Registers the SharePoint clients with settings bound from configuration.
    /// </summary>
    /// <param name="services">The service collection.</param>
    /// <param name="configuration">
    /// The <c>SharePoint</c> section: <c>SiteBaseUrl</c>, <c>HttpTimeoutMinutes</c>, <c>UseFormDigestCaching</c>,
    /// plus optional <c>Identity</c> (<c>Domain</c>, <c>ClaimsPrefix</c>) and <c>StorageScope</c> (<c>BaseFolderServerRelativeUrl</c>) subsections.
    /// </param>
    /// <param name="configureHttpClient">Customizes the shared HTTP client, typically to set credentials.</param>
    public static IServiceCollection AddSharePointOnPrem(
        this IServiceCollection services,
        IConfiguration configuration,
        Action<IHttpClientBuilder>? configureHttpClient = null)
    {
        ArgumentNullException.ThrowIfNull(services);
        ArgumentNullException.ThrowIfNull(configuration);

        services.Configure<SharePointOnPremOptions>(configuration);
        services.Configure<SharePointIdentityOptions>(configuration.GetSection("Identity"));
        services.Configure<SharePointStorageScopeOptions>(configuration.GetSection("StorageScope"));

        return services.AddSharePointOnPrem(configureHttpClient: configureHttpClient);
    }

    private static void ConfigureHttpClient(IServiceProvider serviceProvider, HttpClient client)
    {
        var options = serviceProvider.GetRequiredService<SharePointOnPremOptions>();

        var siteBaseUrl = options.SiteBaseUrl.EndsWith('/')
            ? options.SiteBaseUrl
            : options.SiteBaseUrl + "/";

        client.BaseAddress = new Uri(siteBaseUrl, UriKind.Absolute);
        client.Timeout = TimeSpan.FromMinutes(options.HttpTimeoutMinutes);

        if (!client.DefaultRequestHeaders.Contains("Accept"))
            client.DefaultRequestHeaders.Add("Accept", "application/json;odata=nometadata");
    }
}
