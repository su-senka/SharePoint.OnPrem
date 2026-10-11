using AwesomeAssertions;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using SharePoint.OnPrem.Abstractions;
using SharePoint.OnPrem.Core;
using Xunit;

namespace SharePoint.OnPrem.DependencyInjection.Tests;

public class ServiceCollectionExtensionsTests
{
    [Fact]
    public void AddSharePointOnPrem_RegistersNormalizerAndScope()
    {
        var provider = BuildProvider();

        var normalizer = provider.GetRequiredService<ILoginNameNormalizer>();
        var pathScope = provider.GetRequiredService<ISharePointServerRelativePathScope>();
#pragma warning disable CS0618
        var legacyPathScope = provider.GetRequiredService<ISharePointPathScope>();
#pragma warning restore CS0618

        normalizer.Normalize("jdoe").Should().Be("i:0#.w|CONTOSO\\jdoe");
        pathScope.ToServerRelativePath("customers/1042/2026").Should().Be("/sites/team/Attachments/customers/1042/2026");
        legacyPathScope.ToServerRelativePath("customers/1042/2026").Should().Be("/sites/team/Attachments/customers/1042/2026");
    }

    [Fact]
    public void AddSharePointOnPrem_RegistersCapabilityClients()
    {
        var provider = BuildProvider();

        provider.GetRequiredService<ISharePointFileClient>().Should().NotBeNull();
        provider.GetRequiredService<ISharePointFolderClient>().Should().NotBeNull();
        provider.GetRequiredService<ISharePointSecurityClient>().Should().NotBeNull();
        provider.GetRequiredService<ISharePointSiteClient>().Should().NotBeNull();
    }

    [Fact]
    public void AddSharePointOnPrem_ResolvesConfiguredCoreOptions()
    {
        var provider = BuildProvider();

        var options = provider.GetRequiredService<SharePointOnPremOptions>();

        options.SiteBaseUrl.Should().Be("https://sharepoint.contoso.local");
        options.HttpTimeoutMinutes.Should().Be(5);
    }

    [Fact]
    public void AddSharePointOnPrem_ConfigureHttpClient_AppliesToSharedNamedClient()
    {
        var services = new ServiceCollection();
        var configuredClientNames = new List<string>();

        services.AddSharePointOnPrem(
            configureCore: options => options.SiteBaseUrl = "https://sharepoint.contoso.local/sites/team",
            configureHttpClient: builder => configuredClientNames.Add(builder.Name));

        using var provider = services.BuildServiceProvider();
        var client = provider.GetRequiredService<IHttpClientFactory>().CreateClient(ServiceCollectionExtensions.HttpClientName);

        configuredClientNames.Should().Equal(ServiceCollectionExtensions.HttpClientName);
        client.BaseAddress.Should().Be(new Uri("https://sharepoint.contoso.local/sites/team/"));
        provider.GetRequiredService<ISharePointFolderClient>().Should().NotBeNull();
    }

    [Fact]
    public void AddSharePointOnPrem_FromConfiguration_BindsAllSections()
    {
        var configuration = new ConfigurationBuilder()
            .AddInMemoryCollection(new Dictionary<string, string?>
            {
                ["SharePoint:SiteBaseUrl"] = "https://sharepoint.contoso.local/sites/team",
                ["SharePoint:HttpTimeoutMinutes"] = "10",
                ["SharePoint:Identity:Domain"] = "CONTOSO",
                ["SharePoint:StorageScope:BaseFolderServerRelativeUrl"] = "/sites/team/Attachments"
            })
            .Build();

        var services = new ServiceCollection();
        services.AddSharePointOnPrem(configuration.GetSection("SharePoint"));
        using var provider = services.BuildServiceProvider();

        provider.GetRequiredService<SharePointOnPremOptions>().HttpTimeoutMinutes.Should().Be(10);
        provider.GetRequiredService<ILoginNameNormalizer>().Normalize("jdoe").Should().Be("i:0#.w|CONTOSO\\jdoe");
        provider.GetRequiredService<ISharePointServerRelativePathScope>().ToServerRelativePath("2026").Should().Be("/sites/team/Attachments/2026");
    }

    [Fact]
    public void AddSharePointOnPrem_KeepsNormalizerRegisteredBeforehand()
    {
        var services = new ServiceCollection();
        services.AddSingleton<ILoginNameNormalizer, UpperCaseNormalizer>();
        services.AddSharePointOnPrem(configureCore: options => options.SiteBaseUrl = "https://sharepoint.contoso.local");

        using var provider = services.BuildServiceProvider();

        provider.GetRequiredService<ILoginNameNormalizer>().Should().BeOfType<UpperCaseNormalizer>();
    }

    private sealed class UpperCaseNormalizer : ILoginNameNormalizer
    {
        public string Normalize(string input) => input.ToUpperInvariant();
    }

    private static ServiceProvider BuildProvider()
    {
        var services = new ServiceCollection();

        services.AddSharePointOnPrem(
            configureCore: options =>
            {
                options.SiteBaseUrl = "https://sharepoint.contoso.local";
                options.HttpTimeoutMinutes = 5;
            },
            configureIdentity: options =>
            {
                options.Domain = "CONTOSO";
                options.ClaimsPrefix = "i:0#.w|";
            },
            configureStorageScope: options =>
            {
                options.BaseFolderServerRelativeUrl = "/sites/team/Attachments";
            });


        return services.BuildServiceProvider();
    }
}

