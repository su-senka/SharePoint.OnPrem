using FluentAssertions;
using Microsoft.Extensions.DependencyInjection;
using SharePoint.OnPrem.Abstractions;
using SharePoint.OnPrem.Core;

namespace SharePoint.OnPrem.DependencyInjection.Tests;

public class ServiceCollectionExtensionsTests
{
    [Fact]
    public void AddSharePointOnPrem_RegistersNormalizerAndScope()
    {
        var provider = BuildProvider();

        var normalizer = provider.GetRequiredService<ILoginNameNormalizer>();
        var pathScope = provider.GetRequiredService<ISharePointPathScope>();

        normalizer.Normalize("jnovak").Should().Be("i:0#.w|ACR\\jnovak");
        pathScope.ToServerRelativePath("3255/26/001").Should().Be("/sites/pp/Attachments/3255/26/001");
    }

    [Fact]
    public void AddSharePointOnPrem_RegistersCapabilityClients()
    {
        var provider = BuildProvider();

        provider.GetRequiredService<ISharePointFileClient>().Should().NotBeNull();
        provider.GetRequiredService<ISharePointFolderClient>().Should().NotBeNull();
        provider.GetRequiredService<ISharePointSecurityClient>().Should().NotBeNull();
    }

    [Fact]
    public void AddSharePointOnPrem_ResolvesConfiguredCoreOptions()
    {
        var provider = BuildProvider();

        var options = provider.GetRequiredService<SharePointOnPremOptions>();

        options.SiteBaseUrl.Should().Be("https://sharepoint.local");
        options.HttpTimeoutMinutes.Should().Be(5);
    }

    private static ServiceProvider BuildProvider()
    {
        var services = new ServiceCollection();

        services.AddSharePointOnPrem(
            configureCore: options =>
            {
                options.SiteBaseUrl = "https://sharepoint.local";
                options.HttpTimeoutMinutes = 5;
            },
            configureIdentity: options =>
            {
                options.Domain = "ACR";
                options.ClaimsPrefix = "i:0#.w|";
            },
            configureStorageScope: options =>
            {
                options.BaseFolderServerRelativeUrl = "/sites/pp/Attachments";
            });


        return services.BuildServiceProvider();
    }
}
