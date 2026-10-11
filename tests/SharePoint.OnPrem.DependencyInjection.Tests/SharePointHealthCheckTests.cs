using System.Net;
using AwesomeAssertions;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Diagnostics.HealthChecks;
using SharePoint.OnPrem.Abstractions;
using SharePoint.OnPrem.HealthChecks;

namespace SharePoint.OnPrem.DependencyInjection.Tests;

public class SharePointHealthCheckTests
{
    [Fact]
    public async Task CheckHealth_WhenSharePointAnswers_IsHealthyAndReportsAccount()
    {
        var report = await RunAsync(new FakeSiteClient());

        var entry = report.Entries["sharepoint"];
        entry.Status.Should().Be(HealthStatus.Healthy);
        entry.Description.Should().Contain("i:0#.w|contoso\\svc-sharepoint");
        entry.Data["product"].Should().Be("SharePoint 2019");
    }

    [Fact]
    public async Task CheckHealth_WhenCredentialsRejected_IsUnhealthyWithCredentialHint()
    {
        var report = await RunAsync(new FakeSiteClient(new SharePointUnauthorizedException("401") { StatusCode = HttpStatusCode.Unauthorized }));

        var entry = report.Entries["sharepoint"];
        entry.Status.Should().Be(HealthStatus.Unhealthy);
        entry.Description.Should().Contain("credentials");
    }

    [Fact]
    public async Task CheckHealth_WhenServerUnreachable_UsesConfiguredFailureStatus()
    {
        var report = await RunAsync(new FakeSiteClient(new HttpRequestException("No such host")), HealthStatus.Degraded);

        report.Entries["sharepoint"].Status.Should().Be(HealthStatus.Degraded);
        report.Entries["sharepoint"].Description.Should().Contain("unreachable");
    }

    private static async Task<HealthReport> RunAsync(ISharePointSiteClient siteClient, HealthStatus? failureStatus = null)
    {
        var services = new ServiceCollection();
        services.AddLogging();
        services.AddSingleton(siteClient);
        services.AddHealthChecks().AddSharePointOnPrem(failureStatus: failureStatus);

        await using var provider = services.BuildServiceProvider();
        return await provider.GetRequiredService<HealthCheckService>().CheckHealthAsync();
    }

    private sealed class FakeSiteClient(Exception? failure = null) : ISharePointSiteClient
    {
        public Task<SharePointWebInfo> GetWebInfoAsync(CancellationToken ct = default)
            => Task.FromResult(new SharePointWebInfo("Team", "https://sharepoint.contoso.local/sites/team", "/sites/team", 1033, "16.0.0.10337"));

        public Task<SharePointUserInfo> GetCurrentUserAsync(CancellationToken ct = default)
            => failure is null
                ? Task.FromResult(new SharePointUserInfo(12, "i:0#.w|contoso\\svc-sharepoint", "SharePoint Service", null, false))
                : Task.FromException<SharePointUserInfo>(failure);
    }
}
