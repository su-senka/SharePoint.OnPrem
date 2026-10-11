using AwesomeAssertions;
using SharePoint.OnPrem.Abstractions;
using SharePoint.OnPrem.Core.Tests.Contracts;

namespace SharePoint.OnPrem.Core.Tests.Core;

public class SharePointSiteClientTests
{
    [Fact]
    public async Task GetWebInfoAsync_ReadsWebAndServerBuild()
    {
        var handler = new QueueMessageHandler(_ =>
        {
            var response = SharePointContractTestHelpers.JsonResponse(
                "{\"Title\":\"Team\",\"Url\":\"https://sharepoint.contoso.local/sites/team\",\"ServerRelativeUrl\":\"/sites/team\",\"Language\":1033}");
            response.Headers.Add("MicrosoftSharePointTeamServices", "16.0.0.10337");
            return response;
        });

        var sut = CreateSut(handler);

        var web = await sut.GetWebInfoAsync();

        web.Should().Be(new SharePointWebInfo("Team", "https://sharepoint.contoso.local/sites/team", "/sites/team", 1033, "16.0.0.10337"));
        web.ProductName.Should().Be("SharePoint 2019");
        handler.Requests[0].Uri.Should().EndWith("_api/web?$select=Title,Url,ServerRelativeUrl,Language");
    }

    [Fact]
    public async Task GetCurrentUserAsync_ReadsAuthenticatedUser()
    {
        var handler = new QueueMessageHandler(_ => SharePointContractTestHelpers.JsonResponse(
            "{\"Id\":12,\"LoginName\":\"i:0#.w|contoso\\\\svc-sharepoint\",\"Title\":\"SharePoint Service\",\"Email\":\"\",\"IsSiteAdmin\":true}"));

        var user = await CreateSut(handler).GetCurrentUserAsync();

        user.Should().Be(new SharePointUserInfo(12, "i:0#.w|contoso\\svc-sharepoint", "SharePoint Service", null, true));
    }

    [Theory]
    [InlineData("15.0.0.4569", "SharePoint 2013")]
    [InlineData("16.0.0.5317", "SharePoint 2016")]
    [InlineData("16.0.0.10337", "SharePoint 2019")]
    [InlineData("16.0.0.17928", "SharePoint Server Subscription Edition")]
    [InlineData("garbage", null)]
    [InlineData(null, null)]
    public void GetProductName_MapsBuildToProduct(string? build, string? expected)
    {
        SharePointProductVersion.GetProductName(build).Should().Be(expected);
    }

    private static SharePointSiteClient CreateSut(QueueMessageHandler handler)
        => new(
            SharePointContractTestHelpers.CreateHttpClient(handler),
            new SharePointRequestExecutor(new SharePointFormDigestProvider(new SharePointOnPremOptions
            {
                SiteBaseUrl = "https://sharepoint.contoso.local/sites/team"
            })));
}
