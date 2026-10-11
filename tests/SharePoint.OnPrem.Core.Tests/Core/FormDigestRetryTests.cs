using System.Net;
using System.Text;
using AwesomeAssertions;
using SharePoint.OnPrem.Abstractions;
using SharePoint.OnPrem.Core.Tests.Contracts;
using SharePoint.OnPrem.Files;

namespace SharePoint.OnPrem.Core.Tests.Core;

public class FormDigestRetryTests
{
    private const string Digest1 = "{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}";
    private const string Digest2 = "{\"FormDigestValue\":\"digest-2\",\"FormDigestTimeoutSeconds\":1800}";

    [Fact]
    public async Task Write_WhenCachedDigestIsRejected_RefreshesDigestAndRetriesOnce()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse(Digest1),
            _ => new HttpResponseMessage(HttpStatusCode.Forbidden),
            _ => SharePointContractTestHelpers.JsonResponse(Digest2),
            _ => new HttpResponseMessage(HttpStatusCode.OK));

        var sut = SharePointContractTestHelpers.CreateFileClient(SharePointContractTestHelpers.CreateHttpClient(handler));

        await sut.DeleteAsync("/sites/team/Docs/a.txt");

        handler.Requests.Should().HaveCount(4);
        handler.Requests[3].Headers["X-RequestDigest"].Should().ContainSingle().Which.Should().Be("digest-2");
        handler.Requests[3].Headers["X-HTTP-Method"].Should().ContainSingle().Which.Should().Be("DELETE");
    }

    [Fact]
    public async Task Write_WhenRetriedWithJsonBody_ResendsSameBody()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse(Digest1),
            _ => new HttpResponseMessage(HttpStatusCode.Forbidden),
            _ => SharePointContractTestHelpers.JsonResponse(Digest2),
            _ => new HttpResponseMessage(HttpStatusCode.NoContent));

        var sut = SharePointContractTestHelpers.CreateFileClient(SharePointContractTestHelpers.CreateHttpClient(handler));

        await sut.UpdateMetadataAsync("/sites/team/Docs/a.txt", new Dictionary<string, object?> { ["Title"] = "Hello" });

        handler.Requests[3].Body.Should().Be(handler.Requests[1].Body).And.Contain("Hello");
        handler.Requests[3].ContentHeaders["Content-Type"].Should().ContainSingle().Which.Should().Contain("odata=nometadata");
    }

    [Fact]
    public async Task Write_WhenStreamedUploadIsRejected_DoesNotRetry()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse(Digest1),
            _ => new HttpResponseMessage(HttpStatusCode.Forbidden));

        var sut = SharePointContractTestHelpers.CreateFileClient(SharePointContractTestHelpers.CreateHttpClient(handler));
        await using var content = new MemoryStream(Encoding.UTF8.GetBytes("hello"));

        var act = () => sut.UploadAsync(new UploadFileRequest("/sites/team/Docs", "a.txt", content));

        await act.Should().ThrowAsync<SharePointUnauthorizedException>();
        handler.Requests.Should().HaveCount(2);
    }

    [Fact]
    public async Task Write_WhenDigestCachingDisabled_DoesNotRetry()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse(Digest1),
            _ => new HttpResponseMessage(HttpStatusCode.Forbidden));

        var sut = new SharePointFileClient(
            SharePointContractTestHelpers.CreateHttpClient(handler),
            new SharePointRequestExecutor(new SharePointFormDigestProvider(new SharePointOnPremOptions
            {
                SiteBaseUrl = "https://sharepoint.contoso.local/sites/team",
                UseFormDigestCaching = false
            })));

        var act = () => sut.DeleteAsync("/sites/team/Docs/a.txt");

        await act.Should().ThrowAsync<SharePointUnauthorizedException>();
        handler.Requests.Should().HaveCount(2);
    }

    [Fact]
    public async Task Read_WhenThrottled_ExposesRetryAfter()
    {
        var handler = new QueueMessageHandler(_ =>
        {
            var response = new HttpResponseMessage(HttpStatusCode.TooManyRequests);
            response.Headers.RetryAfter = new System.Net.Http.Headers.RetryConditionHeaderValue(TimeSpan.FromSeconds(30));
            return response;
        });

        var sut = SharePointContractTestHelpers.CreateFileClient(SharePointContractTestHelpers.CreateHttpClient(handler));

        var act = () => sut.GetFileWebUrlAsync("/sites/team/Docs/a.txt");

        (await act.Should().ThrowAsync<SharePointTransientException>()).Which.RetryAfter.Should().Be(TimeSpan.FromSeconds(30));
    }
}
