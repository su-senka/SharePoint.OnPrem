using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using AwesomeAssertions;
using SharePoint.OnPrem.Abstractions;
using SharePoint.OnPrem.Files;
using SharePoint.OnPrem.Security;
using Xunit;

namespace SharePoint.OnPrem.Core.Tests.Contracts;

public class SharePointFormDigestProviderTests
{
    [Fact]
    public async Task GetDigestAsync_WhenCachingEnabled_RequestsContextInfoOnlyOnce()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = new SharePointFormDigestProvider(new SharePointOnPremOptions
        {
            SiteBaseUrl = "https://sharepoint.contoso.local/sites/team",
            UseFormDigestCaching = true
        });

        var first = await sut.GetDigestAsync(client);
        var second = await sut.GetDigestAsync(client);

        first.Should().Be("digest-1");
        second.Should().Be("digest-1");
        handler.Requests.Should().HaveCount(1);
        handler.Requests.Single().Uri.Should().EndWith("/_api/contextinfo");
    }

    [Fact]
    public async Task GetDigestAsync_WhenCachingDisabled_RequestsContextInfoEachTime()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-2\",\"FormDigestTimeoutSeconds\":1800}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = new SharePointFormDigestProvider(new SharePointOnPremOptions
        {
            SiteBaseUrl = "https://sharepoint.contoso.local/sites/team",
            UseFormDigestCaching = false
        });

        var first = await sut.GetDigestAsync(client);
        var second = await sut.GetDigestAsync(client);

        first.Should().Be("digest-1");
        second.Should().Be("digest-2");
        handler.Requests.Should().HaveCount(2);
    }
}

public class SharePointRequestExecutorTests
{
    [Fact]
    public async Task SendAsync_WhenResponseIsNotSuccessful_ThrowsTypedExceptionWithDetail()
    {
        var handler = new QueueMessageHandler(_ => SharePointContractTestHelpers.JsonResponse("{\"detail\":\"File missing\"}", HttpStatusCode.NotFound));
        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var digestProvider = new SharePointFormDigestProvider(new SharePointOnPremOptions { SiteBaseUrl = "https://sharepoint.contoso.local/sites/team" });
        var sut = new SharePointRequestExecutor(digestProvider);

        var act = async () =>
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, "_api/web/test");
            using var _ = await sut.SendAsync(client, request);
        };

        await act.Should().ThrowAsync<SharePointNotFoundException>()
            .WithMessage("*File missing*");
    }
    [Fact]
    public async Task SendAsync_WhenJsonLightErrorWithMessageObject_UsesMessageValueAndStatusCode()
    {
        var handler = new QueueMessageHandler(_ => SharePointContractTestHelpers.JsonResponse(
            "{\"odata.error\":{\"code\":\"-2130575338, Microsoft.SharePoint.SPException\",\"message\":{\"lang\":\"en-US\",\"value\":\"The file does not exist.\"}}}",
            HttpStatusCode.BadRequest));
        var sut = CreateExecutor();

        var act = async () =>
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, "_api/web/test");
            using var _ = await sut.SendAsync(SharePointContractTestHelpers.CreateHttpClient(handler), request);
        };

        var exception = (await act.Should().ThrowAsync<SharePointException>()).Which;
        exception.Message.Should().EndWith("Detail=The file does not exist.");
        exception.StatusCode.Should().Be(HttpStatusCode.BadRequest);
    }

    [Fact]
    public async Task SendAsync_WhenVerboseErrorWithMessageObject_DoesNotCrashErrorMapping()
    {
        var handler = new QueueMessageHandler(_ => SharePointContractTestHelpers.JsonResponse(
            "{\"error\":{\"code\":\"x\",\"message\":{\"lang\":\"en-US\",\"value\":\"Access denied.\"}}}",
            HttpStatusCode.Forbidden));
        var sut = CreateExecutor();

        var act = async () =>
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, "_api/web/test");
            using var _ = await sut.SendAsync(SharePointContractTestHelpers.CreateHttpClient(handler), request);
        };

        await act.Should().ThrowAsync<SharePointUnauthorizedException>().WithMessage("*Access denied.*");
    }

    [Theory]
    [InlineData(HttpStatusCode.TooManyRequests)]
    [InlineData(HttpStatusCode.BadGateway)]
    [InlineData(HttpStatusCode.ServiceUnavailable)]
    [InlineData(HttpStatusCode.GatewayTimeout)]
    public async Task SendAsync_WhenTransientStatus_ThrowsTransientException(HttpStatusCode statusCode)
    {
        var handler = new QueueMessageHandler(_ => new HttpResponseMessage(statusCode) { Content = new StringContent(new string('x', 5000)) });
        var sut = CreateExecutor();

        var act = async () =>
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, "_api/web/test");
            using var _ = await sut.SendAsync(SharePointContractTestHelpers.CreateHttpClient(handler), request);
        };

        var exception = (await act.Should().ThrowAsync<SharePointTransientException>()).Which;
        exception.Message.Length.Should().BeLessThan(1500);
    }

    private static SharePointRequestExecutor CreateExecutor()
        => new(new SharePointFormDigestProvider(new SharePointOnPremOptions { SiteBaseUrl = "https://sharepoint.contoso.local/sites/team" }));
}

public class SharePointFileClientTests
{
    [Fact]
    public async Task UploadAsync_PostsToFilesAddEndpointAndReturnsStoredFile()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => SharePointContractTestHelpers.JsonResponse("{}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        await using var content = new MemoryStream(Encoding.UTF8.GetBytes("hello"));
        var result = await sut.UploadAsync(new UploadFileRequest(
            "/sites/team/Attachments",
            "test file.txt",
            content,
            "text/plain"));

        result.ServerRelativeUrl.Should().Be("/sites/team/Attachments/test file.txt");
        result.FileName.Should().Be("test file.txt");

        handler.Requests.Should().HaveCount(2);
        handler.Requests[1].Method.Should().Be(HttpMethod.Post);
        handler.Requests[1].Uri.Should().Contain("GetFolderByServerRelativeUrl('%2Fsites%2Fteam%2FAttachments')/Files/add(url='test%20file.txt',overwrite=true)");
        handler.Requests[1].Headers["X-RequestDigest"].Should().ContainSingle().Which.Should().Be("digest-1");
        handler.Requests[1].ContentHeaders["Content-Type"].Should().ContainSingle().Which.Should().StartWith("text/plain");
    }

    [Fact]
    public async Task UploadAsync_DoesNotDisposeCallerStream()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => SharePointContractTestHelpers.JsonResponse("{}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        await using var content = new MemoryStream(Encoding.UTF8.GetBytes("hello"));
        await sut.UploadAsync(new UploadFileRequest("/sites/team/Attachments", "a.txt", content));

        content.CanRead.Should().BeTrue();
        handler.Requests[1].Body.Should().Be("hello");
    }

    [Fact]
    public async Task ExistsAsync_WhenPathContainsApostrophe_DoublesItInsideODataLiteral()
    {
        var handler = new QueueMessageHandler(_ => SharePointContractTestHelpers.JsonResponse("{}"));
        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        await sut.ExistsAsync("/sites/team/Docs/O'Brien.docx");

        handler.Requests[0].Uri.Should().Contain("GetFileByServerRelativeUrl('%2Fsites%2Fteam%2FDocs%2FO%27%27Brien.docx')");
    }

    [Fact]
    public async Task UploadAsync_WhenServerReturnsStoredFileData_UsesResponseValues()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => SharePointContractTestHelpers.JsonResponse("{\"ServerRelativeUrl\":\"/sites/team/Attachments/server-name.txt\",\"Name\":\"server-name.txt\"}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        await using var content = new MemoryStream(Encoding.UTF8.GetBytes("hello"));
        var result = await sut.UploadAsync(new UploadFileRequest(
            "/sites/team/Attachments",
            "client-name.txt",
            content));

        result.ServerRelativeUrl.Should().Be("/sites/team/Attachments/server-name.txt");
        result.FileName.Should().Be("server-name.txt");
    }

    [Fact]
    public async Task DownloadAsync_ReturnsStreamContentTypeAndLeafFileName()
    {
        var response = new HttpResponseMessage(HttpStatusCode.OK)
        {
            Content = new ByteArrayContent(Encoding.UTF8.GetBytes("hello"))
        };
        response.Content.Headers.ContentType = new MediaTypeHeaderValue("text/plain");

        var handler = new QueueMessageHandler(_ => response);
        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        await using var result = await sut.DownloadAsync("/sites/team/Attachments/test.txt");

        using var reader = new StreamReader(result.Content);
        var text = await reader.ReadToEndAsync();
        text.Should().Be("hello");
        result.ContentType.Should().Be("text/plain");
        result.FileName.Should().Be("test.txt");
    }

    [Fact]
    public async Task ExistsAsync_WhenFileExists_ReturnsTrue()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"ServerRelativeUrl\":\"/sites/team/Attachments/test.txt\"}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        var result = await sut.ExistsAsync("/sites/team/Attachments/test.txt");

        result.Should().BeTrue();
        handler.Requests.Should().ContainSingle();
        handler.Requests[0].Uri.Should().Contain("GetFileByServerRelativeUrl('%2Fsites%2Fteam%2FAttachments%2Ftest.txt')");
    }

    [Fact]
    public async Task ExistsAsync_WhenFileIsMissing_ReturnsFalse()
    {
        var handler = new QueueMessageHandler(_ => new HttpResponseMessage(HttpStatusCode.NotFound));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        var result = await sut.ExistsAsync("/sites/team/Attachments/missing.txt");

        result.Should().BeFalse();
    }

    [Fact]
    public async Task GetWebUrlAsync_WhenLinkingUrlMissing_FallsBackToAbsoluteServerRelativeUrl()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"ServerRelativeUrl\":\"/sites/team/Attachments/test.docx\"}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        var result = await sut.GetFileWebUrlAsync("/sites/team/Attachments/test.docx");

        result.Should().Be("https://sharepoint.contoso.local/sites/team/Attachments/test.docx");
    }

    [Fact]
    public async Task DeleteAsync_SendsDeleteSemanticsViaPostWithDigest()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => SharePointContractTestHelpers.JsonResponse("{}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        await sut.DeleteAsync("/sites/team/Attachments/test.txt");

        handler.Requests.Should().HaveCount(2);
        handler.Requests[1].Headers["X-HTTP-Method"].Should().ContainSingle().Which.Should().Be("DELETE");
        handler.Requests[1].Headers["IF-MATCH"].Should().ContainSingle().Which.Should().Be("*");
    }

    [Fact]
    public async Task CopyAsync_SendsCopyToEndpointWithDigestAndReturnsDestination()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => SharePointContractTestHelpers.JsonResponse("{}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        var result = await sut.CopyAsync(
            "/sites/team/Attachments/source.txt",
            "/sites/team/Archive",
            "copy.txt",
            overwrite: true);

        result.ServerRelativeUrl.Should().Be("/sites/team/Archive/copy.txt");
        result.FileName.Should().Be("copy.txt");
        handler.Requests.Should().HaveCount(2);
        handler.Requests[1].Uri.Should().Contain("copyto(strnewurl='%2Fsites%2Fteam%2FArchive%2Fcopy.txt',boverwrite=true)");
    }

    [Fact]
    public async Task MoveAsync_SendsMoveToEndpointWithFlagsAndReturnsDestination()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => SharePointContractTestHelpers.JsonResponse("{}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        var result = await sut.MoveAsync(
            "/sites/team/Attachments/source.txt",
            "/sites/team/Archive",
            "moved.txt",
            overwrite: false);

        result.ServerRelativeUrl.Should().Be("/sites/team/Archive/moved.txt");
        result.FileName.Should().Be("moved.txt");
        handler.Requests.Should().HaveCount(2);
        handler.Requests[1].Uri.Should().Contain("moveto(newurl='%2Fsites%2Fteam%2FArchive%2Fmoved.txt',flags=0)");
    }

    [Fact]
    public async Task RenameAsync_UsesMoveToWithinSameFolderAndReturnsRenamedFile()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => SharePointContractTestHelpers.JsonResponse("{}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        var result = await sut.RenameAsync("/sites/team/Attachments/source.txt", "renamed.txt");

        result.ServerRelativeUrl.Should().Be("/sites/team/Attachments/renamed.txt");
        result.FileName.Should().Be("renamed.txt");
        handler.Requests.Should().HaveCount(2);
        handler.Requests[1].Uri.Should().Contain("GetFileByServerRelativeUrl('%2Fsites%2Fteam%2FAttachments%2Fsource.txt')/moveto(newurl='%2Fsites%2Fteam%2FAttachments%2Frenamed.txt',flags=1)");
    }

    [Fact]
    public async Task UpdateMetadataAsync_SendsMergeHeadersAndJsonPayload()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => new HttpResponseMessage(HttpStatusCode.NoContent));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        await sut.UpdateMetadataAsync(
            "/sites/team/Attachments/test.txt",
            new Dictionary<string, object?>
            {
                ["Title"] = "Quarterly Report",
                ["Category"] = "Finance"
            });

        handler.Requests.Should().HaveCount(2);
        handler.Requests[1].Method.Should().Be(HttpMethod.Post);
        handler.Requests[1].Uri.Should().Contain("GetFileByServerRelativeUrl('%2Fsites%2Fteam%2FAttachments%2Ftest.txt')/ListItemAllFields");
        handler.Requests[1].Headers["X-HTTP-Method"].Should().ContainSingle().Which.Should().Be("MERGE");
        handler.Requests[1].Headers["IF-MATCH"].Should().ContainSingle().Which.Should().Be("*");
        handler.Requests[1].Headers["X-RequestDigest"].Should().ContainSingle().Which.Should().Be("digest-1");
        handler.Requests[1].Body.Should().Contain("\"Title\":\"Quarterly Report\"");
        handler.Requests[1].Body.Should().Contain("\"Category\":\"Finance\"");
    }

    [Fact]
    public async Task GetMetadataAsync_WithSelectFields_SendsSelectQueryAndReturnsMetadataMap()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"Title\":\"Quarterly Report\",\"Category\":\"Finance\"}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        var result = await sut.GetMetadataAsync(
            "/sites/team/Attachments/test.txt",
            ["Title", "Category"]);

        handler.Requests.Should().HaveCount(1);
        handler.Requests[0].Uri.Should().Contain("GetFileByServerRelativeUrl('%2Fsites%2Fteam%2FAttachments%2Ftest.txt')/ListItemAllFields?$select=Title,Category");
        result.Should().ContainKey("Title");
        result.Should().ContainKey("Category");
        result["Title"].GetString().Should().Be("Quarterly Report");
        result["Category"].GetString().Should().Be("Finance");
    }

    [Fact]
    public async Task TryGetMetadataAsync_WhenFileExists_ReturnsMetadataMap()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"Title\":\"Quarterly Report\"}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        var result = await sut.TryGetMetadataAsync("/sites/team/Attachments/test.txt", ["Title"]);

        result.Should().NotBeNull();
        result.Should().ContainKey("Title");
        result["Title"].GetString().Should().Be("Quarterly Report");
    }

    [Fact]
    public async Task TryGetMetadataAsync_WhenFileIsMissing_ReturnsNull()
    {
        var handler = new QueueMessageHandler(_ => new HttpResponseMessage(HttpStatusCode.NotFound));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        var result = await sut.TryGetMetadataAsync("/sites/team/Attachments/missing.txt", ["Title"]);

        result.Should().BeNull();
    }

    [Fact]
    public async Task GetMetadataValueAsync_WhenFieldExists_ReturnsTypedValue()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"Title\":\"Quarterly Report\"}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        var result = await sut.GetMetadataValueAsync<string>("/sites/team/Attachments/test.txt", "Title");

        result.Should().Be("Quarterly Report");
        handler.Requests.Should().ContainSingle();
        handler.Requests[0].Uri.Should().Contain("$select=Title");
    }

    [Fact]
    public async Task GetMetadataValueAsync_WhenFieldMissing_ReturnsDefault()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"Category\":\"Finance\"}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        var result = await sut.GetMetadataValueAsync<string>("/sites/team/Attachments/test.txt", "Title");

        result.Should().BeNull();
    }

    [Fact]
    public async Task TryGetMetadataValueAsync_WhenFileExists_ReturnsTypedValue()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"Title\":\"Quarterly Report\"}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        var result = await sut.TryGetMetadataValueAsync<string>("/sites/team/Attachments/test.txt", "Title");

        result.Should().Be("Quarterly Report");
    }

    [Fact]
    public async Task TryGetMetadataValueAsync_WhenFileMissing_ReturnsDefault()
    {
        var handler = new QueueMessageHandler(_ => new HttpResponseMessage(HttpStatusCode.NotFound));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        var result = await sut.TryGetMetadataValueAsync<string>("/sites/team/Attachments/missing.txt", "Title");

        result.Should().BeNull();
    }
}

public class SharePointFolderClientTests
{
    [Fact]
    public async Task ExistsAsync_WhenFolderIsMissing_ReturnsFalse()
    {
        var handler = new QueueMessageHandler(_ => new HttpResponseMessage(HttpStatusCode.NotFound));
        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFolderClient(client);

        var result = await sut.ExistsAsync("/sites/team/Attachments/2026");

        result.Should().BeFalse();
    }

    [Fact]
    public async Task CreateAsync_WhenFolderAlreadyExists_DoesNotThrow()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => new HttpResponseMessage(HttpStatusCode.Conflict));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFolderClient(client);

        var act = async () => await sut.CreateAsync("/sites/team/Attachments/2026");

        await act.Should().NotThrowAsync();
        handler.Requests[1].Headers["X-RequestDigest"].Should().ContainSingle().Which.Should().Be("digest-1");
    }

    [Fact]
    public async Task DeleteAsync_SendsDeleteSemanticsViaPostWithDigest()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => new HttpResponseMessage(HttpStatusCode.OK));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFolderClient(client);

        await sut.DeleteAsync("/sites/team/Attachments/2026");

        handler.Requests.Should().HaveCount(2);
        handler.Requests[1].Method.Should().Be(HttpMethod.Post);
        handler.Requests[1].Uri.Should().Contain("GetFolderByServerRelativeUrl('%2Fsites%2Fteam%2FAttachments%2F2026')");
        handler.Requests[1].Headers["X-HTTP-Method"].Should().ContainSingle().Which.Should().Be("DELETE");
        handler.Requests[1].Headers["IF-MATCH"].Should().ContainSingle().Which.Should().Be("*");
        handler.Requests[1].Headers["X-RequestDigest"].Should().ContainSingle().Which.Should().Be("digest-1");
    }

    [Fact]
    public async Task DeleteAsync_WhenFolderDoesNotExist_DoesNotThrow()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => new HttpResponseMessage(HttpStatusCode.NotFound));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFolderClient(client);

        var act = async () => await sut.DeleteAsync("/sites/team/Attachments/missing");

        await act.Should().NotThrowAsync();
    }

    [Fact]
    public async Task EnsureFolderPathAsync_CreatesEachMissingSegmentBelowTheSiteOnce()
    {
        var handler = new QueueMessageHandler(
            _ => new HttpResponseMessage(HttpStatusCode.NotFound),
            _ => SharePointContractTestHelpers.JsonResponse("{\"Exists\":true}"),
            _ => new HttpResponseMessage(HttpStatusCode.NotFound),
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => SharePointContractTestHelpers.JsonResponse("{}", HttpStatusCode.Created),
            _ => new HttpResponseMessage(HttpStatusCode.NotFound),
            _ => SharePointContractTestHelpers.JsonResponse("{}", HttpStatusCode.Created));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFolderClient(client);

        await sut.EnsureFolderPathAsync("/sites/team/Docs/2026/001");

        handler.Requests.Should().NotContain(r => r.Uri.Contains("GetFolderByServerRelativeUrl('%2Fsites')"));
        handler.Requests.Should().NotContain(r => r.Uri.Contains("GetFolderByServerRelativeUrl('%2Fsites%2Fteam')"));
        handler.Requests.Count(r => r.Uri.EndsWith("/_api/contextinfo")).Should().Be(1);
        var createBodies = handler.Requests
            .Where(r => r.Uri.EndsWith("/_api/web/folders"))
            .Select(r => r.Body)
            .ToList();

        createBodies.Should().HaveCount(2);
        createBodies[0].Should().Contain("/sites/team/Docs/2026\"");
        createBodies[1].Should().Contain("/sites/team/Docs/2026/001");
    }

    [Fact]
    public async Task EnsureFolderPathAsync_WhenFolderExists_MakesSingleRequest()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"Exists\":true}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFolderClient(client);

        await sut.EnsureFolderPathAsync("/sites/team/Docs/2026/001");

        handler.Requests.Should().ContainSingle();
    }

    [Fact]
    public async Task GetFilesAsync_ReturnsFilesWithStringEncodedLength()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse(
                "{\"value\":[{\"Name\":\"a.txt\",\"ServerRelativeUrl\":\"/sites/team/Docs/a.txt\",\"Length\":\"1234\",\"TimeCreated\":\"2026-01-02T03:04:05Z\",\"TimeLastModified\":\"2026-02-03T04:05:06Z\"}]}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFolderClient(client);

        var files = await sut.GetFilesAsync("/sites/team/Docs");

        handler.Requests[0].Uri.Should().Contain("GetFolderByServerRelativeUrl('%2Fsites%2Fteam%2FDocs')/Files?$select=");
        files.Should().ContainSingle();
        files[0].Should().Be(new SharePointFileInfo(
            "a.txt",
            "/sites/team/Docs/a.txt",
            1234,
            new DateTimeOffset(2026, 1, 2, 3, 4, 5, TimeSpan.Zero),
            new DateTimeOffset(2026, 2, 3, 4, 5, 6, TimeSpan.Zero)));
    }

    [Fact]
    public async Task GetFoldersAsync_ReturnsSubfolders()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse(
                "{\"value\":[{\"Name\":\"2026\",\"ServerRelativeUrl\":\"/sites/team/Docs/2026\",\"ItemCount\":7}]}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFolderClient(client);

        var folders = await sut.GetFoldersAsync("/sites/team/Docs");

        handler.Requests[0].Uri.Should().Contain("/Folders?$select=");
        folders.Should().ContainSingle().Which.Should().Be(new SharePointFolderInfo("2026", "/sites/team/Docs/2026", 7));
    }

    [Fact]
    public async Task GetFilesAsync_WhenFolderMissing_ThrowsNotFound()
    {
        var handler = new QueueMessageHandler(_ => new HttpResponseMessage(HttpStatusCode.NotFound));
        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFolderClient(client);

        var act = () => sut.GetFilesAsync("/sites/team/Missing");

        await act.Should().ThrowAsync<SharePointNotFoundException>();
    }
}

internal static class SharePointContractTestHelpers
{
    public static SharePointFileClient CreateFileClient(HttpClient client)
        => new(client, new SharePointRequestExecutor(new SharePointFormDigestProvider(new SharePointOnPremOptions
        {
            SiteBaseUrl = "https://sharepoint.contoso.local/sites/team",
            UseFormDigestCaching = true
        })));

    public static SharePointFolderClient CreateFolderClient(HttpClient client)
        => new(client, new SharePointRequestExecutor(new SharePointFormDigestProvider(new SharePointOnPremOptions
        {
            SiteBaseUrl = "https://sharepoint.contoso.local/sites/team",
            UseFormDigestCaching = true
        })));

    public static SharePointSecurityClient CreateSecurityClient(HttpClient client)
        => new(
            client,
            new DefaultLoginNameNormalizer(new SharePointIdentityOptions
            {
                Domain = "CONTOSO",
                ClaimsPrefix = "i:0#.w|"
            }),
            new SharePointRequestExecutor(new SharePointFormDigestProvider(new SharePointOnPremOptions
            {
                SiteBaseUrl = "https://sharepoint.contoso.local/sites/team",
                UseFormDigestCaching = true
            })));

    public static HttpClient CreateHttpClient(HttpMessageHandler handler)
        => new(handler)
        {
            BaseAddress = new Uri("https://sharepoint.contoso.local/sites/team/")
        };

    public static HttpResponseMessage JsonResponse(string json, HttpStatusCode statusCode = HttpStatusCode.OK)
    {
        return new HttpResponseMessage(statusCode)
        {
            Content = new StringContent(json, Encoding.UTF8, "application/json")
        };
    }
}

internal sealed class QueueMessageHandler(params Func<HttpRequestMessage, HttpResponseMessage>[] responders) : HttpMessageHandler
{
    private readonly Queue<Func<HttpRequestMessage, HttpResponseMessage>> _responders = new(responders);

    public List<RecordedRequest> Requests { get; } = [];

    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
        Requests.Add(await RecordedRequest.CreateAsync(request, cancellationToken));

        if (_responders.Count == 0)
            throw new InvalidOperationException("No queued response available for the request.");

        var response = _responders.Dequeue().Invoke(request);
        response.RequestMessage = request;
        return response;
    }
}

internal sealed record RecordedRequest(
    HttpMethod Method,
    string Uri,
    IReadOnlyDictionary<string, string[]> Headers,
    IReadOnlyDictionary<string, string[]> ContentHeaders,
    string? Body)
{
    public static async Task<RecordedRequest> CreateAsync(HttpRequestMessage request, CancellationToken ct)
    {
        var headers = request.Headers.ToDictionary(h => h.Key, h => h.Value.ToArray(), StringComparer.OrdinalIgnoreCase);
        var contentHeaders = request.Content?.Headers.ToDictionary(h => h.Key, h => h.Value.ToArray(), StringComparer.OrdinalIgnoreCase)
            ?? new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
        var body = request.Content is null ? null : await request.Content.ReadAsStringAsync(ct);

        return new RecordedRequest(
            request.Method,
            request.RequestUri?.OriginalString ?? string.Empty,
            headers,
            contentHeaders,
            body);
    }
}


