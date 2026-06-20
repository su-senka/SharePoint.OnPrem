using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using FluentAssertions;
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
            SiteBaseUrl = "https://sharepoint.local/sites/pp",
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
            SiteBaseUrl = "https://sharepoint.local/sites/pp",
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
        var digestProvider = new SharePointFormDigestProvider(new SharePointOnPremOptions { SiteBaseUrl = "https://sharepoint.local/sites/pp" });
        var sut = new SharePointRequestExecutor(digestProvider);

        var act = async () =>
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, "_api/web/test");
            using var _ = await sut.SendAsync(client, request);
        };

        await act.Should().ThrowAsync<SharePointNotFoundException>()
            .WithMessage("*File missing*");
    }
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
            "/sites/pp/Attachments",
            "test file.txt",
            content,
            "text/plain"));

        result.ServerRelativeUrl.Should().Be("/sites/pp/Attachments/test file.txt");
        result.FileName.Should().Be("test file.txt");

        handler.Requests.Should().HaveCount(2);
        handler.Requests[1].Method.Should().Be(HttpMethod.Post);
        handler.Requests[1].Uri.Should().Contain("GetFolderByServerRelativeUrl('%2Fsites%2Fpp%2FAttachments')/Files/add(url='test%20file.txt',overwrite=true)");
        handler.Requests[1].Headers["X-RequestDigest"].Should().ContainSingle().Which.Should().Be("digest-1");
        handler.Requests[1].ContentHeaders["Content-Type"].Should().ContainSingle().Which.Should().StartWith("text/plain");
    }

    [Fact]
    public async Task UploadAsync_WhenServerReturnsStoredFileData_UsesResponseValues()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => SharePointContractTestHelpers.JsonResponse("{\"ServerRelativeUrl\":\"/sites/pp/Attachments/server-name.txt\",\"Name\":\"server-name.txt\"}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        await using var content = new MemoryStream(Encoding.UTF8.GetBytes("hello"));
        var result = await sut.UploadAsync(new UploadFileRequest(
            "/sites/pp/Attachments",
            "client-name.txt",
            content));

        result.ServerRelativeUrl.Should().Be("/sites/pp/Attachments/server-name.txt");
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

        await using var result = await sut.DownloadAsync("/sites/pp/Attachments/test.txt");

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
            _ => SharePointContractTestHelpers.JsonResponse("{\"ServerRelativeUrl\":\"/sites/pp/Attachments/test.txt\"}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        var result = await sut.ExistsAsync("/sites/pp/Attachments/test.txt");

        result.Should().BeTrue();
        handler.Requests.Should().ContainSingle();
        handler.Requests[0].Uri.Should().Contain("GetFileByServerRelativeUrl('%2Fsites%2Fpp%2FAttachments%2Ftest.txt')");
    }

    [Fact]
    public async Task ExistsAsync_WhenFileIsMissing_ReturnsFalse()
    {
        var handler = new QueueMessageHandler(_ => new HttpResponseMessage(HttpStatusCode.NotFound));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        var result = await sut.ExistsAsync("/sites/pp/Attachments/missing.txt");

        result.Should().BeFalse();
    }

    [Fact]
    public async Task GetWebUrlAsync_WhenLinkingUrlMissing_FallsBackToAbsoluteServerRelativeUrl()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"ServerRelativeUrl\":\"/sites/pp/Attachments/test.docx\"}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        var result = await sut.GetFileWebUrlAsync("/sites/pp/Attachments/test.docx");

        result.Should().Be("https://sharepoint.local/sites/pp/Attachments/test.docx");
    }

    [Fact]
    public async Task DeleteAsync_SendsDeleteSemanticsViaPostWithDigest()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => SharePointContractTestHelpers.JsonResponse("{}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        await sut.DeleteAsync("/sites/pp/Attachments/test.txt");

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
            "/sites/pp/Attachments/source.txt",
            "/sites/pp/Archive",
            "copy.txt",
            overwrite: true);

        result.ServerRelativeUrl.Should().Be("/sites/pp/Archive/copy.txt");
        result.FileName.Should().Be("copy.txt");
        handler.Requests.Should().HaveCount(2);
        handler.Requests[1].Uri.Should().Contain("copyto(strnewurl='%2Fsites%2Fpp%2FArchive%2Fcopy.txt',boverwrite=true)");
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
            "/sites/pp/Attachments/source.txt",
            "/sites/pp/Archive",
            "moved.txt",
            overwrite: false);

        result.ServerRelativeUrl.Should().Be("/sites/pp/Archive/moved.txt");
        result.FileName.Should().Be("moved.txt");
        handler.Requests.Should().HaveCount(2);
        handler.Requests[1].Uri.Should().Contain("moveto(newurl='%2Fsites%2Fpp%2FArchive%2Fmoved.txt',flags=0)");
    }

    [Fact]
    public async Task RenameAsync_UsesMoveToWithinSameFolderAndReturnsRenamedFile()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => SharePointContractTestHelpers.JsonResponse("{}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        var result = await sut.RenameAsync("/sites/pp/Attachments/source.txt", "renamed.txt");

        result.ServerRelativeUrl.Should().Be("/sites/pp/Attachments/renamed.txt");
        result.FileName.Should().Be("renamed.txt");
        handler.Requests.Should().HaveCount(2);
        handler.Requests[1].Uri.Should().Contain("GetFileByServerRelativeUrl('%2Fsites%2Fpp%2FAttachments%2Fsource.txt')/moveto(newurl='%2Fsites%2Fpp%2FAttachments%2Frenamed.txt',flags=1)");
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
            "/sites/pp/Attachments/test.txt",
            new Dictionary<string, object?>
            {
                ["Title"] = "Quarterly Report",
                ["Category"] = "Finance"
            });

        handler.Requests.Should().HaveCount(2);
        handler.Requests[1].Method.Should().Be(HttpMethod.Post);
        handler.Requests[1].Uri.Should().Contain("GetFileByServerRelativeUrl('%2Fsites%2Fpp%2FAttachments%2Ftest.txt')/ListItemAllFields");
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
            "/sites/pp/Attachments/test.txt",
            ["Title", "Category"]);

        handler.Requests.Should().HaveCount(1);
        handler.Requests[0].Uri.Should().Contain("GetFileByServerRelativeUrl('%2Fsites%2Fpp%2FAttachments%2Ftest.txt')/ListItemAllFields?$select=Title,Category");
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

        var result = await sut.TryGetMetadataAsync("/sites/pp/Attachments/test.txt", ["Title"]);

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

        var result = await sut.TryGetMetadataAsync("/sites/pp/Attachments/missing.txt", ["Title"]);

        result.Should().BeNull();
    }

    [Fact]
    public async Task GetMetadataValueAsync_WhenFieldExists_ReturnsTypedValue()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"Title\":\"Quarterly Report\"}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        var result = await sut.GetMetadataValueAsync<string>("/sites/pp/Attachments/test.txt", "Title");

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

        var result = await sut.GetMetadataValueAsync<string>("/sites/pp/Attachments/test.txt", "Title");

        result.Should().BeNull();
    }

    [Fact]
    public async Task TryGetMetadataValueAsync_WhenFileExists_ReturnsTypedValue()
    {
        var handler = new QueueMessageHandler(
            _ => SharePointContractTestHelpers.JsonResponse("{\"Title\":\"Quarterly Report\"}"));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        var result = await sut.TryGetMetadataValueAsync<string>("/sites/pp/Attachments/test.txt", "Title");

        result.Should().Be("Quarterly Report");
    }

    [Fact]
    public async Task TryGetMetadataValueAsync_WhenFileMissing_ReturnsDefault()
    {
        var handler = new QueueMessageHandler(_ => new HttpResponseMessage(HttpStatusCode.NotFound));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFileClient(client);

        var result = await sut.TryGetMetadataValueAsync<string>("/sites/pp/Attachments/missing.txt", "Title");

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

        var result = await sut.ExistsAsync("/sites/pp/Attachments/2026");

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

        var act = async () => await sut.CreateAsync("/sites/pp/Attachments/2026");

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

        await sut.DeleteAsync("/sites/pp/Attachments/2026");

        handler.Requests.Should().HaveCount(2);
        handler.Requests[1].Method.Should().Be(HttpMethod.Post);
        handler.Requests[1].Uri.Should().Contain("GetFolderByServerRelativeUrl('%2Fsites%2Fpp%2FAttachments%2F2026')");
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

        var act = async () => await sut.DeleteAsync("/sites/pp/Attachments/missing");

        await act.Should().NotThrowAsync();
    }

    [Fact]
    public async Task EnsurePathAsync_CreatesEachMissingSegmentOnce()
    {
        var handler = new QueueMessageHandler(
            _ => new HttpResponseMessage(HttpStatusCode.NotFound),
            _ => SharePointContractTestHelpers.JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => SharePointContractTestHelpers.JsonResponse("{}", HttpStatusCode.Created),
            _ => new HttpResponseMessage(HttpStatusCode.NotFound),
            _ => SharePointContractTestHelpers.JsonResponse("{}", HttpStatusCode.Created),
            _ => new HttpResponseMessage(HttpStatusCode.NotFound),
            _ => SharePointContractTestHelpers.JsonResponse("{}", HttpStatusCode.Created));

        var client = SharePointContractTestHelpers.CreateHttpClient(handler);
        var sut = SharePointContractTestHelpers.CreateFolderClient(client);

        await sut.EnsureFolderPathAsync("/Docs/2026/001");

        handler.Requests.Count(r => r.Uri.EndsWith("/_api/contextinfo")).Should().Be(1);
        var createBodies = handler.Requests
            .Where(r => r.Uri.EndsWith("/_api/web/folders"))
            .Select(r => r.Body)
            .ToList();

        createBodies.Should().HaveCount(3);
        createBodies[0].Should().Contain("/Docs");
        createBodies[1].Should().Contain("/Docs/2026");
        createBodies[2].Should().Contain("/Docs/2026/001");
    }
}

internal static class SharePointContractTestHelpers
{
    public static SharePointFileClient CreateFileClient(HttpClient client)
        => new(client, new SharePointRequestExecutor(new SharePointFormDigestProvider(new SharePointOnPremOptions
        {
            SiteBaseUrl = "https://sharepoint.local/sites/pp",
            UseFormDigestCaching = true
        })));

    public static SharePointFolderClient CreateFolderClient(HttpClient client)
        => new(client, new SharePointRequestExecutor(new SharePointFormDigestProvider(new SharePointOnPremOptions
        {
            SiteBaseUrl = "https://sharepoint.local/sites/pp",
            UseFormDigestCaching = true
        })));

    public static SharePointSecurityClient CreateSecurityClient(HttpClient client)
        => new(
            client,
            new DefaultLoginNameNormalizer(new SharePointIdentityOptions
            {
                Domain = "ACR",
                ClaimsPrefix = "i:0#.w|"
            }),
            new SharePointRequestExecutor(new SharePointFormDigestProvider(new SharePointOnPremOptions
            {
                SiteBaseUrl = "https://sharepoint.local/sites/pp",
                UseFormDigestCaching = true
            })));

    public static HttpClient CreateHttpClient(HttpMessageHandler handler)
        => new(handler)
        {
            BaseAddress = new Uri("https://sharepoint.local/sites/pp/")
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


