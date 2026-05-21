using System.Net;
using System.Text;
using FluentAssertions;
using Microsoft.Extensions.DependencyInjection;
using PP.Application.Services.SharePoint;
using SharePoint.OnPrem.Abstractions;
using SharePoint.OnPrem.Core;
using SharePoint.OnPrem.DependencyInjection;
using SharePoint.OnPrem.PlanProvozuAdapter;

namespace SharePoint.OnPrem.PlanProvozuAdapter.Tests;

public class PlanProvozuSharePointServiceAdapterTests
{
    [Fact]
    public async Task UploadFileAsync_UsesScopedFolderAndReturnsStoredPath()
    {
        var fileClient = new RecordingFileClient
        {
            UploadResult = new SharePointStoredFile("/sites/pp/Attachments/2026/doc.txt", "doc.txt")
        };

        var sut = CreateAdapter(fileClient: fileClient);

        await using var stream = new MemoryStream(Encoding.UTF8.GetBytes("demo"));
        var result = await sut.UploadFileAsync("2026", "doc.txt", stream);

        result.Should().Be("/sites/pp/Attachments/2026/doc.txt");
        fileClient.UploadRequests.Should().ContainSingle();
        fileClient.UploadRequests[0].FolderServerRelativeUrl.Should().Be("/sites/pp/Attachments/2026");
        fileClient.UploadRequests[0].FileName.Should().Be("doc.txt");
    }

    [Fact]
    public async Task DeleteFileAsync_UsesScopedLeafPath()
    {
        var fileClient = new RecordingFileClient();
        var sut = CreateAdapter(fileClient: fileClient);

        await sut.DeleteFileAsync("legacy.txt");

        fileClient.DeleteRequests.Should().ContainSingle().Which.Should().Be("/sites/pp/Attachments/legacy.txt");
    }

    [Fact]
    public async Task EnsureRootSecurityAsync_ComposesExpectedSecurityCalls()
    {
        var securityClient = new RecordingSecurityClient();
        var sut = CreateAdapter(securityClient: securityClient);

        await sut.EnsureRootSecurityAsync("/sites/pp/Attachments/3255/26/001", "Readers", "Owners");

        securityClient.RecordedCalls.Should().Equal(
            "EnsureGroup:Readers:Readers – read access",
            "EnsureGroup:Owners:Owners – edit access",
            "BreakInheritance:/sites/pp/Attachments/3255/26/001:False",
            "BindRole:/sites/pp/Attachments/3255/26/001:Readers:Read",
            "BindRole:/sites/pp/Attachments/3255/26/001:Owners:Edit");
    }

    [Fact]
    public async Task EnsureAssignmentSecurityAsync_ComposesExpectedSecurityCalls()
    {
        var securityClient = new RecordingSecurityClient();
        var sut = CreateAdapter(securityClient: securityClient);

        await sut.EnsureAssignmentSecurityAsync("/sites/pp/Attachments/3255/26/001/003", "Writers");

        securityClient.RecordedCalls.Should().Equal(
            "EnsureGroup:Writers:Writers – assignment writers",
            "BreakInheritance:/sites/pp/Attachments/3255/26/001/003:True",
            "BindRole:/sites/pp/Attachments/3255/26/001/003:Writers:Edit");
    }

    [Fact]
    public async Task GrantFileReadAsync_UsesLegacyFilePermissionWorkflow()
    {
        var securityClient = new RecordingSecurityClient { EnsureUserResult = 321 };
        var handler = new QueueMessageHandler(
            _ => JsonResponse("{\"FormDigestValue\":\"digest-1\",\"FormDigestTimeoutSeconds\":1800}"),
            _ => new HttpResponseMessage(HttpStatusCode.OK),
            _ => new HttpResponseMessage(HttpStatusCode.OK));

        var httpClient = CreateHttpClient(handler);
        var sut = CreateAdapter(httpClient: httpClient, securityClient: securityClient);

        await sut.GrantFileReadAsync("legacy.txt", "jnovak");

        securityClient.RecordedCalls.Should().ContainSingle().Which.Should().Be("EnsureUser:jnovak");
        handler.Requests.Should().HaveCount(3);
        handler.Requests[1].Uri.Should().Contain("GetFileByServerRelativeUrl('%2Fsites%2Fpp%2FAttachments%2Flegacy.txt')/ListItemAllFields/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)");
        handler.Requests[2].Uri.Should().Contain("addroleassignment(principalid=321,roledefid=1073741826)");
    }

    [Fact]
    public void AddPlanProvozuSharePointCompatibilityAdapter_RegistersLegacyService()
    {
        var services = new ServiceCollection();
        services.AddSharePointOnPrem(
            configureCore: options =>
            {
                options.SiteBaseUrl = "https://sharepoint.local/sites/pp";
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

        services.AddPlanProvozuSharePointCompatibilityAdapter();

        using var provider = services.BuildServiceProvider();

        provider.GetRequiredService<ISharePointService>().Should().BeOfType<PlanProvozuSharePointServiceAdapter>();
    }

    private static PlanProvozuSharePointServiceAdapter CreateAdapter(
        HttpClient? httpClient = null,
        RecordingFileClient? fileClient = null,
        RecordingFolderClient? folderClient = null,
        RecordingSecurityClient? securityClient = null,
        ISharePointPathScope? pathScope = null)
    {
        var client = httpClient ?? CreateHttpClient(new QueueMessageHandler());
        var executor = new SharePointRequestExecutor(new SharePointFormDigestProvider(new SharePointOnPremOptions
        {
            SiteBaseUrl = "https://sharepoint.local/sites/pp",
            UseFormDigestCaching = true
        }));

        return new PlanProvozuSharePointServiceAdapter(
            client,
            fileClient ?? new RecordingFileClient(),
            folderClient ?? new RecordingFolderClient(),
            securityClient ?? new RecordingSecurityClient(),
            pathScope ?? new StaticPathScope("/sites/pp/Attachments"),
            executor);
    }

    private static HttpClient CreateHttpClient(HttpMessageHandler handler)
        => new(handler)
        {
            BaseAddress = new Uri("https://sharepoint.local/sites/pp/")
        };

    private static HttpResponseMessage JsonResponse(string json, HttpStatusCode statusCode = HttpStatusCode.OK)
        => new(statusCode)
        {
            Content = new StringContent(json, Encoding.UTF8, "application/json")
        };
}

internal sealed class RecordingFileClient : ISharePointFileClient
{
    public List<(string FolderServerRelativeUrl, string FileName)> UploadRequests { get; } = [];
    public List<string> DeleteRequests { get; } = [];
    public List<string> DownloadRequests { get; } = [];
    public List<string> WebUrlRequests { get; } = [];

    public SharePointStoredFile UploadResult { get; set; } = new("/sites/pp/Attachments/file.txt", "file.txt");

    public Task<SharePointStoredFile> UploadAsync(UploadFileRequest request, CancellationToken ct = default)
    {
        UploadRequests.Add((request.FolderServerRelativeUrl, request.FileName));
        return Task.FromResult(UploadResult);
    }

    public Task<SharePointFileDownload> DownloadAsync(string serverRelativeUrl, CancellationToken ct = default)
    {
        DownloadRequests.Add(serverRelativeUrl);
        return Task.FromResult(new SharePointFileDownload([], "application/octet-stream", "file.txt"));
    }

    public Task DeleteAsync(string serverRelativeUrl, CancellationToken ct = default)
    {
        DeleteRequests.Add(serverRelativeUrl);
        return Task.CompletedTask;
    }

    public Task<string> GetWebUrlAsync(string serverRelativeUrl, CancellationToken ct = default)
    {
        WebUrlRequests.Add(serverRelativeUrl);
        return Task.FromResult("https://sharepoint.local/web-url");
    }
}

internal sealed class RecordingFolderClient : ISharePointFolderClient
{
    public List<string> EnsuredPaths { get; } = [];

    public Task EnsurePathAsync(string serverRelativeFolder, CancellationToken ct = default)
    {
        EnsuredPaths.Add(serverRelativeFolder);
        return Task.CompletedTask;
    }

    public Task<bool> ExistsAsync(string serverRelativeFolder, CancellationToken ct = default)
        => Task.FromResult(false);

    public Task CreateAsync(string serverRelativeFolder, CancellationToken ct = default)
        => Task.CompletedTask;
}

internal sealed class RecordingSecurityClient : ISharePointSecurityClient
{
    public int EnsureUserResult { get; set; } = 123;
    public int EnsureGroupResult { get; set; } = 456;
    public List<string> RecordedCalls { get; } = [];

    public Task BreakInheritanceAsync(string serverRelativeFolder, bool copyRoleAssignments, CancellationToken ct = default)
    {
        RecordedCalls.Add($"BreakInheritance:{serverRelativeFolder}:{copyRoleAssignments}");
        return Task.CompletedTask;
    }

    public Task ResetInheritanceAsync(string serverRelativeFolder, CancellationToken ct = default)
    {
        RecordedCalls.Add($"ResetInheritance:{serverRelativeFolder}");
        return Task.CompletedTask;
    }

    public Task<int> EnsureGroupAsync(string groupName, string? description = null, CancellationToken ct = default)
    {
        RecordedCalls.Add($"EnsureGroup:{groupName}:{description}");
        return Task.FromResult(EnsureGroupResult);
    }

    public Task DeleteGroupAsync(string groupName, CancellationToken ct = default)
    {
        RecordedCalls.Add($"DeleteGroup:{groupName}");
        return Task.CompletedTask;
    }

    public Task<int> EnsureUserAsync(string loginName, CancellationToken ct = default)
    {
        RecordedCalls.Add($"EnsureUser:{loginName}");
        return Task.FromResult(EnsureUserResult);
    }

    public Task<IReadOnlyList<string>> GetGroupMembersAsync(string groupName, CancellationToken ct = default)
    {
        RecordedCalls.Add($"GetGroupMembers:{groupName}");
        return Task.FromResult<IReadOnlyList<string>>([]);
    }

    public Task AddUsersToGroupAsync(string groupName, IEnumerable<string> loginNames, CancellationToken ct = default)
    {
        RecordedCalls.Add($"AddUsers:{groupName}:{string.Join(',', loginNames)}");
        return Task.CompletedTask;
    }

    public Task RemoveUsersFromGroupAsync(string groupName, IEnumerable<string> loginNames, CancellationToken ct = default)
    {
        RecordedCalls.Add($"RemoveUsers:{groupName}:{string.Join(',', loginNames)}");
        return Task.CompletedTask;
    }

    public Task BindRoleToFolderAsync(string serverRelativeFolder, string principalName, string roleName, CancellationToken ct = default)
    {
        RecordedCalls.Add($"BindRole:{serverRelativeFolder}:{principalName}:{roleName}");
        return Task.CompletedTask;
    }

    public Task RemoveRoleFromFolderAsync(string serverRelativeFolder, string principalName, CancellationToken ct = default)
    {
        RecordedCalls.Add($"RemoveRole:{serverRelativeFolder}:{principalName}");
        return Task.CompletedTask;
    }

    public Task SyncGroupMembershipAsync(string groupName, IEnumerable<string> desiredLoginNames, CancellationToken ct = default)
    {
        RecordedCalls.Add($"SyncMembers:{groupName}:{string.Join(',', desiredLoginNames)}");
        return Task.CompletedTask;
    }
}

internal sealed class StaticPathScope(string baseFolder) : ISharePointPathScope
{
    private readonly string _baseFolder = baseFolder.TrimEnd('/');

    public string ToServerRelativePath(string relativePath)
    {
        var trimmed = relativePath.Trim('/');
        return $"{_baseFolder}/{trimmed}";
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
    string? Body)
{
    public static async Task<RecordedRequest> CreateAsync(HttpRequestMessage request, CancellationToken ct)
    {
        var headers = request.Headers.ToDictionary(h => h.Key, h => h.Value.ToArray(), StringComparer.OrdinalIgnoreCase);
        var body = request.Content is null ? null : await request.Content.ReadAsStringAsync(ct);

        return new RecordedRequest(
            request.Method,
            request.RequestUri?.OriginalString ?? string.Empty,
            headers,
            body);
    }
}
