using SharePoint.OnPrem.Abstractions;
using SharePoint.OnPrem.Core;
using System.Net;
using System.Net.Http.Headers;
using System.Text.Json;

namespace SharePoint.OnPrem.Files;

public sealed class SharePointFileClient(HttpClient httpClient, ISharePointRequestExecutor requestExecutor) : ISharePointFileClient
{
    private readonly HttpClient _httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
    private readonly ISharePointRequestExecutor _requestExecutor = requestExecutor ?? throw new ArgumentNullException(nameof(requestExecutor));

    public async Task<SharePointStoredFile> UploadAsync(UploadFileRequest request, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(request);

        var folderServerRelativeUrl = ServerRelativeUrl.Validate(request.FolderServerRelativeUrl);
        var fileName = ValidateLeafFileName(request.FileName);
        var encodedFolder = Uri.EscapeDataString(folderServerRelativeUrl);
        var encodedFileName = Uri.EscapeDataString(fileName);

        using var message = new HttpRequestMessage(
            HttpMethod.Post,
            $"_api/web/GetFolderByServerRelativeUrl('{encodedFolder}')/Files/add(url='{encodedFileName}',overwrite={request.Overwrite.ToString().ToLowerInvariant()})");

        var streamContent = new StreamContent(request.Content);
        if (!string.IsNullOrWhiteSpace(request.ContentType))
            streamContent.Headers.ContentType = MediaTypeHeaderValue.Parse(request.ContentType);

        message.Content = streamContent;

        using var response = await _requestExecutor.SendAsync(
            _httpClient,
            message,
            new SharePointSendOptions { IncludeFormDigest = true },
            ct);

        return new SharePointStoredFile($"{folderServerRelativeUrl.TrimEnd('/')}/{fileName}", fileName);
    }

    public async Task<SharePointFileDownload> DownloadAsync(string serverRelativeUrl, CancellationToken ct = default)
    {
        var validatedUrl = ServerRelativeUrl.Validate(serverRelativeUrl);
        var encodedUrl = Uri.EscapeDataString(validatedUrl);

        using var message = new HttpRequestMessage(
            HttpMethod.Get,
            $"_api/web/GetFileByServerRelativeUrl('{encodedUrl}')/$value");

        using var response = await _requestExecutor.SendAsync(_httpClient, message, ct: ct);
        var bytes = await response.Content.ReadAsByteArrayAsync(ct);
        var contentType = response.Content.Headers.ContentType?.MediaType ?? "application/octet-stream";

        return new SharePointFileDownload(bytes, contentType, Path.GetFileName(validatedUrl));
    }

    public async Task DeleteAsync(string serverRelativeUrl, CancellationToken ct = default)
    {
        var validatedUrl = ServerRelativeUrl.Validate(serverRelativeUrl);
        var encodedUrl = Uri.EscapeDataString(validatedUrl);

        using var message = new HttpRequestMessage(
            HttpMethod.Post,
            $"_api/web/GetFileByServerRelativeUrl('{encodedUrl}')");

        message.Headers.Add("IF-MATCH", "*");
        message.Headers.Add("X-HTTP-Method", "DELETE");

        using var response = await _requestExecutor.SendAsync(
            _httpClient,
            message,
            new SharePointSendOptions { IncludeFormDigest = true },
            ct);
    }

    public async Task<string> GetWebUrlAsync(string serverRelativeUrl, CancellationToken ct = default)
    {
        var validatedUrl = ServerRelativeUrl.Validate(serverRelativeUrl);
        var encodedUrl = Uri.EscapeDataString(validatedUrl);

        using var message = new HttpRequestMessage(
            HttpMethod.Get,
            $"_api/web/GetFileByServerRelativeUrl('{encodedUrl}')?$select=LinkingUrl,ServerRelativeUrl");

        using var response = await _requestExecutor.SendAsync(_httpClient, message, ct: ct);
        var payload = await response.Content.ReadAsStringAsync(ct);
        using var document = JsonDocument.Parse(payload);

        if (document.RootElement.TryGetProperty("LinkingUrl", out var linkingUrlElement))
        {
            var linkingUrl = linkingUrlElement.GetString();
            if (!string.IsNullOrWhiteSpace(linkingUrl))
                return linkingUrl;
        }

        if (document.RootElement.TryGetProperty("ServerRelativeUrl", out var serverRelativeElement))
        {
            var returnedUrl = serverRelativeElement.GetString();
            if (!string.IsNullOrWhiteSpace(returnedUrl))
                return new Uri(_httpClient.BaseAddress!, returnedUrl).ToString();
        }

        return new Uri(_httpClient.BaseAddress!, validatedUrl).ToString();
    }

    private static string ValidateLeafFileName(string fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName))
            throw new SharePointValidationException("File name must be provided.");

        var trimmed = fileName.Trim();
        if (trimmed.Contains('/') || trimmed.Contains('\\'))
            throw new SharePointValidationException("File name must be a leaf name without path separators.");

        return trimmed;
    }
}

public sealed class SharePointFolderClient(HttpClient httpClient, ISharePointRequestExecutor requestExecutor) : ISharePointFolderClient
{
    private readonly HttpClient _httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
    private readonly ISharePointRequestExecutor _requestExecutor = requestExecutor ?? throw new ArgumentNullException(nameof(requestExecutor));

    public async Task EnsurePathAsync(string serverRelativeFolder, CancellationToken ct = default)
    {
        var validatedFolder = ServerRelativeUrl.Validate(serverRelativeFolder);
        var segments = validatedFolder.Split('/', StringSplitOptions.RemoveEmptyEntries);
        var current = string.Empty;

        foreach (var segment in segments)
        {
            current += "/" + segment;
            if (await ExistsAsync(current, ct))
                continue;

            await CreateAsync(current, ct);
        }
    }

    public async Task<bool> ExistsAsync(string serverRelativeFolder, CancellationToken ct = default)
    {
        var validatedFolder = ServerRelativeUrl.Validate(serverRelativeFolder);
        var encodedFolder = Uri.EscapeDataString(validatedFolder);

        using var message = new HttpRequestMessage(
            HttpMethod.Get,
            $"_api/web/GetFolderByServerRelativeUrl('{encodedFolder}')?$select=Exists");

        using var response = await _requestExecutor.SendAsync(
            _httpClient,
            message,
            new SharePointSendOptions { EnsureSuccessStatusCode = false },
            ct);

        if (response.StatusCode == HttpStatusCode.NotFound || response.StatusCode == HttpStatusCode.InternalServerError)
            return false;

        await _requestExecutor.EnsureSuccessAsync(response, ct);

        var payload = await response.Content.ReadAsStringAsync(ct);
        using var document = JsonDocument.Parse(payload);
        return document.RootElement.TryGetProperty("Exists", out var existsElement)
               && existsElement.ValueKind == JsonValueKind.True;
    }

    public async Task CreateAsync(string serverRelativeFolder, CancellationToken ct = default)
    {
        var validatedFolder = ServerRelativeUrl.Validate(serverRelativeFolder);

        using var message = new HttpRequestMessage(HttpMethod.Post, "_api/web/folders")
        {
            Content = SharePointContentFactory.CreateJsonContent(new
            {
                ServerRelativeUrl = validatedFolder
            })
        };

        using var response = await _requestExecutor.SendAsync(
            _httpClient,
            message,
            new SharePointSendOptions
            {
                IncludeFormDigest = true,
                EnsureSuccessStatusCode = false
            },
            ct);

        if (response.StatusCode == HttpStatusCode.Conflict)
            return;

        await _requestExecutor.EnsureSuccessAsync(response, ct);
    }
}

