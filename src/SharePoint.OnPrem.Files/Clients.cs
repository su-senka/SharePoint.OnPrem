using SharePoint.OnPrem.Abstractions;
using SharePoint.OnPrem.Core;
using System.Net;
using System.Net.Http.Headers;
using System.Text.Json;

namespace SharePoint.OnPrem.Files;

internal sealed class SharePointFileClient(HttpClient httpClient, ISharePointRequestExecutor requestExecutor) : ISharePointFileClient
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

        // Read server-confirmed name and URL; fall back to constructed values if absent.
        string? responseServerRelativeUrl = null;
        string? responseFileName = null;
        try
        {
            var payload = await response.Content.ReadAsStringAsync(ct);
            if (!string.IsNullOrWhiteSpace(payload))
            {
                using var document = JsonDocument.Parse(payload);
                if (document.RootElement.TryGetProperty("ServerRelativeUrl", out var sruEl))
                    responseServerRelativeUrl = sruEl.GetString();
                if (document.RootElement.TryGetProperty("Name", out var nameEl))
                    responseFileName = nameEl.GetString();
            }
        }
        catch (JsonException) { }

        var storedUrl = !string.IsNullOrWhiteSpace(responseServerRelativeUrl)
            ? responseServerRelativeUrl
            : $"{folderServerRelativeUrl.TrimEnd('/')}/{fileName}";
        var storedName = !string.IsNullOrWhiteSpace(responseFileName)
            ? responseFileName
            : fileName;

        return new SharePointStoredFile(storedUrl, storedName);
    }

    public async Task<SharePointFileDownload> DownloadAsync(string serverRelativeUrl, CancellationToken ct = default)
    {
        var validatedUrl = ServerRelativeUrl.Validate(serverRelativeUrl);
        var encodedUrl = Uri.EscapeDataString(validatedUrl);

        using var message = new HttpRequestMessage(
            HttpMethod.Get,
            $"_api/web/GetFileByServerRelativeUrl('{encodedUrl}')/$value");

        // Do NOT 'using' the response here — caller disposes SharePointFileDownload which
        // disposes the HttpResponseStream wrapper, which in turn disposes the response.
        var response = await _requestExecutor.SendAsync(
            _httpClient,
            message,
            new SharePointSendOptions { UseResponseHeadersRead = true },
            ct);

        var contentType = response.Content.Headers.ContentType?.MediaType ?? "application/octet-stream";
        var stream = new HttpResponseStream(response, await response.Content.ReadAsStreamAsync(ct));
        return new SharePointFileDownload(stream, contentType, Path.GetFileName(validatedUrl));
    }

    public async Task<bool> ExistsAsync(string serverRelativeUrl, CancellationToken ct = default)
    {
        var validatedUrl = ServerRelativeUrl.Validate(serverRelativeUrl);
        var encodedUrl = Uri.EscapeDataString(validatedUrl);

        using var message = new HttpRequestMessage(
            HttpMethod.Get,
            $"_api/web/GetFileByServerRelativeUrl('{encodedUrl}')?$select=ServerRelativeUrl");

        using var response = await _requestExecutor.SendAsync(
            _httpClient,
            message,
            new SharePointSendOptions { EnsureSuccessStatusCode = false },
            ct);

        if (response.StatusCode == HttpStatusCode.NotFound)
            return false;

        await _requestExecutor.EnsureSuccessAsync(response, ct);
        return true;
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

    public async Task<string> GetFileWebUrlAsync(string serverRelativeUrl, CancellationToken ct = default)
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

    public async Task<SharePointStoredFile> CopyAsync(
        string sourceServerRelativeUrl,
        string destinationFolderServerRelativeUrl,
        string destinationFileName,
        bool overwrite = true,
        CancellationToken ct = default)
    {
        var sourceUrl = ServerRelativeUrl.Validate(sourceServerRelativeUrl);
        var destinationUrl = BuildDestinationFileUrl(destinationFolderServerRelativeUrl, destinationFileName);
        var encodedSource = Uri.EscapeDataString(sourceUrl);
        var encodedDestination = Uri.EscapeDataString(destinationUrl);

        using var message = new HttpRequestMessage(
            HttpMethod.Post,
            $"_api/web/GetFileByServerRelativeUrl('{encodedSource}')/copyto(strnewurl='{encodedDestination}',boverwrite={overwrite.ToString().ToLowerInvariant()})");

        using var response = await _requestExecutor.SendAsync(
            _httpClient,
            message,
            new SharePointSendOptions { IncludeFormDigest = true },
            ct);

        return new SharePointStoredFile(destinationUrl, Path.GetFileName(destinationUrl));
    }

    public async Task<SharePointStoredFile> MoveAsync(
        string sourceServerRelativeUrl,
        string destinationFolderServerRelativeUrl,
        string destinationFileName,
        bool overwrite = true,
        CancellationToken ct = default)
    {
        var sourceUrl = ServerRelativeUrl.Validate(sourceServerRelativeUrl);
        var destinationUrl = BuildDestinationFileUrl(destinationFolderServerRelativeUrl, destinationFileName);
        var encodedSource = Uri.EscapeDataString(sourceUrl);
        var encodedDestination = Uri.EscapeDataString(destinationUrl);
        var flags = overwrite ? 1 : 0;

        using var message = new HttpRequestMessage(
            HttpMethod.Post,
            $"_api/web/GetFileByServerRelativeUrl('{encodedSource}')/moveto(newurl='{encodedDestination}',flags={flags})");

        using var response = await _requestExecutor.SendAsync(
            _httpClient,
            message,
            new SharePointSendOptions { IncludeFormDigest = true },
            ct);

        return new SharePointStoredFile(destinationUrl, Path.GetFileName(destinationUrl));
    }

    public Task<SharePointStoredFile> RenameAsync(
        string serverRelativeUrl,
        string newFileName,
        bool overwrite = true,
        CancellationToken ct = default)
    {
        var sourceUrl = ServerRelativeUrl.Validate(serverRelativeUrl);
        var separatorIndex = sourceUrl.LastIndexOf('/');
        if (separatorIndex <= 0)
            throw new SharePointValidationException("File path must include a parent folder.");

        var destinationFolder = sourceUrl[..separatorIndex];
        return MoveAsync(sourceUrl, destinationFolder, newFileName, overwrite, ct);
    }

    public async Task UpdateMetadataAsync(
        string serverRelativeUrl,
        IReadOnlyDictionary<string, object?> fields,
        CancellationToken ct = default)
    {
        var validatedUrl = ServerRelativeUrl.Validate(serverRelativeUrl);
        ArgumentNullException.ThrowIfNull(fields);

        if (fields.Count == 0)
            throw new SharePointValidationException("At least one metadata field must be provided.");

        var encodedUrl = Uri.EscapeDataString(validatedUrl);

        using var message = new HttpRequestMessage(
            HttpMethod.Post,
            $"_api/web/GetFileByServerRelativeUrl('{encodedUrl}')/ListItemAllFields");

        message.Headers.Add("IF-MATCH", "*");
        message.Headers.Add("X-HTTP-Method", "MERGE");
        message.Content = SharePointContentFactory.CreateJsonContent(fields);

        using var response = await _requestExecutor.SendAsync(
            _httpClient,
            message,
            new SharePointSendOptions { IncludeFormDigest = true },
            ct);
    }

    public async Task<IReadOnlyDictionary<string, JsonElement>> GetMetadataAsync(
        string serverRelativeUrl,
        IEnumerable<string>? selectFields = null,
        CancellationToken ct = default)
    {
        var validatedUrl = ServerRelativeUrl.Validate(serverRelativeUrl);
        var encodedUrl = Uri.EscapeDataString(validatedUrl);
        var selectQuery = BuildSelectQuery(selectFields);

        using var message = new HttpRequestMessage(
            HttpMethod.Get,
            $"_api/web/GetFileByServerRelativeUrl('{encodedUrl}')/ListItemAllFields{selectQuery}");

        using var response = await _requestExecutor.SendAsync(_httpClient, message, ct: ct);
        return await ReadMetadataMapAsync(response, ct);
    }

    public async Task<IReadOnlyDictionary<string, JsonElement>?> TryGetMetadataAsync(
        string serverRelativeUrl,
        IEnumerable<string>? selectFields = null,
        CancellationToken ct = default)
    {
        var validatedUrl = ServerRelativeUrl.Validate(serverRelativeUrl);
        var encodedUrl = Uri.EscapeDataString(validatedUrl);
        var selectQuery = BuildSelectQuery(selectFields);

        using var message = new HttpRequestMessage(
            HttpMethod.Get,
            $"_api/web/GetFileByServerRelativeUrl('{encodedUrl}')/ListItemAllFields{selectQuery}");

        using var response = await _requestExecutor.SendAsync(
            _httpClient,
            message,
            new SharePointSendOptions { EnsureSuccessStatusCode = false },
            ct);

        if (response.StatusCode == HttpStatusCode.NotFound)
            return null;

        await _requestExecutor.EnsureSuccessAsync(response, ct);
        return await ReadMetadataMapAsync(response, ct);
    }

    public async Task<T?> GetMetadataValueAsync<T>(
        string serverRelativeUrl,
        string fieldName,
        CancellationToken ct = default)
    {
        ValidateMetadataFieldName(fieldName);

        var metadata = await GetMetadataAsync(serverRelativeUrl, [fieldName], ct);
        return ConvertMetadataValue<T>(metadata, fieldName);
    }

    public async Task<T?> TryGetMetadataValueAsync<T>(
        string serverRelativeUrl,
        string fieldName,
        CancellationToken ct = default)
    {
        ValidateMetadataFieldName(fieldName);

        var metadata = await TryGetMetadataAsync(serverRelativeUrl, [fieldName], ct);
        if (metadata is null)
            return default;

        return ConvertMetadataValue<T>(metadata, fieldName);
    }

    [Obsolete("Use GetFileWebUrlAsync instead.")]
    public Task<string> GetWebUrlAsync(string serverRelativeUrl, CancellationToken ct = default)
        => GetFileWebUrlAsync(serverRelativeUrl, ct);

    private static string BuildDestinationFileUrl(string destinationFolderServerRelativeUrl, string destinationFileName)
    {
        var destinationFolder = ServerRelativeUrl.Validate(destinationFolderServerRelativeUrl).TrimEnd('/');
        var fileName = ValidateLeafFileName(destinationFileName);
        return $"{destinationFolder}/{fileName}";
    }

    private static async Task<IReadOnlyDictionary<string, JsonElement>> ReadMetadataMapAsync(HttpResponseMessage response, CancellationToken ct)
    {
        var payload = await response.Content.ReadAsStringAsync(ct);
        using var document = JsonDocument.Parse(payload);

        if (document.RootElement.ValueKind != JsonValueKind.Object)
            return new Dictionary<string, JsonElement>(StringComparer.OrdinalIgnoreCase);

        var metadata = new Dictionary<string, JsonElement>(StringComparer.OrdinalIgnoreCase);
        foreach (var property in document.RootElement.EnumerateObject())
        {
            metadata[property.Name] = property.Value.Clone();
        }

        return metadata;
    }

    private static void ValidateMetadataFieldName(string fieldName)
    {
        if (string.IsNullOrWhiteSpace(fieldName))
            throw new SharePointValidationException("Metadata field name must be provided.");
    }

    private static T? ConvertMetadataValue<T>(IReadOnlyDictionary<string, JsonElement> metadata, string fieldName)
    {
        if (!metadata.TryGetValue(fieldName, out var value)
            || value.ValueKind == JsonValueKind.Null
            || value.ValueKind == JsonValueKind.Undefined)
        {
            return default;
        }

        try
        {
            return value.Deserialize<T>();
        }
        catch (JsonException ex)
        {
            throw new SharePointValidationException($"Metadata field '{fieldName}' could not be converted to {typeof(T).Name}: {ex.Message}");
        }
    }


    private static string BuildSelectQuery(IEnumerable<string>? selectFields)
    {
        if (selectFields is null)
            return string.Empty;

        var fields = selectFields
            .Where(field => !string.IsNullOrWhiteSpace(field))
            .Select(field => field.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (fields.Length == 0)
            return string.Empty;

        if (fields.Any(field => field.Contains(',')))
            throw new SharePointValidationException("Metadata select field names must not contain commas.");

        return $"?$select={string.Join(',', fields)}";
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

internal sealed class SharePointFolderClient(HttpClient httpClient, ISharePointRequestExecutor requestExecutor) : ISharePointFolderClient
{
    private readonly HttpClient _httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
    private readonly ISharePointRequestExecutor _requestExecutor = requestExecutor ?? throw new ArgumentNullException(nameof(requestExecutor));

    public async Task EnsureFolderPathAsync(string serverRelativeFolder, CancellationToken ct = default)
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

    [Obsolete("Use EnsureFolderPathAsync instead.")]
    public Task EnsurePathAsync(string serverRelativeFolder, CancellationToken ct = default)
        => EnsureFolderPathAsync(serverRelativeFolder, ct);

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

        // SharePoint On-Prem may return 500 for non-existent folders in addition to 404.
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

        using var message = new HttpRequestMessage(HttpMethod.Post, "_api/web/folders");
        message.Content = SharePointContentFactory.CreateJsonContent(new
        {
            ServerRelativeUrl = validatedFolder
        });

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

    public async Task DeleteAsync(string serverRelativeFolder, CancellationToken ct = default)
    {
        var validatedFolder = ServerRelativeUrl.Validate(serverRelativeFolder);
        var encodedFolder = Uri.EscapeDataString(validatedFolder);

        using var message = new HttpRequestMessage(
            HttpMethod.Post,
            $"_api/web/GetFolderByServerRelativeUrl('{encodedFolder}')");

        message.Headers.Add("IF-MATCH", "*");
        message.Headers.Add("X-HTTP-Method", "DELETE");

        using var response = await _requestExecutor.SendAsync(
            _httpClient,
            message,
            new SharePointSendOptions
            {
                IncludeFormDigest = true,
                EnsureSuccessStatusCode = false
            },
            ct);

        if (response.StatusCode == HttpStatusCode.NotFound || response.StatusCode == HttpStatusCode.InternalServerError)
            return;

        await _requestExecutor.EnsureSuccessAsync(response, ct);
    }
}
