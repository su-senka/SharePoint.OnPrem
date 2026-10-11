using SharePoint.OnPrem.Abstractions;
using SharePoint.OnPrem.Core;
using System.Globalization;
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
        var encodedFolder = ODataLiteral.Encode(folderServerRelativeUrl);
        var encodedFileName = ODataLiteral.Encode(fileName);

        using var message = new HttpRequestMessage(
            HttpMethod.Post,
            $"_api/web/GetFolderByServerRelativeUrl('{encodedFolder}')/Files/add(url='{encodedFileName}',overwrite={request.Overwrite.ToString().ToLowerInvariant()})");

        ArgumentNullException.ThrowIfNull(request.Content);

        var streamContent = SharePointContentFactory.CreateNonOwningStreamContent(request.Content);
        if (!string.IsNullOrWhiteSpace(request.ContentType))
            streamContent.Headers.ContentType = MediaTypeHeaderValue.Parse(request.ContentType);

        message.Content = streamContent;

        using var response = await _requestExecutor.SendAsync(
            _httpClient,
            message,
            new SharePointSendOptions { IncludeFormDigest = true },
            ct);

        return await ReadStoredFileAsync(response, $"{folderServerRelativeUrl.TrimEnd('/')}/{fileName}", fileName, ct);
    }

    public async Task<SharePointStoredFile> UploadChunkedAsync(
        UploadFileRequest request,
        int chunkSizeBytes = 10 * 1024 * 1024,
        IProgress<long>? progress = null,
        CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(request);
        ArgumentNullException.ThrowIfNull(request.Content);
        ArgumentOutOfRangeException.ThrowIfLessThan(chunkSizeBytes, 1);

        // Two buffers: we must know whether a chunk is the last one before sending it (FinishUpload vs ContinueUpload),
        // and the stream may not be seekable.
        var current = new byte[chunkSizeBytes];
        var next = new byte[chunkSizeBytes];
        var currentLength = await ReadChunkAsync(request.Content, current, ct);
        var nextLength = await ReadChunkAsync(request.Content, next, ct);

        if (nextLength == 0)
        {
            // Fits in a single request: a plain upload is cheaper and also works on SharePoint 2013.
            using var single = new MemoryStream(current, 0, currentLength, writable: false);
            var stored = await UploadAsync(request with { Content = single }, ct);
            progress?.Report(currentLength);
            return stored;
        }

        // The upload session targets an existing file, so create an empty one first (this also enforces Overwrite).
        var placeholder = await UploadAsync(request with { Content = Stream.Null }, ct);
        var fileEndpoint = $"_api/web/GetFileByServerRelativeUrl('{ODataLiteral.Encode(placeholder.ServerRelativeUrl)}')";
        var uploadId = Guid.NewGuid();
        long offset = 0;

        try
        {
            (await SendChunkAsync($"{fileEndpoint}/StartUpload(uploadId=guid'{uploadId}')", current, currentLength, ct)).Dispose();

            offset += currentLength;
            progress?.Report(offset);
            (current, next, currentLength) = (next, current, nextLength);

            while ((nextLength = await ReadChunkAsync(request.Content, next, ct)) > 0)
            {
                (await SendChunkAsync($"{fileEndpoint}/ContinueUpload(uploadId=guid'{uploadId}',fileOffset={offset})", current, currentLength, ct)).Dispose();

                offset += currentLength;
                progress?.Report(offset);
                (current, next, currentLength) = (next, current, nextLength);
            }

            using var finish = await SendChunkAsync($"{fileEndpoint}/FinishUpload(uploadId=guid'{uploadId}',fileOffset={offset})", current, currentLength, ct);
            offset += currentLength;
            progress?.Report(offset);

            return await ReadStoredFileAsync(finish, placeholder.ServerRelativeUrl, placeholder.FileName, ct);
        }
        catch (Exception)
        {
            await TryCancelUploadAsync(fileEndpoint, uploadId);
            throw;
        }
    }

    public async Task CheckOutAsync(string serverRelativeUrl, CancellationToken ct = default)
        => await PostToFileAsync(serverRelativeUrl, "CheckOut()", ct);

    public async Task CheckInAsync(
        string serverRelativeUrl,
        string? comment = null,
        SharePointCheckInType checkInType = SharePointCheckInType.MinorCheckIn,
        CancellationToken ct = default)
        => await PostToFileAsync(serverRelativeUrl, $"CheckIn(comment='{ODataLiteral.Encode(comment ?? string.Empty)}',checkintype={(int)checkInType})", ct);

    public async Task UndoCheckOutAsync(string serverRelativeUrl, CancellationToken ct = default)
        => await PostToFileAsync(serverRelativeUrl, "UndoCheckOut()", ct);

    public async Task<IReadOnlyList<SharePointFileVersion>> GetVersionsAsync(string serverRelativeUrl, CancellationToken ct = default)
    {
        var encodedUrl = ODataLiteral.Encode(ServerRelativeUrl.Validate(serverRelativeUrl));

        using var message = new HttpRequestMessage(
            HttpMethod.Get,
            $"_api/web/GetFileByServerRelativeUrl('{encodedUrl}')/Versions?$select=ID,VersionLabel,Created,Size,CheckInComment,IsCurrentVersion");

        using var response = await _requestExecutor.SendAsync(_httpClient, message, ct: ct);
        using var document = JsonDocument.Parse(await response.Content.ReadAsStringAsync(ct));

        if (!document.RootElement.TryGetProperty("value", out var value) || value.ValueKind != JsonValueKind.Array)
            return [];

        return value.EnumerateArray()
            .Select(version =>
            {
                var comment = JsonReading.ReadString(version, "CheckInComment");
                return new SharePointFileVersion(
                    (int)JsonReading.ReadInt64(version, "ID"),
                    JsonReading.ReadString(version, "VersionLabel"),
                    JsonReading.ReadDateTime(version, "Created"),
                    JsonReading.ReadInt64(version, "Size"),
                    comment.Length == 0 ? null : comment,
                    version.TryGetProperty("IsCurrentVersion", out var isCurrent) && isCurrent.ValueKind == JsonValueKind.True);
            })
            .ToList();
    }

    private async Task PostToFileAsync(string serverRelativeUrl, string action, CancellationToken ct)
    {
        var encodedUrl = ODataLiteral.Encode(ServerRelativeUrl.Validate(serverRelativeUrl));

        using var message = new HttpRequestMessage(HttpMethod.Post, $"_api/web/GetFileByServerRelativeUrl('{encodedUrl}')/{action}");
        using var response = await _requestExecutor.SendAsync(
            _httpClient,
            message,
            new SharePointSendOptions { IncludeFormDigest = true },
            ct);
    }

    private async Task<HttpResponseMessage> SendChunkAsync(string endpoint, byte[] buffer, int length, CancellationToken ct)
    {
        using var message = new HttpRequestMessage(HttpMethod.Post, endpoint)
        {
            Content = new ByteArrayContent(buffer, 0, length)
        };

        return await _requestExecutor.SendAsync(
            _httpClient,
            message,
            new SharePointSendOptions { IncludeFormDigest = true },
            ct);
    }

    private async Task TryCancelUploadAsync(string fileEndpoint, Guid uploadId)
    {
        try
        {
            // Not tied to the caller's token: the upload may have failed precisely because it was cancelled.
            using var message = new HttpRequestMessage(HttpMethod.Post, $"{fileEndpoint}/CancelUpload(uploadId=guid'{uploadId}')");
            using var response = await _requestExecutor.SendAsync(
                _httpClient,
                message,
                new SharePointSendOptions { IncludeFormDigest = true, EnsureSuccessStatusCode = false },
                CancellationToken.None);
        }
        catch (Exception)
        {
            // Best effort; the original failure is more useful to the caller.
        }
    }

    private static async Task<int> ReadChunkAsync(Stream stream, byte[] buffer, CancellationToken ct)
    {
        var total = 0;
        int read;
        while (total < buffer.Length && (read = await stream.ReadAsync(buffer.AsMemory(total), ct)) > 0)
            total += read;

        return total;
    }

    // Reads the server-confirmed name and URL; falls back to the expected values if absent.
    private static async Task<SharePointStoredFile> ReadStoredFileAsync(
        HttpResponseMessage response,
        string fallbackServerRelativeUrl,
        string fallbackFileName,
        CancellationToken ct)
    {
        string? responseServerRelativeUrl = null;
        string? responseFileName = null;
        try
        {
            var payload = await response.Content.ReadAsStringAsync(ct);
            if (!string.IsNullOrWhiteSpace(payload))
            {
                using var document = JsonDocument.Parse(payload);
                if (document.RootElement.ValueKind == JsonValueKind.Object)
                {
                    responseServerRelativeUrl = JsonReading.ReadString(document.RootElement, "ServerRelativeUrl");
                    responseFileName = JsonReading.ReadString(document.RootElement, "Name");
                }
            }
        }
        catch (JsonException) { }

        return new SharePointStoredFile(
            !string.IsNullOrWhiteSpace(responseServerRelativeUrl) ? responseServerRelativeUrl : fallbackServerRelativeUrl,
            !string.IsNullOrWhiteSpace(responseFileName) ? responseFileName : fallbackFileName);
    }

    public async Task<SharePointFileDownload> DownloadAsync(string serverRelativeUrl, CancellationToken ct = default)
    {
        var validatedUrl = ServerRelativeUrl.Validate(serverRelativeUrl);
        var encodedUrl = ODataLiteral.Encode(validatedUrl);

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
        var encodedUrl = ODataLiteral.Encode(validatedUrl);

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
        var encodedUrl = ODataLiteral.Encode(validatedUrl);

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
        var encodedUrl = ODataLiteral.Encode(validatedUrl);

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
        var encodedSource = ODataLiteral.Encode(sourceUrl);
        var encodedDestination = ODataLiteral.Encode(destinationUrl);

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
        var encodedSource = ODataLiteral.Encode(sourceUrl);
        var encodedDestination = ODataLiteral.Encode(destinationUrl);
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

        var encodedUrl = ODataLiteral.Encode(validatedUrl);

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
        var encodedUrl = ODataLiteral.Encode(validatedUrl);
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
        var encodedUrl = ODataLiteral.Encode(validatedUrl);
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
            throw new SharePointValidationException($"Metadata field '{fieldName}' could not be converted to {typeof(T).Name}: {ex.Message}", ex);
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
        var validatedFolder = ServerRelativeUrl.Validate(serverRelativeFolder).TrimEnd('/');
        if (validatedFolder.Length == 0 || await ExistsAsync(validatedFolder, ct))
            return;

        // The segments of the site's own URL (e.g. /sites/team) are not folders in this web and cannot be created here.
        var webPath = _httpClient.BaseAddress is null
            ? string.Empty
            : Uri.UnescapeDataString(_httpClient.BaseAddress.AbsolutePath).TrimEnd('/');

        var segments = validatedFolder.Split('/', StringSplitOptions.RemoveEmptyEntries);
        var current = string.Empty;

        foreach (var segment in segments)
        {
            current += "/" + segment;
            if (IsSameOrAncestorPath(current, webPath))
                continue;

            if (await ExistsAsync(current, ct))
                continue;

            await CreateAsync(current, ct);
        }
    }

    private static bool IsSameOrAncestorPath(string candidate, string path)
        => path.Equals(candidate, StringComparison.OrdinalIgnoreCase)
           || path.StartsWith(candidate + "/", StringComparison.OrdinalIgnoreCase);

    [Obsolete("Use EnsureFolderPathAsync instead.")]
    public Task EnsurePathAsync(string serverRelativeFolder, CancellationToken ct = default)
        => EnsureFolderPathAsync(serverRelativeFolder, ct);

    public async Task<bool> ExistsAsync(string serverRelativeFolder, CancellationToken ct = default)
    {
        var validatedFolder = ServerRelativeUrl.Validate(serverRelativeFolder);
        var encodedFolder = ODataLiteral.Encode(validatedFolder);

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

    public async Task<IReadOnlyList<SharePointFileInfo>> GetFilesAsync(string serverRelativeFolder, CancellationToken ct = default)
    {
        var items = await GetChildrenAsync(serverRelativeFolder, "Files?$select=Name,ServerRelativeUrl,Length,TimeCreated,TimeLastModified", ct);
        return items
            .Select(item => new SharePointFileInfo(
                JsonReading.ReadString(item, "Name"),
                JsonReading.ReadString(item, "ServerRelativeUrl"),
                JsonReading.ReadInt64(item, "Length"),
                JsonReading.ReadDateTime(item, "TimeCreated"),
                JsonReading.ReadDateTime(item, "TimeLastModified")))
            .ToList();
    }

    public async Task<IReadOnlyList<SharePointFolderInfo>> GetFoldersAsync(string serverRelativeFolder, CancellationToken ct = default)
    {
        var items = await GetChildrenAsync(serverRelativeFolder, "Folders?$select=Name,ServerRelativeUrl,ItemCount", ct);
        return items
            .Select(item => new SharePointFolderInfo(
                JsonReading.ReadString(item, "Name"),
                JsonReading.ReadString(item, "ServerRelativeUrl"),
                (int)JsonReading.ReadInt64(item, "ItemCount")))
            .ToList();
    }

    private async Task<IReadOnlyList<JsonElement>> GetChildrenAsync(string serverRelativeFolder, string childQuery, CancellationToken ct)
    {
        var validatedFolder = ServerRelativeUrl.Validate(serverRelativeFolder);
        var encodedFolder = ODataLiteral.Encode(validatedFolder);

        using var message = new HttpRequestMessage(
            HttpMethod.Get,
            $"_api/web/GetFolderByServerRelativeUrl('{encodedFolder}')/{childQuery}");

        using var response = await _requestExecutor.SendAsync(_httpClient, message, ct: ct);
        var payload = await response.Content.ReadAsStringAsync(ct);
        using var document = JsonDocument.Parse(payload);

        if (!document.RootElement.TryGetProperty("value", out var value) || value.ValueKind != JsonValueKind.Array)
            return [];

        return value.EnumerateArray().Select(item => item.Clone()).ToList();
    }

    public async Task DeleteAsync(string serverRelativeFolder, CancellationToken ct = default)
    {
        var validatedFolder = ServerRelativeUrl.Validate(serverRelativeFolder);
        var encodedFolder = ODataLiteral.Encode(validatedFolder);

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

internal static class JsonReading
{
    public static string ReadString(JsonElement element, string propertyName)
        => element.TryGetProperty(propertyName, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? string.Empty
            : string.Empty;

    // SharePoint serializes Edm.Int64 values (e.g. file Length) as JSON strings.
    public static long ReadInt64(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var value))
            return 0;

        return value.ValueKind switch
        {
            JsonValueKind.Number when value.TryGetInt64(out var number) => number,
            JsonValueKind.String when long.TryParse(value.GetString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) => parsed,
            _ => 0
        };
    }

    public static DateTimeOffset? ReadDateTime(JsonElement element, string propertyName)
        => element.TryGetProperty(propertyName, out var value)
           && value.ValueKind == JsonValueKind.String
           && DateTimeOffset.TryParse(value.GetString(), CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var parsed)
            ? parsed
            : null;
}
