using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using SharePoint.OnPrem.Abstractions;

namespace SharePoint.OnPrem.Core;

public sealed class SharePointSendOptions
{
    public bool IncludeFormDigest { get; init; }
    public bool EnsureSuccessStatusCode { get; init; } = true;
    public bool UseResponseHeadersRead { get; init; }
}

internal interface IFormDigestProvider
{
    Task<string> GetDigestAsync(HttpClient httpClient, CancellationToken ct = default);
}

public interface ISharePointRequestExecutor
{
    Task<HttpResponseMessage> SendAsync(
        HttpClient httpClient,
        HttpRequestMessage request,
        SharePointSendOptions? options = null,
        CancellationToken ct = default);

    Task EnsureSuccessAsync(HttpResponseMessage response, CancellationToken ct = default);
}

internal sealed class SharePointFormDigestProvider(SharePointOnPremOptions options) : IFormDigestProvider, IDisposable
{
    private readonly SharePointOnPremOptions _options = options ?? throw new ArgumentNullException(nameof(options));
    private readonly SemaphoreSlim _gate = new(1, 1);

    private string? _cachedDigest;
    private DateTimeOffset _expiresAtUtc;

    public async Task<string> GetDigestAsync(HttpClient httpClient, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(httpClient);
        _options.Validate();

        if (TryGetCachedDigest(out var cached))
            return cached;

        await _gate.WaitAsync(ct);
        try
        {
            if (TryGetCachedDigest(out cached))
                return cached;

            using var request = new HttpRequestMessage(HttpMethod.Post, "_api/contextinfo")
            {
                Content = new StringContent(string.Empty)
            };

            using var response = await httpClient.SendAsync(request, ct);
            if (!response.IsSuccessStatusCode)
                throw await SharePointErrorFactory.CreateAsync(response, ct);

            var payload = await response.Content.ReadAsStringAsync(ct);
            using var document = JsonDocument.Parse(payload);

            var digest = ReadDigest(document.RootElement)
                ?? throw new SharePointValidationException("Failed to load form digest value from SharePoint response.")
                {
                    RequestUrl = response.RequestMessage?.RequestUri?.ToString()
                };

            var timeoutSeconds = ReadTimeoutSeconds(document.RootElement);
            if (_options.UseFormDigestCaching)
            {
                _cachedDigest = digest;
                _expiresAtUtc = DateTimeOffset.UtcNow.AddSeconds(Math.Max(30, timeoutSeconds - 30));
            }

            return digest;
        }
        catch (JsonException ex)
        {
            throw new SharePointValidationException($"Malformed response when retrieving form digest: {ex.Message}");
        }
        finally
        {
            _gate.Release();
        }
    }

    public void Dispose()
    {
        _gate.Dispose();
    }

    private bool TryGetCachedDigest(out string digest)
    {
        if (_options.UseFormDigestCaching
            && !string.IsNullOrWhiteSpace(_cachedDigest)
            && DateTimeOffset.UtcNow < _expiresAtUtc)
        {
            digest = _cachedDigest;
            return true;
        }

        digest = string.Empty;
        return false;
    }

    private static string? ReadDigest(JsonElement root)
    {
        if (root.ValueKind == JsonValueKind.Object)
        {
            if (root.TryGetProperty("FormDigestValue", out var digestValue))
                return digestValue.GetString();

            if (root.TryGetProperty("d", out var d)
                && d.ValueKind == JsonValueKind.Object
                && d.TryGetProperty("GetContextWebInformation", out var info)
                && info.ValueKind == JsonValueKind.Object
                && info.TryGetProperty("FormDigestValue", out var nestedDigest))
            {
                return nestedDigest.GetString();
            }
        }

        return null;
    }

    private static int ReadTimeoutSeconds(JsonElement root)
    {
        if (root.ValueKind == JsonValueKind.Object)
        {
            if (root.TryGetProperty("FormDigestTimeoutSeconds", out var timeout)
                && timeout.TryGetInt32(out var direct))
            {
                return direct;
            }

            if (root.TryGetProperty("d", out var d)
                && d.ValueKind == JsonValueKind.Object
                && d.TryGetProperty("GetContextWebInformation", out var info)
                && info.ValueKind == JsonValueKind.Object
                && info.TryGetProperty("FormDigestTimeoutSeconds", out var nestedTimeout)
                && nestedTimeout.TryGetInt32(out var nested))
            {
                return nested;
            }
        }

        return 1800;
    }
}

internal sealed class SharePointRequestExecutor(IFormDigestProvider formDigestProvider) : ISharePointRequestExecutor
{
    private readonly IFormDigestProvider _formDigestProvider = formDigestProvider ?? throw new ArgumentNullException(nameof(formDigestProvider));

    public async Task<HttpResponseMessage> SendAsync(
        HttpClient httpClient,
        HttpRequestMessage request,
        SharePointSendOptions? options = null,
        CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(httpClient);
        ArgumentNullException.ThrowIfNull(request);

        var effectiveOptions = options ?? new SharePointSendOptions();

        if (effectiveOptions.IncludeFormDigest && !request.Headers.Contains("X-RequestDigest"))
        {
            var digest = await _formDigestProvider.GetDigestAsync(httpClient, ct);
            request.Headers.Add("X-RequestDigest", digest);
        }

        var completionOption = effectiveOptions.UseResponseHeadersRead
            ? HttpCompletionOption.ResponseHeadersRead
            : HttpCompletionOption.ResponseContentRead;

        var response = await httpClient.SendAsync(request, completionOption, ct);
        if (!effectiveOptions.EnsureSuccessStatusCode)
            return response;

        if (response.IsSuccessStatusCode)
            return response;

        var exception = await SharePointErrorFactory.CreateAsync(response, ct);
        response.Dispose();
        throw exception;
    }

    public async Task EnsureSuccessAsync(HttpResponseMessage response, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(response);

        if (response.IsSuccessStatusCode)
            return;

        throw await SharePointErrorFactory.CreateAsync(response, ct);
    }
}

internal static class SharePointContentFactory
{
    public static HttpContent CreateJsonContent(object payload)
    {
        var json = JsonSerializer.Serialize(payload);
        var content = new StringContent(json, Encoding.UTF8);
        content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=nometadata");
        return content;
    }
}

/// <summary>
/// A Stream that keeps the owning HttpResponseMessage alive until the stream itself is disposed.
/// Used to implement true streaming downloads without eagerly buffering the response body.
/// </summary>
internal sealed class HttpResponseStream : Stream
{
    private readonly HttpResponseMessage _response;
    private readonly Stream _inner;

    internal HttpResponseStream(HttpResponseMessage response, Stream inner)
    {
        _response = response;
        _inner = inner;
    }

    public override bool CanRead => _inner.CanRead;
    public override bool CanSeek => _inner.CanSeek;
    public override bool CanWrite => _inner.CanWrite;
    public override long Length => _inner.Length;
    public override long Position { get => _inner.Position; set => _inner.Position = value; }
    public override void Flush() => _inner.Flush();
    public override Task FlushAsync(CancellationToken ct) => _inner.FlushAsync(ct);
    public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
    public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken ct) => _inner.ReadAsync(buffer, offset, count, ct);
    public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken ct = default) => _inner.ReadAsync(buffer, ct);
    public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
    public override void SetLength(long value) => _inner.SetLength(value);
    public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _inner.Dispose();
            _response.Dispose();
        }
        base.Dispose(disposing);
    }

    public override async ValueTask DisposeAsync()
    {
        await _inner.DisposeAsync();
        _response.Dispose();
        await base.DisposeAsync();
    }
}

internal static class SharePointErrorFactory
{
    public static async Task<SharePointException> CreateAsync(HttpResponseMessage response, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(response);

        var requestUrl = response.RequestMessage?.RequestUri?.ToString();
        var correlationId = response.Headers.TryGetValues("SPRequestGuid", out var values)
            ? values.FirstOrDefault()
            : null;

        var detail = await TryReadProblemDetailAsync(response, ct);
        var message = string.IsNullOrWhiteSpace(detail)
            ? $"SharePoint request failed. Status={(int)response.StatusCode} {response.StatusCode}, Url={requestUrl}, SPRequestGuid={correlationId}"
            : $"SharePoint request failed. Status={(int)response.StatusCode} {response.StatusCode}, Url={requestUrl}, SPRequestGuid={correlationId}. Detail={detail}";

        return response.StatusCode switch
        {
            HttpStatusCode.Unauthorized or HttpStatusCode.Forbidden
                => new SharePointUnauthorizedException(message) { RequestUrl = requestUrl, CorrelationId = correlationId },
            HttpStatusCode.NotFound
                => new SharePointNotFoundException(message) { RequestUrl = requestUrl, CorrelationId = correlationId },
            HttpStatusCode.Conflict
                => new SharePointConflictException(message) { RequestUrl = requestUrl, CorrelationId = correlationId },
            (HttpStatusCode)429 or HttpStatusCode.ServiceUnavailable
                => new SharePointTransientException(message) { RequestUrl = requestUrl, CorrelationId = correlationId },
            _ => new SharePointException(message) { RequestUrl = requestUrl, CorrelationId = correlationId }
        };
    }

    private static async Task<string?> TryReadProblemDetailAsync(HttpResponseMessage response, CancellationToken ct)
    {
        if (response.Content is null)
            return null;

        var payload = await response.Content.ReadAsStringAsync(ct);
        if (string.IsNullOrWhiteSpace(payload))
            return null;

        try
        {
            using var document = JsonDocument.Parse(payload);
            if (document.RootElement.ValueKind == JsonValueKind.Object)
            {
                if (document.RootElement.TryGetProperty("error", out var error)
                    && error.ValueKind == JsonValueKind.Object
                    && error.TryGetProperty("message", out var messageElement))
                {
                    return messageElement.GetString();
                }

                if (document.RootElement.TryGetProperty("Detail", out var detailUpper))
                    return detailUpper.GetString();

                if (document.RootElement.TryGetProperty("detail", out var detailLower))
                    return detailLower.GetString();
            }
        }
        catch (JsonException)
        {
            return payload;
        }

        return payload;
    }
}

