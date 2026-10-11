using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using SharePoint.OnPrem.Abstractions;

namespace SharePoint.OnPrem.Core;

/// <summary>Per-request options for <see cref="ISharePointRequestExecutor.SendAsync"/>.</summary>
public sealed class SharePointSendOptions
{
    /// <summary>Attach an <c>X-RequestDigest</c> header. Required for every write (POST) request.</summary>
    public bool IncludeFormDigest { get; init; }

    /// <summary>Throw a typed <see cref="SharePointException"/> for non-success responses. Default: true.</summary>
    public bool EnsureSuccessStatusCode { get; init; } = true;

    /// <summary>Return as soon as headers arrive, so the body can be streamed.</summary>
    public bool UseResponseHeadersRead { get; init; }
}

internal interface IFormDigestProvider
{
    Task<string> GetDigestAsync(HttpClient httpClient, CancellationToken ct = default);

    /// <summary>Drops <paramref name="digest"/> from the cache. Returns false when caching is disabled (a retry would not help).</summary>
    bool Invalidate(string digest);
}

/// <summary>
/// Sends requests to SharePoint with form digest handling and typed error mapping. Use it to call REST endpoints the
/// typed clients do not cover, with the same pipeline: inject it together with an <see cref="HttpClient"/> from
/// <c>IHttpClientFactory.CreateClient(ServiceCollectionExtensions.HttpClientName)</c>.
/// </summary>
public interface ISharePointRequestExecutor
{
    /// <summary>
    /// Sends <paramref name="request"/>. When a write fails with 403 because a cached form digest expired early
    /// (e.g. after an application pool recycle), the digest is refreshed and a replayable request is retried once.
    /// </summary>
    Task<HttpResponseMessage> SendAsync(
        HttpClient httpClient,
        HttpRequestMessage request,
        SharePointSendOptions? options = null,
        CancellationToken ct = default);

    /// <summary>Throws the typed <see cref="SharePointException"/> for a non-success response.</summary>
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
            throw new SharePointValidationException($"Malformed response when retrieving form digest: {ex.Message}", ex);
        }
        finally
        {
            _gate.Release();
        }
    }

    public bool Invalidate(string digest)
    {
        if (!_options.UseFormDigestCaching)
            return false;

        // Only clear our own digest; another caller may already have cached a fresh one.
        Interlocked.CompareExchange(ref _cachedDigest, null, digest);
        return true;
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

        string? digest = null;
        if (effectiveOptions.IncludeFormDigest && !request.Headers.Contains(DigestHeader))
        {
            digest = await _formDigestProvider.GetDigestAsync(httpClient, ct);
            request.Headers.Add(DigestHeader, digest);
        }

        var completionOption = effectiveOptions.UseResponseHeadersRead
            ? HttpCompletionOption.ResponseHeadersRead
            : HttpCompletionOption.ResponseContentRead;

        var response = await httpClient.SendAsync(request, completionOption, ct);

        // A cached digest stops being valid before its timeout when the application pool recycles; SharePoint then
        // answers 403 ("The security validation for this page is invalid"). Refresh it and retry once if we can replay the body.
        if (digest is not null
            && response.StatusCode == HttpStatusCode.Forbidden
            && IsReplayable(request.Content)
            && _formDigestProvider.Invalidate(digest))
        {
            var retry = await CloneAsync(request, ct);
            retry.Headers.Add(DigestHeader, await _formDigestProvider.GetDigestAsync(httpClient, ct));
            response.Dispose();
            response = await httpClient.SendAsync(retry, completionOption, ct);
        }

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

    private const string DigestHeader = "X-RequestDigest";

    // Streamed uploads cannot be re-read; JSON and chunk bodies are byte arrays and can be.
    private static bool IsReplayable(HttpContent? content) => content is null or ByteArrayContent;

    private static async Task<HttpRequestMessage> CloneAsync(HttpRequestMessage request, CancellationToken ct)
    {
        var clone = new HttpRequestMessage(request.Method, request.RequestUri)
        {
            Version = request.Version,
            VersionPolicy = request.VersionPolicy
        };

        foreach (var header in request.Headers.Where(h => !h.Key.Equals(DigestHeader, StringComparison.OrdinalIgnoreCase)))
            clone.Headers.TryAddWithoutValidation(header.Key, header.Value);

        if (request.Content is not null)
        {
            clone.Content = new ByteArrayContent(await request.Content.ReadAsByteArrayAsync(ct));
            foreach (var header in request.Content.Headers.Where(h => !h.Key.Equals("Content-Length", StringComparison.OrdinalIgnoreCase)))
                clone.Content.Headers.TryAddWithoutValidation(header.Key, header.Value);
        }

        return clone;
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

    /// <summary>
    /// Wraps a caller-owned stream as request content. Unlike <see cref="StreamContent"/>, disposing the request
    /// does not dispose the stream, so the caller keeps control of its lifetime.
    /// </summary>
    public static HttpContent CreateNonOwningStreamContent(Stream content) => new NonOwningStreamContent(content);

    private sealed class NonOwningStreamContent(Stream content) : HttpContent
    {
        protected override Task SerializeToStreamAsync(Stream stream, TransportContext? context)
            => content.CopyToAsync(stream);

        protected override Task SerializeToStreamAsync(Stream stream, TransportContext? context, CancellationToken cancellationToken)
            => content.CopyToAsync(stream, cancellationToken);

        protected override bool TryComputeLength(out long length)
        {
            if (content.CanSeek)
            {
                length = content.Length - content.Position;
                return true;
            }

            length = 0;
            return false;
        }
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

        var statusCode = response.StatusCode;
        return statusCode switch
        {
            HttpStatusCode.Unauthorized or HttpStatusCode.Forbidden
                => new SharePointUnauthorizedException(message) { RequestUrl = requestUrl, CorrelationId = correlationId, StatusCode = statusCode },
            HttpStatusCode.NotFound
                => new SharePointNotFoundException(message) { RequestUrl = requestUrl, CorrelationId = correlationId, StatusCode = statusCode },
            HttpStatusCode.Conflict
                => new SharePointConflictException(message) { RequestUrl = requestUrl, CorrelationId = correlationId, StatusCode = statusCode },
            HttpStatusCode.TooManyRequests or HttpStatusCode.BadGateway or HttpStatusCode.ServiceUnavailable or HttpStatusCode.GatewayTimeout
                => new SharePointTransientException(message)
                {
                    RequestUrl = requestUrl,
                    CorrelationId = correlationId,
                    StatusCode = statusCode,
                    RetryAfter = ReadRetryAfter(response.Headers.RetryAfter)
                },
            _ => new SharePointException(message) { RequestUrl = requestUrl, CorrelationId = correlationId, StatusCode = statusCode }
        };
    }

    // Error bodies that are not SharePoint JSON (IIS or proxy HTML pages) can be large; keep exception messages readable.
    private const int MaxDetailLength = 1000;

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
            var root = document.RootElement;
            if (root.ValueKind == JsonValueKind.Object)
            {
                // JSON light ("odata.error") and verbose ("error") both carry message either as a string or as { lang, value }.
                if ((root.TryGetProperty("odata.error", out var error) || root.TryGetProperty("error", out error))
                    && error.ValueKind == JsonValueKind.Object
                    && error.TryGetProperty("message", out var messageElement))
                {
                    var errorMessage = ReadText(messageElement);
                    if (!string.IsNullOrWhiteSpace(errorMessage))
                        return errorMessage;
                }

                if (root.TryGetProperty("Detail", out var detail) || root.TryGetProperty("detail", out detail))
                {
                    var detailText = ReadText(detail);
                    if (!string.IsNullOrWhiteSpace(detailText))
                        return detailText;
                }
            }
        }
        catch (JsonException)
        {
        }

        return payload.Length <= MaxDetailLength ? payload : payload[..MaxDetailLength] + "…";
    }

    private static TimeSpan? ReadRetryAfter(RetryConditionHeaderValue? retryAfter)
    {
        if (retryAfter?.Delta is { } delta)
            return delta;

        if (retryAfter?.Date is { } date)
        {
            var wait = date - DateTimeOffset.UtcNow;
            return wait > TimeSpan.Zero ? wait : TimeSpan.Zero;
        }

        return null;
    }

    private static string? ReadText(JsonElement element) => element.ValueKind switch
    {
        JsonValueKind.String => element.GetString(),
        JsonValueKind.Object when element.TryGetProperty("value", out var value) && value.ValueKind == JsonValueKind.String
            => value.GetString(),
        _ => null
    };
}

