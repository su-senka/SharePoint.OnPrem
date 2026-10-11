using System.Net;

namespace SharePoint.OnPrem.Abstractions;

/// <summary>
/// Base type for SharePoint integration failures with optional transport metadata.
/// </summary>
public class SharePointException(string message, Exception? innerException = null) : Exception(message, innerException)
{
    /// <summary>The <c>SPRequestGuid</c> response header; use it to find the request in the SharePoint ULS logs.</summary>
    public string? CorrelationId { get; init; }

    /// <summary>The request URL that failed, when the failure came from an HTTP response.</summary>
    public string? RequestUrl { get; init; }

    /// <summary>The HTTP status code returned by SharePoint, when the failure came from an HTTP response.</summary>
    public HttpStatusCode? StatusCode { get; init; }
}

/// <summary>SharePoint returned 401 or 403. Check the credentials configured on the HTTP client and the principal's permissions.</summary>
public sealed class SharePointUnauthorizedException(string message, Exception? innerException = null) : SharePointException(message, innerException);

/// <summary>The requested file, folder, group, or user does not exist (HTTP 404).</summary>
public sealed class SharePointNotFoundException(string message, Exception? innerException = null) : SharePointException(message, innerException);

/// <summary>The operation conflicts with the current state of the resource (HTTP 409).</summary>
public sealed class SharePointConflictException(string message, Exception? innerException = null) : SharePointException(message, innerException);

/// <summary>SharePoint is throttling or temporarily unavailable (HTTP 429, 502, 503, or 504). The operation can usually be retried.</summary>
public sealed class SharePointTransientException(string message, Exception? innerException = null) : SharePointException(message, innerException)
{
    /// <summary>How long SharePoint asked the client to wait (the <c>Retry-After</c> header), when it sent one.</summary>
    public TimeSpan? RetryAfter { get; init; }
}

/// <summary>An argument failed client-side validation, or SharePoint returned a response the library could not interpret.</summary>
public sealed class SharePointValidationException(string message, Exception? innerException = null) : SharePointException(message, innerException);
