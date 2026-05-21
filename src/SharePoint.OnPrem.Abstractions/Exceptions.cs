namespace SharePoint.OnPrem.Abstractions;

/// <summary>
/// Base type for SharePoint integration failures with optional transport metadata.
/// </summary>
public class SharePointException(string message) : Exception(message)
{
    public string? CorrelationId { get; init; }
    public string? RequestUrl { get; init; }
}

public sealed class SharePointUnauthorizedException(string message) : SharePointException(message);
public sealed class SharePointNotFoundException(string message) : SharePointException(message);
public sealed class SharePointConflictException(string message) : SharePointException(message);
public sealed class SharePointTransientException(string message) : SharePointException(message);
public sealed class SharePointValidationException(string message) : SharePointException(message);

