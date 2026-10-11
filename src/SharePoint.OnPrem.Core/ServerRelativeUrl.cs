using SharePoint.OnPrem.Abstractions;

namespace SharePoint.OnPrem.Core;

/// <summary>
/// Validation and composition helpers for server-relative SharePoint paths such as <c>/sites/team/Shared Documents/report.docx</c>.
/// Paths are passed decoded; the library encodes them when building requests.
/// </summary>
public static class ServerRelativeUrl
{
    /// <summary>
    /// Validates a server-relative path and returns it trimmed.
    /// </summary>
    /// <exception cref="SharePointValidationException">
    /// The path is empty, does not start with '/', is already URL-encoded, or contains '.' or '..' segments.
    /// </exception>
    public static string Validate(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new SharePointValidationException("Server-relative path must be provided.");

        var trimmed = value.Trim();
        if (!trimmed.StartsWith('/'))
            throw new SharePointValidationException("Server-relative path must start with '/'.");

        if (trimmed.Contains('%'))
            throw new SharePointValidationException("Server-relative path must not already be URL encoded.");

        EnsureNoDotSegments(trimmed);
        return trimmed;
    }

    /// <summary>
    /// Appends a relative path to a server-relative base folder, e.g. <c>Combine("/sites/team/Docs", "2026/Q1")</c>.
    /// </summary>
    public static string Combine(string baseFolder, string relativePath)
    {
        var validatedBase = Validate(baseFolder).TrimEnd('/');
        var validatedRelative = ValidateRelative(relativePath);
        return $"{validatedBase}/{validatedRelative}";
    }

    /// <summary>
    /// Validates a path relative to some base folder and returns it with surrounding slashes trimmed.
    /// </summary>
    public static string ValidateRelative(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new SharePointValidationException("Relative path must be provided.");

        var trimmed = value.Trim().Trim('/');
        if (trimmed.Length == 0)
            throw new SharePointValidationException("Relative path must not be empty.");

        if (trimmed.Contains('%'))
            throw new SharePointValidationException("Relative path must not already be URL encoded.");

        EnsureNoDotSegments(trimmed);
        return trimmed;
    }

    // Rejecting dot segments keeps a scoped relative path (often derived from user input) inside its base folder.
    private static void EnsureNoDotSegments(string path)
    {
        foreach (var segment in path.Split('/', '\\'))
        {
            if (segment.Trim() is "." or "..")
                throw new SharePointValidationException("Path must not contain '.' or '..' segments.");
        }
    }
}

#pragma warning disable CS0618
internal sealed class BaseFolderPathScope(SharePointStorageScopeOptions options) : ISharePointServerRelativePathScope, ISharePointPathScope
#pragma warning restore CS0618
{
    private readonly SharePointStorageScopeOptions _options = options ?? throw new ArgumentNullException(nameof(options));

    public string ToServerRelativePath(string relativePath)
    {
        if (string.IsNullOrWhiteSpace(_options.BaseFolderServerRelativeUrl))
            throw new InvalidOperationException("BaseFolderServerRelativeUrl is not configured for path scoping.");

        return ServerRelativeUrl.Combine(_options.BaseFolderServerRelativeUrl, relativePath);
    }
}

