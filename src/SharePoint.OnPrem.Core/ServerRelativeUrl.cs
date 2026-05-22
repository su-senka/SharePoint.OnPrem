using SharePoint.OnPrem.Abstractions;

namespace SharePoint.OnPrem.Core;

public static class ServerRelativeUrl
{
    public static string Validate(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new SharePointValidationException("Server-relative path must be provided.");

        var trimmed = value.Trim();
        if (!trimmed.StartsWith('/'))
            throw new SharePointValidationException("Server-relative path must start with '/'.");

        if (trimmed.Contains('%'))
            throw new SharePointValidationException("Server-relative path must not already be URL encoded.");

        return trimmed;
    }

    public static string Combine(string baseFolder, string relativePath)
    {
        var validatedBase = Validate(baseFolder).TrimEnd('/');
        var validatedRelative = ValidateRelative(relativePath);
        return $"{validatedBase}/{validatedRelative}";
    }

    public static string ValidateRelative(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new SharePointValidationException("Relative path must be provided.");

        var trimmed = value.Trim().Trim('/');
        if (trimmed.Length == 0)
            throw new SharePointValidationException("Relative path must not be empty.");

        if (trimmed.Contains('%'))
            throw new SharePointValidationException("Relative path must not already be URL encoded.");

        return trimmed;
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

