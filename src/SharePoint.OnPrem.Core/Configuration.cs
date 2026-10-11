namespace SharePoint.OnPrem.Core;

/// <summary>
/// Connection settings. Bind from the <c>SharePoint</c> configuration section or configure in code.
/// </summary>
public sealed class SharePointOnPremOptions
{
    /// <summary>Absolute URL of the site (web) all requests are sent to, e.g. <c>https://sharepoint.contoso.local/sites/team</c>.</summary>
    public string SiteBaseUrl { get; set; } = string.Empty;

    /// <summary>HTTP request timeout in minutes. Raise it for large uploads. Default: 5.</summary>
    public int HttpTimeoutMinutes { get; set; } = 5;

    /// <summary>Reuse the form digest until shortly before it expires instead of requesting one per write. Default: true.</summary>
    public bool UseFormDigestCaching { get; set; } = true;

    /// <exception cref="InvalidOperationException">A setting is missing or invalid.</exception>
    public void Validate()
    {
        if (string.IsNullOrWhiteSpace(SiteBaseUrl))
            throw new InvalidOperationException("SharePoint:SiteBaseUrl configuration is required.");

        if (!Uri.TryCreate(SiteBaseUrl, UriKind.Absolute, out var uri) || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
            throw new InvalidOperationException("SharePoint:SiteBaseUrl must be an absolute http(s) URI.");

        if (HttpTimeoutMinutes < 1)
            throw new InvalidOperationException("SharePoint:HttpTimeoutMinutes must be greater than or equal to 1.");
    }
}

/// <summary>
/// Settings used to turn login names into the form SharePoint expects. Only the security client uses them.
/// </summary>
public sealed class SharePointIdentityOptions
{
    /// <summary>
    /// NetBIOS domain prepended to bare logins (<c>jdoe</c> → <c>CONTOSO\jdoe</c>). Optional; when empty, bare logins are rejected.
    /// </summary>
    public string Domain { get; set; } = string.Empty;

    /// <summary>
    /// Claims prefix prepended to <c>DOMAIN\user</c> logins. Default: <c>i:0#.w|</c> (Windows claims authentication).
    /// Set to an empty string for classic-mode web applications.
    /// </summary>
    public string ClaimsPrefix { get; set; } = "i:0#.w|";

    /// <exception cref="InvalidOperationException">A setting is invalid.</exception>
    public void Validate()
    {
        if (!string.IsNullOrEmpty(Domain) && Domain.Contains('\\'))
            throw new InvalidOperationException("SharePoint:Identity:Domain must be a domain name only, without '\\'.");
    }
}

/// <summary>
/// Optional base folder for <see cref="Abstractions.ISharePointServerRelativePathScope"/>.
/// </summary>
public sealed class SharePointStorageScopeOptions
{
    /// <summary>Server-relative base folder, e.g. <c>/sites/team/Shared Documents</c>.</summary>
    public string? BaseFolderServerRelativeUrl { get; set; }

    /// <exception cref="Abstractions.SharePointValidationException">The base folder is not a valid server-relative path.</exception>
    public void Validate()
    {
        if (string.IsNullOrWhiteSpace(BaseFolderServerRelativeUrl))
            return;

        ServerRelativeUrl.Validate(BaseFolderServerRelativeUrl);
    }
}
