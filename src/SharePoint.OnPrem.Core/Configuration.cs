namespace SharePoint.OnPrem.Core;

public sealed class SharePointOnPremOptions
{
    public string SiteBaseUrl { get; set; } = string.Empty;
    public int HttpTimeoutMinutes { get; set; } = 5;
    public bool UseFormDigestCaching { get; set; } = true;

    public void Validate()
    {
        if (string.IsNullOrWhiteSpace(SiteBaseUrl))
            throw new InvalidOperationException("SharePoint:SiteBaseUrl configuration is required.");

        if (!Uri.TryCreate(SiteBaseUrl, UriKind.Absolute, out _))
            throw new InvalidOperationException("SharePoint:SiteBaseUrl must be an absolute URI.");

        if (HttpTimeoutMinutes < 1)
            throw new InvalidOperationException("SharePoint:HttpTimeoutMinutes must be greater than or equal to 1.");
    }
}

public sealed class SharePointIdentityOptions
{
    public string Domain { get; set; } = string.Empty;
    public string ClaimsPrefix { get; set; } = string.Empty;

    public void Validate()
    {
        if (string.IsNullOrWhiteSpace(Domain))
            throw new InvalidOperationException("SharePoint:Identity:Domain configuration is required.");

        if (string.IsNullOrWhiteSpace(ClaimsPrefix))
            throw new InvalidOperationException("SharePoint:Identity:ClaimsPrefix configuration is required.");
    }
}

public sealed class SharePointStorageScopeOptions
{
    public string? BaseFolderServerRelativeUrl { get; set; }

    public void Validate()
    {
        if (string.IsNullOrWhiteSpace(BaseFolderServerRelativeUrl))
            return;

        ServerRelativeUrl.Validate(BaseFolderServerRelativeUrl);
    }
}

