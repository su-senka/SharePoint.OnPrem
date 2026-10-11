namespace SharePoint.OnPrem.Abstractions;

/// <summary>
/// Site-level information, mainly for connectivity and authentication diagnostics.
/// </summary>
public interface ISharePointSiteClient
{
    /// <summary>Reads the configured web's title, URLs, language, and the SharePoint server build.</summary>
    Task<SharePointWebInfo> GetWebInfoAsync(CancellationToken ct = default);

    /// <summary>Returns the user SharePoint authenticated the requests as. Answers "which account am I really using?".</summary>
    Task<SharePointUserInfo> GetCurrentUserAsync(CancellationToken ct = default);
}

/// <summary>Basic information about a SharePoint web.</summary>
/// <param name="Title">Web title.</param>
/// <param name="Url">Absolute web URL.</param>
/// <param name="ServerRelativeUrl">Server-relative web URL, e.g. <c>/sites/team</c>.</param>
/// <param name="Language">Web locale id (LCID), e.g. 1033 for English.</param>
/// <param name="ServerVersion">Build from the <c>MicrosoftSharePointTeamServices</c> response header, e.g. <c>16.0.0.10337</c>.</param>
public sealed record SharePointWebInfo(
    string Title,
    string Url,
    string ServerRelativeUrl,
    int Language,
    string? ServerVersion)
{
    /// <summary>
    /// Best-effort product name derived from <see cref="ServerVersion"/>: SharePoint 2013, 2016, 2019, or Subscription Edition.
    /// </summary>
    public string? ProductName => SharePointProductVersion.GetProductName(ServerVersion);
}

/// <summary>A SharePoint site user.</summary>
/// <param name="Id">Site user id.</param>
/// <param name="LoginName">Login name, e.g. <c>i:0#.w|contoso\svc-sharepoint</c>.</param>
/// <param name="Title">Display name.</param>
/// <param name="Email">E-mail address, if SharePoint knows it.</param>
/// <param name="IsSiteAdmin">Whether the user is a site collection administrator.</param>
public sealed record SharePointUserInfo(
    int Id,
    string LoginName,
    string Title,
    string? Email,
    bool IsSiteAdmin);

/// <summary>Maps SharePoint server build numbers to product names.</summary>
public static class SharePointProductVersion
{
    /// <summary>
    /// Returns the product name for a build such as <c>16.0.0.10337</c>, or <see langword="null"/> when it cannot be determined.
    /// </summary>
    public static string? GetProductName(string? serverVersion)
    {
        if (!Version.TryParse(serverVersion, out var version))
            return null;

        // 2016, 2019 and Subscription Edition all report major version 16; they differ by build number
        // (16.0.4xxx-5xxx, 16.0.10xxx, 16.0.14xxx and later). The build is in the last field ("16.0.0.10337").
        var build = version.Revision > 0 ? version.Revision : version.Build;
        return version.Major switch
        {
            15 => "SharePoint 2013",
            16 when build >= 14000 => "SharePoint Server Subscription Edition",
            16 when build >= 10000 => "SharePoint 2019",
            16 => "SharePoint 2016",
            _ => null
        };
    }
}
