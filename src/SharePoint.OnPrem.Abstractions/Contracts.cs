namespace SharePoint.OnPrem.Abstractions;

/// <summary>
/// Normalizes consumer-supplied principals into the login format expected by SharePoint on-prem.
/// </summary>
public interface ILoginNameNormalizer
{
    string Normalize(string input);
}

/// <summary>
/// Optional path-scoping abstraction for consumers that want to work relative to a fixed SharePoint base folder.
/// </summary>
public interface ISharePointPathScope
{
    string ToServerRelativePath(string relativePath);
}

public interface ISharePointFileClient
{
    Task<SharePointStoredFile> UploadAsync(UploadFileRequest request, CancellationToken ct = default);
    Task<SharePointFileDownload> DownloadAsync(string serverRelativeUrl, CancellationToken ct = default);
    Task DeleteAsync(string serverRelativeUrl, CancellationToken ct = default);
    Task<string> GetWebUrlAsync(string serverRelativeUrl, CancellationToken ct = default);
}

public interface ISharePointFolderClient
{
    Task EnsurePathAsync(string serverRelativeFolder, CancellationToken ct = default);
    Task<bool> ExistsAsync(string serverRelativeFolder, CancellationToken ct = default);
    Task CreateAsync(string serverRelativeFolder, CancellationToken ct = default);
}

public interface ISharePointSecurityClient
{
    Task BreakInheritanceAsync(string serverRelativeFolder, bool copyRoleAssignments, CancellationToken ct = default);
    Task ResetInheritanceAsync(string serverRelativeFolder, CancellationToken ct = default);

    Task<int> EnsureGroupAsync(string groupName, string? description = null, CancellationToken ct = default);
    Task DeleteGroupAsync(string groupName, CancellationToken ct = default);

    Task<int> EnsureUserAsync(string loginName, CancellationToken ct = default);
    Task<IReadOnlyList<string>> GetGroupMembersAsync(string groupName, CancellationToken ct = default);
    Task AddUsersToGroupAsync(string groupName, IEnumerable<string> loginNames, CancellationToken ct = default);
    Task RemoveUsersFromGroupAsync(string groupName, IEnumerable<string> loginNames, CancellationToken ct = default);

    Task BindRoleToFolderAsync(string serverRelativeFolder, string principalName, string roleName, CancellationToken ct = default);
    Task RemoveRoleFromFolderAsync(string serverRelativeFolder, string principalName, CancellationToken ct = default);
    Task SyncGroupMembershipAsync(string groupName, IEnumerable<string> desiredLoginNames, CancellationToken ct = default);
}

public sealed record UploadFileRequest(
    string FolderServerRelativeUrl,
    string FileName,
    Stream Content,
    string? ContentType = null,
    bool Overwrite = true);

public sealed record SharePointStoredFile(
    string ServerRelativeUrl,
    string FileName);

public sealed record SharePointFileDownload(
    byte[] Content,
    string ContentType,
    string FileName);

