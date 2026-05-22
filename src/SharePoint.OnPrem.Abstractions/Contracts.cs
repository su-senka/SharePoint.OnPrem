using System.Text.Json;

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
public interface ISharePointServerRelativePathScope
{
    string ToServerRelativePath(string relativePath);
}

[Obsolete("Use ISharePointServerRelativePathScope instead.")]
public interface ISharePointPathScope : ISharePointServerRelativePathScope;

public interface ISharePointFileClient
{
    Task<SharePointStoredFile> UploadAsync(UploadFileRequest request, CancellationToken ct = default);
    Task<SharePointFileDownload> DownloadAsync(string serverRelativeUrl, CancellationToken ct = default);
    Task DeleteAsync(string serverRelativeUrl, CancellationToken ct = default);
    Task<string> GetFileWebUrlAsync(string serverRelativeUrl, CancellationToken ct = default);
    Task<SharePointStoredFile> CopyAsync(
        string sourceServerRelativeUrl,
        string destinationFolderServerRelativeUrl,
        string destinationFileName,
        bool overwrite = true,
        CancellationToken ct = default);
    Task<SharePointStoredFile> MoveAsync(
        string sourceServerRelativeUrl,
        string destinationFolderServerRelativeUrl,
        string destinationFileName,
        bool overwrite = true,
        CancellationToken ct = default);
    Task<SharePointStoredFile> RenameAsync(
        string serverRelativeUrl,
        string newFileName,
        bool overwrite = true,
        CancellationToken ct = default);
    Task UpdateMetadataAsync(
        string serverRelativeUrl,
        IReadOnlyDictionary<string, object?> fields,
        CancellationToken ct = default);
    Task<IReadOnlyDictionary<string, JsonElement>> GetMetadataAsync(
        string serverRelativeUrl,
        IEnumerable<string>? selectFields = null,
        CancellationToken ct = default);
    Task<IReadOnlyDictionary<string, JsonElement>?> TryGetMetadataAsync(
        string serverRelativeUrl,
        IEnumerable<string>? selectFields = null,
        CancellationToken ct = default);
    Task<T?> GetMetadataValueAsync<T>(
        string serverRelativeUrl,
        string fieldName,
        CancellationToken ct = default);
    Task<T?> TryGetMetadataValueAsync<T>(
        string serverRelativeUrl,
        string fieldName,
        CancellationToken ct = default);

    [Obsolete("Use GetFileWebUrlAsync instead.")]
    Task<string> GetWebUrlAsync(string serverRelativeUrl, CancellationToken ct = default);
}

public interface ISharePointFolderClient
{
    Task EnsureFolderPathAsync(string serverRelativeFolder, CancellationToken ct = default);

    [Obsolete("Use EnsureFolderPathAsync instead.")]
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
    Task<IReadOnlyList<SharePointPrincipalRoleAssignment>> GetFolderRoleAssignmentsAsync(string serverRelativeFolder, CancellationToken ct = default);
    Task BindRoleToFileAsync(string serverRelativeFile, string principalName, string roleName, CancellationToken ct = default);
    Task RemoveRoleFromFileAsync(string serverRelativeFile, string principalName, CancellationToken ct = default);
    Task<IReadOnlyList<SharePointPrincipalRoleAssignment>> GetFileRoleAssignmentsAsync(string serverRelativeFile, CancellationToken ct = default);
}

public sealed record SharePointRoleDefinition(
    int Id,
    string Name);

public sealed record SharePointPrincipalRoleAssignment(
    int PrincipalId,
    string? PrincipalTitle,
    string? PrincipalLoginName,
    int? PrincipalType,
    IReadOnlyList<SharePointRoleDefinition> Roles);

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

