using System.Text.Json;

namespace SharePoint.OnPrem.Abstractions;

/// <summary>
/// Normalizes consumer-supplied principals into the login format expected by SharePoint on-prem.
/// </summary>
public interface ILoginNameNormalizer
{
    /// <summary>
    /// Converts a login such as <c>jdoe</c> or <c>CONTOSO\jdoe</c> into the form SharePoint stores, e.g. <c>i:0#.w|CONTOSO\jdoe</c>.
    /// </summary>
    string Normalize(string input);
}

/// <summary>
/// Optional path-scoping abstraction for consumers that want to work relative to a fixed SharePoint base folder.
/// </summary>
public interface ISharePointServerRelativePathScope
{
    /// <summary>
    /// Combines the configured base folder with <paramref name="relativePath"/>. Rejects '.' and '..' segments,
    /// so the result always stays inside the base folder.
    /// </summary>
    string ToServerRelativePath(string relativePath);
}

[Obsolete("Use ISharePointServerRelativePathScope instead.")]
public interface ISharePointPathScope : ISharePointServerRelativePathScope;

/// <summary>
/// File operations against a SharePoint site. All paths are server-relative and not URL-encoded,
/// e.g. <c>/sites/team/Shared Documents/report.docx</c>.
/// </summary>
public interface ISharePointFileClient
{
    /// <summary>Uploads a file into an existing folder. The caller keeps ownership of <see cref="UploadFileRequest.Content"/>.</summary>
    Task<SharePointStoredFile> UploadAsync(UploadFileRequest request, CancellationToken ct = default);

    /// <summary>
    /// Uploads a large file in chunks (<c>StartUpload</c> / <c>ContinueUpload</c> / <c>FinishUpload</c>), so it is not
    /// limited by the size of a single request. Requires SharePoint 2016 or later; files that fit in one chunk use a
    /// plain upload and work everywhere. If a chunk fails, the upload session is cancelled; the empty file created
    /// at the start of the session remains.
    /// </summary>
    /// <param name="request">The file to upload. <see cref="UploadFileRequest.Content"/> may be non-seekable and is not disposed.</param>
    /// <param name="chunkSizeBytes">Bytes per request. Default: 10 MiB. Two buffers of this size are allocated.</param>
    /// <param name="progress">Receives the total number of bytes uploaded after each chunk.</param>
    /// <param name="ct">Cancellation token.</param>
    Task<SharePointStoredFile> UploadChunkedAsync(
        UploadFileRequest request,
        int chunkSizeBytes = 10 * 1024 * 1024,
        IProgress<long>? progress = null,
        CancellationToken ct = default);

    /// <summary>Streams a file's content. Dispose the result to release the HTTP connection.</summary>
    /// <exception cref="SharePointNotFoundException">The file does not exist.</exception>
    Task<SharePointFileDownload> DownloadAsync(string serverRelativeUrl, CancellationToken ct = default);

    /// <summary>Returns whether a file exists at the given path.</summary>
    Task<bool> ExistsAsync(string serverRelativeUrl, CancellationToken ct = default);

    /// <summary>Deletes a file (it goes to the site recycle bin according to SharePoint settings).</summary>
    Task DeleteAsync(string serverRelativeUrl, CancellationToken ct = default);

    /// <summary>Returns a browser URL for the file: SharePoint's <c>LinkingUrl</c> when available, otherwise the absolute file URL.</summary>
    Task<string> GetFileWebUrlAsync(string serverRelativeUrl, CancellationToken ct = default);

    /// <summary>Copies a file within the same site collection.</summary>
    Task<SharePointStoredFile> CopyAsync(
        string sourceServerRelativeUrl,
        string destinationFolderServerRelativeUrl,
        string destinationFileName,
        bool overwrite = true,
        CancellationToken ct = default);

    /// <summary>Moves a file within the same site collection.</summary>
    Task<SharePointStoredFile> MoveAsync(
        string sourceServerRelativeUrl,
        string destinationFolderServerRelativeUrl,
        string destinationFileName,
        bool overwrite = true,
        CancellationToken ct = default);

    /// <summary>Renames a file in place.</summary>
    Task<SharePointStoredFile> RenameAsync(
        string serverRelativeUrl,
        string newFileName,
        bool overwrite = true,
        CancellationToken ct = default);

    /// <summary>
    /// Updates list item fields of the file. Keys are field internal names (e.g. <c>Title</c>, not the display name).
    /// </summary>
    Task UpdateMetadataAsync(
        string serverRelativeUrl,
        IReadOnlyDictionary<string, object?> fields,
        CancellationToken ct = default);

    /// <summary>
    /// Reads list item fields of the file. Pass <paramref name="selectFields"/> (internal names) to limit the result.
    /// </summary>
    /// <exception cref="SharePointNotFoundException">The file does not exist.</exception>
    Task<IReadOnlyDictionary<string, JsonElement>> GetMetadataAsync(
        string serverRelativeUrl,
        IEnumerable<string>? selectFields = null,
        CancellationToken ct = default);

    /// <summary>Same as <see cref="GetMetadataAsync"/>, but returns <see langword="null"/> when the file does not exist.</summary>
    Task<IReadOnlyDictionary<string, JsonElement>?> TryGetMetadataAsync(
        string serverRelativeUrl,
        IEnumerable<string>? selectFields = null,
        CancellationToken ct = default);

    /// <summary>Reads one field and deserializes it to <typeparamref name="T"/>; returns default when the field is missing or null.</summary>
    /// <exception cref="SharePointNotFoundException">The file does not exist.</exception>
    Task<T?> GetMetadataValueAsync<T>(
        string serverRelativeUrl,
        string fieldName,
        CancellationToken ct = default);

    /// <summary>Same as <see cref="GetMetadataValueAsync{T}"/>, but returns default when the file does not exist.</summary>
    Task<T?> TryGetMetadataValueAsync<T>(
        string serverRelativeUrl,
        string fieldName,
        CancellationToken ct = default);

    /// <summary>Checks the file out to the current user.</summary>
    Task CheckOutAsync(string serverRelativeUrl, CancellationToken ct = default);

    /// <summary>Checks a checked-out file back in, creating a new version.</summary>
    Task CheckInAsync(
        string serverRelativeUrl,
        string? comment = null,
        SharePointCheckInType checkInType = SharePointCheckInType.MinorCheckIn,
        CancellationToken ct = default);

    /// <summary>Discards the check-out and any changes made while the file was checked out.</summary>
    Task UndoCheckOutAsync(string serverRelativeUrl, CancellationToken ct = default);

    /// <summary>
    /// Lists the file's previous versions, oldest first. The current version is not included; read it with
    /// <see cref="GetMetadataAsync"/>. Returns an empty list when versioning is disabled.
    /// </summary>
    Task<IReadOnlyList<SharePointFileVersion>> GetVersionsAsync(string serverRelativeUrl, CancellationToken ct = default);

    [Obsolete("Use GetFileWebUrlAsync instead.")]
    Task<string> GetWebUrlAsync(string serverRelativeUrl, CancellationToken ct = default);
}

/// <summary>How <see cref="ISharePointFileClient.CheckInAsync"/> records the new version (SP.CheckinType).</summary>
public enum SharePointCheckInType
{
    /// <summary>Creates a minor (draft) version, e.g. 1.1.</summary>
    MinorCheckIn = 0,

    /// <summary>Creates a major (published) version, e.g. 2.0.</summary>
    MajorCheckIn = 1,

    /// <summary>Replaces the current version without creating a new one.</summary>
    OverwriteCheckIn = 2
}

/// <summary>A previous version of a file.</summary>
/// <param name="Id">Version id (512 × major + minor, e.g. 513 for 1.1).</param>
/// <param name="VersionLabel">Version label, e.g. <c>1.1</c>.</param>
/// <param name="Created">When the version was created (UTC).</param>
/// <param name="Size">Size in bytes.</param>
/// <param name="CheckInComment">Check-in comment, if any.</param>
/// <param name="IsCurrentVersion">Whether this is the current version.</param>
public sealed record SharePointFileVersion(
    int Id,
    string VersionLabel,
    DateTimeOffset? Created,
    long Size,
    string? CheckInComment,
    bool IsCurrentVersion);

/// <summary>
/// Folder operations against a SharePoint site. All paths are server-relative and not URL-encoded.
/// </summary>
public interface ISharePointFolderClient
{
    /// <summary>
    /// Creates the folder and any missing parent folders. Does nothing when the folder already exists.
    /// The path must be inside the configured site, e.g. <c>/sites/team/Shared Documents/2026/Q1</c>.
    /// </summary>
    Task EnsureFolderPathAsync(string serverRelativeFolder, CancellationToken ct = default);

    [Obsolete("Use EnsureFolderPathAsync instead.")]
    Task EnsurePathAsync(string serverRelativeFolder, CancellationToken ct = default);

    /// <summary>Returns whether a folder exists at the given path.</summary>
    Task<bool> ExistsAsync(string serverRelativeFolder, CancellationToken ct = default);

    /// <summary>Creates a single folder whose parent already exists. Does nothing when the folder already exists.</summary>
    Task CreateAsync(string serverRelativeFolder, CancellationToken ct = default);

    /// <summary>Deletes a folder and its contents. Does nothing when the folder does not exist.</summary>
    Task DeleteAsync(string serverRelativeFolder, CancellationToken ct = default);

    /// <summary>Lists the files directly inside a folder (not recursive).</summary>
    /// <exception cref="SharePointNotFoundException">The folder does not exist.</exception>
    Task<IReadOnlyList<SharePointFileInfo>> GetFilesAsync(string serverRelativeFolder, CancellationToken ct = default);

    /// <summary>
    /// Lists the folders directly inside a folder (not recursive). At a library root this includes SharePoint's hidden <c>Forms</c> folder.
    /// </summary>
    /// <exception cref="SharePointNotFoundException">The folder does not exist.</exception>
    Task<IReadOnlyList<SharePointFolderInfo>> GetFoldersAsync(string serverRelativeFolder, CancellationToken ct = default);
}

/// <summary>
/// Site groups, users, and item-level permissions.
/// </summary>
/// <remarks>
/// Principal names passed to role methods are resolved as a site group first. If no group matches and the name contains
/// '\', '@' or '|', it is treated as a user login; otherwise a site group with that name is created.
/// Role names are resolved by SharePoint role type for <c>Read</c>, <c>Contribute</c>, <c>Edit</c>, <c>Design</c> and
/// <c>Full Control</c> (works on localized servers), and by role definition name for anything else.
/// </remarks>
public interface ISharePointSecurityClient
{
    /// <summary>Gives the folder its own permissions. Optionally copies the inherited role assignments first.</summary>
    Task BreakInheritanceAsync(string serverRelativeFolder, bool copyRoleAssignments, CancellationToken ct = default);

    /// <summary>Makes the folder inherit permissions from its parent again, discarding unique role assignments.</summary>
    Task ResetInheritanceAsync(string serverRelativeFolder, CancellationToken ct = default);

    /// <summary>Gives the file its own permissions. Optionally copies the inherited role assignments first.</summary>
    Task BreakFileInheritanceAsync(string serverRelativeFile, bool copyRoleAssignments, CancellationToken ct = default);

    /// <summary>Makes the file inherit permissions from its parent again, discarding unique role assignments.</summary>
    Task ResetFileInheritanceAsync(string serverRelativeFile, CancellationToken ct = default);

    /// <summary>Returns the Id of the site group, creating it when it does not exist.</summary>
    Task<int> EnsureGroupAsync(string groupName, string? description = null, CancellationToken ct = default);

    /// <summary>Deletes the site group. Does nothing when the group does not exist.</summary>
    Task DeleteGroupAsync(string groupName, CancellationToken ct = default);

    /// <summary>Returns the site user Id for a login, adding the user to the site when needed (<c>ensureuser</c>).</summary>
    Task<int> EnsureUserAsync(string loginName, CancellationToken ct = default);

    /// <summary>Returns the login names of the group's members, or an empty list when the group does not exist.</summary>
    Task<IReadOnlyList<string>> GetGroupMembersAsync(string groupName, CancellationToken ct = default);

    /// <summary>Adds users to the group. Users that are already members are skipped.</summary>
    Task AddUsersToGroupAsync(string groupName, IEnumerable<string> loginNames, CancellationToken ct = default);

    /// <summary>Removes users from the group. Users that are not members are skipped.</summary>
    Task RemoveUsersFromGroupAsync(string groupName, IEnumerable<string> loginNames, CancellationToken ct = default);

    /// <summary>Grants <paramref name="roleName"/> on the folder to a group or user. The folder must have unique permissions.</summary>
    Task BindRoleToFolderAsync(string serverRelativeFolder, string principalName, string roleName, CancellationToken ct = default);

    /// <summary>Removes all of the principal's role assignments from the folder. Does nothing when the principal is unknown.</summary>
    Task RemoveRoleFromFolderAsync(string serverRelativeFolder, string principalName, CancellationToken ct = default);

    /// <summary>Adds and removes members so the group contains exactly <paramref name="desiredLoginNames"/>.</summary>
    Task SyncGroupMembershipAsync(string groupName, IEnumerable<string> desiredLoginNames, CancellationToken ct = default);

    /// <summary>Lists the principals that have permissions on the folder, with their roles.</summary>
    Task<IReadOnlyList<SharePointPrincipalRoleAssignment>> GetFolderRoleAssignmentsAsync(string serverRelativeFolder, CancellationToken ct = default);

    /// <summary>Grants <paramref name="roleName"/> on the file to a group or user. The file must have unique permissions.</summary>
    Task BindRoleToFileAsync(string serverRelativeFile, string principalName, string roleName, CancellationToken ct = default);

    /// <summary>Removes all of the principal's role assignments from the file. Does nothing when the principal is unknown.</summary>
    Task RemoveRoleFromFileAsync(string serverRelativeFile, string principalName, CancellationToken ct = default);

    /// <summary>Lists the principals that have permissions on the file, with their roles.</summary>
    Task<IReadOnlyList<SharePointPrincipalRoleAssignment>> GetFileRoleAssignmentsAsync(string serverRelativeFile, CancellationToken ct = default);
}

/// <summary>A SharePoint role definition (permission level), e.g. Read or Contribute.</summary>
public sealed record SharePointRoleDefinition(
    int Id,
    string Name);

/// <summary>A principal's role assignment on a securable item.</summary>
/// <param name="PrincipalId">Site principal Id.</param>
/// <param name="PrincipalTitle">Display name.</param>
/// <param name="PrincipalLoginName">Login name, e.g. <c>i:0#.w|CONTOSO\jdoe</c>, or the group name.</param>
/// <param name="PrincipalType">SharePoint principal type: 1 = user, 4 = security group, 8 = SharePoint group.</param>
/// <param name="Roles">Role definitions bound to the principal on this item.</param>
public sealed record SharePointPrincipalRoleAssignment(
    int PrincipalId,
    string? PrincipalTitle,
    string? PrincipalLoginName,
    int? PrincipalType,
    IReadOnlyList<SharePointRoleDefinition> Roles);

/// <summary>Arguments for <see cref="ISharePointFileClient.UploadAsync"/>.</summary>
/// <param name="FolderServerRelativeUrl">Existing target folder, e.g. <c>/sites/team/Shared Documents</c>.</param>
/// <param name="FileName">Leaf file name without path separators.</param>
/// <param name="Content">File content. Read from its current position; not disposed by the upload.</param>
/// <param name="ContentType">Optional MIME type sent with the upload.</param>
/// <param name="Overwrite">Replace an existing file with the same name.</param>
public sealed record UploadFileRequest(
    string FolderServerRelativeUrl,
    string FileName,
    Stream Content,
    string? ContentType = null,
    bool Overwrite = true);

/// <summary>A file's location after an upload, copy, move, or rename.</summary>
public sealed record SharePointStoredFile(
    string ServerRelativeUrl,
    string FileName);

/// <summary>A file returned by <see cref="ISharePointFolderClient.GetFilesAsync"/>.</summary>
/// <param name="Name">Leaf file name.</param>
/// <param name="ServerRelativeUrl">Server-relative URL of the file.</param>
/// <param name="Length">Size in bytes.</param>
/// <param name="TimeCreated">Creation time (UTC).</param>
/// <param name="TimeLastModified">Last modification time (UTC).</param>
public sealed record SharePointFileInfo(
    string Name,
    string ServerRelativeUrl,
    long Length,
    DateTimeOffset? TimeCreated,
    DateTimeOffset? TimeLastModified);

/// <summary>A folder returned by <see cref="ISharePointFolderClient.GetFoldersAsync"/>.</summary>
/// <param name="Name">Leaf folder name.</param>
/// <param name="ServerRelativeUrl">Server-relative URL of the folder.</param>
/// <param name="ItemCount">Number of files and folders directly inside the folder.</param>
public sealed record SharePointFolderInfo(
    string Name,
    string ServerRelativeUrl,
    int ItemCount);

/// <summary>
/// Represents a downloaded file. Dispose to release the underlying network stream and HTTP response.
/// </summary>
public sealed class SharePointFileDownload : IDisposable, IAsyncDisposable
{
    /// <summary>The response body, streamed from the server. Read it before disposing this object.</summary>
    public Stream Content { get; }

    /// <summary>The response media type, or <c>application/octet-stream</c> when SharePoint does not send one.</summary>
    public string ContentType { get; }

    /// <summary>The leaf file name taken from the requested path.</summary>
    public string FileName { get; }

    public SharePointFileDownload(Stream content, string contentType, string fileName)
    {
        Content = content;
        ContentType = contentType;
        FileName = fileName;
    }

    public void Dispose() => Content.Dispose();
    public ValueTask DisposeAsync() => Content.DisposeAsync();
}
