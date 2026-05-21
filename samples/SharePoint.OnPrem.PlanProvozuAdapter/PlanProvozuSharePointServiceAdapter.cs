using PP.Application.Services.SharePoint;
using SharePoint.OnPrem.Abstractions;
using SharePoint.OnPrem.Core;

namespace SharePoint.OnPrem.PlanProvozuAdapter;

/// <summary>
/// Compatibility adapter that implements PlanProvozu's legacy <see cref="ISharePointService"/>
/// on top of the standalone <c>SharePoint.OnPrem</c> package.
/// This adapter is intended for side-by-side validation before swapping production code.
/// </summary>
public sealed class PlanProvozuSharePointServiceAdapter(
    HttpClient httpClient,
    ISharePointFileClient fileClient,
    ISharePointFolderClient folderClient,
    ISharePointSecurityClient securityClient,
    ISharePointPathScope pathScope,
    ISharePointRequestExecutor requestExecutor)
    : ISharePointService
{
    private const int ReadRoleDefinitionId = 1073741826;
    private const int EditRoleDefinitionId = 1073741827;

    private readonly HttpClient _httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
    private readonly ISharePointFileClient _fileClient = fileClient ?? throw new ArgumentNullException(nameof(fileClient));
    private readonly ISharePointFolderClient _folderClient = folderClient ?? throw new ArgumentNullException(nameof(folderClient));
    private readonly ISharePointSecurityClient _securityClient = securityClient ?? throw new ArgumentNullException(nameof(securityClient));
    private readonly ISharePointPathScope _pathScope = pathScope ?? throw new ArgumentNullException(nameof(pathScope));
    private readonly ISharePointRequestExecutor _requestExecutor = requestExecutor ?? throw new ArgumentNullException(nameof(requestExecutor));

    public async Task<string> UploadFileAsync(string folder, string fileName, Stream stream, CancellationToken cancellationToken = default)
    {
        var folderServerRelativeUrl = _pathScope.ToServerRelativePath(folder);
        var file = await _fileClient.UploadAsync(
            new UploadFileRequest(folderServerRelativeUrl, ValidateLeafFileName(fileName), stream, Overwrite: true),
            cancellationToken);

        return file.ServerRelativeUrl;
    }

    public Task EnsureFolderPathAsync(string relativeFolder, CancellationToken cancellationToken)
        => _folderClient.EnsurePathAsync(_pathScope.ToServerRelativePath(relativeFolder), cancellationToken);

    public async Task<byte[]> DownloadFileAsync(string storageIdentifier, CancellationToken cancellationToken = default)
    {
        var file = await _fileClient.DownloadAsync(storageIdentifier, cancellationToken);
        return file.Content;
    }

    public Task DeleteFileAsync(string fileName, CancellationToken cancellationToken = default)
        => _fileClient.DeleteAsync(_pathScope.ToServerRelativePath(ValidateLeafFileName(fileName)), cancellationToken);

    public async Task GrantFileEditAsync(string fileName, string usernameOrLogin, CancellationToken cancellationToken = default)
    {
        var principalId = await _securityClient.EnsureUserAsync(usernameOrLogin, cancellationToken);
        var serverRelativeFile = _pathScope.ToServerRelativePath(ValidateLeafFileName(fileName));

        await BreakFileInheritanceAsync(serverRelativeFile, copyRoleAssignments: false, cancellationToken);
        await AssignRoleToFileAsync(serverRelativeFile, principalId, EditRoleDefinitionId, cancellationToken);
    }

    public async Task RemoveFileEditAsync(string fileName, string usernameOrLogin, CancellationToken cancellationToken = default)
    {
        var principalId = await _securityClient.EnsureUserAsync(usernameOrLogin, cancellationToken);
        var serverRelativeFile = _pathScope.ToServerRelativePath(ValidateLeafFileName(fileName));

        await BreakFileInheritanceAsync(serverRelativeFile, copyRoleAssignments: false, cancellationToken);
        await RemoveRoleFromFileAsync(serverRelativeFile, principalId, EditRoleDefinitionId, cancellationToken);
    }

    public async Task GrantFileReadAsync(string fileName, string usernameOrLogin, CancellationToken cancellationToken = default)
    {
        var principalId = await _securityClient.EnsureUserAsync(usernameOrLogin, cancellationToken);
        var serverRelativeFile = _pathScope.ToServerRelativePath(ValidateLeafFileName(fileName));

        await BreakFileInheritanceAsync(serverRelativeFile, copyRoleAssignments: false, cancellationToken);
        await AssignRoleToFileAsync(serverRelativeFile, principalId, ReadRoleDefinitionId, cancellationToken);
    }

    public async Task RemoveFileReadAsync(string fileName, string usernameOrLogin, CancellationToken cancellationToken = default)
    {
        var principalId = await _securityClient.EnsureUserAsync(usernameOrLogin, cancellationToken);
        var serverRelativeFile = _pathScope.ToServerRelativePath(ValidateLeafFileName(fileName));

        await BreakFileInheritanceAsync(serverRelativeFile, copyRoleAssignments: false, cancellationToken);
        await RemoveRoleFromFileAsync(serverRelativeFile, principalId, ReadRoleDefinitionId, cancellationToken);
    }

    public Task BreakInheritanceAsync(string folderServerRelativeUrl, bool copyRoleAssignments, CancellationToken cancellationToken = default)
        => _securityClient.BreakInheritanceAsync(folderServerRelativeUrl, copyRoleAssignments, cancellationToken);

    public Task ResetInheritanceAsync(string folderServerRelativeUrl, CancellationToken cancellationToken = default)
        => _securityClient.ResetInheritanceAsync(folderServerRelativeUrl, cancellationToken);

    public Task GrantFolderReadAsync(string folderServerRelativeUrl, string usernameOrLogin, CancellationToken cancellationToken = default)
        => _securityClient.BindRoleToFolderAsync(folderServerRelativeUrl, usernameOrLogin, "Read", cancellationToken);

    public Task RemoveFolderReadAsync(string folderServerRelativeUrl, string usernameOrLogin, CancellationToken cancellationToken = default)
        => _securityClient.RemoveRoleFromFolderAsync(folderServerRelativeUrl, usernameOrLogin, cancellationToken);

    public Task<int> EnsureGroupAsync(string groupName, string description, CancellationToken cancellationToken = default)
        => _securityClient.EnsureGroupAsync(groupName, description, cancellationToken);

    public Task DeleteGroupAsync(string groupName, CancellationToken cancellationToken = default)
        => _securityClient.DeleteGroupAsync(groupName, cancellationToken);

    public Task<int> EnsureUserAsync(string loginOrUpn, CancellationToken cancellationToken = default)
        => _securityClient.EnsureUserAsync(loginOrUpn, cancellationToken);

    public Task<IReadOnlyList<string>> GetGroupMembersAsync(string groupName, CancellationToken cancellationToken = default)
        => _securityClient.GetGroupMembersAsync(groupName, cancellationToken);

    public Task AddUsersToGroupAsync(string groupName, IEnumerable<string> userLogins, CancellationToken cancellationToken = default)
        => _securityClient.AddUsersToGroupAsync(groupName, userLogins, cancellationToken);

    public Task RemoveUsersFromGroupAsync(string groupName, IEnumerable<string> userLogins, CancellationToken cancellationToken = default)
        => _securityClient.RemoveUsersFromGroupAsync(groupName, userLogins, cancellationToken);

    public Task BindRoleToFolderAsync(string folderServerRelativeUrl, string groupName, string roleName, CancellationToken cancellationToken = default)
        => _securityClient.BindRoleToFolderAsync(folderServerRelativeUrl, groupName, roleName, cancellationToken);

    public Task RemoveRoleFromFolderAsync(string folderServerRelativeUrl, string groupName, CancellationToken cancellationToken = default)
        => _securityClient.RemoveRoleFromFolderAsync(folderServerRelativeUrl, groupName, cancellationToken);

    public Task<string> GetFileWebUrlAsync(string serverRelativeUrl, CancellationToken cancellationToken = default)
        => _fileClient.GetWebUrlAsync(serverRelativeUrl, cancellationToken);

    public async Task EnsureRootSecurityAsync(string rootFolderUrl, string readersGroup, string ownersGroup, CancellationToken cancellationToken = default)
    {
        await _securityClient.EnsureGroupAsync(readersGroup, $"{readersGroup} – read access", cancellationToken);
        await _securityClient.EnsureGroupAsync(ownersGroup, $"{ownersGroup} – edit access", cancellationToken);
        await _securityClient.BreakInheritanceAsync(rootFolderUrl, copyRoleAssignments: false, cancellationToken);
        await _securityClient.BindRoleToFolderAsync(rootFolderUrl, readersGroup, "Read", cancellationToken);
        await _securityClient.BindRoleToFolderAsync(rootFolderUrl, ownersGroup, "Edit", cancellationToken);
    }

    public async Task EnsureAssignmentSecurityAsync(string assignmentFolderUrl, string writersGroup, CancellationToken cancellationToken = default)
    {
        await _securityClient.EnsureGroupAsync(writersGroup, $"{writersGroup} – assignment writers", cancellationToken);
        await _securityClient.BreakInheritanceAsync(assignmentFolderUrl, copyRoleAssignments: true, cancellationToken);
        await _securityClient.BindRoleToFolderAsync(assignmentFolderUrl, writersGroup, "Edit", cancellationToken);
    }

    public Task SyncGroupMembershipAsync(string groupName, IEnumerable<string> desiredUserLogins, CancellationToken cancellationToken = default)
        => _securityClient.SyncGroupMembershipAsync(groupName, desiredUserLogins, cancellationToken);

    private async Task BreakFileInheritanceAsync(string serverRelativeFileUrl, bool copyRoleAssignments, CancellationToken cancellationToken)
    {
        var encodedUrl = Uri.EscapeDataString(ServerRelativeUrl.Validate(serverRelativeFileUrl));
        using var request = new HttpRequestMessage(
            HttpMethod.Post,
            $"_api/web/GetFileByServerRelativeUrl('{encodedUrl}')/ListItemAllFields/breakroleinheritance(copyRoleAssignments={copyRoleAssignments.ToString().ToLowerInvariant()},clearSubscopes=true)");

        using var response = await _requestExecutor.SendAsync(
            _httpClient,
            request,
            new SharePointSendOptions { IncludeFormDigest = true },
            cancellationToken);
    }

    private async Task AssignRoleToFileAsync(string serverRelativeFileUrl, int principalId, int roleDefinitionId, CancellationToken cancellationToken)
    {
        var encodedUrl = Uri.EscapeDataString(ServerRelativeUrl.Validate(serverRelativeFileUrl));
        using var request = new HttpRequestMessage(
            HttpMethod.Post,
            $"_api/web/GetFileByServerRelativeUrl('{encodedUrl}')/ListItemAllFields/roleassignments/addroleassignment(principalid={principalId},roledefid={roleDefinitionId})");

        using var response = await _requestExecutor.SendAsync(
            _httpClient,
            request,
            new SharePointSendOptions { IncludeFormDigest = true },
            cancellationToken);
    }

    private async Task RemoveRoleFromFileAsync(string serverRelativeFileUrl, int principalId, int roleDefinitionId, CancellationToken cancellationToken)
    {
        var encodedUrl = Uri.EscapeDataString(ServerRelativeUrl.Validate(serverRelativeFileUrl));
        using var request = new HttpRequestMessage(
            HttpMethod.Post,
            $"_api/web/GetFileByServerRelativeUrl('{encodedUrl}')/ListItemAllFields/roleassignments/removeroleassignment(principalid={principalId},roledefid={roleDefinitionId})");

        using var response = await _requestExecutor.SendAsync(
            _httpClient,
            request,
            new SharePointSendOptions { IncludeFormDigest = true },
            cancellationToken);
    }

    private static string ValidateLeafFileName(string fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName))
            throw new SharePointValidationException("File name must be provided.");

        var trimmed = fileName.Trim();
        if (trimmed.Contains('/') || trimmed.Contains('\\'))
            throw new SharePointValidationException("File name must be a leaf name without path separators.");

        return trimmed;
    }
}

