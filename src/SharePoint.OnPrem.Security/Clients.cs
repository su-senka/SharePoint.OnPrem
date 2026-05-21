using SharePoint.OnPrem.Abstractions;
using SharePoint.OnPrem.Core;
using System.Collections.Concurrent;
using System.Net;
using System.Text.Json;

namespace SharePoint.OnPrem.Security;

public sealed class SharePointSecurityClient(
    HttpClient httpClient,
    ILoginNameNormalizer loginNameNormalizer,
    ISharePointRequestExecutor requestExecutor) : ISharePointSecurityClient
{
    private readonly HttpClient _httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
    private readonly ILoginNameNormalizer _loginNameNormalizer = loginNameNormalizer ?? throw new ArgumentNullException(nameof(loginNameNormalizer));
    private readonly ISharePointRequestExecutor _requestExecutor = requestExecutor ?? throw new ArgumentNullException(nameof(requestExecutor));

    private static readonly ConcurrentDictionary<string, int> RoleDefinitionIdCache = new(StringComparer.OrdinalIgnoreCase);

    private static readonly Dictionary<string, int> RoleNameToType = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Read"] = 2,
        ["Číst"] = 2,
        ["Edit"] = 6,
        ["Contribute"] = 3
    };

    public async Task BreakInheritanceAsync(string serverRelativeFolder, bool copyRoleAssignments, CancellationToken ct = default)
    {
        var validatedFolder = ServerRelativeUrl.Validate(serverRelativeFolder);
        var encodedFolder = Uri.EscapeDataString(validatedFolder);

        using var message = new HttpRequestMessage(
            HttpMethod.Post,
            $"_api/web/GetFolderByServerRelativeUrl('{encodedFolder}')/ListItemAllFields/breakroleinheritance(copyRoleAssignments={copyRoleAssignments.ToString().ToLowerInvariant()},clearSubscopes=true)");

        using var response = await _requestExecutor.SendAsync(
            _httpClient,
            message,
            new SharePointSendOptions { IncludeFormDigest = true },
            ct);
    }

    public async Task ResetInheritanceAsync(string serverRelativeFolder, CancellationToken ct = default)
    {
        var validatedFolder = ServerRelativeUrl.Validate(serverRelativeFolder);
        var encodedFolder = Uri.EscapeDataString(validatedFolder);

        using var message = new HttpRequestMessage(
            HttpMethod.Post,
            $"_api/web/GetFolderByServerRelativeUrl('{encodedFolder}')/ListItemAllFields/resetroleinheritance");

        using var response = await _requestExecutor.SendAsync(
            _httpClient,
            message,
            new SharePointSendOptions { IncludeFormDigest = true },
            ct);
    }

    public async Task<int> EnsureGroupAsync(string groupName, string? description = null, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(groupName))
            throw new SharePointValidationException("Group name must be provided.");

        var existingId = await TryGetGroupIdAsync(groupName, ct);
        if (existingId.HasValue)
            return existingId.Value;

        using var message = new HttpRequestMessage(HttpMethod.Post, "_api/web/sitegroups")
        {
            Content = SharePointContentFactory.CreateJsonContent(new
            {
                __metadata = new { type = "SP.Group" },
                Title = groupName,
                Description = description ?? string.Empty
            })
        };

        using var response = await _requestExecutor.SendAsync(
            _httpClient,
            message,
            new SharePointSendOptions
            {
                IncludeFormDigest = true,
                EnsureSuccessStatusCode = false
            },
            ct);

        if (response.StatusCode == HttpStatusCode.Conflict)
        {
            var resolvedId = await TryGetGroupIdAsync(groupName, ct);
            return resolvedId ?? throw new SharePointConflictException($"Group '{groupName}' conflict but the existing Id could not be resolved.");
        }

        await _requestExecutor.EnsureSuccessAsync(response, ct);

        var payload = await response.Content.ReadAsStringAsync(ct);
        using var document = JsonDocument.Parse(payload);
        return document.RootElement.GetProperty("Id").GetInt32();
    }

    public async Task DeleteGroupAsync(string groupName, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(groupName))
            throw new SharePointValidationException("Group name must be provided.");

        var groupId = await TryGetGroupIdAsync(groupName, ct);
        if (!groupId.HasValue)
            return;

        using var message = new HttpRequestMessage(HttpMethod.Post, $"_api/web/sitegroups/removebyid({groupId.Value})");

        using var response = await _requestExecutor.SendAsync(
            _httpClient,
            message,
            new SharePointSendOptions
            {
                IncludeFormDigest = true,
                EnsureSuccessStatusCode = false
            },
            ct);

        if (response.StatusCode == HttpStatusCode.NotFound)
            return;

        await _requestExecutor.EnsureSuccessAsync(response, ct);
    }

    public async Task<int> EnsureUserAsync(string loginName, CancellationToken ct = default)
    {
        var normalizedLogin = NormalizeLogin(loginName);

        using var message = new HttpRequestMessage(HttpMethod.Post, "_api/web/ensureuser")
        {
            Content = SharePointContentFactory.CreateJsonContent(new { logonName = normalizedLogin })
        };

        using var response = await _requestExecutor.SendAsync(
            _httpClient,
            message,
            new SharePointSendOptions { IncludeFormDigest = true },
            ct);

        var payload = await response.Content.ReadAsStringAsync(ct);
        using var document = JsonDocument.Parse(payload);
        return document.RootElement.GetProperty("Id").GetInt32();
    }

    public async Task<IReadOnlyList<string>> GetGroupMembersAsync(string groupName, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(groupName))
            throw new SharePointValidationException("Group name must be provided.");

        var encodedName = Uri.EscapeDataString(groupName);
        using var message = new HttpRequestMessage(HttpMethod.Get, $"_api/web/sitegroups/getbyname('{encodedName}')/users?$select=LoginName");

        using var response = await _requestExecutor.SendAsync(
            _httpClient,
            message,
            new SharePointSendOptions { EnsureSuccessStatusCode = false },
            ct);

        if (response.StatusCode == HttpStatusCode.NotFound)
            return Array.Empty<string>();

        await _requestExecutor.EnsureSuccessAsync(response, ct);

        var payload = await response.Content.ReadAsStringAsync(ct);
        using var document = JsonDocument.Parse(payload);

        if (!document.RootElement.TryGetProperty("value", out var value)
            || value.ValueKind != JsonValueKind.Array)
        {
            return Array.Empty<string>();
        }

        return value.EnumerateArray()
            .Select(entry => entry.TryGetProperty("LoginName", out var login) ? login.GetString() : null)
            .Where(login => !string.IsNullOrWhiteSpace(login))
            .Select(login => login!)
            .ToList();
    }

    public async Task AddUsersToGroupAsync(string groupName, IEnumerable<string> loginNames, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(groupName))
            throw new SharePointValidationException("Group name must be provided.");

        var encodedName = Uri.EscapeDataString(groupName);

        foreach (var loginName in NormalizeLogins(loginNames))
        {
            using var message = new HttpRequestMessage(HttpMethod.Post, $"_api/web/sitegroups/getbyname('{encodedName}')/users")
            {
                Content = SharePointContentFactory.CreateJsonContent(new
                {
                    __metadata = new { type = "SP.User" },
                    LoginName = loginName
                })
            };

            using var response = await _requestExecutor.SendAsync(
                _httpClient,
                message,
                new SharePointSendOptions
                {
                    IncludeFormDigest = true,
                    EnsureSuccessStatusCode = false
                },
                ct);

            if (response.StatusCode == HttpStatusCode.Conflict)
                continue;

            await _requestExecutor.EnsureSuccessAsync(response, ct);
        }
    }

    public async Task RemoveUsersFromGroupAsync(string groupName, IEnumerable<string> loginNames, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(groupName))
            throw new SharePointValidationException("Group name must be provided.");

        var encodedName = Uri.EscapeDataString(groupName);

        foreach (var loginName in NormalizeLogins(loginNames))
        {
            var encodedLogin = Uri.EscapeDataString(loginName);
            using var message = new HttpRequestMessage(
                HttpMethod.Post,
                $"_api/web/sitegroups/getbyname('{encodedName}')/users/removebyloginname(@v)?@v='{encodedLogin}'");

            using var response = await _requestExecutor.SendAsync(
                _httpClient,
                message,
                new SharePointSendOptions
                {
                    IncludeFormDigest = true,
                    EnsureSuccessStatusCode = false
                },
                ct);

            if (response.StatusCode == HttpStatusCode.NotFound)
                continue;

            await _requestExecutor.EnsureSuccessAsync(response, ct);
        }
    }

    public async Task BindRoleToFolderAsync(string serverRelativeFolder, string principalName, string roleName, CancellationToken ct = default)
    {
        var validatedFolder = ServerRelativeUrl.Validate(serverRelativeFolder);
        var principalId = await ResolvePrincipalIdAsync(principalName, createMissingGroups: true, ct);
        var roleDefinitionId = await GetRoleDefinitionIdAsync(roleName, ct);
        var encodedFolder = Uri.EscapeDataString(validatedFolder);

        using var message = new HttpRequestMessage(
            HttpMethod.Post,
            $"_api/web/GetFolderByServerRelativeUrl('{encodedFolder}')/ListItemAllFields/roleassignments/addroleassignment(principalid={principalId},roledefid={roleDefinitionId})");

        using var response = await _requestExecutor.SendAsync(
            _httpClient,
            message,
            new SharePointSendOptions { IncludeFormDigest = true },
            ct);
    }

    public async Task RemoveRoleFromFolderAsync(string serverRelativeFolder, string principalName, CancellationToken ct = default)
    {
        var validatedFolder = ServerRelativeUrl.Validate(serverRelativeFolder);
        var principalId = await ResolvePrincipalIdOrNullAsync(principalName, ct);
        if (!principalId.HasValue)
            return;

        var encodedFolder = Uri.EscapeDataString(validatedFolder);
        using var message = new HttpRequestMessage(
            HttpMethod.Post,
            $"_api/web/GetFolderByServerRelativeUrl('{encodedFolder}')/ListItemAllFields/roleassignments/getbyprincipalid({principalId.Value})");

        message.Headers.Add("IF-MATCH", "*");
        message.Headers.Add("X-HTTP-Method", "DELETE");

        using var response = await _requestExecutor.SendAsync(
            _httpClient,
            message,
            new SharePointSendOptions
            {
                IncludeFormDigest = true,
                EnsureSuccessStatusCode = false
            },
            ct);

        if (response.StatusCode == HttpStatusCode.NotFound)
            return;

        await _requestExecutor.EnsureSuccessAsync(response, ct);
    }

    public async Task SyncGroupMembershipAsync(string groupName, IEnumerable<string> desiredLoginNames, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(groupName))
            throw new SharePointValidationException("Group name must be provided.");

        var desired = NormalizeLogins(desiredLoginNames).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var current = (await GetGroupMembersAsync(groupName, ct)).ToHashSet(StringComparer.OrdinalIgnoreCase);

        var toAdd = desired.Except(current).ToArray();
        var toRemove = current.Except(desired).ToArray();

        if (toAdd.Length > 0)
            await AddUsersToGroupAsync(groupName, toAdd, ct);

        if (toRemove.Length > 0)
            await RemoveUsersFromGroupAsync(groupName, toRemove, ct);
    }

    private async Task<int?> TryGetGroupIdAsync(string groupName, CancellationToken ct)
    {
        var encodedName = Uri.EscapeDataString(groupName);
        using var message = new HttpRequestMessage(HttpMethod.Get, $"_api/web/sitegroups/getbyname('{encodedName}')?$select=Id");

        using var response = await _requestExecutor.SendAsync(
            _httpClient,
            message,
            new SharePointSendOptions { EnsureSuccessStatusCode = false },
            ct);

        if (response.StatusCode == HttpStatusCode.NotFound || response.StatusCode == HttpStatusCode.InternalServerError)
            return null;

        await _requestExecutor.EnsureSuccessAsync(response, ct);

        var payload = await response.Content.ReadAsStringAsync(ct);
        using var document = JsonDocument.Parse(payload);
        return document.RootElement.GetProperty("Id").GetInt32();
    }

    private async Task<int?> TryGetUserPrincipalIdAsync(string loginName, CancellationToken ct)
    {
        var normalizedLogin = NormalizeLogin(loginName);
        var encodedLogin = Uri.EscapeDataString(normalizedLogin);
        using var message = new HttpRequestMessage(HttpMethod.Get, $"_api/web/siteusers(@v)?@v='{encodedLogin}'");

        using var response = await _requestExecutor.SendAsync(
            _httpClient,
            message,
            new SharePointSendOptions { EnsureSuccessStatusCode = false },
            ct);

        if (response.StatusCode == HttpStatusCode.NotFound)
            return null;

        await _requestExecutor.EnsureSuccessAsync(response, ct);

        var payload = await response.Content.ReadAsStringAsync(ct);
        using var document = JsonDocument.Parse(payload);
        return document.RootElement.GetProperty("Id").GetInt32();
    }

    private async Task<int> ResolvePrincipalIdAsync(string principalName, bool createMissingGroups, CancellationToken ct)
    {
        var existingGroupId = await TryGetGroupIdAsync(principalName, ct);
        if (existingGroupId.HasValue)
            return existingGroupId.Value;

        if (LooksLikeLogin(principalName))
        {
            var existingUserId = await TryGetUserPrincipalIdAsync(principalName, ct);
            return existingUserId ?? await EnsureUserAsync(principalName, ct);
        }

        return createMissingGroups
            ? await EnsureGroupAsync(principalName, string.Empty, ct)
            : (await TryGetGroupIdAsync(principalName, ct)
                ?? throw new SharePointNotFoundException($"SharePoint principal '{principalName}' was not found."));
    }

    private async Task<int?> ResolvePrincipalIdOrNullAsync(string principalName, CancellationToken ct)
    {
        var existingGroupId = await TryGetGroupIdAsync(principalName, ct);
        if (existingGroupId.HasValue)
            return existingGroupId.Value;

        if (LooksLikeLogin(principalName))
            return await TryGetUserPrincipalIdAsync(principalName, ct);

        return null;
    }

    private async Task<int> GetRoleDefinitionIdAsync(string roleName, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(roleName))
            throw new SharePointValidationException("Role name must be provided.");

        if (RoleDefinitionIdCache.TryGetValue(roleName, out var cached))
            return cached;

        if (!RoleNameToType.TryGetValue(roleName, out var roleType))
            throw new SharePointValidationException($"Unknown role name '{roleName}'. Use Read, Edit, or Contribute.");

        using var message = new HttpRequestMessage(HttpMethod.Get, $"_api/web/roledefinitions/getbytype({roleType})?$select=Id");
        using var response = await _requestExecutor.SendAsync(_httpClient, message, ct: ct);

        var payload = await response.Content.ReadAsStringAsync(ct);
        using var document = JsonDocument.Parse(payload);
        var roleDefinitionId = document.RootElement.GetProperty("Id").GetInt32();

        RoleDefinitionIdCache[roleName] = roleDefinitionId;
        return roleDefinitionId;
    }

    private IEnumerable<string> NormalizeLogins(IEnumerable<string> loginNames)
    {
        ArgumentNullException.ThrowIfNull(loginNames);

        return loginNames
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Select(NormalizeLogin)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private string NormalizeLogin(string loginName)
    {
        if (string.IsNullOrWhiteSpace(loginName))
            throw new SharePointValidationException("Login name must be provided.");

        return _loginNameNormalizer.Normalize(loginName);
    }

    private static bool LooksLikeLogin(string principalName)
    {
        return principalName.Contains('\\')
               || principalName.Contains('@')
               || principalName.Contains('|');
    }
}

