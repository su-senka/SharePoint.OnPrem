# Security (Groups, Users, Roles)

## Capabilities
Available through `ISharePointSecurityClient`:
- folder inheritance: `BreakInheritanceAsync`, `ResetInheritanceAsync`
- file inheritance: `BreakFileInheritanceAsync`, `ResetFileInheritanceAsync`
- groups: `EnsureGroupAsync`, `DeleteGroupAsync`, `GetGroupMembersAsync`
- users: `EnsureUserAsync`
- membership: `AddUsersToGroupAsync`, `RemoveUsersFromGroupAsync`, `SyncGroupMembershipAsync`
- role binding: `BindRoleToFolderAsync`, `RemoveRoleFromFolderAsync`
- file role binding: `BindRoleToFileAsync`, `RemoveRoleFromFileAsync`
- permission inspection: `GetFolderRoleAssignmentsAsync`, `GetFileRoleAssignmentsAsync`

## Membership sync example
```csharp
using System.Net;
using Microsoft.Extensions.DependencyInjection;
using SharePoint.OnPrem.Abstractions;
using SharePoint.OnPrem.DependencyInjection;

var services = new ServiceCollection();
services.AddSharePointOnPrem(
    configureCore: o => o.SiteBaseUrl = "https://sharepoint.contoso.local/sites/team",
    configureIdentity: o => o.Domain = "CONTOSO",
    configureHttpClient: http => http.ConfigurePrimaryHttpMessageHandler(
        () => new HttpClientHandler { Credentials = CredentialCache.DefaultCredentials }));

using var provider = services.BuildServiceProvider();
var security = provider.GetRequiredService<ISharePointSecurityClient>();

const string group = "Team Readers";
await security.EnsureGroupAsync(group, "Read access for root assignment tree");

await security.SyncGroupMembershipAsync(group, new[]
{
    "jdoe",                      // bare login, expanded with the configured Domain
    "CONTOSO\\asmith",            // DOMAIN\user
    "i:0#.w|contoso\\bjones"     // claims login, as SharePoint reports it
});
```

## Folder role binding example
```csharp
await security.BreakInheritanceAsync("/sites/team/Shared Documents/Contracts", copyRoleAssignments: false);
await security.BindRoleToFolderAsync("/sites/team/Shared Documents/Contracts", "Team Readers", "Read");
await security.BindRoleToFolderAsync("/sites/team/Shared Documents/Contracts", "Team Owners", "Edit");

// Break inheritance on a specific file before setting file-level permissions
await security.BreakFileInheritanceAsync("/sites/team/Shared Documents/Contracts/report.xlsx", copyRoleAssignments: false);
await security.BindRoleToFileAsync("/sites/team/Shared Documents/Contracts/report.xlsx", "Team Owners", "Edit");

var assignments = await security.GetFolderRoleAssignmentsAsync("/sites/team/Shared Documents/Contracts");
foreach (var assignment in assignments)
{
    Console.WriteLine($"{assignment.PrincipalTitle}: {string.Join(", ", assignment.Roles.Select(r => r.Name))}");
}

var fileAssignments = await security.GetFileRoleAssignmentsAsync("/sites/team/Shared Documents/Contracts/report.xlsx");
```

## Role names
These names resolve by SharePoint role type, so they work on servers in any language:
- `Read`
- `Contribute`
- `Edit`
- `Design`
- `Full Control`

Any other name, such as a custom permission level or a localized name like `Lesen` (German "Read"), is resolved with `roledefinitions/getbyname`. An unknown name throws `SharePointValidationException`.

## Behavior notes
- Membership operations are idempotent for conflicts/not-found removal paths.
- `SyncGroupMembershipAsync` performs set reconciliation (+add, -remove).
- Principal names may be group names or user login names. A name is resolved as a site group first. If no group exists and the name contains `\`, `@` or `|`, it is treated as a user (added to the site with `ensureuser` if needed).
- **`BindRoleToFolderAsync` / `BindRoleToFileAsync` create a missing site group** when the principal name doesn't look like a login. A typo in a group name therefore creates a new, empty group. Also, a bare login such as `jdoe` is treated as a group name here; pass `CONTOSO\jdoe` for users.
- Binding a role requires the item to have unique permissions; call `BreakInheritanceAsync` / `BreakFileInheritanceAsync` first.
- `SyncGroupMembershipAsync` compares normalized logins case-insensitively, so `CONTOSO\jdoe` matches the `i:0#.w|contoso\jdoe` SharePoint reports.
- `GetFolderRoleAssignmentsAsync` returns expanded principal + role bindings from `RoleAssignments` on the folder list item.
- `GetFileRoleAssignmentsAsync` returns expanded principal + role bindings from `RoleAssignments` on the file list item.

