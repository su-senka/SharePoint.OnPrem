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
using Microsoft.Extensions.DependencyInjection;
using SharePoint.OnPrem.Abstractions;
using SharePoint.OnPrem.DependencyInjection;

var services = new ServiceCollection();
services.AddSharePointOnPrem(
    configureCore: o => o.SiteBaseUrl = "https://sharepoint.local/sites/pp",
    configureIdentity: o =>
    {
        o.Domain = "ACR";
        o.ClaimsPrefix = "i:0#.w|";
    });

using var provider = services.BuildServiceProvider();
var security = provider.GetRequiredService<ISharePointSecurityClient>();

const string group = "PP Readers";
await security.EnsureGroupAsync(group, "Read access for root assignment tree");

await security.SyncGroupMembershipAsync(group, new[]
{
    "jnovak",
    "ACR\\svoboda",
    "i:0#.w|ACR\\kral"
});
```

## Folder role binding example
```csharp
await security.BreakInheritanceAsync("/sites/pp/Attachments/3255/26/001", copyRoleAssignments: false);
await security.BindRoleToFolderAsync("/sites/pp/Attachments/3255/26/001", "PP Readers", "Read");
await security.BindRoleToFolderAsync("/sites/pp/Attachments/3255/26/001", "PP Owners", "Edit");

// Break inheritance on a specific file before setting file-level permissions
await security.BreakFileInheritanceAsync("/sites/pp/Attachments/3255/26/001/report.xlsx", copyRoleAssignments: false);
await security.BindRoleToFileAsync("/sites/pp/Attachments/3255/26/001/report.xlsx", "PP Owners", "Edit");

var assignments = await security.GetFolderRoleAssignmentsAsync("/sites/pp/Attachments/3255/26/001");
foreach (var assignment in assignments)
{
    Console.WriteLine($"{assignment.PrincipalTitle}: {string.Join(", ", assignment.Roles.Select(r => r.Name))}");
}

var fileAssignments = await security.GetFileRoleAssignmentsAsync("/sites/pp/Attachments/3255/26/001/report.xlsx");
```

## Role names
Supported out of the box:
- `Read`
- `Edit`
- `Contribute`
- `Číst`

## Behavior notes
- Membership operations are idempotent for conflicts/not-found removal paths.
- `SyncGroupMembershipAsync` performs set reconciliation (+add, -remove).
- Principal names may be group names or user login names.
- `GetFolderRoleAssignmentsAsync` returns expanded principal + role bindings from `RoleAssignments` on the folder list item.
- `GetFileRoleAssignmentsAsync` returns expanded principal + role bindings from `RoleAssignments` on the file list item.

