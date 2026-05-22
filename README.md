# SharePoint.OnPrem

`SharePoint.OnPrem` is a .NET SDK for SharePoint On-Prem REST integrations (files, folders, groups, users, and permissions).

The repository currently targets a first public release and is built as modular packages.

## Minimum requirements
- .NET SDK 10.0+
- SharePoint On-Prem environment with REST API access
- Host-configured authentication for outgoing `HttpClient` requests

## Package map
- `SharePoint.OnPrem.Abstractions`: public contracts, DTOs, and exception types.
- `SharePoint.OnPrem.Core`: options, URL/path helpers, digest provider, request executor, error mapping.
- `SharePoint.OnPrem.Files`: file upload/download/delete, web URL lookup, folder create/exists/ensure.
- `SharePoint.OnPrem.Security`: groups, users, membership sync, inheritance, role binding.
- `SharePoint.OnPrem.DependencyInjection`: `IServiceCollection` registration helpers.

## Supported scenarios
- File and folder lifecycle operations in SharePoint (`upload`, `download`, `delete`, `ensure folder path`).
- File relocation workflows (`copy`, `move`, `rename`) with server-relative paths.
- File metadata update/read workflows (`UpdateMetadataAsync`, `GetMetadataAsync`, typed helpers).
- Security operations for groups/users/membership and folder/file role assignment.
- Permission inspection on folder/file list items.

## Non-goals
- App-specific orchestration layers and legacy adapters.
- Broad SharePoint platform abstraction beyond REST-oriented primitives.
- Host authentication middleware; auth stays a host application responsibility.

## Quick start (local source build)
```bash
cd /Users/oleksandr/RiderProjects/SharePoint.OnPrem

dotnet restore SharePoint.OnPrem.slnx
dotnet build SharePoint.OnPrem.slnx -c Debug --no-restore
dotnet test SharePoint.OnPrem.slnx -c Debug --no-build
```

## Quick start (DI usage)
```csharp
using Microsoft.Extensions.DependencyInjection;
using SharePoint.OnPrem.DependencyInjection;

var services = new ServiceCollection();

services.AddSharePointOnPrem(
    configureCore: options =>
    {
        options.SiteBaseUrl = "https://sharepoint.local/sites/pp";
        options.HttpTimeoutMinutes = 5;
        options.UseFormDigestCaching = true;
    },
    configureIdentity: options =>
    {
        options.Domain = "ACR";
        options.ClaimsPrefix = "i:0#.w|";
    },
    configureStorageScope: options =>
    {
        options.BaseFolderServerRelativeUrl = "/sites/pp/Attachments";
    });
```

## Documentation
- `docs/architecture.md`
- `docs/authentication.md`
- `docs/files.md`
- `docs/security.md`
- `docs/troubleshooting.md`
- `docs/contributing.md`
- `CHANGELOG.md`

## Release status
- Core/files/security/DI functionality is implemented and covered by tests.
- Package metadata and symbols are configured for packing.
- NuGet artifacts can be generated from `src/*` projects.

## CI/CD workflows
- `.github/workflows/ci.yml`: restore, build, and test on pushes/PRs.

Package publishing to NuGet is performed manually by maintainers.

## Notes
- Package metadata repository links target `https://github.com/su-senka/SharePoint.OnPrem`.

