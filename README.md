# SharePoint.OnPrem

`SharePoint.OnPrem` is a .NET SDK for SharePoint On-Prem REST integrations (files, folders, groups, users, and permissions).

The repository currently targets a first public release and is built as modular packages.

## Package map
- `SharePoint.OnPrem.Abstractions`: public contracts, DTOs, and exception types.
- `SharePoint.OnPrem.Core`: options, URL/path helpers, digest provider, request executor, error mapping.
- `SharePoint.OnPrem.Files`: file upload/download/delete, web URL lookup, folder create/exists/ensure.
- `SharePoint.OnPrem.Security`: groups, users, membership sync, inheritance, role binding.
- `SharePoint.OnPrem.DependencyInjection`: `IServiceCollection` registration helpers.

## Quick start (local source build)
```bash
cd /Users/oleksandr/RiderProjects/PlanProvozu/SharePoint.OnPrem

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

## Example programs
- Base sample: `samples/SharePoint.OnPrem.SampleConsole`
- Legacy compatibility adapter sample (not part of default solution build): `samples/SharePoint.OnPrem.PlanProvozuAdapter`

Run the base sample:
```bash
cd /Users/oleksandr/RiderProjects/PlanProvozu/SharePoint.OnPrem
dotnet run --project samples/SharePoint.OnPrem.SampleConsole/SharePoint.OnPrem.SampleConsole.csproj
```

## Documentation
- `docs/architecture.md`
- `docs/authentication.md`
- `docs/files.md`
- `docs/security.md`
- `docs/troubleshooting.md`
- `docs/migration.md`
- `CHANGELOG.md`
- `SECURITY.md`

## Release status
- Core/files/security/DI functionality is implemented and covered by tests.
- Package metadata and symbols are configured for packing.
- NuGet artifacts can be generated from `src/*` projects.

## CI/CD workflows
- `.github/workflows/ci.yml`: restore, build, and test on pushes/PRs.
- `.github/workflows/pack.yml`: create `.nupkg` and `.snupkg` artifacts (tag or manual run).
- `.github/workflows/publish.yml`: manual/release publish to NuGet.

Publish workflow requires repository secret:
- `NUGET_API_KEY`

## Notes
- Package metadata repository links target `https://github.com/su-senka/SharePoint.OnPrem`.
- The `PlanProvozu` compatibility adapter is for migration/validation; it is not the main public SDK surface.

