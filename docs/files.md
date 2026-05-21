# Files and Folders

## File operations
Available through `ISharePointFileClient`:
- `UploadAsync`
- `DownloadAsync`
- `DeleteAsync`
- `GetWebUrlAsync`

## Folder operations
Available through `ISharePointFolderClient`:
- `ExistsAsync`
- `CreateAsync`
- `EnsurePathAsync`

## Example
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
    },
    configureStorageScope: o => o.BaseFolderServerRelativeUrl = "/sites/pp/Attachments");

using var provider = services.BuildServiceProvider();

var files = provider.GetRequiredService<ISharePointFileClient>();
var folders = provider.GetRequiredService<ISharePointFolderClient>();
var scope = provider.GetRequiredService<ISharePointPathScope>();

var relativeFolder = "3255/26/001";
var folder = scope.ToServerRelativePath(relativeFolder);

await folders.EnsurePathAsync(folder);

await using var stream = File.OpenRead("local-report.xlsx");
var stored = await files.UploadAsync(new UploadFileRequest(folder, "report.xlsx", stream, ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));

var webUrl = await files.GetWebUrlAsync(stored.ServerRelativeUrl);
```

## Behavior notes
- File name validation enforces leaf names (no path separators).
- Paths must be server-relative (start with `/`) and must not be pre-encoded.
- `GetWebUrlAsync` attempts SharePoint linking URL first, then falls back to absolute URL composition.

