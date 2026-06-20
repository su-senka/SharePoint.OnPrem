# Files and Folders

## File operations
Available through `ISharePointFileClient`:
- `UploadAsync`
- `DownloadAsync` — returns `SharePointFileDownload` (implements `IDisposable`/`IAsyncDisposable`); dispose when done reading the stream
- `ExistsAsync`
- `DeleteAsync`
- `GetFileWebUrlAsync`
- `CopyAsync`
- `MoveAsync`
- `RenameAsync`
- `UpdateMetadataAsync`
- `GetMetadataAsync`
- `TryGetMetadataAsync`
- `GetMetadataValueAsync<T>`
- `TryGetMetadataValueAsync<T>`

## Folder operations
Available through `ISharePointFolderClient`:
- `ExistsAsync`
- `CreateAsync`
- `DeleteAsync`
- `EnsureFolderPathAsync`

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
var scope = provider.GetRequiredService<ISharePointServerRelativePathScope>();

var relativeFolder = "3255/26/001";
var folder = scope.ToServerRelativePath(relativeFolder);

await folders.EnsureFolderPathAsync(folder);

await using var stream = File.OpenRead("local-report.xlsx");
var stored = await files.UploadAsync(new UploadFileRequest(folder, "report.xlsx", stream, ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));

await using var download = await files.DownloadAsync(stored.ServerRelativeUrl);
// read download.Content (Stream) before disposing

var exists = await files.ExistsAsync(stored.ServerRelativeUrl);

var webUrl = await files.GetFileWebUrlAsync(stored.ServerRelativeUrl);

var copied = await files.CopyAsync(stored.ServerRelativeUrl, "/sites/pp/Archive", "report-copy.xlsx");
var moved = await files.MoveAsync(copied.ServerRelativeUrl, "/sites/pp/Archive/2026", "report-final.xlsx");
var renamed = await files.RenameAsync(moved.ServerRelativeUrl, "report-approved.xlsx");
await files.UpdateMetadataAsync(
    renamed.ServerRelativeUrl,
    new Dictionary<string, object?>
    {
        ["Title"] = "Report Approved",
        ["Category"] = "Finance"
    });

var metadata = await files.GetMetadataAsync(renamed.ServerRelativeUrl, ["Title", "Category"]);
var maybeMetadata = await files.TryGetMetadataAsync("/sites/pp/Attachments/missing.xlsx", ["Title"]);
var title = await files.GetMetadataValueAsync<string>(renamed.ServerRelativeUrl, "Title");
var maybeTitle = await files.TryGetMetadataValueAsync<string>("/sites/pp/Attachments/missing.xlsx", "Title");
```

## Behavior notes
- File name validation enforces leaf names (no path separators).
- Paths must be server-relative (start with `/`) and must not be pre-encoded.
- `DownloadAsync` returns a `SharePointFileDownload` that wraps the response stream. Dispose it (or use `await using`) to release the network connection.
- `GetFileWebUrlAsync` attempts SharePoint linking URL first, then falls back to absolute URL composition.
- `CopyAsync` uses SharePoint `copyto(...)`; `MoveAsync`/`RenameAsync` use `moveto(...)` with overwrite flags.
- `UpdateMetadataAsync` sends a list-item `MERGE` against `ListItemAllFields` for the file.
- `GetMetadataAsync` reads `ListItemAllFields` and returns selected fields as a `JsonElement` map.
- `TryGetMetadataAsync` reads metadata like `GetMetadataAsync`, but returns `null` when the file is not found.
- `GetMetadataValueAsync<T>` reads one field and converts it to a typed value, returning default when the field is missing/null.
- `TryGetMetadataValueAsync<T>` combines not-found-safe lookup with typed field conversion.
- `ISharePointFolderClient.DeleteAsync` is idempotent — deleting a non-existent folder does not throw.

