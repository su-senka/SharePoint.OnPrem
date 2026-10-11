# Files and Folders

## File operations
Available through `ISharePointFileClient`:
- `UploadAsync`: single request
- `UploadChunkedAsync`: upload session for large files, with progress (SharePoint 2016+)
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
- `CheckOutAsync`, `CheckInAsync`, `UndoCheckOutAsync`
- `GetVersionsAsync`

## Folder operations
Available through `ISharePointFolderClient`:
- `ExistsAsync`
- `CreateAsync`
- `DeleteAsync`
- `EnsureFolderPathAsync`: creates missing parents; a single request when the folder already exists
- `GetFilesAsync`: files directly in a folder (name, URL, size, created/modified time)
- `GetFoldersAsync`: subfolders directly in a folder (name, URL, item count)

## Example
```csharp
using System.Net;
using Microsoft.Extensions.DependencyInjection;
using SharePoint.OnPrem.Abstractions;
using SharePoint.OnPrem.DependencyInjection;

var services = new ServiceCollection();
services.AddSharePointOnPrem(
    configureCore: o => o.SiteBaseUrl = "https://sharepoint.contoso.local/sites/team",
    configureStorageScope: o => o.BaseFolderServerRelativeUrl = "/sites/team/Attachments",
    configureHttpClient: http => http.ConfigurePrimaryHttpMessageHandler(
        () => new HttpClientHandler { Credentials = CredentialCache.DefaultCredentials }));

using var provider = services.BuildServiceProvider();

var files = provider.GetRequiredService<ISharePointFileClient>();
var folders = provider.GetRequiredService<ISharePointFolderClient>();
var scope = provider.GetRequiredService<ISharePointServerRelativePathScope>();

var relativeFolder = "Customers/1042/2026";
var folder = scope.ToServerRelativePath(relativeFolder);

await folders.EnsureFolderPathAsync(folder);

var existing = await folders.GetFilesAsync(folder);
var subfolders = await folders.GetFoldersAsync(folder);

await using var stream = File.OpenRead("local-report.xlsx");
var stored = await files.UploadAsync(new UploadFileRequest(folder, "report.xlsx", stream, ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));

await using var download = await files.DownloadAsync(stored.ServerRelativeUrl);
// read download.Content (Stream) before disposing

var exists = await files.ExistsAsync(stored.ServerRelativeUrl);

var webUrl = await files.GetFileWebUrlAsync(stored.ServerRelativeUrl);

var copied = await files.CopyAsync(stored.ServerRelativeUrl, "/sites/team/Archive", "report-copy.xlsx");
var moved = await files.MoveAsync(copied.ServerRelativeUrl, "/sites/team/Archive/2026", "report-final.xlsx");
var renamed = await files.RenameAsync(moved.ServerRelativeUrl, "report-approved.xlsx");
await files.UpdateMetadataAsync(
    renamed.ServerRelativeUrl,
    new Dictionary<string, object?>
    {
        ["Title"] = "Report Approved",
        ["Category"] = "Finance"
    });

var metadata = await files.GetMetadataAsync(renamed.ServerRelativeUrl, ["Title", "Category"]);
var maybeMetadata = await files.TryGetMetadataAsync("/sites/team/Attachments/missing.xlsx", ["Title"]);
var title = await files.GetMetadataValueAsync<string>(renamed.ServerRelativeUrl, "Title");
var maybeTitle = await files.TryGetMetadataValueAsync<string>("/sites/team/Attachments/missing.xlsx", "Title");
```

## Large files
```csharp
await using var iso = File.OpenRead("install.iso");
var progress = new Progress<long>(sent => Console.WriteLine($"{sent * 100 / iso.Length}%"));

var stored = await files.UploadChunkedAsync(
    new UploadFileRequest("/sites/team/Shared Documents/Media", "install.iso", iso),
    chunkSizeBytes: 20 * 1024 * 1024,
    progress: progress);
```

- Files no larger than one chunk go through a plain `UploadAsync`, which also works on SharePoint 2013.
- The content stream may be non-seekable (e.g. a network or decompression stream). The library reads ahead one chunk to know which one is last.
- If a chunk fails, the upload session is cancelled and the original exception is rethrown. The empty file created at the start of the session stays in the library.

## Check-in, check-out and versions
```csharp
await files.CheckOutAsync(url);
await files.UploadAsync(new UploadFileRequest(folder, "contract.docx", updated));
await files.CheckInAsync(url, "Signed by both parties", SharePointCheckInType.MajorCheckIn);

foreach (var version in await files.GetVersionsAsync(url))
    Console.WriteLine($"{version.VersionLabel} {version.Created:d} {version.CheckInComment}");
```

`GetVersionsAsync` returns previous versions only; the current version is the file itself.

## Behavior notes
- File name validation enforces leaf names (no path separators).
- Paths must be server-relative (start with `/`), must not be pre-encoded, and must not contain `.` or `..` segments. `ISharePointServerRelativePathScope` therefore cannot escape its base folder.
- Names containing apostrophes (`O'Brien.docx`) are supported. Names containing `%` or `#` are not.
- `UploadAsync` reads `Content` from its current position and does **not** dispose it.
- `EnsureFolderPathAsync` only creates folders below the configured site (it never tries to create `/sites` or the site's own path).
- `GetFoldersAsync` on a library root also returns SharePoint's hidden `Forms` folder.
- `DownloadAsync` returns a `SharePointFileDownload` that wraps the response stream. Dispose it (or use `await using`) to release the network connection.
- `GetFileWebUrlAsync` attempts SharePoint linking URL first, then falls back to absolute URL composition.
- `CopyAsync` uses SharePoint `copyto(...)`; `MoveAsync`/`RenameAsync` use `moveto(...)` with overwrite flags.
- `UpdateMetadataAsync` sends a list-item `MERGE` against `ListItemAllFields` for the file.
- `GetMetadataAsync` reads `ListItemAllFields` and returns selected fields as a `JsonElement` map.
- `TryGetMetadataAsync` reads metadata like `GetMetadataAsync`, but returns `null` when the file is not found.
- `GetMetadataValueAsync<T>` reads one field and converts it to a typed value, returning default when the field is missing/null.
- `TryGetMetadataValueAsync<T>` combines not-found-safe lookup with typed field conversion.
- `ISharePointFolderClient.DeleteAsync` is idempotent — deleting a non-existent folder does not throw.

