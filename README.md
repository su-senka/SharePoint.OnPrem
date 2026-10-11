# SharePoint.OnPrem

A typed .NET client for the **SharePoint Server (on-premises) REST API**: files, folders, metadata, site groups, users, and item-level permissions.

The SharePoint SDKs Microsoft maintains today (PnP Core, Microsoft Graph) target SharePoint Online. `SharePoint.OnPrem` targets on-premises farms and wraps the `_api/web` REST endpoints behind small, testable interfaces that integrate with `IHttpClientFactory` and dependency injection.

## Why this exists

Calling the on-prem REST API directly looks easy until you hit its quirks:

- Every write needs a **form digest**. The digest silently expires when the application pool recycles, and writes then fail with *"The security validation for this page is invalid"*.
- Names inside `GetFileByServerRelativeUrl('...')` are OData literals, so a file called `O'Brien.docx` breaks the request unless the quote is doubled **before** URL encoding.
- SharePoint stores users as claims (`i:0#.w|contoso\jdoe`) while people type `CONTOSO\JDoe`, so naive membership comparisons add and remove the wrong people.
- A missing group or folder often comes back as **HTTP 500**, not 404.
- Large files need the `StartUpload` / `ContinueUpload` / `FinishUpload` session API, and Windows authentication has to be configured on the right `HttpClient` handler.

This library handles all of these behind a typed API, and every behavior is pinned by contract tests against recorded REST calls.

## Features

- **Files**: upload (single request or chunked, with progress), streaming download, exists, delete, copy, move, rename, web URL
- **Metadata**: read and update list item fields, typed value helpers
- **Versioning**: check out, check in (minor/major/overwrite), undo check-out, version history
- **Folders**: ensure a path (creates missing parents), exists, create, delete, list files and subfolders
- **Security**: site groups, users, idempotent membership sync, break/reset inheritance, role assignments on folders and files
- **Diagnostics**: current user, web info, server product and build, ASP.NET Core health check
- **Reliability**: typed exceptions with `SPRequestGuid` correlation ids, `Retry-After` on throttling, automatic form digest refresh

## Packages

| Package | Contents |
| --- | --- |
| `SharePoint.OnPrem.DependencyInjection` | `AddSharePointOnPrem(...)`. **Install this one**; it brings in the rest. |
| `SharePoint.OnPrem.Files` | `ISharePointFileClient`, `ISharePointFolderClient` implementations |
| `SharePoint.OnPrem.Security` | `ISharePointSecurityClient` implementation |
| `SharePoint.OnPrem.HealthChecks` | `AddHealthChecks().AddSharePointOnPrem()`: verifies connectivity and credentials |
| `SharePoint.OnPrem.Core` | Options, path helpers, form digest handling, request executor, error mapping, `ISharePointSiteClient` |
| `SharePoint.OnPrem.Abstractions` | Interfaces, DTOs, exceptions. Reference it from projects that only consume the interfaces. |

```bash
dotnet add package SharePoint.OnPrem.DependencyInjection
```

## Requirements

- .NET 10
- SharePoint Server with REST API access and support for `application/json;odata=nometadata` responses (SharePoint 2013 SP1 and later)

## Quick start

```csharp
using System.Net;
using Microsoft.Extensions.DependencyInjection;
using SharePoint.OnPrem.Abstractions;
using SharePoint.OnPrem.DependencyInjection;

var services = new ServiceCollection();

services.AddSharePointOnPrem(
    configureCore: o => o.SiteBaseUrl = "https://sharepoint.contoso.local/sites/team",
    configureIdentity: o => o.Domain = "CONTOSO",
    // On-prem SharePoint usually needs Windows authentication (NTLM/Kerberos).
    configureHttpClient: http => http.ConfigurePrimaryHttpMessageHandler(
        () => new HttpClientHandler { Credentials = CredentialCache.DefaultCredentials }));

using var provider = services.BuildServiceProvider();

var folders = provider.GetRequiredService<ISharePointFolderClient>();
var files = provider.GetRequiredService<ISharePointFileClient>();

await folders.EnsureFolderPathAsync("/sites/team/Shared Documents/Reports/2026");

await using var stream = File.OpenRead("report.pdf");
var stored = await files.UploadAsync(new UploadFileRequest(
    "/sites/team/Shared Documents/Reports/2026", "report.pdf", stream, "application/pdf"));

foreach (var file in await folders.GetFilesAsync("/sites/team/Shared Documents/Reports/2026"))
    Console.WriteLine($"{file.Name} ({file.Length} bytes)");
```

### Configuration binding

```json
{
  "SharePoint": {
    "SiteBaseUrl": "https://sharepoint.contoso.local/sites/team",
    "HttpTimeoutMinutes": 10,
    "Identity": { "Domain": "CONTOSO" },
    "StorageScope": { "BaseFolderServerRelativeUrl": "/sites/team/Shared Documents" }
  }
}
```

```csharp
builder.Services.AddSharePointOnPrem(builder.Configuration.GetSection("SharePoint"));
```

## Conventions

- **Paths are server-relative and decoded**: `/sites/team/Shared Documents/a b.docx`, not `Shared%20Documents`. The library encodes them. `.` and `..` segments are rejected.
- **Errors are typed**: `SharePointNotFoundException`, `SharePointUnauthorizedException`, `SharePointConflictException`, `SharePointTransientException` (429/502/503/504), `SharePointValidationException`. All derive from `SharePointException`, which carries `StatusCode`, `RequestUrl`, and `CorrelationId` (the `SPRequestGuid` to look up in the ULS logs).
- **Idempotent where it makes sense**: `EnsureFolderPathAsync`, `EnsureGroupAsync`, folder `DeleteAsync`, `DeleteGroupAsync`, and membership add/remove tolerate "already exists" or "already gone".
- **Form digests** for write requests are fetched and cached automatically.

## Documentation

- [Authentication and HTTP pipeline](docs/auth.md)
- [Files and folders](docs/files.md)
- [Security: groups, users, permissions](docs/security.md)
- [Diagnostics and troubleshooting](docs/troubleshooting.md)
- [Changelog](CHANGELOG.md)

## Known limitations

- File and folder names containing `%` or `#` are not supported (the library uses the `*ByServerRelativeUrl` endpoints).
- `UploadAsync` sends the file in one request. Use `UploadChunkedAsync` for large files (SharePoint 2016 and later).
- No built-in retries. Add a resilience handler through `configureHttpClient` if you need them (see [docs/auth.md](docs/auth.md)).

## Try it against your farm

The sample CLI is a quick smoke test. It authenticates as the current Windows user, or with `SP_USERNAME` / `SP_PASSWORD` / `SP_DOMAIN`:

```bash
dotnet run --project samples/SharePoint.OnPrem.Sample -- https://sharepoint.contoso.local/sites/team whoami
# User:    i:0#.w|contoso\jdoe (John Doe)
# Web:     Team <https://sharepoint.contoso.local/sites/team>
# Server:  SharePoint 2019 (16.0.0.10337)

dotnet run --project samples/SharePoint.OnPrem.Sample -- <siteUrl> ls "/sites/team/Shared Documents"
dotnet run --project samples/SharePoint.OnPrem.Sample -- <siteUrl> upload ./big.iso "/sites/team/Shared Documents"
```

Commands: `whoami`, `health`, `ls`, `upload`, `download`, `versions`.

## Building from source

```bash
dotnet test SharePoint.OnPrem.slnx
dotnet pack SharePoint.OnPrem.slnx -c Release -o artifacts
```

## License

[MIT](LICENSE)
