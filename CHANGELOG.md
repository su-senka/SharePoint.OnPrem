# Changelog

All notable changes to this project are documented here. The project follows [Semantic Versioning](https://semver.org/); while the version is `0.x`, minor versions may contain breaking changes.

## 0.1.0 (unreleased)

First public release. Changes relative to the earlier internal builds:

### Added
- `ISharePointFileClient.UploadChunkedAsync`: chunked upload sessions for large files, with progress reporting and non-seekable stream support.
- `CheckOutAsync`, `CheckInAsync`, `UndoCheckOutAsync`, and `GetVersionsAsync` on `ISharePointFileClient`.
- `ISharePointSiteClient`: current user and web information, including the server build and product name (`SharePointProductVersion`).
- `SharePoint.OnPrem.HealthChecks` package: `AddHealthChecks().AddSharePointOnPrem()`.
- Automatic form digest refresh: a write rejected with 403 after the cached digest expired early (e.g. an application pool recycle) is retried once with a new digest.
- `SharePointTransientException.RetryAfter`.
- Sample CLI (`samples/SharePoint.OnPrem.Sample`) for smoke-testing a farm.
- `ISharePointFolderClient.GetFilesAsync` and `GetFoldersAsync` for listing folder contents (`SharePointFileInfo`, `SharePointFolderInfo`).
- `AddSharePointOnPrem(IConfiguration)` overload that binds the `SharePoint`, `SharePoint:Identity` and `SharePoint:StorageScope` sections.
- `configureHttpClient` callback on `AddSharePointOnPrem` for configuring credentials, proxies, and handlers. All clients now share one named `HttpClient` (`ServiceCollectionExtensions.HttpClientName`).
- `SharePointException.StatusCode`; all exception types accept an inner exception.
- Role names other than the built-in ones (custom or localized permission levels) are resolved by name. `Design` and `Full Control` were added to the built-in names.
- XML documentation for the public API.

### Changed
- `SharePointIdentityOptions.ClaimsPrefix` defaults to `i:0#.w|`, and `Domain` is optional (only needed to expand bare logins).
- `DOMAIN\user` logins are normalized to `{ClaimsPrefix}DOMAIN\user`. Claims-encoded logins (containing `|`) are left unchanged.
- HTTP 502 and 504 map to `SharePointTransientException`.
- Core services are registered with `TryAdd`, so replacements registered earlier take precedence.

### Fixed
- `SyncGroupMembershipAsync` removed members whose login differed only in casing or form (e.g. `CONTOSO\jdoe` vs. `i:0#.w|contoso\jdoe`) from what SharePoint reports.
- Paths, file names, group names, and logins containing an apostrophe produced malformed requests.
- `EnsureFolderPathAsync` tried to probe and create the site's own path segments (e.g. `/sites`). It now starts below the site and makes a single request when the folder already exists.
- `UploadAsync` disposed the caller's content stream.
- Error responses in OData JSON-light (`odata.error`) or verbose format with an object `message` are parsed correctly; the verbose form previously caused an `InvalidOperationException`. Non-JSON error bodies are truncated in exception messages.
- Relative paths containing `..` could escape the configured base folder of `ISharePointServerRelativePathScope`.
