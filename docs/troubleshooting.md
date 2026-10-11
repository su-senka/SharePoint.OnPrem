# Diagnostics and troubleshooting

## Check the connection first

Most integration problems with on-prem SharePoint are authentication or URL problems. Before debugging application code, ask SharePoint who it thinks you are:

```csharp
var site = provider.GetRequiredService<ISharePointSiteClient>();
var user = await site.GetCurrentUserAsync();   // i:0#.w|contoso\svc-sharepoint
var web = await site.GetWebInfoAsync();        // title, URL, LCID, "SharePoint 2019 (16.0.0.10337)"
```

The sample CLI does the same from the command line: `sp-onprem <siteUrl> whoami`.

## Health checks

```csharp
builder.Services.AddSharePointOnPrem(builder.Configuration.GetSection("SharePoint"), http => /* credentials */);
builder.Services.AddHealthChecks().AddSharePointOnPrem(tags: ["ready"]);

app.MapHealthChecks("/health/ready", new() { Predicate = r => r.Tags.Contains("ready") });
```

The check is **Healthy** when SharePoint answers and accepts the credentials. Its data contains `login`, `webUrl`, `serverVersion` and `product`. Rejected credentials, unreachable hosts, and timeouts each produce a distinct description.

## Correlation ids

Every `SharePointException` created from an HTTP response carries `CorrelationId`, the `SPRequestGuid` header. Give it to the farm administrator: `Get-SPLogEvent | ? Correlation -eq <guid>` or ULS Viewer shows the server-side error behind a vague REST message.

## Common errors

| Symptom | Cause | What to do |
| --- | --- | --- |
| `SharePointUnauthorizedException` (401) on every call | No credentials on the HTTP handler. | Configure `configureHttpClient` (see [auth.md](auth.md)). On Linux/macOS, install `gss-ntlmssp` / Kerberos libraries. |
| 403 *"The security validation for this page is invalid"* | Expired or foreign form digest (application pool recycle, digest from another site). | Handled automatically for replayable writes. Streamed `UploadAsync` cannot be replayed, so retry the upload yourself. |
| 403 on writes only | The account can read but not contribute. | Check permissions with `GetFolderRoleAssignmentsAsync`, or test with `whoami`. |
| `SharePointNotFoundException` for a path you can see in the browser | Path does not belong to `SiteBaseUrl`'s web, or contains a typo in the library URL (`Shared Documents` vs. `Documents`). | Use the server-relative URL from the browser address bar, decoded. |
| 500 for a group, folder, or role that does not exist | Older SharePoint versions report "not found" as 500. | The library treats these as "not found" where the operation is idempotent. |
| Upload fails for large files | Single-request limit (web application maximum upload size) or `HttpTimeoutMinutes`. | Use `UploadChunkedAsync`; raise the timeout. |
| `SharePointTransientException` (429/503) | Throttling or farm under load. | Wait for `RetryAfter`, or add a resilience handler for GET requests. |
| Validation error about `%` | Paths must be passed decoded. | Pass `/sites/team/Shared Documents`, not `Shared%20Documents`. |
