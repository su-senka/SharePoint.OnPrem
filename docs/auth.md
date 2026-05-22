# Authentication

This SDK uses `HttpClient` and does not impose a single auth model. Most SharePoint On-Prem deployments use integrated auth variants.

## Recommended setup
Configure authentication at `HttpClient` handler level in the host app/environment. Then register SDK services with `AddSharePointOnPrem(...)`.

## Example (Windows integrated auth style)
```csharp
using Microsoft.Extensions.DependencyInjection;
using SharePoint.OnPrem.DependencyInjection;

var services = new ServiceCollection();

services.AddSharePointOnPrem(
    configureCore: options =>
    {
        options.SiteBaseUrl = "https://sharepoint.local/sites/pp";
        options.HttpTimeoutMinutes = 5;
    },
    configureIdentity: options =>
    {
        options.Domain = "ACR";
        options.ClaimsPrefix = "i:0#.w|";
    });
```

## Login normalization
`ISharePointSecurityClient` operations normalize login names via `ILoginNameNormalizer`.

Default behavior:
- already claims-qualified login -> unchanged
- `DOMAIN\\user` or `user@domain` -> unchanged
- bare `user` -> expanded to `{ClaimsPrefix}{Domain}\\user`

You can replace the default normalizer by registering your own `ILoginNameNormalizer` implementation.

## Notes
- Ensure `SiteBaseUrl` points to the SharePoint site root used by your REST calls.
- Keep `BaseFolderServerRelativeUrl` server-relative and not URL-encoded.
- For environments with reverse proxies/custom auth, validate end-to-end with integration tests.

