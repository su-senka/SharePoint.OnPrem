# Authentication and HTTP pipeline

The library sends all requests through one named `HttpClient` (`ServiceCollectionExtensions.HttpClientName`, `"SharePoint.OnPrem"`). It does not implement authentication itself. You configure credentials on that client's handler through the `configureHttpClient` callback, which receives the `IHttpClientBuilder`.

## Windows authentication (NTLM / Kerberos)

Most on-prem farms use Windows authentication.

Run as the current process identity (service account, app pool identity):

```csharp
services.AddSharePointOnPrem(
    configureCore: o => o.SiteBaseUrl = "https://sharepoint.contoso.local/sites/team",
    configureHttpClient: http => http.ConfigurePrimaryHttpMessageHandler(
        () => new HttpClientHandler { Credentials = CredentialCache.DefaultCredentials }));
```

Use explicit credentials:

```csharp
configureHttpClient: http => http.ConfigurePrimaryHttpMessageHandler(
    () => new HttpClientHandler { Credentials = new NetworkCredential("svc-sharepoint", password, "CONTOSO") })
```

On Linux and macOS, NTLM and Kerberos need the GSSAPI libraries (`libgssapi-krb5` / `gss-ntlmssp`). Kerberos also needs a valid ticket or keytab.

## Other handlers

`configureHttpClient` exposes the full `IHttpClientBuilder`, so you can add logging, proxies, or resilience:

```csharp
// Microsoft.Extensions.Http.Resilience
configureHttpClient: http => http.AddStandardResilienceHandler()
```

Upload requests stream the caller's `Stream` and cannot be replayed. If you add retries, limit them to idempotent (GET) requests or buffer uploads.

## Login normalization

`ISharePointSecurityClient` normalizes login names through `ILoginNameNormalizer`, so callers can pass short forms and membership sync can compare them with what SharePoint stores:

| Input | Result (defaults: `ClaimsPrefix = "i:0#.w|"`, `Domain = "CONTOSO"`) |
| --- | --- |
| `jdoe` | `i:0#.w\|CONTOSO\jdoe` |
| `CONTOSO\jdoe` | `i:0#.w\|CONTOSO\jdoe` |
| `i:0#.w\|CONTOSO\jdoe`, `c:0+.w\|s-1-5-21-...` | unchanged (already claims-encoded) |
| `jdoe@contoso.com` | unchanged |

- `Domain` is optional. Without it, bare logins like `jdoe` are rejected.
- For classic-mode (non-claims) web applications, set `ClaimsPrefix` to an empty string.
- To use different rules, register your own `ILoginNameNormalizer` before or after calling `AddSharePointOnPrem`.

## Notes

- `SiteBaseUrl` is the web all requests go to. Paths you pass must belong to that web's site collection.
- The default timeout is 5 minutes (`HttpTimeoutMinutes`). Raise it for large uploads.
