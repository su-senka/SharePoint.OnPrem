# Architecture

## Design goals
- Keep transport and SharePoint protocol details in a reusable core.
- Split capabilities by domain (files vs security) to avoid one large service class.
- Keep app-specific workflows (for example project-specific orchestration) outside the public SDK.

## Package boundaries
- `SharePoint.OnPrem.Abstractions`
  - interfaces (`ISharePointFileClient`, `ISharePointFolderClient`, `ISharePointSecurityClient`)
  - request/response records
  - exception hierarchy
- `SharePoint.OnPrem.Core`
  - options (`SharePointOnPremOptions`, `SharePointIdentityOptions`, `SharePointStorageScopeOptions`)
  - login normalization
  - server-relative path validation/combination
  - digest provider and request executor
  - HTTP error mapping to typed exceptions
- `SharePoint.OnPrem.Files`
  - file upload/download/delete
  - web URL resolution
  - folder exists/create/ensure path
- `SharePoint.OnPrem.Security`
  - groups and users
  - membership add/remove/sync
  - inheritance break/reset
  - role binding/removal
- `SharePoint.OnPrem.DependencyInjection`
  - service registration for all package interfaces

## Dependency direction
- `Files` and `Security` depend on `Abstractions` + `Core`.
- `DependencyInjection` depends on all package libraries.
- `Abstractions` has no project dependencies.

## Runtime flow (typical)
1. Consumer resolves a typed client from DI.
2. Client constructs SharePoint REST request.
3. `ISharePointRequestExecutor` optionally injects form digest and executes request.
4. Non-success responses are translated into typed `SharePointException` variants.

## Non-goals for first release
- SharePoint list/document-library high-level workflows.
- End-to-end app workflow orchestration (kept in application layer).

