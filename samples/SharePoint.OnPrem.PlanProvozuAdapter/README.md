# SharePoint.OnPrem.PlanProvozuAdapter

Compatibility layer for `PlanProvozu`'s legacy `PP.Application.Services.SharePoint.ISharePointService`.

## Purpose
This project is **not** the new public SharePoint SDK surface. It is a bridge that lets you validate whether the standalone `SharePoint.OnPrem` package can satisfy the old service contract used by `PlanProvozu`.

## What it does
- implements the legacy `ISharePointService` interface
- maps legacy file/folder/group methods to the new package clients
- preserves legacy high-level orchestration helpers:
  - `EnsureRootSecurityAsync`
  - `EnsureAssignmentSecurityAsync`
- preserves legacy file-permission flows by issuing file-level role assignment REST calls through the shared core executor

## Registration
Register the standalone package first, then register the compatibility adapter.

```csharp
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
    },
    configureStorageScope: options =>
    {
        options.BaseFolderServerRelativeUrl = "/sites/pp/Attachments";
    });

services.AddPlanProvozuSharePointCompatibilityAdapter();
```

## Validation status
Behavior is covered by `tests/SharePoint.OnPrem.PlanProvozuAdapter.Tests`.

These tests verify:
- scoped upload/delete mapping
- root/assignment security composition
- legacy file read-permission flow
- DI registration of the legacy interface

## Intended use
Use this adapter only for migration and side-by-side verification. Once `PlanProvozu` is updated to consume the standalone package directly, this adapter can be retired.

