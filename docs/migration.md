# Migration Guide

This guide helps move existing custom SharePoint service implementations to `SharePoint.OnPrem` incrementally.

## Recommended migration strategy
1. Keep existing app service contracts unchanged.
2. Introduce a compatibility adapter in your application layer.
3. Route adapter internals to `SharePoint.OnPrem` typed clients.
4. Validate behavior side-by-side in tests.
5. Swap DI binding once confidence is high.

## Mapping template
Typical legacy to new mapping:
- file upload/download/delete -> `ISharePointFileClient`
- folder create/exists/ensure -> `ISharePointFolderClient`
- group/user/membership/roles -> `ISharePointSecurityClient`
- path conventions -> `ISharePointPathScope`

## Compatibility adapter note
This repository contains an example adapter under:
- `samples/SharePoint.OnPrem.PlanProvozuAdapter`

It demonstrates how to keep a legacy service interface while using the new package internally.

## Validation checklist
- Compare generated request URIs for critical operations.
- Compare role/membership side effects in SharePoint.
- Verify exception mapping expected by your app.
- Verify any app-specific folder conventions.

## After successful migration
- Remove compatibility adapter layer.
- Resolve package interfaces directly in application services.

