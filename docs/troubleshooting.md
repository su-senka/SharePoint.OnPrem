# Troubleshooting

## SourceLink warning during local build
Message example:
- `Source control information is not available - the generated source link is empty.`

Meaning:
- local build environment does not expose expected git metadata for SourceLink.

Impact:
- package still builds and packs.

Action:
- verify SourceLink in CI on GitHub-hosted build runners.

## `SharePointValidationException` for paths
Typical causes:
- path does not start with `/`
- path already contains URL-encoded `%`
- file name contains path separators

Action:
- pass server-relative raw paths (for example `/sites/pp/Attachments/3255/26/001`).
- pass leaf file names only.

## Unauthorized/Forbidden responses
Symptoms:
- `SharePointUnauthorizedException`

Action:
- verify app/service account permissions in SharePoint site and library.
- verify host authentication setup for outgoing `HttpClient` calls.

## Conflict on group/user operations
Symptoms:
- add-member returns conflict

Behavior:
- operations are designed to treat expected conflicts idempotently.

Action:
- usually no action needed unless conflict indicates wrong principal identity.

## JSON parsing errors from SharePoint
Symptoms:
- `SharePointValidationException` during digest or response parsing

Action:
- inspect response payload and verify API mode / endpoint shape.
- ensure proxy/security layers are not returning HTML error pages to API calls.

