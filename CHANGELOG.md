# Changelog

All notable changes to this project are documented in this file.

The format is based on Keep a Changelog and this project follows Semantic Versioning.

## [Unreleased]

### Added
- Sprint D release assets:
  - `CHANGELOG.md`
  - `SECURITY.md`

### Changed
- Prepared repository-level governance and release process documentation for first public publish.

## [0.1.0] - 2026-05-22

### Added
- Modular package split:
  - `SharePoint.OnPrem.Abstractions`
  - `SharePoint.OnPrem.Core`
  - `SharePoint.OnPrem.Files`
  - `SharePoint.OnPrem.Security`
  - `SharePoint.OnPrem.DependencyInjection`
- Core runtime primitives:
  - options model and validation
  - login normalization
  - server-relative path helpers
  - form digest provider with caching
  - request executor and typed error mapping
- File and folder operations:
  - upload/download/delete
  - web URL lookup
  - folder exists/create/ensure
- Security operations:
  - group and user ensure/delete/list
  - membership add/remove/sync
  - inheritance break/reset
  - folder role binding/removal
- Dependency injection registration extensions.
- Contract-style tests with stubbed `HttpMessageHandler`.
- Compatibility adapter sample for legacy `PlanProvozu` service interface.
- GitHub Actions workflows:
  - CI (`.github/workflows/ci.yml`)
  - Pack (`.github/workflows/pack.yml`)
  - Publish (`.github/workflows/publish.yml`)

[Unreleased]: https://github.com/su-senka/SharePoint.OnPrem/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/su-senka/SharePoint.OnPrem/releases/tag/v0.1.0

