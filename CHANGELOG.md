# Changelog

All notable changes to this repository will be documented in this file.

The format is based on Keep a Changelog, and this project currently tracks changes under a rolling `Unreleased` section until the first tagged release.

## [Unreleased]

### Added

- Added `scripts/build-report.ps1` as a single local entry point for validation, field-map generation, template filling, image insertion, and final style formatting.
- Added `scripts/convert-docx-template-frame.ps1` and template-frame output support for the local, URL, and Feishu build paths.
- Added `profiles/report-profile.schema.json` and `scripts/validate-report-profiles.ps1` to make built-in report profiles schema-backed and smoke-tested.
- Added `scripts/new-report-profile.ps1` to scaffold schema-valid report profiles before manual profile tuning.
- Added the built-in `software-test-report` profile for software testing reports with test-case, result, and defect-analysis sections.
- Added the built-in `deployment-report` profile for deployment and operations reports with deployment, verification, and rollback sections.
- Added `examples/profile-presets/` with reusable `weekly-report` and `meeting-minutes` custom profile presets for adjacent document experiments.
- Added `scripts/report-defaults.ps1` so generated runs can remember the last course name and experiment name.
- Added `CONTRIBUTING.md` with repository workflow, testing expectations, and contribution scope.
- Added `CODE_OF_CONDUCT.md`, `SECURITY.md`, `SUPPORT.md`, and `ROADMAP.md` for GitHub-facing repository completeness.
- Added issue templates, a PR template, and a matrix quality workflow under `.github/`.
- Added `demo/` assets and a demo guide suitable for GitHub repository previews.
- Added `.gitattributes` to keep Office files and demo images treated as binary content.

### Changed

- Expanded `README.md` with a quick-start build flow, demo links, and contributor-oriented repository structure notes.
- Expanded `README.md` with repository health notes and a future profile-driven document roadmap.
- Clarified the GitHub-facing documentation for the expansion from experiment reports into course-design and internship report profiles.
- Documented validation/risk output files, summary fields, and machine-readable structural/pagination risk codes.
- Expanded smoke tests to cover the new local build entry point and required repository files.
- Improved the final docx style formatter so body table rows can flow more naturally in common report templates instead of preserving awkward row-splitting constraints.
- Improved direct-chat image handling so OpenClaw-staged relative attachment paths such as `media/inbound/example.png` can resolve into the final docx image pipeline.
- Propagated structural validation and pagination-risk summaries through URL and Feishu wrapper outputs.
- Expanded smoke coverage for structural validation and pagination-risk codes across experiment and internship report profiles.
- Added end-to-end smoke coverage for pagination-risk warning propagation through build-report, URL wrapper, and Feishu wrapper summaries.
- Added smoke coverage for template-frame docx generation and wrapper summary propagation.

### Fixed

- Fixed the report-style formatter so it no longer leaks XML attribute return values into the PowerShell pipeline.
- Fixed title detection for common report titles such as `计算机网络实验报告`.
