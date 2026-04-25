# Changelog

All notable changes to this repository will be documented in this file.

The format is based on Keep a Changelog, and this project currently tracks changes under a rolling `Unreleased` section until the first tagged release.

## [Unreleased]

### Added

- Added `scripts/run-one-click-demo.ps1` so the repository now has a deterministic demo that can be run without preparing external templates or screenshots.
- Added documentation indexes under `docs/`, `examples/`, `scripts/`, `profiles/`, and `references/` to make the repository easier to navigate on first open.
- Added `docs/one-click-demo.md` and `docs/social-launch-kit.md` to cover the demo flow and public-facing launch materials.
- Added `examples/demo-one-click/` and `examples/report-templates/experiment-report-template.docx` as a self-contained onboarding bundle.
- Added `scripts/build-report.ps1` as a single local entry point for validation, field-map generation, template filling, image insertion, and final style formatting.
- Added `scripts/report-defaults.ps1` so generated runs can remember the last course name and experiment name.
- Added `CONTRIBUTING.md` with repository workflow, testing expectations, and contribution scope.
- Added `CODE_OF_CONDUCT.md`, `SECURITY.md`, `SUPPORT.md`, and `ROADMAP.md` for GitHub-facing repository completeness.
- Added issue templates, a PR template, and a matrix quality workflow under `.github/`.
- Added `demo/` assets and a demo guide suitable for GitHub repository previews.
- Added `.gitattributes` to keep Office files and demo images treated as binary content.

### Changed

- Reworked `README.md` into a Chinese-first project homepage with clear positioning, quick-start commands, directory overview, and documentation navigation.
- Refreshed the demo and example documentation so the repo reads like a complete open-source project instead of a loose collection of helper scripts.
- Updated the example image JSON files to reference repo-contained demo assets instead of non-existent placeholder paths.
- Expanded `README.md` with a quick-start build flow, demo links, and contributor-oriented repository structure notes.
- Expanded `README.md` with repository health notes and a future profile-driven document roadmap.
- Expanded smoke tests to cover the new local build entry point and required repository files.
- Improved the final docx style formatter so body table rows can flow more naturally in common report templates instead of preserving awkward row-splitting constraints.
- Improved direct-chat image handling so OpenClaw-staged relative attachment paths such as `media/inbound/example.png` can resolve into the final docx image pipeline.

### Fixed

- Fixed `install-skill.ps1` so editor metadata and Python cache artifacts are not copied into installed skill directories.
- Fixed repository hygiene by ignoring Python cache artifacts such as `__pycache__/` and `*.pyc`.
- Fixed the report-style formatter so it no longer leaks XML attribute return values into the PowerShell pipeline.
- Fixed title detection for common report titles such as `计算机网络实验报告`.
