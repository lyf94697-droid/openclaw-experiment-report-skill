# Contributing

## Scope

This repository is an OpenClaw skill plus supporting PowerShell helpers for deterministic Chinese university experiment-report workflows.

Good contributions usually improve one of these areas:

- report-writing guidance in `SKILL.md`
- deterministic `docx` parsing and filling
- screenshot or image layout handling
- validation and smoke-test coverage
- documentation, examples, and reproducible demos

## Before You Change Anything

1. Read [`README.md`](README.md) for the current workflow and repository scope.
2. Keep the repository focused on OpenClaw skill usage. Do not turn it into a generic desktop app unless that is an explicit project decision.
3. Prefer deterministic document transformations over GUI automation claims.

## Development Workflow

1. Make the smallest coherent change that solves the problem.
2. Keep PowerShell scripts ASCII unless the file already uses Chinese content or Unicode is clearly required.
3. When editing scripts, preserve Windows PowerShell compatibility.
4. Add or update examples when behavior changes in a user-visible way.
5. Update [`CHANGELOG.md`](CHANGELOG.md) under `Unreleased`.

## Required Checks

Run local smoke tests before opening a PR:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-smoke-tests.ps1
```

If you changed the end-to-end local build flow, also run a focused build with [`scripts/build-report.ps1`](scripts/build-report.ps1) or [`scripts/run-e2e-sample.ps1`](scripts/run-e2e-sample.ps1).

## Testing Expectations

- Bug fixes should add or tighten an assertion when practical.
- New scripts should be covered by syntax checks and at least one smoke-test path.
- Changes to examples should keep JSON parseable and aligned with the current script contracts.

## Documentation Expectations

Update the relevant docs when you change:

- command-line flags
- output file names
- supported template patterns
- image layout behavior
- repository layout

## Pull Request Notes

A strong PR description should include:

- what changed
- why the change was needed
- how it was tested
- any remaining limitations or assumptions

## Privacy

Do not commit:

- real student identity data
- private templates
- private screenshots
- generated report outputs from real coursework

Use sanitized demo inputs whenever you need repository fixtures.
