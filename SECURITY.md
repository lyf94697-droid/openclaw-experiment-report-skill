# Security Policy

## Supported Scope

This repository is a local OpenClaw skill plus PowerShell tooling for report generation and `docx` processing.

Security-relevant areas include:

- handling of local file paths
- document template processing
- web reference fetching helpers
- direct-chat wrapper scripts that execute local commands

## Reporting a Vulnerability

Do not open a public issue for a security-sensitive problem until the maintainer has had a chance to assess it.

Please report issues such as:

- path traversal or unsafe file write behavior
- unintended command execution
- unsafe handling of untrusted webpage content
- leakage of local file paths or sensitive local content
- document-processing behavior that could overwrite files outside the intended output scope

When reporting, include:

- affected script path
- exact command or prompt used
- expected behavior
- actual behavior
- minimal reproduction steps
- whether the issue requires local access, crafted input files, or a specific chat runtime

## Disclosure Guidance

- Prefer private reporting first.
- After a fix is available, coordinated public disclosure is welcome.
- Avoid including personal documents, student data, or private templates in the initial report.

## Hardening Notes

This repository already tries to reduce common risks by:

- preferring deterministic local scripts over ad-hoc chat orchestration
- treating webpage references as untrusted source material
- keeping output rooted in explicit directories
- validating image and template inputs before editing `docx` packages
