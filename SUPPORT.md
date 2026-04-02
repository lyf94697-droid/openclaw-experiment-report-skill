# Support

## What to open as an issue

Open an issue when you have:

- a reproducible bug
- a missing report-generation feature
- a template compatibility problem
- a request for a new document profile or workflow
- a documentation gap that blocks normal use

## What to include

Good reports usually include:

- the script you ran
- whether you used local files, direct chat, or URL-driven generation
- PowerShell version
- whether you used Word, WPS, or only raw `docx`
- a minimal sanitized template or extracted template outline if possible
- the exact failure text

## Before opening an issue

Please check:

1. `README.md`
2. `CONTRIBUTING.md`
3. `scripts\run-smoke-tests.ps1`
4. existing issues

## Scope limits

This repository does not currently promise:

- full GUI automation of WPS or Word
- perfect support for every arbitrary school template
- stable cloud-only behavior without a local OpenClaw-compatible runtime

## Feature direction

Requests for non-experiment documents are welcome when they can be framed as reusable document profiles rather than one-off hardcoded hacks.
