# Fast Report Session

This helper is for day-to-day report generation speed. It avoids running the full smoke test before every single report when the repository and templates have not changed.

## Command

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\check-fast-report-session.ps1
```

If it says the cache is a hit, skip `run-smoke-tests.ps1` and go straight to the report-specific checks:

- draft validation when requirements are present
- DOCX layout check
- PDF export
- rendered preview image check

## First Run Or After Changes

Run this once when the cache is missing, stale, or critical files changed:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\check-fast-report-session.ps1 -RunSmokeWhenNeeded
```

After the smoke test passes, the script records a cache file under:

```text
tests-output/fast-report-session-cache.json
```

`tests-output/` is ignored by git, so this cache stays local.

## Template-Aware Cache

When using a real school template outside the repository, include it in the fingerprint:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\check-fast-report-session.ps1 `
  -TemplatePath "E:\实验报告\00-模板\实验报告模版1.docx","E:\新建文件夹\课程设计-模板.doc"
```

If one of those templates changes, the helper will recommend a fresh smoke run.

## Rule Of Thumb

- Same day, same repo, same templates: skip full smoke.
- Build scripts, profiles, template-frame logic, image insertion, or templates changed: run smoke once.
- Final deliverables still need PDF rendering and preview-image inspection.
