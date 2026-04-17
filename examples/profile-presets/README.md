# Custom Profile Presets

This directory contains example report-profile presets for adjacent document types that still benefit from `-ReportProfilePath`, including snapshots that mirror a built-in profile.

Current examples:

- `weekly-report.json` (kept as a path-based example snapshot; a built-in profile also exists under `profiles/`)
- `meeting-minutes.json`

Recommended workflow:

1. Validate the example presets before using them:

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\scripts\validate-report-profiles.ps1 `
     -ProfileDir ".\examples\profile-presets" `
     -Format text
   ```

2. Try one preset directly through `-ReportProfilePath` without copying it into `profiles/`:

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\scripts\generate-report-inputs.ps1 `
     -ReportProfilePath ".\examples\profile-presets\meeting-minutes.json" `
     -CourseName "Campus Navigation App" `
     -ExperimentName "Iteration Review Meeting" `
     -StudentName "Li Si" `
     -StudentId "20261234" `
     -ClassName "SE 2302" `
     -TeacherName "Wang" `
     -ExperimentDate "2026-04-12" `
     -ExperimentLocation "GitHub + Feishu + Meeting Room" `
     -DetailLevel full `
     -OutputDir ".\tests-output\meeting-minutes-sample"
   ```

3. If the preset works for your document family, either keep using it through `-ReportProfilePath` or move a copy into `profiles/` and continue tuning aliases, captions, and prompt guidance until it is ready for built-in promotion.

Each preset can also tune `paginationRiskThresholds` so validation warnings match the document family's normal section length and screenshot density.

You can also generate sample input bundles for every preset in this directory:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-profile-preset-samples.ps1 `
  -OutputDir ".\tests-output\profile-preset-samples"
```

That command writes one subdirectory per preset, each containing:

- `prompt.txt`
- `metadata.auto.json`
- `requirements.auto.json`
- `report-inputs-summary.json`

It also writes a top-level `profile-preset-samples.md` index so you can quickly preview all generated sample bundles without opening each JSON file.

These presets are intentionally examples. `meeting-minutes` still represents an external adjacent document type, while `weekly-report` stays here as a reusable path-based sample even though a built-in profile now exists.
