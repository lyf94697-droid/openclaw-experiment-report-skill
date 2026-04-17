# Custom Profile Presets

This directory contains example report-profile presets kept as path-based sample snapshots and customization starting points.

Current examples:

- `weekly-report.json` (a built-in profile also exists under `profiles/`)
- `meeting-minutes.json` (a built-in profile also exists under `profiles/`)
- `monthly-report.json` (currently remains a preset example; built-in promotion is still future work)

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
     -ReportProfilePath ".\examples\profile-presets\monthly-report.json" `
     -CourseName "Campus Navigation App" `
     -ExperimentName "April Monthly Progress Report" `
     -StudentName "Li Si" `
     -StudentId "20261234" `
     -ClassName "SE 2302" `
     -TeacherName "Wang" `
     -ExperimentProperty "Project Monthly Report" `
     -ExperimentDate "2026-04" `
     -ExperimentLocation "GitHub + Feishu + Local Development" `
     -DetailLevel full `
     -OutputDir ".\tests-output\monthly-report-sample"
   ```

3. If the preset works for your document family, either keep using it through `-ReportProfilePath`, switch to the matching built-in `-ReportProfileName`, or fork a copy and continue tuning aliases, captions, and prompt guidance.

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

These presets are intentionally examples. `weekly-report` and `meeting-minutes` already exist as built-in profiles, while `monthly-report` remains an external adjacent preset that can be promoted later if the shape proves out.
