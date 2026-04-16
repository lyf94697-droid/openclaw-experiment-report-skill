# Custom Profile Presets

This directory contains example report-profile presets for adjacent document types that are not promoted to built-in profiles yet.

Current examples:

- `weekly-report.json`
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
     -ReportProfilePath ".\examples\profile-presets\weekly-report.json" `
     -CourseName "校园导览小程序" `
     -ExperimentName "第 6 周迭代周报" `
     -StudentName "李四" `
     -StudentId "20261234" `
     -ClassName "软工 2302" `
     -TeacherName "王老师" `
     -ExperimentDate "第 6 周" `
     -ExperimentLocation "GitHub + 飞书 + 本地开发环境" `
     -DetailLevel full `
     -OutputDir ".\tests-output\weekly-preset-sample"
   ```

3. If the preset works for your document family, either keep using it through `-ReportProfilePath` or move a copy into `profiles/` and continue tuning aliases, captions, and prompt guidance.

These presets are intentionally examples, not built-in defaults. They are meant to shorten the path from "this repo is close to my document type" to "I can prototype my own profile today".
