# Course Design Fastline

This document keeps the course-design report path usable without pulling it into the default experiment-report path.

## Status

`course-design-report` is a candidate stable line, not the default line.

Use it explicitly when the target is a school course-design report. Do not make it run by default for normal lab reports.

## Existing Assets

- Profile: `profiles/course-design-report.json`
- Minimal template: `examples/report-templates/course-design-report-template.docx`
- Realistic fixture: `examples/realistic-report-fixtures/course-design-full-example.docx`
- Real template notes: `docs/real-template-patterns.md`
- Local reference import for `附件6：课程设计报告模板及示例.doc`: `tests-output/real-template-references-20260419-200112/`

The existing profile already covers:

- course-design metadata: student name, id, class, teacher, course name, topic, design category, date, location
- core sections: design goal, development environment, requirement analysis, design and implementation, result, issues and improvement, summary
- image placement defaults for implementation screenshots and result screenshots
- profile-specific prompt, requirements, field-map, image-map, style, and wrapper summaries

## Default Command

Use the URL wrapper when source material is a web tutorial or task page:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report-from-url.ps1 `
  -ReferenceUrls "https://example.com/course-design-reference" `
  -TemplatePath "E:\reports\course-design-template.docx" `
  -ImagePaths "E:\reports\screenshots\step.png","E:\reports\screenshots\result.png" `
  -OutputDir "E:\reports\course-design-output" `
  -ReportProfileName course-design-report `
  -PipelineMode fast `
  -StyleProfile auto `
  -CourseName "软件工程综合实践" `
  -ExperimentName "校园导览小程序设计" `
  -StudentName "李亦非" `
  -StudentId "244100198" `
  -ClassName "24C" `
  -TeacherName "李老师" `
  -ExperimentProperty "课程设计" `
  -ExperimentDate "2026年4月" `
  -ExperimentLocation "实验楼"
```

Use `build-report-from-feishu.ps1` when the report body already exists or when direct chat hands over the full prompt and local paths.

## Fastline Rules

- Always pass `-ReportProfileName course-design-report`.
- Keep `-PipelineMode fast` unless the fast output is visibly wrong.
- Do not run full validation, template diagnostics, or layout checking by default.
- Do not change `experiment-report` defaults while working on course-design reports.
- Treat `附件6` as a visual and structural reference, not as a committed source artifact.

## When To Escalate To Slow Checks

Run layout or profile diagnostics only when one of these happens:

- missing final docx
- metadata not filled
- placeholder text remains
- screenshots or captions are missing
- images are placed in the wrong section
- cover/grading-table structure is visibly broken
- a new real template is being adapted for the first time

Useful commands:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\check-docx-layout.ps1 `
  -DocxPath "E:\reports\course-design-output\final.docx" `
  -ReportProfileName course-design-report `
  -ExpectedImageCount 4 `
  -ExpectedCaptionCount 4
```

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\check-report-profile-template-fit.ps1 `
  -TemplatePath "E:\reports\course-design-template.docx" `
  -ReportProfileName course-design-report
```

## Gap To Attachment 6 Quality

The current fastline can produce a submit-ready course-design report, but `附件6` quality needs extra first-class support for:

- abstract
- keywords
- grading/evaluation table
- reference list
- deeper numbered subsections
- database-design tables
- flowchart and screenshot-heavy sections
- longer body generation than the current 1400-2200 character target

These belong to the course-design candidate line. They must not be added as defaults to the normal experiment-report fastline.
