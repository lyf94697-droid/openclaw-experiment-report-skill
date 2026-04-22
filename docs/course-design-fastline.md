# Course Design Fastline

This document keeps the course-design report path usable without pulling it into the default experiment-report path.

## Status

`course-design-report` is an explicit opt-in fastline, not the default experiment-report line.

Use it when the target is a school course-design report. Keep normal lab reports on `experiment-report`.

## Existing Assets

- Profile: `profiles/course-design-report.json`
- Minimal template: `examples/report-templates/course-design-report-template.docx`
- Realistic fixture: `examples/realistic-report-fixtures/course-design-full-example.docx`
- Real template notes: `docs/real-template-patterns.md`
- Local reference import for `附件6：课程设计报告模板及示例.doc`: `tests-output/real-template-references-20260419-200112/`

The existing profile already covers:

- course-design metadata: student name, id, class, teacher, course name, topic, design category, date, location
- Attachment 6 style sections: abstract, keywords, design goal, development environment, requirement analysis, design and implementation, result, issues and improvement, summary, references
- image placement defaults for implementation screenshots and result screenshots
- profile-specific prompt, requirements, field-map, image-map, style, and wrapper summaries
- paragraph composite filling for the short school template: the four body placeholders can now carry the richer Attachment 6 style section pack without changing the normal experiment-report path
- fixed text subsection numbering for course-design body sections, such as `3.1`, `4.1`, and `4.2`
- black-and-white generated PNG flowcharts for reports that need a simple process diagram
- course-design table insertion for common module, database, and test-result tables when the report content provides enough structure

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
- `.doc` school templates are accepted on Windows when Word COM is available; the build step converts them into `converted-templates/*.docx` inside the output directory before filling.
- Do not run full validation, template diagnostics, or layout checking by default; run them only when adapting a new template or when the output is visibly wrong.
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

The current fastline is still explicit opt-in, but it now carries the main `附件6`-style structure:

- abstract
- keywords
- reference list
- deeper numbered subsections
- database-design tables
- black-and-white PNG flowchart and screenshot-heavy sections
- longer full-mode body generation

Known limits:

- Generated flowcharts are PNG images, not editable Word-native shapes.
- The current flowchart style is closer to the school examples than the original simple version, but still needs real-topic tuning for complex multi-branch designs.
- If the provided school template has a grading/evaluation table, it is preserved and filled where fields exist; if the blank template does not include it, the fastline does not synthesize a new grading page by default.
- `.doc` conversion still depends on local Word/WPS COM support.
