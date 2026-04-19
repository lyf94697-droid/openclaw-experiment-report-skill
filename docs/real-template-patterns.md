# Real Template Patterns

This note summarizes anonymized structure patterns observed from local university experiment-report and course-design references. The original files may contain school templates, copyrighted examples, student names, IDs, screenshots, and teacher comments, so they should stay in local `tests-output/` or private input folders and should not be committed.

Use this document as design guidance for profiles, field maps, image placement, and future template fixtures.

## Local Import Workflow

Run the importer against local references when adapting a new school template:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\import-report-template-references.ps1 `
  -Path "E:\reports\experiment-template.doc","E:\reports\course-design-template.doc","E:\reports\filled-example.pdf"
```

The importer:

- copies original references into an ignored `tests-output/real-template-references-*` directory
- unblocks copied files before conversion
- converts `.doc` references to `.docx` with WPS COM first, then Word COM if WPS is unavailable
- extracts `.docx` outlines with `extract-docx-template.ps1`
- copies `.pdf` references by default without conversion, because Word PDF conversion can hang or require interactive confirmation
- writes `import-summary.json` with conversion status, paths, table counts, and shape counts

If a PDF really needs a Word-converted outline, rerun with `-ConvertPdf` on a machine where Word can convert that file reliably. Keep the result local unless the source is safe to publish.

## Pattern 1: Single-Table Framed Experiment Report

Common shape:

- one title paragraph such as an information-college experiment report title
- one large table containing all metadata and body fields
- metadata labels such as student ID, name, class, course name, experiment content, experiment property, date, location, and equipment
- one large body cell for the whole report
- teacher comment and signature row at the bottom

Typical body headings inside the large cell:

- experiment purpose
- experiment method
- experiment principle
- experiment result
- experiment summary

Implementation implication:

- Do not treat this as a normal paragraph template.
- The field-map layer needs composite-body filling rules that can place multiple sections into one table cell.
- The style formatter should preserve the table frame, keep the body cell top-aligned, and avoid forcing row splits that create large blank areas.

## Pattern 2: Filled Single-Table Experiment Example

Common shape:

- same one-table frame as the blank experiment template
- body content is already embedded inside the large body cell
- headings often use Chinese numbered sections such as `一、实验目的`, `二、实验方法`, `三、实验结果`, and `四、实验小结`
- screenshots and captions appear as figure references inside the body text

Implementation implication:

- The extractor and diagnostics should understand that realistic final reports may be table-contained but still have a full chapter structure.
- Figure placement should support inserting screenshots inside or after the body cell instead of assuming each section has its own paragraph anchor.

## Pattern 3: Cover Plus Body Course-Design Report

Common shape:

- cover-style title page for a course-design report
- semester line
- metadata table for topic, major, class, student ID, name, and time
- body headings after the cover

Typical body headings:

- problem analysis or requirement analysis
- system design
- implementation result
- summary

Implementation implication:

- `course-design-report` should not be a renamed experiment report.
- It needs cover metadata filling plus long body-section filling.
- Generated content should emphasize design rationale, algorithms, data structures, test results, and improvement analysis.

## Pattern 4: Multi-Table Integrated Experiment Report

Common shape:

- cover page with multiple student rows, college, major, class, course name, teacher/title, semester, and fill date
- printed-office footer text
- a first table for experiment design plan
- a second table for methods, data processing, and references
- a third or fourth table for experiment phenomena, result analysis, conclusion, summary, teacher comment, and score

Typical design-plan headings:

- experiment number, name, time, and lab
- experiment purpose
- experiment principle, flow, or device diagram
- equipment and materials
- methods, steps, and precautions
- data processing method
- references

Typical report headings:

- experiment phenomena and results
- analysis and conclusion
- experiment summary
- teacher comment and score

Implementation implication:

- The profile model should allow multi-table body blocks, not only one metadata table plus paragraphs.
- Template-fit diagnostics should report missing or combined section targets at table-cell granularity.
- Future fixtures should include multi-student cover fields and teacher scoring rows.

## Pattern 5: Full Course-Design Template With Example

Common shape:

- course-design cover
- evaluation or grading table
- abstract
- keywords
- purpose and requirements
- design body with many numbered subsections
- architecture diagrams, E-R diagrams, data-flow diagrams, module tables, database schema tables, and procedure diagrams
- references and optional appendices

Implementation implication:

- Long course-design outputs need abstract and keywords support.
- Table and figure captions need independent handling, not just image captions.
- A realistic course-design profile should support chapter-depth beyond one-level headings, including database-design tables and implementation-flow diagrams.

## Pattern 6: Screenshot-Heavy Filled Networking Experiment Report

Common shape:

- school information-college experiment report frame
- metadata table
- long body with multiple protocol sections
- many figures, screenshots, routing tables, command outputs, and verification captions
- captions are often embedded as `图1`, `图2`, `图3`, and so on

Implementation implication:

- Validation should tolerate figure-heavy sections but still warn about pagination risk.
- Image placement should group related screenshots and keep captions close to images.
- The report generator should produce specific verification analysis instead of only saying that the experiment succeeded.

## What To Build Next

Most useful next upgrades:

- add explicit composite-body rules for the single-table experiment-report frame
- add multi-table integrated-experiment fixtures with anonymized synthetic content
- expand `course-design-report` with abstract, keywords, grading table, figure/table caption rules, and deeper numbered subsections
- add template-fit diagnostics for cover-only fields, grading tables, and large body cells
- keep real user references local, then commit only anonymized structure summaries and synthetic fixtures
