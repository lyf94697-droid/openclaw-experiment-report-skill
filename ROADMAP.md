# Roadmap

## Direction

The repository should not grow into a generic office tool.

Its strongest path is narrower and more practical:

- keep one reliable document pipeline
- make that pipeline profile-driven
- expand only into document types that share the same core workflow

The core workflow remains:

1. gather references and user inputs
2. generate or clean the report body
3. fit content into a local docx template
4. plan image placement
5. insert screenshots and captions
6. apply final styling
7. validate the output

That makes this repository best suited to structured Chinese documents that:

- usually have a fixed chapter outline
- often come with a school or team template
- benefit from screenshot or evidence insertion
- need a submit-ready `docx` result

## Phase 0: Stabilize The Experiment-Report Pipeline

Before adding more document types, the current experiment-report flow should become easier to trust, debug, and reuse.

Priority work:

- integrate image-placement planning into the main wrappers instead of exposing it only as a lower-level script option
- improve template compatibility diagnostics and explain why a template did or did not fit automatic filling rules
- keep strengthening pagination and layout risk checks for grouped images, especially WPS-sensitive cases
- keep expanding regression fixtures for mixed cover/body templates, grouped screenshots, and section-end image insertion
- move more hardcoded experiment-report rules into explicit profile metadata

Implemented stabilization baseline:

- profile-specific structural validation now reports machine-readable finding codes for missing required headings, duplicate headings, section-order anomalies, empty sections, placeholder-only sections, and short sections
- pagination-risk warnings now report machine-readable codes for long sections, dense section blocks, and figure-heavy sections
- pagination-risk thresholds now live in profile metadata and flow into validation summaries, generated requirements, build summaries, and wrapper traces
- validation findings now include machine-readable remediation guidance for structural issues, content misses, and pagination risks
- `build-report.ps1`, `build-report-from-url.ps1`, and `build-report-from-feishu.ps1` now propagate validation and pagination-risk summaries into their summary JSON files and pipeline traces
- `build-report.ps1`, `build-report-from-url.ps1`, and `build-report-from-feishu.ps1` now support optional template-frame docx delivery while keeping the normal final docx path stable
- smoke coverage now includes passing fixtures, structural-risk fixtures, and end-to-end pagination-warning propagation through the local build, URL wrapper, and Feishu wrapper paths

Definition of done for this phase:

- a user can run the main wrappers with less manual intervention
- failures are explained in actionable terms
- layout regressions are easier to catch before opening WPS or Word

## Phase 1: Introduce Real Document Profiles

This phase is now partially implemented. The repository no longer treats `experiment-report` as the only first-class document type.

Introduce reusable document profiles with profile-owned metadata such as:

- required sections
- metadata labels
- caption rules
- validation thresholds
- default style profile
- prompt preset
- image placement defaults

Implemented baseline:

- `experiment-report`
- `course-design-report`
- `internship-report`
- `software-test-report`
- `deployment-report`
- `scripts/new-report-profile.ps1` now scaffolds schema-valid report profile drafts
- profile JSON now has schema-backed validation and smoke coverage
- profile-backed validation, field-map generation, image-map generation, insertion, styling, and smoke coverage

Continuing goal:

- keep one shared implementation pipeline
- switch behavior by profile metadata rather than by more script-specific branching

This phase remains the prerequisite for healthy expansion.

## Phase 2: Expand To The Closest Document Families

The first new document types should be the ones that reuse the most of the current logic.

### 1. Course Design Reports

Status: built-in profile is in place.

Why first:

- closest to experiment reports
- same need for template filling, screenshots, and structured sections
- high reuse of current field-map, image-map, and style logic

Likely profile differences:

- richer design/implementation sections
- more code explanation
- more emphasis on architecture and results analysis

### 2. Software Test Reports

Status: built-in profile is in place.

Why second:

- naturally evidence-driven
- screenshot and result insertion matter a lot
- good fit for validation-oriented output

Likely profile differences:

- test case structure
- expected vs actual result sections
- defect or issue tracking sections
- richer validation rules

### 3. Deployment Or Operations Reports

Status: built-in profile is in place.

Why third:

- still highly procedural
- often template-based
- command blocks, verification steps, and screenshots already fit the current pipeline

Likely profile differences:

- environment tables
- deployment steps
- verification and rollback sections
- operations-focused captions and summaries

## Phase 3: Expand To Adjacent School And Team Documents

After the evidence-heavy document profiles are stable, the repository can move into less screenshot-centric but still structured document types.

### 4. Internship Reports

Status: built-in profile is in place.

Good fit because they still benefit from:

- fixed chapter structures
- metadata fields
- school templates
- optional screenshot evidence

### 5. Weekly And Monthly Reports

Status: built-in `weekly-report` profile is in place; `monthly-report` remains future work.

Good fit when the structure is profile-driven rather than free-form:

- progress
- completed work
- problems
- next steps

### 6. Meeting Minutes With Template Filling

Good fit later, not sooner:

- less image-heavy
- more table- and action-item-oriented
- still compatible with template filling and profile-specific validation

## Recommended Profile Order

The expansion order should stay disciplined:

1. `experiment-report` stabilization
2. `course-design-report`
3. `software-test-report`
4. `deployment-report`
5. `internship-report`
6. `weekly-report`
7. `meeting-minutes`
8. `monthly-report`

The first six are now built-in. The next adjacent candidate in this order is `meeting-minutes`, followed by `monthly-report` if a reusable preset proves out.

## Supporting Platform Work

Alongside new profiles, the repository should keep investing in a few shared capabilities.

### Template Fit And Diagnostics

- better extraction of template structure
- clearer field-lock and blank-block detection
- explicit reasons when automatic filling falls back or skips content

### Image Planning And Layout

- stronger automatic section inference
- confidence scoring for image placement
- grouped-image chunking such as stable 2x2 blocks
- more layout strategies beyond a single row-table approach

### Validation And Risk Detection

- richer output checks per profile
- more profile-specific pagination-risk presets for document families with very different length and screenshot density
- more Word/WPS-sensitive layout heuristics for image-heavy reports
- richer remediation guidance for profile-specific and template-fit diagnostics

### Prompt Assets And Examples

- one-shot prompt examples
- plan-first prompt examples
- profile-specific prompt presets
- more realistic end-to-end fixtures

Implemented baseline:

- `profiles/weekly-report.json`
- `examples/profile-presets/weekly-report.json`
- `examples/profile-presets/meeting-minutes.json`
- `scripts/run-profile-preset-samples.ps1`
- `scripts/analyze-roadmap-next-step.ps1`
- `.github/workflows/roadmap-triage.yml`
- schema-backed `paginationRiskThresholds` in report profiles and preset examples

`meeting-minutes` still stays outside `profiles/` on purpose so adjacent document types can be prototyped through `-ReportProfilePath` before they are promoted to built-in profiles. `weekly-report` has crossed that threshold and now lives in `profiles/`, while the preset copy remains as a reusable example snapshot and sample-runner input.

## Non-Goals For Now

- turning the repository into a general-purpose office suite
- supporting arbitrary unstructured documents with one giant universal prompt
- expanding to unrelated document types before the profile model is stable
- shipping a GUI before the profile-driven pipeline is mature

## Practical Success Metric

This repository is on the right track if it becomes:

- easy to run for repeatable school and engineering documents
- predictable when templates and screenshots are involved
- explainable when automation falls short
- extensible by adding profiles rather than rewriting the pipeline
