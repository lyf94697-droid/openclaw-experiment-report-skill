# OpenClaw Experiment Report Skill

[![Quality Checks](https://github.com/lyf94697-droid/openclaw-experiment-report-skill/actions/workflows/quality.yml/badge.svg)](https://github.com/lyf94697-droid/openclaw-experiment-report-skill/actions/workflows/quality.yml)
[![Smoke Tests](https://github.com/lyf94697-droid/openclaw-experiment-report-skill/actions/workflows/smoke-tests.yml/badge.svg)](https://github.com/lyf94697-droid/openclaw-experiment-report-skill/actions/workflows/smoke-tests.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

Portable OpenClaw skill for writing Chinese university experiment reports from experiment topics, requirements, code, screenshots, data, tutorial pages, and blank WPS/Word/docx templates.

See [`demo/README.md`](demo/README.md) for GitHub-friendly preview assets and a ready-made 2x2 image layout example.
See [`docs/GITHUB_LAUNCH.md`](docs/GITHUB_LAUNCH.md) for the recommended GitHub description, topics, release notes, and social-preview setup.

## At A Glance

This repository is for people who want to go from:

- an experiment topic or tutorial URL
- local screenshots
- a blank `docx` template

to:

- a generated Chinese lab report body
- a deterministically filled `docx`
- grouped screenshots with captions
- a cleaner final submission-style document

## Preview

| Step Screenshot | Result Screenshot |
| --- | --- |
| ![Network config preview](demo/assets/step-network-config.png) | ![Ping result preview](demo/assets/result-ping.png) |

For a full 2x2 preview block, see [`demo/README.md`](demo/README.md).

## Typical Workflow

1. Provide a topic, tutorial URL, screenshots, or an existing report body.
2. Generate a complete Chinese experiment report body through OpenClaw.
3. Fill a blank `docx` template with deterministic field and body mapping.
4. Insert grouped screenshots with captions and apply final submission-style formatting.

## What It Does

- Writes the full report before touching template formatting.
- Extracts local `docx` template structure deterministically before field mapping.
- Generates a deterministic `docx` field map from the finished report body plus template structure.
- Generates a deterministic image map from a filled `docx` copy plus local screenshots or photos.
- Fills common `docx` report templates deterministically from a field map.
- Fills both short fields and common section-body paragraphs in `docx` templates.
- Detects option rows such as `实验性质： ①综合性实验 ②设计性实验 ③验证性实验` and marks the selected item with `√`.
- Inserts local screenshots or experiment photos into a filled `docx` copy with captions.
- Supports both stacked screenshots and grouped row layouts such as two images per row.
- Supports shared row-group anchors, so one 2x2 or 2x3 image block can stay together under a single section or explicit anchor.
- Applies a final report-style pass to `docx` titles, headings, metadata lines, body paragraphs, captions, and image blocks.
- Supports both direct anchors such as `P8` / `T1R6C1` and stable section anchors such as `实验步骤` / `实验结果`.
- Fits content into blank templates through field-by-field mapping.
- Uses screenshots as factual evidence and generates figure captions plus insertion guidance.
- Supports tutorial or CSDN article driven workflows without copying long passages verbatim.
- Can fetch public tutorial pages and append them as structured reference material before report generation.
- Includes a chat-friendly one-shot wrapper so Feishu or other direct-chat runtimes can call one stable local script instead of improvising the whole pipeline step by step.
- Includes PowerShell helpers for installation, docx template analysis, article extraction, and smoke tests.

## Current Scope

This repository is designed for OpenClaw users. It is not a standalone desktop app.

Included:

- OpenClaw skill instructions in `SKILL.md`
- reusable reference files in `references/`
- practical helpers in `scripts/`
- example prompts in `examples/`
- repository governance files for GitHub collaboration

Not included yet:

- guaranteed WPS or Word GUI auto-fill
- fully automatic binary template editing without user review
- non-OpenClaw standalone application packaging

## Repository Health

This repository now includes the basic files expected from a maintainable open-source project:

- `README.md`
- `LICENSE`
- `CONTRIBUTING.md`
- `CHANGELOG.md`
- `CODE_OF_CONDUCT.md`
- `SECURITY.md`
- `SUPPORT.md`
- issue templates and a PR template under `.github/`
- Windows CI workflows under `.github/workflows/`

That does not magically make the project finished, but it does make the repo reviewable, clonable, and easier for outside contributors to use without guessing the workflow.

## Future Direction

The current implementation is intentionally focused on experiment reports, but the core pipeline is broader than that:

- reference gathering
- body generation
- validation
- template filling
- image mapping
- final styling

That means the project can grow into a more general document-generation toolkit, as long as future document types are added as reusable profiles instead of one-off hardcoded branches.

See [`ROADMAP.md`](ROADMAP.md) for the planned path from an experiment-report skill to a profile-driven document workflow.

## Install

Option 1: clone into the personal Agent Skills directory that OpenClaw actually loads.

```powershell
git clone <your-repo-url> "$env:USERPROFILE\.agents\skills\experiment-report"
```

Option 2: run the bundled installer from a checked-out repo.

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install-skill.ps1
```

Then run `openclaw skills list` and confirm `experiment-report` appears. Restart OpenClaw or reload skills only if your runtime caches the skill list.

The skill is intended to trigger when prompts mention words such as `实验报告`, `实验模板`, `WPS模板`, `Word模板`, `docx模板`, or `lab report`.
For the most deterministic behavior, especially on busy long-lived sessions, start your request with `/experiment-report`.

## Quick Start

If you want the most stable chat-friendly local entry point, use the Feishu wrapper:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report-from-feishu.ps1 `
  -ReferenceUrls "https://blog.csdn.net/..." `
  -CourseName "计算机网络" `
  -ExperimentName "局域网搭建与常用 DOS 命令使用" `
  -TemplatePath "E:\reports\template.docx" `
  -StudentName "张三" `
  -StudentId "20260001" `
  -ClassName "计科 2201" `
  -ImagePaths "E:\reports\step-1.png","E:\reports\step-2.png","E:\reports\result-1.png","E:\reports\result-2.png" `
  -OutputDir "E:\reports\final-output"
```

After the first generated run, `build-report-from-feishu.ps1` remembers the last `CourseName` and `ExperimentName`, so later runs can omit one or both if they stay the same.

That wrapper keeps the final deliverables in the chosen output directory and stores intermediate files under `artifacts\`, which makes direct-chat runs much easier to inspect and much less cluttered.

If you want to go directly from a tutorial URL to a final `docx` without the extra wrapper layer, use the URL build entry point:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report-from-url.ps1 `
  -ReferenceUrls "https://blog.csdn.net/..." `
  -CourseName "计算机网络" `
  -ExperimentName "局域网搭建与常用 DOS 命令使用" `
  -TemplatePath "E:\reports\template.docx" `
  -StudentName "张三" `
  -StudentId "20260001" `
  -ClassName "计科 2201"
```

The same remembered-default behavior also applies to `build-report-from-url.ps1`: the first generated run sets the defaults, and later runs can omit `-ExperimentName` when the experiment stays the same.

That one command can fetch the public tutorial page, generate a report body through OpenClaw, clean the saved report text, auto-generate metadata and baseline validation rules, fill the template, and emit a final style-formatted `docx`.

If you already have a finished report body plus a blank template, use the local build entry point:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report.ps1 `
  -TemplatePath "E:\reports\template.docx" `
  -ReportPath ".\examples\sample-report.txt" `
  -MetadataPath ".\examples\docx-report-metadata.json" `
  -ImageSpecsPath ".\examples\docx-image-specs-row.json" `
  -RequirementsPath ".\examples\e2e-sample-requirements.json" `
  -StyleFinalDocx `
  -StyleProfile auto
```

That one command can validate the report text, generate a field map, fill the template, insert grouped screenshots, and optionally emit a final style-formatted `docx`.

## Practical Scripts

### `scripts\install-skill.ps1`

- Copies this repo into `$HOME\.agents\skills\experiment-report` by default, which is the personal skills directory OpenClaw loads.
- Supports `-Force` to back up an existing install and reinstall cleanly.
- Stores force-reinstall backups outside the scanned `skills/` directory so OpenClaw does not accidentally load backup copies as duplicate skills.

### `scripts\build-report.ps1`

- Provides a single local entry point for the deterministic `docx` build pipeline.
- Starts from a finished report body plus a blank template.
- Optionally validates the report against explicit requirements before packaging outputs.
- Generates the field map, filled template, image map, image-embedded `docx`, and optional final style-formatted `docx`.
- Writes a machine-readable `summary.json` so CI jobs or wrapper scripts can consume the output paths.
- When `-StyleFinalDocx` is enabled, supports `-StyleProfile auto|default|compact|school`.
- Also supports `-StyleProfilePath` to load a custom JSON style profile file.
- `auto` chooses `compact` for cover-table templates, `school` for paragraph-cover templates, and otherwise falls back to `default`.

Example:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report.ps1 `
  -TemplatePath "E:\reports\template.docx" `
  -ReportPath "E:\reports\report.txt" `
  -MetadataPath ".\examples\docx-report-metadata.json" `
  -ImageSpecsPath ".\examples\docx-image-specs-row.json" `
  -StyleFinalDocx `
  -StyleProfile auto `
  -StyleProfilePath ".\examples\style-profile-custom.json"
```

### `scripts\build-report-from-url.ps1`

- Provides a higher-level local entry point for the URL-driven workflow.
- Fetches one or more public tutorial pages and appends them as reference material before report generation.
- Generates a report body through OpenClaw gateway chat, saves both the raw and cleaned report text, and trims a known mojibake tail pattern when it appears in the local runtime output.
- Can auto-generate a baseline `metadata.auto.json` and `requirements.auto.json` so users do not need to prepare those files for common runs.
- Supports `-DetailLevel standard|full`; `full` is the default and asks OpenClaw for a more substantial, less terse report body.
- When `-TemplatePath` is provided, continues into the deterministic `docx` build pipeline and writes a friendly final filename such as `学号-姓名-实验名-最终版.docx`.

Example:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report-from-url.ps1 `
  -ReferenceUrls "https://blog.csdn.net/..." `
  -CourseName "Windows 网络管理 / 网络操作系统 / 服务器配置与管理" `
  -ExperimentName "Windows Server 2022/2025 搭建 Web 服务器" `
  -TemplatePath "E:\reports\template.docx" `
  -StudentName "张三" `
  -StudentId "20260001" `
  -ClassName "计科 2201" `
  -TeacherName "李老师" `
  -DetailLevel full
```

### `scripts\build-report-from-feishu.ps1`

- Provides a direct-chat-oriented one-shot local wrapper around the stable report pipeline.
- Accepts either `-ReportPath` for an existing finished body, or generation inputs such as `-ReferenceUrls`, `-ReferenceTextPaths`, `-PromptPath`, or `-PromptText`.
- Keeps the output root tidy by copying the final `docx` and `report.txt` to the selected output directory while storing intermediate files under `artifacts\`.
- Reuses the same metadata, image, validation, and final-style options as the lower-level scripts.
- Defaults to `-DetailLevel full` when it needs to generate the report body.
- Remembers the last generated `CourseName` and `ExperimentName` under the local agents home, so repeated runs do not have to keep retyping them.

Example:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report-from-feishu.ps1 `
  -ReferenceUrls "https://blog.csdn.net/..." `
  -CourseName "Windows 网络管理 / 网络操作系统 / 服务器配置与管理" `
  -ExperimentName "Windows Server 2022/2025 搭建 Web 服务器" `
  -TemplatePath "E:\reports\template.docx" `
  -StudentName "张三" `
  -StudentId "20260001" `
  -ClassName "计科 2201" `
  -TeacherName "李老师" `
  -ImagePaths "E:\reports\step-1.png","E:\reports\step-2.png","E:\reports\result-1.png","E:\reports\result-2.png" `
  -OutputDir "E:\reports\final-output"
```

### `scripts\extract-docx-template.ps1`

- Reads a local `.docx` template and outputs a markdown or JSON outline.
- Captures paragraph order, table cells, and likely fillable fields.
- Reduces template guesswork before OpenClaw writes a field mapping.

Example:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\extract-docx-template.ps1 -Path "E:\reports\template.docx"
```

JSON output:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\extract-docx-template.ps1 -Path "E:\reports\template.docx" -Format json
```

### `scripts\apply-docx-field-map.ps1`

- Fills common report templates from a JSON field map.
- Supports two input styles:
  - label keys such as `课程名称`, `姓名`, `学号`
  - location keys such as `P2`, `T1R1C2`
- Supports scalar values, paragraph arrays, and objects such as `{ "mode": "after", "paragraphs": [...] }`.
- Label keys are conservative by default: they only fill blank or placeholder-shaped targets.
- Location keys are explicit: they can overwrite an existing paragraph or cell.
- Common section headings such as `实验目的` or `实验步骤` can keep the heading paragraph and fill the blank paragraph after it.

Example:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\apply-docx-field-map.ps1 `
  -TemplatePath "E:\reports\template.docx" `
  -MappingPath ".\examples\docx-field-map.json" `
  -OutPath "E:\reports\template.filled.docx"
```

See [`examples/docx-field-map.json`](examples/docx-field-map.json) for the expected JSON shape.

Block example:

```json
{
  "实验目的": {
    "mode": "after",
    "paragraphs": [
      "掌握网络拓扑搭建流程。",
      "理解常用 DOS 命令的作用。"
    ]
  },
  "实验步骤": [
    "配置虚拟机网络参数。",
    "执行 ipconfig 与 ping 验证连通性。"
  ]
}
```

Recommended workflow:

1. Run `extract-docx-template.ps1` on the blank template.
2. Let OpenClaw generate the full report body first.
3. Run `generate-docx-field-map.ps1` from the blank template plus the finished report body.
4. For section-body paragraphs, prefer `paragraphs` arrays and use `mode: "after"` when the template has a fixed heading followed by a blank paragraph.
5. Review the generated JSON and adjust only when the template is unusual.
6. Run `apply-docx-field-map.ps1` to produce a filled copy.
7. If screenshots should be embedded into the final `docx`, run `generate-docx-image-map.ps1` from the filled copy plus your local image list.
8. Run `insert-docx-images.ps1` on the filled copy.
9. If the final copy needs cleaner report typography, run `format-docx-report-style.ps1` on the filled or image-embedded docx.

### `scripts\generate-docx-field-map.ps1`

- Reads a `.docx` template outline, a finished report, and optional metadata JSON.
- Generates a deterministic field map wrapper with `summary`, `fieldMap`, and `notes`.
- Emits section-body mappings in the shape that `apply-docx-field-map.ps1` already accepts.
- Preserves fixed section headings by generating `mode: "after"` when the template uses heading-plus-blank-paragraph layout.
- The generated JSON can be passed directly to `apply-docx-field-map.ps1`; the fill script accepts the wrapper and automatically uses its `fieldMap` property.

Example:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\generate-docx-field-map.ps1 `
  -TemplatePath "E:\reports\template.docx" `
  -ReportPath ".\tests-output\report.txt" `
  -MetadataPath ".\examples\docx-report-metadata.json" `
  -OutFile ".\tests-output\generated-field-map.json"
```

The repository example metadata file is [`examples/docx-report-metadata.json`](examples/docx-report-metadata.json).

### `scripts\generate-docx-image-map.ps1`

- Reads a filled `.docx` plus either image specs JSON or plain image paths.
- Generates a deterministic image-map wrapper with `summary`, `images`, and `notes`.
- Prefers stable section anchors such as `实验步骤` or `实验结果`, so image placement does not drift when paragraph counts change after template filling.
- Can infer a section and caption from the file name when you only provide image paths.
- Preserves optional image `layout` objects such as `{ "mode": "row", "columns": 2, "group": "results-grid" }`.
- When one grouped row block spans multiple sections, it auto-emits a shared `layout.groupAnchor` so the whole block stays together.
- The generated JSON can be passed directly to `insert-docx-images.ps1`.

Example:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\generate-docx-image-map.ps1 `
  -DocxPath "E:\reports\template.filled.docx" `
  -ImageSpecsPath ".\examples\docx-image-specs.json" `
  -OutFile ".\tests-output\generated-image-map.json"
```

Plain image-path example:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\generate-docx-image-map.ps1 `
  -DocxPath "E:\reports\template.filled.docx" `
  -ImagePaths "E:\reports\images\step-1.png","E:\reports\images\result-1.png"
```

The repository example image specs file is [`examples/docx-image-specs.json`](examples/docx-image-specs.json).
For grouped row layout, see [`examples/docx-image-specs-row.json`](examples/docx-image-specs-row.json).

### `scripts\insert-docx-images.ps1`

- Inserts local `png/jpg/jpeg/gif/bmp` images into an existing `.docx`.
- Supports direct anchors:
  - paragraph anchors such as `P8`
  - table-cell anchors such as `T1R6C1`
- Supports stable section anchors such as `实验步骤` or `实验结果`; the script resolves them against the current filled `docx` and inserts after the matching section heading paragraph, even inside table cells.
- Adds centered image paragraphs plus optional centered captions.
- Supports grouped row layouts via `layout.mode = "row"` and `layout.columns = 2`, producing a borderless image table for side-by-side screenshots.
- Supports `layout.groupAnchor` for grouped row layouts, so one shared 2x2 or 2x3 block can be pinned under a single section or explicit anchor even when the member images have different section tags.
- Keeps image insertion deterministic by writing standard OpenXML media entries and relationships directly.
- Works well as the last step after `apply-docx-field-map.ps1`.

Example:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\insert-docx-images.ps1 `
  -DocxPath "E:\reports\template.filled.docx" `
  -MappingPath ".\examples\docx-image-map.json" `
  -OutPath "E:\reports\template.filled.images.docx"
```

See [`examples/docx-image-map.json`](examples/docx-image-map.json) for the expected JSON shape.
For grouped row layout, see [`examples/docx-image-map-row.json`](examples/docx-image-map-row.json).

### `scripts\format-docx-report-style.ps1`

- Applies a simple report-style formatting pass to an existing `.docx`.
- Centers report titles, figure captions, and image paragraphs.
- Adds bold spacing rules for common report section headings.
- Keeps metadata paragraphs left-aligned without first-line indentation.
- Applies a configurable first-line indent and line spacing to normal body paragraphs.
- Supports built-in style profiles:
  - `auto`: picks a built-in profile from the current document structure
  - `default`: balanced report spacing
  - `compact`: tighter title, heading, body, caption, and image spacing
  - `school`: looser submission-style spacing
- Also supports `-ProfilePath` for a custom JSON profile file.
- Merge order is: resolved built-in profile -> custom profile file -> explicit command-line `*Twips` overrides.
- Works well as the last step after template filling and optional image insertion.

Example:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\format-docx-report-style.ps1 `
  -DocxPath "E:\reports\template.filled.images.docx" `
  -OutPath "E:\reports\template.filled.images.styled.docx" `
  -Profile auto `
  -Overwrite
```

Custom profile example:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\format-docx-report-style.ps1 `
  -DocxPath "E:\reports\template.filled.images.docx" `
  -OutPath "E:\reports\template.filled.images.custom.docx" `
  -ProfilePath ".\examples\style-profile-custom.json" `
  -Overwrite
```

See [`examples/style-profile-custom.json`](examples/style-profile-custom.json) for the supported shape:

- `baseProfile`: optional `auto|default|compact|school`
- `settings`: optional overrides for `BodyFirstLineTwips`, `BodyLineTwips`, `BodyAfterTwips`, `HeadingBeforeTwips`, `HeadingAfterTwips`, `CaptionAfterTwips`, `TitleAfterTwips`, `ImageBeforeTwips`, and `ImageAfterTwips`

### `scripts\prepare-report-prompt.ps1`

- Builds a final report-generation prompt from a base prompt plus optional local reference text files or public tutorial URLs.
- Appends extracted reference material with explicit anti-copying and fact-priority guidance.
- Lets the same prompt flow be reused by `run-e2e-sample.ps1` and `generate-report-chat.ps1`.

Example:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\prepare-report-prompt.ps1 `
  -PromptPath ".\examples\e2e-sample-prompt.txt" `
  -ReferenceTextPaths ".\local-inputs\tutorial-reference.txt" `
  -ReferenceUrls "https://blog.csdn.net/..." `
  -OutFile ".\tests-output\prepared-prompt.txt"
```

### `scripts\fetch-web-article.ps1`

- Uses the local OpenClaw browser to extract article text from a public webpage.
- Prefers article or main-content selectors and falls back to `document.body.innerText`.

Example:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\fetch-web-article.ps1 -Url "https://example.com/tutorial"
```

### `scripts\fetch-csdn-article.ps1`

- Compatibility wrapper around `fetch-web-article.ps1` for existing CSDN-focused workflows.

Example:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\fetch-csdn-article.ps1 -Url "https://blog.csdn.net/..."
```

Supported environment variables:

- `OPENCLAW_CMD`: absolute path to `openclaw.cmd` if it is not on `PATH`
- `OPENCLAW_BROWSER_PROFILE`: browser profile name, default is `openclaw`

### `scripts\self-check.ps1`

- Verifies the OpenClaw CLI path and browser profile status.

### `scripts\run-smoke-tests.ps1`

- Runs repeatable local smoke tests against syntax, docx extraction, installation, and optional OpenClaw self-check.
- Also verifies deterministic docx field-map generation, image-map generation, image insertion, final docx style formatting, and docx fill behavior.

Example:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-smoke-tests.ps1
```

### `scripts\validate-report-draft.ps1`

- Validates a generated report against explicit requirements.
- Checks required sections, minimum length, required keywords, forbidden phrases, and optional figure references.
- Works with either a report file or inline text.

Example:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\validate-report-draft.ps1 `
  -Path ".\tests-output\report.txt" `
  -RequirementsPath ".\examples\e2e-sample-requirements.json" `
  -Format json
```

### `scripts\run-e2e-sample.ps1`

- Installs the skill, asks OpenClaw to generate a full sample report, and validates the output end to end.
- Records raw agent output, extracted report text, validation JSON, and a summary JSON under `tests-output/`.
- When `-TemplatePath` is provided, it also generates a `docx` field map, fills the template, and saves an extracted outline of the filled result for review.
- When you also provide `-ImageSpecsPath`, `-ImageSpecsJson`, or `-ImagePaths`, it automatically generates an image map from the filled copy and writes a final docx with embedded images.
- Supports `-ReferenceTextPaths` and `-ReferenceUrls`, so local tutorial notes or public webpages can be appended to the generated request before report generation.
- When you add `-StyleFinalDocx`, it runs `format-docx-report-style.ps1` on the final filled copy, or on the image-embedded copy when images were included.
- When style formatting is enabled, `-StyleProfile auto|default|compact|school` and optional `-StyleProfilePath` are forwarded to the formatter.
- Defaults to `guided-chat` mode, which sends the prompt through the local OpenClaw gateway chat API together with the skill policy. This is the most stable path on Windows because it avoids `openclaw.cmd --message` newline and encoding issues.
- Still supports `native-agent` mode when you want to diagnose native slash-command behavior.
- Resets the target OpenClaw session before the run by default so previous chat history does not contaminate the sample.

Example:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-e2e-sample.ps1 `
  -Agent gpt `
  -Mode guided-chat `
  -ReferenceUrls "https://blog.csdn.net/..." `
  -TimeoutSeconds 300
```

Template fill example:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-e2e-sample.ps1 `
  -Agent gpt `
  -Mode guided-chat `
  -TemplatePath "E:\reports\template.docx" `
  -MetadataPath ".\examples\docx-report-metadata.json" `
  -TimeoutSeconds 300
```

Template + images example:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-e2e-sample.ps1 `
  -Agent gpt `
  -Mode guided-chat `
  -TemplatePath "E:\reports\template.docx" `
  -MetadataPath ".\examples\docx-report-metadata.json" `
  -ImagePaths "E:\reports\images\step-1.png","E:\reports\images\result-1.png" `
  -StyleFinalDocx `
  -StyleProfile auto `
  -StyleProfilePath ".\examples\style-profile-custom.json" `
  -TimeoutSeconds 300
```

### `scripts\generate-report-chat.ps1`

- Sends a report-generation request directly through the local OpenClaw gateway chat API.
- Injects the repository `SKILL.md` policy into the message so report generation remains usable even when native skill slash commands are unreliable in the local CLI path.
- Supports `-ReferenceTextPaths` and `-ReferenceUrls` so tutorial material can be appended to the final request automatically.
- Writes the final report body to a target file and can reset the session before the run.

### `scripts\reset-openclaw-session.ps1`

- Resets a specific OpenClaw session key through the local gateway using the configured gateway token.
- Useful when repeated tests would otherwise inherit old slash-command instructions or stale chat history.

Example:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\reset-openclaw-session.ps1 -SessionKey "agent:gpt:main"
```

Example:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-e2e-sample.ps1 `
  -Agent gpt `
  -Thinking medium `
  -TimeoutSeconds 300
```

## Feishu Direct Chat

Feishu direct chat can be made reliable, but it is still less deterministic than the local script entry points. The safe position is:

- Do not assume the chat runtime can read local Windows paths unless it actually proves it.
- Put path access checking into the prompt itself.
- Ask it to either finish the local workflow or explicitly say which path was inaccessible.
- Give explicit image grouping such as `图1、图2属于实验步骤；图3、图4属于实验结果` instead of relying on pure guessing.
- If it needs temporary JSON with Chinese captions, paths, or section names, make it use explicit UTF-8 output instead of generic ad-hoc file writes.
- If it needs multiple PowerShell actions, prefer separate commands or `;` instead of `&&`.

The most stable direct-chat pattern is now: ask Feishu to run `scripts\build-report-from-feishu.ps1` with explicit arguments, instead of asking it to improvise every intermediate JSON and shell step on its own.

If Feishu direct chat can see uploaded images and you also provide local image paths, use them together:

- Let the uploaded images serve as the semantic reference, so the model can describe what each screenshot actually shows.
- Let the local image paths serve as the file source for the final `docx`, so the wrapper script can still embed the images deterministically.
- If the attachment view and the local file disagree, trust the uploaded image content first and then fix the local file selection.
- If Feishu direct chat exposes prompt-injected attachment note lines such as `[media attached 1/4: media/inbound/example.png (image/png)]`, those injected paths can also be reused as `-ImagePaths` for the final `docx`.

Recommended Feishu prompt:

```text
请在仓库根目录直接运行本地脚本 .\scripts\build-report-from-feishu.ps1，不要自己临时拼接一长串中间 JSON。

如果你已经在之前的生成任务里设置过课程名和实验名，而且这次不变，那么 `-CourseName` 和 `-ExperimentName` 可以省略。

如果你当前工作目录不是仓库根目录，请先切换到：
E:\游戏\openclaw-experiment-report-skill

先确认下面这些路径或网址能访问；如果有任何一个不能访问，请明确说出“无法访问哪个路径或网址”，不要假装已经读取成功。

然后直接运行：

powershell -ExecutionPolicy Bypass -File .\scripts\build-report-from-feishu.ps1 `
  -ReferenceUrls "https://blog.csdn.net/你的文章链接" `
  -CourseName "计算机网络" `
  -ExperimentName "局域网搭建与常用 DOS 命令使用" `
  -TemplatePath "E:\实验报告\实验报告模版1.docx" `
  -StudentName "李亦非" `
  -StudentId "244100198" `
  -ClassName "24c" `
  -TeacherName "李老师" `
  -ExperimentProperty "③验证性实验" `
  -ExperimentDate "2026年4月2日" `
  -ExperimentLocation "睿智楼四栋212" `
  -ImagePaths "E:\实验报告\step-1.png","E:\实验报告\step-2.png","E:\实验报告\result-1.png","E:\实验报告\result-2.png" `
  -OutputDir "E:\实验报告\新建文件夹" `
  -StyleProfile auto `
  -DetailLevel full

要求：
- 图1、图2属于实验步骤
- 图3、图4属于实验结果
- 四张图按 2x2 连续图片块排版
- 实验性质这一项要表现为勾选 ③验证性实验
- 结果部分以我的截图和已知事实为准
- 正文要比简略版更充实，但不要编造不存在的数据、截图细节或报错
```

The same prompt is also saved in [`examples/feishu-one-shot-script-prompt.md`](examples/feishu-one-shot-script-prompt.md). A lower-level direct-file prompt remains available in [`examples/feishu-local-files-prompt.md`](examples/feishu-local-files-prompt.md).

Hybrid attachment + local-path prompt:

```text
我会在飞书里直接上传实验截图，同时也会给你这些图片在我电脑上的本地路径。

请这样处理：
- 先以我上传的图片附件为准，识别每张图到底展示了什么内容
- 再使用本地路径作为最终 docx 插图文件来源
- 如果你看到了附件，但本地路径无法访问，请明确说出哪个路径不能访问
- 如果本地路径可访问，就直接运行本地脚本生成最终 docx

工作目录：
E:\游戏\openclaw-experiment-report-skill

请运行：

powershell -ExecutionPolicy Bypass -File .\scripts\build-report-from-feishu.ps1 `
  -ReferenceUrls "https://blog.csdn.net/你的文章链接" `
  -CourseName "计算机网络" `
  -ExperimentName "局域网搭建与常用 DOS 命令使用" `
  -TemplatePath "E:\实验报告\实验报告模版1.docx" `
  -StudentName "李亦非" `
  -StudentId "244100198" `
  -ClassName "24c" `
  -TeacherName "李老师" `
  -ExperimentProperty "③验证性实验" `
  -ExperimentDate "2026年4月2日" `
  -ExperimentLocation "睿智楼四栋212" `
  -ImagePaths "E:\实验报告\step-1.png","E:\实验报告\step-2.png","E:\实验报告\result-1.png","E:\实验报告\result-2.png" `
  -OutputDir "E:\实验报告\新建文件夹" `
  -StyleProfile auto `
  -DetailLevel full

额外要求：
- 图1、图2属于实验步骤
- 图3、图4属于实验结果
- 四张图按 2x2 连续图片块排版
- 结果部分优先根据我上传的图片附件来写
- 最终插图仍使用我给出的本地路径文件
```

Uploaded attachments only, but still insert them into the final `docx`:

- If the runtime injects attachment note lines such as `[media attached ...]` into the prompt, extract the paths from those lines and pass them into `-ImagePaths`.
- The image scripts now resolve OpenClaw-staged relative attachment paths such as `media/inbound/example.png` against the session workspace and the repo parent directory, so the final insertion flow does not require you to handwrite absolute paths every time.
- If the runtime does not expose any actual attachment file path, say that clearly instead of pretending the attachment was embedded.

The uploaded-image prompt is saved in [`examples/feishu-uploaded-images-docx-prompt.md`](examples/feishu-uploaded-images-docx-prompt.md).

## Recommended Prompt Patterns

Basic report:

```text
写一份完整的实验报告。
课程名：计算机网络
实验名：局域网搭建与常用 DOS 命令使用
要求：正式、真实、适合大学课程提交。
先输出完整正文，不要先做模板映射。
```

Template mode:

```text
根据这个空白模板完成实验报告：
E:\实验报告\实验报告模板.docx

课程名：计算机网络
实验名：局域网搭建与常用 DOS 命令使用
姓名：张三
学号：20260001
班级：计科 2201

请先用模板解析结果识别字段顺序，再写完整实验报告，最后输出字段映射。
```

Tutorial plus screenshots:

```text
根据这个教程页面和我的实验截图写实验报告，不要照抄原文：
https://blog.csdn.net/...

课程名：计算机网络
实验名：局域网搭建与常用 DOS 命令使用
截图路径：
- E:\实验报告\截图1.png
- E:\实验报告\截图2.png
- E:\实验报告\截图3.png

结果部分请以截图为准，并为每张图给出图号、图题和插入位置建议。
```

## Repository Layout

```text
openclaw-experiment-report-skill/
├─ .gitattributes
├─ SKILL.md
├─ .gitignore
├─ CHANGELOG.md
├─ CONTRIBUTING.md
├─ agents/openai.yaml
├─ demo/
├─ references/
├─ scripts/
├─ examples/
├─ README.md
└─ LICENSE
```

## Development

Run the repository smoke tests before publishing changes:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-smoke-tests.ps1
```

Contributor workflow and review expectations are documented in [`CONTRIBUTING.md`](CONTRIBUTING.md). Ongoing repository history is tracked in [`CHANGELOG.md`](CHANGELOG.md).

## Stability Notes

- `docx` template structure is handled deterministically by `extract-docx-template.ps1`.
- `docx` field-map generation is handled deterministically by `generate-docx-field-map.ps1`.
- Common `docx` template filling is handled deterministically by `apply-docx-field-map.ps1`.
- Common section-heading plus blank-paragraph templates are handled deterministically by `apply-docx-field-map.ps1`.
- Installation is handled deterministically by `install-skill.ps1`.
- Repeated local verification is handled by `run-smoke-tests.ps1`.
- The stable auto-fill path is intentionally scoped to common paragraph, section-heading, and table-field templates. For unusual layouts, the reliable fallback remains: full report body plus field mapping.
- Screenshot insertion is intentionally scoped to standard inline image paragraphs plus captions. It does not attempt floating images, text wrapping, or arbitrary WPS desktop automation.

## Privacy Notes

- Do not commit real student identity data, screenshots, or private templates.
- Put local materials under `local-inputs/` or `outputs/`; both are ignored by `.gitignore`.
- If a report must be factual and key data is missing, the skill should ask for it or label the result as a sample draft.

## License

MIT
