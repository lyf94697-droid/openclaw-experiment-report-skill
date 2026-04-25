---
name: experiment-report
description: Write Chinese university lab reports and course-design reports, or fit report content into a user-provided WPS, Word, or docx blank template. Use when the user asks for an experiment report, lab report, course design report, experiment summary, or a WPS, Word, or docx template to be filled from experiment topic, requirements, code, data, screenshots, tutorial pages, or results.
---

## When this skill applies

- The user asks to write or complete an experiment report from zero.
- The user asks for a course design report that follows a fixed school-style structure.
- The user has a blank WPS, Word, or docx template and wants it filled.
- The user provides an experiment title, requirements, screenshots, code, data, outputs, or conclusions and wants a report draft or filled report.
- The user provides a tutorial article or CSDN page as the main procedural reference.

## Core workflow

1. Collect the minimum useful inputs:
   - course name
   - experiment name
   - template path or screenshots if formatting matters
   - experiment requirements or task description
   - actual steps, code, screenshots, outputs, data, or conclusions
   - whether the user wants a factual report or a clearly labeled sample version
   - If the user already provides enough facts to write the report body, do not stop to ask for optional metadata such as name, class, date, or template files.
2. Write the full report content before touching template formatting.
3. If critical facts are missing, do not fabricate exact data, screenshots, or measurements.
4. If a local docx template exists and shell execution is available, run `scripts/extract-docx-template.ps1 -Path <template.docx>` first.
5. Use the extracted outline, table cells, and likely fields as the template source of truth.
6. If a template exists, adapt the finished content to the template order and field names.
7. If the user explicitly wants a filled local docx output and the template matches common report patterns, generate a field map with `scripts/generate-docx-field-map.ps1` and then run `scripts/apply-docx-field-map.ps1`.
   - Use label keys for normal blank-field filling.
   - Use `paragraphs` arrays for section-body content.
   - Use `mode: "after"` when the template keeps a fixed heading paragraph and the actual content should go into the following blank paragraph.
   - Use location keys such as `P2` or `T1R1C2` only when explicit overwrite is needed.
8. If the user also provides screenshots or experiment photos and wants them embedded into the final docx, prefer `scripts/generate-docx-image-map.ps1` on the filled copy and then run `scripts/insert-docx-images.ps1`.
9. If the user wants a cleaner final docx layout, run `scripts/format-docx-report-style.ps1` after template filling and optional image insertion.
10. Prefer content-first completion over fragile GUI wandering.
11. If screenshots are provided, treat them as factual evidence and layout assets.
12. If a tutorial page is provided, treat it as procedural reference and rewrite it into report-style Chinese instead of copying it.
13. When a local workflow should fetch tutorial references before generation, prefer `scripts/prepare-report-prompt.ps1` with `-ReferenceUrls` or `-ReferenceTextPaths`.
14. When the user provides local file paths in a direct chat workflow, inspect those files first if tool access is available; if a path cannot actually be opened, explicitly say which path was inaccessible instead of pretending it was read.
15. When the user asks for a final local `docx` result and the required paths are already present, prefer finishing the end-to-end local workflow over stopping at a body-only draft.
16. On Windows PowerShell, do not chain shell commands with `&&`; use separate executions or `;` so the command remains valid in legacy PowerShell hosts.
17. When intermediate JSON or text files contain Chinese paths, captions, or section names, prefer writing them through PowerShell with explicit UTF-8 encoding or through the bundled scripts; do not rely on generic editor-style writes that may corrupt non-ASCII content.
18. When direct chat already has a template path, screenshots, identity metadata, and either a finished report body or tutorial references, prefer the one-shot local wrapper `scripts/build-report-from-feishu.ps1` over ad-hoc multi-step shell orchestration.
19. If direct chat includes uploaded image attachments and the user also provides local image paths, use the uploaded images to understand the visible content and use the local image paths as the actual files for deterministic `docx` embedding.
20. If direct chat includes uploaded image attachments but no manual local image paths, check whether the runtime injected attachment note lines such as `[media attached ...]` into the prompt. If those lines contain usable image file paths, extract them and pass them into `-ImagePaths` for the local wrapper instead of stopping at body-only output.

## Fixed visual standards

- When the user says to use the previous standard, the fixed standard is: experiment reports keep the original-template outer frame, course-design reports use a large standalone flowchart, and generated flowchart titles have no side decoration lines.
- For experiment-report template-frame output, keep normal table lines in the top metadata table. Put the body into one full-width framed body area, without horizontal separator lines between paragraphs or sections.
- For `course-design-report`, render the overall design flowchart / process diagram near full body width. The default lower bound is `15.8 cm`.
- Do not auto-place course-design flowcharts side by side with screenshots. Flowcharts are standalone design diagrams; screenshots can still use row layouts when they are clearly paired.
- For generated black-and-white flowcharts, keep the title as centered text only. Do not draw left/right decorative horizontal lines around the title.
- After generating a local `docx`, run the strongest practical layout check available. For important deliverables, export or render pages and visually check the frame, flowchart size, and line overlap before claiming completion.
- Keep these rules scoped to Chinese lab reports and course-design reports. Do not turn them into a generic document platform or graph engine unless explicitly requested.

## Writing rules

- Use clear Chinese suitable for university reports.
- Keep claims consistent with the provided requirements, code, data, screenshots, and outputs.
- Avoid empty filler and generic AI phrasing.
- If the user provides the course name or experiment name, write them explicitly into the final report instead of assuming the surrounding chat context is enough.
- If the user already supplied the experiment topic, environment, steps, results, and required headings, write the report immediately instead of asking for more materials.
- Ask follow-up questions only when missing facts would make the result materially wrong, or when the user explicitly wants template filling but no template is available.
- If the user only wants the report body, missing personal identity fields must not block generation.
- If direct chat is being used with local file paths, avoid optimistic assumptions about file access; either read the files or clearly state that file access was not available.
- When the experiment is software-related, include environment, implementation steps, results, analysis, and conclusion.
- When the template has fixed headings, preserve them exactly.
- If webpage instructions and user screenshots differ, trust the user screenshots and outputs.
- If both uploaded image attachments and local image paths are available, use the attachments as the semantic reference for what each image shows, but use the local paths for the final `docx` image insertion workflow.
- If uploaded image attachments are present without manual local paths, prefer the prompt-injected attachment paths from `[media attached ...]` notes as the `ImagePaths` input for `scripts/build-report-from-feishu.ps1` or `scripts/generate-docx-image-map.ps1`. If the runtime does not expose any real attachment path, say that clearly instead of pretending direct `docx` insertion succeeded.
- When screenshots are provided without explicit grouping, infer `实验环境`, `实验步骤`, `实验结果`, or `问题分析` from filenames and visible content, but do not invent unseen details.
- If a local workflow needs temporary JSON such as field maps or image maps and those files include Chinese text, write them in explicit UTF-8 and retry from that stage if parsing fails.

## Output modes

- Default: final report content with headings ready to paste.
- Template mode: exact field-to-content mapping in template order.
- Template mode can include block values such as `{"section-body": ["paragraph one", "paragraph two"]}` when the template body has multi-paragraph sections.
- Image mode: when screenshots should be embedded into a docx, prefer image specs or an image insertion map that can be passed to `scripts/generate-docx-image-map.ps1` or `scripts/insert-docx-images.ps1`. Stable section anchors such as `实验步骤` or `实验结果` are preferred over fragile paragraph numbers when the filled docx may add or move paragraphs.
- Completion mode: if the user explicitly asks to complete the template, first finish the content, then attempt template filling only if tooling is actually available.

## Optional helpers

- For local docx templates, run `scripts/extract-docx-template.ps1` to capture the actual field order before producing a field mapping.
- For local docx templates that should be machine-filled, run `scripts/generate-docx-field-map.ps1` after the report body is ready, then run `scripts/apply-docx-field-map.ps1`.
- For local screenshots or experiment photos that should be embedded into the filled docx, prefer `scripts/generate-docx-image-map.ps1` first and then run `scripts/insert-docx-images.ps1` on the already-filled copy.
- For a cleaner formatted docx copy, optionally run `scripts/format-docx-report-style.ps1` after filling fields and inserting images.
- For chat-driven local execution, prefer `scripts/build-report-from-feishu.ps1` so the wrapper can keep the final deliverable in the output root and move intermediate files into an `artifacts/` subdirectory.
- The image pipeline can resolve OpenClaw-staged relative attachment paths such as `media/inbound/example.png`, so when those paths appear in prompt-injected media notes you can reuse them directly in `-ImagePaths`.
- When the template has fixed section headings plus blank paragraphs, prefer block mappings over flattening long body content into a single field.
- For public tutorial pages, prefer `scripts/fetch-web-article.ps1`; keep `scripts/fetch-csdn-article.ps1` as the compatibility wrapper for CSDN-specific workflows.
- When a tutorial page should flow directly into report generation, prefer `scripts/prepare-report-prompt.ps1` so the extracted reference text is appended to the final request deterministically.
- The helpers are optional. If they are unavailable, still finish the report from the information already provided.

## Read references as needed

- Read `references/common-structures.md` when choosing a report outline.
- Read `references/template-fit.md` when the user provides a WPS, Word, or docx template path or template screenshots.
- Read `references/image-handling.md` when the user provides experiment screenshots or process images.
