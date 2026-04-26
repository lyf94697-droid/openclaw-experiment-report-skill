# Experiment Report Images

## 1. Role of images

- Treat experiment screenshots and process images as evidence and layout assets, not generic attachments.
- Use the visible content only. Do not invent image details that are unclear or absent.

## 2. Section mapping

- First map each image to the most likely section: 实验环境, 实验步骤, 实验结果, 问题分析.
- If multiple screenshots belong to one step, keep them in chronological order under the same step.

## 3. Figure labels and captions

- Prefer explicit figure labels such as 图1, 图2, 图3.
- Use short academic captions such as:
  - 图1 软件安装或启动界面
  - 图2 命令执行结果
  - 图3 网络拓扑或配置结果

## 4. Placement guidance

- Reference each image near the relevant paragraph, for example: “如图1所示”.
- If the final deliverable is a filled local docx, prefer `scripts/generate-docx-image-map.ps1` before `scripts/insert-docx-images.ps1`.
- Use paragraph anchors such as `P8` when the image should appear after a specific body paragraph.
- Use table-cell anchors such as `T1R6C1` when the template keeps screenshots inside a large merged cell.
- Prefer stable section anchors such as `实验步骤` or `实验结果` when the filled docx may add or move paragraphs after template filling.
- When two screenshots should appear side by side, use a shared layout block such as `{ "mode": "row", "columns": 2, "group": "results-grid" }` on each related image item.
- When one row group should stay together under a single location, add `layout.groupAnchor`, for example `{ "mode": "row", "columns": 2, "group": "results-grid", "groupAnchor": "实验结果" }` or `{ "groupAnchor": "anchor:P8" }`.
- If a grouped row block mixes screenshots from different sections, prefer one shared `groupAnchor` over splitting the block across multiple headings.
- If automatic insertion is unavailable, still return:
  - the full report body
  - image insertion order
  - figure captions
  - recommended insertion positions
- Prefer one screenshot per line when images are dense. Use two per row only when they are clearly paired and the template has enough width.
- In the current `experiment-report` default, imitate the teacher excellent example: large centered single screenshots with a `15.8 cm` lower-bound width. Automatic two-column grouping is disabled; use row/2x2 only when the input explicitly says the images should be side by side.

## 5. Course-design flowcharts

- In `course-design-report`, treat the generated flowchart or overall design diagram as a standalone design figure, not as a normal screenshot.
- Use a near-full-width placement for course-design flowcharts. The fixed lower bound is `15.8 cm`.
- Do not auto-group course-design flowcharts into two-column row layouts with screenshots.
- Keep generated black-and-white flowchart titles as centered text only, without left/right decorative lines.
- Before final delivery, check the rendered page when possible: the diagram should be large enough, lines should not overlap nodes, and the title should stay clean.
