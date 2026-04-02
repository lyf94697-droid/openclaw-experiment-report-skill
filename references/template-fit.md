# Template Fit Mode

## 1. Template-first mapping

- 先识别模板里的固定标题、编号、表格字段、封面字段。
- 输出时按模板顺序组织，不要擅自改编号。
- 如果模板标题和常规模板不同，优先服从模板。

## 2. WPS or Word handling

- 本地 WPS 桌面自动化不是主路径，先完成内容，再尝试填写。
- 如果模板是本地 docx，优先先跑 `scripts/extract-docx-template.ps1 -Path <template.docx>`，用提取出的段落顺序、表格单元格和疑似字段做映射。
- 如果正文已经写完且需要机器生成字段映射，优先跑 `scripts/generate-docx-field-map.ps1`，不要手工拼大段 JSON。
- 如果用户明确需要机器生成一个已填内容的 docx 副本，再用 `scripts/apply-docx-field-map.ps1` 执行回填。
- 标签键映射适合保守填空，只会填充空白位或占位符。
- 如果模板里是固定章节标题加空白正文段，优先给出 `paragraphs` 数组。
- 需要保留标题、把正文写到下一段时，用 `mode: "after"`。
- 位置键映射适合显式覆盖，比如 `P2`、`T1R1C2`。
- 如果模板是本地 docx 且可编辑，优先按字段填充。
- 如果只有截图或用户只给了空白模板界面，输出“字段 -> 内容”映射，便于直接粘贴。
- 若没有可靠自动化，就明确说明返回的是正文和字段映射，而不是虚构“已自动填好模板”。

## 3. Field mapping guidance

- 封面字段：课程名、实验名、姓名、学号、班级、日期分开写。
- 表格字段：优先输出单元格级内容，不要把整段话塞进短字段。
- 正文大段落：优先按段落数组输出，避免把多段正文压成一个字段。
- 对“实验目的 / 实验步骤 / 实验结果”这类固定标题，尽量保持标题原样，只映射后续正文。

## 4. Stop conditions

- 模板内容或实验要求缺失到无法保证真实性时，先补信息再继续。
- 若桌面模板填写能力不可用，仍需完成整份报告正文和字段映射，不得直接中断。

## 5. Tutorial article plus screenshots mode

- When the user gives a tutorial article plus their own screenshots or results, treat the article as procedural reference and the screenshots or results as factual evidence.
- Fill missing explanatory sections from the article, but keep the result section aligned with the user's actual outputs.
- If the article includes code and the user has not provided their own code, mark it as reference implementation instead of pretending it is the user's exact work.
