# 示例文件总览

这个目录放的是“看得懂、改得动、能直接套用”的样例，而不是只做测试占位。

## 推荐先看

- [demo-one-click/README.md](demo-one-click/README.md)
- [report-templates/README.md](report-templates/README.md)

## 目录说明

- `demo-one-click/`
  自带演示用的正文、metadata 和 requirements，配合 `scripts/run-one-click-demo.ps1` 直接可跑。

- `report-templates/`
  仓库内置的示例模板，目前包含实验报告模板和课程设计报告模板。

- `docx-field-map.json`
  字段映射示例，适合看模板填充格式。

- `docx-image-specs.json` / `docx-image-specs-row.json`
  插图规格示例，已经改成引用仓库内置的 `demo/assets`。

- `docx-image-map.json` / `docx-image-map-row.json`
  正式 image map 示例，适合看 anchor 和 layout 的写法。

- `docx-report-metadata.json`
  metadata 示例。

- `e2e-sample-requirements.json`
  正文校验需求示例。

- `feishu-uploaded-images-docx-prompt.md`
- `local-uploaded-images-docx-prompt.md`
- `one-shot-uploaded-images-docx-prompt.md`
  面向聊天 / 飞书场景的 prompt 样例。

## 使用建议

- 想最快感受效果：先跑 `scripts/run-one-click-demo.ps1`
- 想理解 JSON 格式：看 `docx-field-map.json` 和 `docx-image-specs*.json`
- 想适配你自己的模板：先看 `report-templates/README.md`
