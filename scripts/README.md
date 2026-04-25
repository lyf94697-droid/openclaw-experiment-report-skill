# 脚本目录说明

这个目录已经不只是“零散 helper”，而是一套相对完整的本地流水线。

## 最常用入口

- `build-report.ps1`
  已有正文时的主入口。负责校验、字段映射、填模板、插图、样式收尾和 layout check。

- `build-report-from-url.ps1`
  适合“教程链接 -> 正文 -> 模板 -> docx”的链路。

- `build-report-from-feishu.ps1`
  适合飞书 / 直聊 / 附件路径场景。

- `run-one-click-demo.ps1`
  仓库自带的一键演示，不需要自己准备素材。

## 模板处理

- `extract-docx-template.ps1`
- `generate-docx-field-map.ps1`
- `apply-docx-field-map.ps1`
- `convert-docx-template-frame.ps1`

## 图片处理

- `generate-docx-image-map.ps1`
- `insert-docx-images.ps1`
- `render-vertical-lab-flowchart.py`

## 质量检查

- `validate-report-draft.ps1`
- `check-docx-layout.ps1`
- `check-report-profile-template-fit.ps1`
- `run-smoke-tests.ps1`
- `self-check.ps1`

## 环境与安装

- `install-skill.ps1`
- `reset-openclaw-session.ps1`
- `report-defaults.ps1`
- `report-profiles.ps1`

## 参考建议

第一次阅读建议顺序：

1. `run-one-click-demo.ps1`
2. `build-report.ps1`
3. `build-report-from-url.ps1`
4. `run-smoke-tests.ps1`
