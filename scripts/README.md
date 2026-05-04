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

- `check-project-readiness.ps1`
  面向展示 / 发布 / 简历场景的轻量检查，确认 README、关键 docs、典型案例和案例 JSON 齐全。

## 模板处理

- `extract-docx-template.ps1`
- `generate-docx-field-map.ps1`
- `apply-docx-field-map.ps1`
- `convert-docx-template-frame.ps1`

## 图片处理

- `generate-docx-image-map.ps1`
- `insert-docx-images.ps1`
- `render-vertical-lab-flowchart.py`

## 固化版式行为

- `convert-docx-template-frame.ps1`：实验报告套原模板外框时，顶部信息表保留表格线，正文合并到一个外框区域，正文段落之间不额外加横线。
- `build-report.ps1`：课程设计报告自动生成的总体设计图默认放在图片列表最前，并使用大图宽度。
- `generate-docx-image-map.ps1` / `insert-docx-images.ps1`：`course-design-report` 的流程图宽度下限为 `15.8 cm`，且不参与自动两列并排布局。
- `render-vertical-lab-flowchart.py`：黑白流程图标题只保留居中文字，不绘制标题左右两侧装饰横线。
- `run-smoke-tests.ps1`：烟测会检查课程设计流程图宽度和非并排布局，避免后续改动破坏固定效果。

## 质量检查

- `validate-report-draft.ps1`
- `check-docx-layout.ps1`
- `check-report-profile-template-fit.ps1`
- `run-smoke-tests.ps1`
- `check-project-readiness.ps1`
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
