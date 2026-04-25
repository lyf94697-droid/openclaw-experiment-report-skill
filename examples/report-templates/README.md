# 模板示例

这个目录放的是仓库内置、可直接引用的示例模板，主要用于：

- 一键演示
- 新用户上手
- 自定义模板前的字段结构参考

## 当前内置模板

| Profile | 展示名称 | 模板文件 | 用途 |
| --- | --- | --- | --- |
| `experiment-report` | 实验报告 | `experiment-report-template.docx` | 默认实验报告演示模板 |
| `course-design-report` | 课程设计报告 | `course-design-report-template.docx` | 课程设计快线最小模板 |

## 使用建议

- 想最快跑通项目：优先用 `experiment-report-template.docx`
- 想体验课程设计快线：用 `course-design-report-template.docx` 并显式传 `-ReportProfileName course-design-report`
- 真实学校模板适配前，先跑 `scripts/check-report-profile-template-fit.ps1`
