# Report Profiles

`profiles/` 用来定义不同文档类型的结构、标题、默认样式和图片落位规则。

## 当前内置 profile

- `experiment-report.json`
  默认实验报告主线，适合中文大学实验报告。

- `course-design-report.json`
  显式 opt-in 的课程设计快线，适合课程设计报告和学校固定模板。

- `report-profile.schema.json`
  profile 结构说明。

## 什么时候要关心 profile

- 你想从普通实验报告切到课程设计报告
- 你想改章节标题、图片默认落位或样式默认值
- 你要给新的文档类型做模板适配

## 相关脚本

- `scripts/report-profiles.ps1`
- `scripts/check-report-profile-template-fit.ps1`
- `scripts/build-report.ps1 -ReportProfileName ...`
