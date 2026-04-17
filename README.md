# OpenClaw Experiment Report Skill
> Generate Chinese lab reports with OpenClaw: drafting, docx template filling, screenshot insertion, and final formatting.

[![Quality Checks](https://github.com/lyf94697-droid/openclaw-experiment-report-skill/actions/workflows/quality.yml/badge.svg)](https://github.com/lyf94697-droid/openclaw-experiment-report-skill/actions/workflows/quality.yml)
[![Smoke Tests](https://github.com/lyf94697-droid/openclaw-experiment-report-skill/actions/workflows/smoke-tests.yml/badge.svg)](https://github.com/lyf94697-droid/openclaw-experiment-report-skill/actions/workflows/smoke-tests.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

## 项目简介

这是一个基于 OpenClaw 的实验报告 skill 和 PowerShell 文档流水线。

它的目标不是只生成一段“实验报告正文”，而是把实验主题、要求、教程链接、代码、截图、数据、空白 `docx`/WPS/Word 模板这些材料串成一条更完整的本地工作流：先生成结构化中文实验报告，再填充模板、插入截图、生成图注，最后做一轮提交风格的排版优化，输出更接近可直接提交的 `docx` 成品。

当前仓库优先解决“中文大学实验报告”这个具体场景，并保持技术路径务实可落地：以 OpenClaw 为生成入口，以本地脚本做模板处理、插图和最终排版。

## 解决什么问题

实验报告最耗时间的部分，通常不是单纯“写字”，而是材料整理和成品收尾：

- 实验主题、教程网页、代码、截图、数据分散在不同地方
- 正文结构要手动整理成“实验目的、步骤、结果、总结”这类固定章节
- 学校模板往往是空白 `docx`，填字段和填正文都很麻烦
- 截图要决定插在哪里、图题怎么写、几张图怎么排
- 最后标题、正文、图注、封面间距这些排版细节很耗时间

这个仓库就是把这些步骤收拢成一条尽量稳定、可重复的本地流程。

## 主要功能

- 根据实验主题、要求、教程链接、代码、截图、数据等材料生成结构化中文实验报告正文
- 对空白 `docx` / WPS / Word 模板做字段映射和正文填充
- 插入实验截图并生成图注，支持按章节或稳定锚点落位
- 支持图片分组布局，例如两图一行、2x2 图片块这类常见报告排版
- 对标题、章节标题、正文、图注、图片块做最终提交风格排版优化
- 支持通过 OpenClaw skill 或本地 PowerShell 入口脚本调用
- 支持基于教程页或参考材料生成内容，但目标是改写和整理，不是长篇照抄

## 工作流程

1. 提供实验主题、要求、教程链接、截图、模板或已有正文
2. 生成结构化中文实验报告正文
3. 填充 `docx` 模板中的字段和章节内容
4. 插入截图并生成图注，必要时做分组布局
5. 对最终文档做排版优化，输出更完整的提交版 `docx`

## 演示效果

下面保留一个简化预览；完整的 2x2 图片块示例和 GitHub 友好的演示素材见 [demo/README.md](demo/README.md)。

| 步骤截图预览 | 结果截图预览 |
| --- | --- |
| ![Network config preview](demo/assets/step-network-config.png) | ![Ping result preview](demo/assets/result-ping.png) |

如果你想看完整四图联排效果，直接打开 [demo/README.md](demo/README.md)。

## 当前范围

- 当前版本已经从单一中文大学实验报告，扩展为面向“学校报告类文档”的 profile-driven 本地流水线
- 当前仓库主要面向 OpenClaw 用户，不是独立桌面应用
- 仓库已经包含 `SKILL.md`、`references/`、`scripts/`、`examples/`、`demo/` 和 GitHub 协作治理文件
- 当前稳定路径是“OpenClaw 生成 + 本地脚本处理模板、图片和排版”
- 常见空白模板、章节正文、截图插入和最终样式处理已经可以跑通
- 当前暂时不保证 Word/WPS GUI 自动化填写，也不承诺任意复杂模板都能无人工确认处理

## 已经扩展到哪些方面

仓库现在的核心变化，是从“实验报告专用脚本”扩展成“同一条报告构建流水线 + 可切换 report profile”。

- 已内置六类报告 profile：`experiment-report`、`course-design-report`、`internship-report`、`software-test-report`、`deployment-report`、`weekly-report`
- `experiment-report` 覆盖常见实验目的、实验环境、实验原理或任务要求、实验步骤、实验结果、问题分析、实验总结
- `course-design-report` 覆盖设计目标、开发环境、需求分析、方案设计与实现、运行结果、问题与改进、设计总结
- `internship-report` 覆盖实习目的、实习单位与环境、实习任务与要求、实习过程与内容、实习成果、问题分析与改进、实习总结
- `software-test-report` 覆盖测试目标、测试环境、测试范围与依据、测试用例设计与执行、测试结果、缺陷分析与改进、测试总结
- `deployment-report` 覆盖部署目标、部署环境、部署方案与架构、部署步骤与配置、验证结果、问题处理与回滚预案、部署总结
- `weekly-report` 覆盖工作目标、协作环境、本周任务与输入、本周完成事项、阶段成果、风险与改进、下周计划
- profile 现在不只是章节名列表，还承载 metadata 标签、默认样式、prompt 文案、章节最小长度、分页风险阈值、图注规则、图片落位优先级和复合模板填充规则
- 模板适配能力也被抽象出来，可以用 `check-report-profile-template-fit.ps1` 诊断新模板缺字段、缺章节、缺 alias 或是否需要新增复合填充规则
- 插图流程已经支持按 profile 识别章节、生成图片落位预案、写入图注、连续图片 2 列分组布局，以及最终 layout check
- 构建过程现在会写出 `summary.json`、`pipeline-trace.json`、`pipeline-trace.md`、`image-placement-plan.md` 和 `layout-check.json`，方便复盘每次运行到底用了什么输入、生成了什么产物
- validation 从“检查有没有章节”扩展到 profile-specific structural validation，包括缺必需章节、章节顺序异常、重复标题、空节、占位符节、过短章节和分页风险 warning
- 现阶段仍然优先服务学校报告类文档；`weekly-report` 已经作为结构稳定、能被 smoke 覆盖的相邻场景纳入内置 profile

## Validation 与风险输出

`validate-report-draft.ps1`、`build-report.ps1`、`build-report-from-url.ps1` 和 `build-report-from-feishu.ps1` 现在都会把结构校验和分页风险写进 machine-readable JSON，方便自动化判断报告是否能继续进入模板填充和交付检查。

主要输出文件：

| 文件 | 作用 |
| --- | --- |
| `validation.json` | `validate-report-draft.ps1` 的完整校验结果，包含章节命中、finding 列表和汇总计数 |
| `summary.json` | `build-report.ps1` 的构建摘要，会透出 validation、layout check、image plan 和最终 docx 路径 |
| `url-build-summary.json` | URL wrapper 摘要，会透出下游 build validation/risk 汇总 |
| `feishu-build-summary.json` | Feishu/local wrapper 摘要，会透出下游 build validation/risk 汇总 |
| `pipeline-trace.json` / `pipeline-trace.md` | wrapper 级调试视图，聚合 generation mode、input mode、validation 状态和关键产物路径 |

常用字段：

| 字段 | 含义 |
| --- | --- |
| `validationPassed` | 没有 error 时为 `true`；pagination risk 目前是 warning，不会单独导致失败 |
| `validationErrorCount` / `validationWarningCount` | validation finding 的 error / warning 数量 |
| `validationPaginationRiskCount` | `category = pagination` 的风险 warning 数量 |
| `validationPaginationRiskThresholds` | 当前使用的分页风险阈值，来自 active report profile 或外部 requirements |
| `validationStructuralIssueCount` | `category = structure` 的结构问题数量 |
| `validationErrorCodes` / `validationWarningCodes` | 去重后的机器可读 code 列表 |
| `validationWarningSummary` | warning 的轻量摘要，包含 `severity`、`code`、`category`、`message`、`remediation` |
| `validationFindingCountsByCode` / `validationFindingCountsByCategory` | 按 code 和 category 聚合的计数表 |
| `templateFrameDocxPath` | 可选模板边框版 `docx` 路径；普通最终稿仍保留在 `finalDocxPath` |

当前结构校验 code：

| Code | Severity | 含义 |
| --- | --- | --- |
| `missing-profile-required-heading` | error | 使用内置或 profile-backed 规则时，缺少 profile 要求的章节标题 |
| `missing-required-section` | error | 使用外部 requirements 且没有 profile 标记时，缺少要求章节 |
| `duplicate-section-heading` | error | 同一章节标题重复出现 |
| `section-order-anomaly` | error | 章节出现顺序和 profile / requirements 期望顺序不一致 |
| `empty-section` | error | 章节存在但没有正文内容 |
| `placeholder-only-section` | error | 章节正文只有占位符，例如 `__________`、`TODO`、`placeholder` |
| `short-section` | error | 章节正文低于该章节的 `minChars` 要求 |

当前分页风险 code：

| Code | Severity | 含义 |
| --- | --- | --- |
| `pagination-risk-long-section` | warning | 单个章节正文较长，进入 Word/WPS 模板后更容易跨页 |
| `pagination-risk-dense-section-block` | warning | 单个章节文本密集、段落数少，分页时容易形成大块断裂 |
| `pagination-risk-figure-cluster` | warning | 单个章节引用较多图片，图片和图注可能造成分页压力 |

这些分页风险阈值可以在 profile 的 `paginationRiskThresholds` 里调整。字段包括 `longSectionChars`、`denseSectionChars`、`denseSectionParagraphs` 和 `figureClusterRefs`；`validate-report-draft.ps1` 的 `summary.paginationRiskThresholds` 会写出本次实际使用的值。

每条 validation finding 现在还会带 `remediation` 字段。它是给自动化和人工排查看的下一步建议，例如补齐缺失章节、合并重复标题、替换占位符、拆分过长段落，或在 profile 中调整分页阈值。

## 快速开始

### 1. 安装 skill

方式一：直接 clone 到 OpenClaw 实际加载的 skills 目录。

```powershell
git clone <your-repo-url> "$env:USERPROFILE\.agents\skills\experiment-report"
```

方式二：在已检出的仓库里运行安装脚本。

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install-skill.ps1
```

安装后可以运行：

```powershell
openclaw skills list
```

确认 `experiment-report` 已出现。  
如果你的运行时会缓存 skill 列表，再重启 OpenClaw 或刷新 skills。

这个 skill 的常见触发词包括：`实验报告`、`实验模板`、`WPS模板`、`Word模板`、`docx模板`、`lab report`。  
如果想要更稳定地触发，建议直接以 `/experiment-report` 开头。

### 2. 常用入口脚本

最常用的本地入口有 4 个：

- `scripts/build-report-from-feishu.ps1`
  适合飞书或直接聊天场景，负责把生成、模板填充、插图和最终输出串起来
- `scripts/build-report-from-url.ps1`
  适合“教程链接 -> 报告正文 -> 模板填充 -> 最终 docx”这类流程
- `scripts/build-report.ps1`
  适合你已经有正文和模板，只想走确定性的本地 `docx` 打包流程
- `scripts/generate-report-inputs.ps1`
  适合先单独导出 `prompt.txt`、`metadata.auto.json`、`requirements.auto.json`，再手动调试生成或对接外部流水线

如果你需要拆开流水线逐步处理，仓库里也已经提供模板抽取、字段映射生成、图片映射生成、插图、样式优化、网页抓取、提示词准备和端到端验证脚本，入口都在 [scripts](scripts) 目录。

其中 `build-report-from-feishu.ps1` 和 `build-report-from-url.ps1` 在需要自动生成正文时默认使用 `-DetailLevel full`，也就是默认要求输出更完整、不是提纲式的正文。

### 3. 一条常见用法

如果你想走聊天友好的本地封装入口，可以直接用：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report-from-feishu.ps1 `
  -ReferenceUrls "https://blog.csdn.net/..." `
  -CourseName "计算机网络" `
  -TemplatePath "E:\reports\template.docx" `
  -StudentName "张三" `
  -StudentId "20260001" `
  -ClassName "计科 2201" `
  -ImagePaths "E:\reports\step-1.png","E:\reports\step-2.png","E:\reports\result-1.png","E:\reports\result-2.png" `
  -OutputDir "E:\reports\final-output"
```

如果你想从教程链接直接走到最终 `docx`，可以用：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report-from-url.ps1 `
  -ReferenceUrls "https://blog.csdn.net/..." `
  -CourseName "计算机网络" `
  -TemplatePath "E:\reports\template.docx" `
  -StudentName "张三" `
  -StudentId "20260001" `
  -ClassName "计科 2201"
```

如果你已经有整理好的正文和模板，直接走本地 `docx` 流程：

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

补充说明：

- `build-report-from-feishu.ps1` 和 `build-report-from-url.ps1` 会优先从提示词、参考文本或 URL 片段推断 `ExperimentName`；推断不到时才复用最近一次保存的题目名
- 在教程标题或参考文本已经包含实验名称时，can omit `-ExperimentName`
- `build-report.ps1` 支持 `-StyleProfile auto|default|compact|school`
- 如果你想加载自定义排版配置，可以配合 `-StyleProfilePath` 使用
- 如果你想额外生成保留模板表格边框的一体版，可以对 `build-report.ps1` 传 `-CreateTemplateFrameDocx`，或对 `build-report-from-url.ps1` / `build-report-from-feishu.ps1` 传 `-CreateTemplateFrameDocx`；wrapper 会把普通最终稿写到 `finalDocxPath`，把边框版写到 `templateFrameDocxPath`
- 自动生成正文、默认 validation / layout-check、模板 field-map 生成，以及 `generate-docx-image-map.ps1` / `insert-docx-images.ps1` 的章节识别现在都会从 `profiles/experiment-report.json` 读取实验报告 profile；需要切换或覆盖时，可对 `build-report.ps1` / `build-report-from-url.ps1` / `build-report-from-feishu.ps1` / `generate-docx-image-map.ps1` / `insert-docx-images.ps1` 使用 `-ReportProfileName` 或 `-ReportProfilePath`
- 仓库现在内置六个 profile：`experiment-report`、`course-design-report`、`internship-report`、`software-test-report`、`deployment-report` 和 `weekly-report`；如果你要生成课程设计报告、专业实习报告、软件测试报告、部署运维报告或结构化项目周报，可以直接传对应的 `-ReportProfileName`
- `build-report-from-url.ps1` 的自动 prompt 也会跟随 active report profile 调整文案，像 `课程名称` / `课题名称` 这类字段标签会直接从 profile 里取，而不再固定写成实验报告措辞
- 最近一次保存的 `CourseName` / `ExperimentName` 默认值现在会按 report profile 隔离保存，`course-design-report` 不会再复用 `experiment-report` 的最近一次题目
- 如果你只想先拿到自动生成的输入物，不想立刻跑 OpenClaw 或 `docx` 流水线，可以先运行 `scripts/generate-report-inputs.ps1`，它会单独导出 `prompt.txt`、`metadata.auto.json`、`requirements.auto.json` 和一份 summary
- 如果你已经有一份可复用的正文，想离线回放后续 validation / 模板填充 / 插图流程，`build-report-from-feishu.ps1`、`build-report-from-url.ps1`、`run-e2e-sample.ps1` 和 `generate-report-chat.ps1` 现在都支持显式 `-PreGeneratedReportPath`
- 相关 summary JSON 现在会额外写出 `generationMode`，用于区分 `live`、`replay` 和 `none`（直接传本地正文）的运行路径
- `build-report.ps1`、`generate-docx-field-map.ps1`、`check-report-profile-template-fit.ps1`、`generate-docx-image-map.ps1`、`insert-docx-images.ps1` 的输出现在也会补充 `inputMode` 字段，用于区分报告、metadata、requirements、图片规格和 image-map 分别来自文件、内联 JSON 还是直接图片列表
- `build-report-from-url.ps1` 和 `build-report-from-feishu.ps1` 的 wrapper summary 现在也会透出下游 `buildReportInputMode` / `buildMetadataInputMode` / `buildRequirementsInputMode` / `buildImageInputMode`，用于把 wrapper 的 `generationMode` 和实际 docx 构建输入来源串成一条完整链路
- 这两个 wrapper 现在还会额外写出一份 `pipeline-trace.json`，把 `wrapper.mode`、`wrapper.generationMode`、下游 `build.*InputMode` 和关键产物路径聚合到一份更短的调试视图里
- 同时也会生成一份更适合人工快速查看的 `pipeline-trace.md`
- `generate-docx-field-map.ps1` 的 JSON 输出现在会额外带 `diagnostics` 和 `summary.diagnosticCountsByCode`，用于解释模板里哪些章节标题、metadata 标签或复合正文单元格没有命中自动映射规则
- 如果你正在适配新模板或准备新增一个 report profile，可以先跑 `scripts/check-report-profile-template-fit.ps1`，它会基于 field-map diagnostics 汇总出缺 metadata、缺章节内容、建议补的 `sectionFields` alias，以及建议新增的 `fieldMapCompositeRules`
- 新增 report profile 时，可以先用 `scripts/new-report-profile.ps1` 生成 schema-valid 草稿，再按具体文档类型调整标题、alias、图注和 prompt 文案
- 如果某类文档天然章节更长或图片更多，可以在 profile 里调高 `paginationRiskThresholds`，避免把正常结构误报成分页风险；反过来也可以调低，用于更早捕捉 WPS/Word 模板风险
- 新增或修改 report profile 后，运行 `scripts/validate-report-profiles.ps1`；profile 结构约束集中在 `profiles/report-profile.schema.json`
- 如果你还不想把新文档类型直接升级成内置 profile，可以先看 `examples/profile-presets/`：目前保留 `meeting-minutes.json` 作为相邻文档类型的自定义 preset，并保留 `weekly-report.json` 作为 path-based 示例快照，适合先验证“这条 pipeline 能不能复用”
- 自定义 preset 不需要先拷进 `profiles/`，可以直接对 `generate-report-inputs.ps1`、`build-report.ps1`、`build-report-from-url.ps1`、`build-report-from-feishu.ps1` 传 `-ReportProfilePath`
- 想一次性查看所有示例 preset 会生成什么 prompt、metadata 和 requirements，可以运行 `scripts/run-profile-preset-samples.ps1`
- 只要传入图片，`build-report.ps1`、`build-report-from-feishu.ps1`、`build-report-from-url.ps1` 都会自动额外写出 `image-placement-plan.md`；如需改位置，可用 `-ImagePlanOutPath`
- 正文排版会单独识别步骤编号和 DOS/终端命令，步骤段不做首行缩进，命令段使用等宽字体、浅灰底和更紧凑的单倍行距
- 最终排版会统一标题、正文、图注的字号；表格型模板会尽量保留模板默认字体观感，避免额外强制字体导致成品不像原模板
- 表格型实验报告模板会使用更接近模板默认观感的字号，把正文单元格改为顶部对齐，并减少普通正文的强制分页保持，避免留下过多空白
- 不确定多张截图会被放到哪里时，可以先运行 `scripts/generate-docx-image-map.ps1 -PlanOnly -Format markdown` 输出图片分配预案，确认章节、图注和布局后再生成正式 image map
- 多张图片连续归入同一实验章节时，插图流程会默认使用每行 2 张的自动分组布局；显式 `ImageSpecs` 里的 `layout` 配置仍然优先生效
- 生成最终 `docx` 后会写出 `layout-check.json`，检查图片数、图注数、残留占位符和常见实验报告章节，summary 里也会记录 `layoutCheckPassed`、错误数和警告数
- `layout-check.json` 会检查图注编号是否连续，summary 里会给出 `layoutCheckMessage`，便于不打开 JSON 也能快速判断排版是否过关

例如，直接用内置 `weekly-report` 生成周报输入包：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\generate-report-inputs.ps1 `
  -ReportProfileName "weekly-report" `
  -CourseName "校园导览小程序" `
  -ExperimentName "第 6 周迭代周报" `
  -StudentName "李四" `
  -StudentId "20261234" `
  -ClassName "软工 2302" `
  -TeacherName "王老师" `
  -ExperimentProperty "项目周报" `
  -ExperimentDate "第 6 周" `
  -ExperimentLocation "GitHub + 飞书 + 本地开发环境" `
  -DetailLevel full `
  -OutputDir ".\tests-output\weekly-report-sample"
```

如果你想同时生成全部示例 preset 的样例输入包：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-profile-preset-samples.ps1 `
  -OutputDir ".\tests-output\profile-preset-samples"
```

这个命令会额外写出 `profile-preset-samples.md`，方便直接预览每个 preset 对应的 `prompt.txt`、`metadata.auto.json` 和 `requirements.auto.json` 路径。

### 4. 飞书或直聊场景补充

如果你走 Feishu 或其他直接聊天场景，有几条经验是稳定有效的：

- 最稳的方式不是让模型临场拼很多中间 JSON，而是直接调用 `scripts/build-report-from-feishu.ps1`
- 飞书里手机直接上传截图时，可以参考 `examples/feishu-uploaded-images-docx-prompt.md`
- 电脑本地直接上传截图时，可以参考 `examples/local-uploaded-images-docx-prompt.md`
- 想一次生成、不等待图片分配确认时，可以参考 `examples/one-shot-uploaded-images-docx-prompt.md`
- 如果你 uploaded images and you also provide local image paths，建议把上传图片当作语义参考，把本地路径当作最终 `docx` 插图文件来源
- 如果运行时把附件提示注入成类似 `media/inbound/example.png` 这样的相对路径，这些路径也可以继续作为 `-ImagePaths` 传给插图流程
- 对于未标注章节的多张截图，脚本会按上传顺序优先把前半归入实验步骤、后半归入实验结果，再对同章节连续图片应用 2 列布局
- 如果聊天运行时根本没有暴露真实附件路径，就应该明确说不能直接插图，而不是假装已经写进 `docx`

### 5. 本地验证

在提交修改或排查问题前，建议先跑一遍烟测：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-smoke-tests.ps1
```

仓库还包含一个每天运行的 roadmap triage 自动化：`.github/workflows/roadmap-triage.yml`。它会调用 `scripts/analyze-roadmap-next-step.ps1`，读取 [ROADMAP.md](ROADMAP.md)，输出下一批更适合小步实现、且能被 smoke 覆盖的候选升级点。

本地也可以直接运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\analyze-roadmap-next-step.ps1 `
  -OutputDir ".\tests-output\roadmap-triage" `
  -Format markdown
```

## 项目现状

当前版本已经可以跑通一条完整的学校报告类文档工作流：生成正文、填模板、插图片、补图注、做最终排版。  
实验报告、课程设计报告、专业实习报告、软件测试报告、部署运维报告和结构化项目周报已经进入内置 profile 范围，并且都有 smoke 覆盖。
仓库还在持续迭代中，但当前优先级不是盲目扩场景，而是先把 profile-driven 这条链路做得更稳、更好用、更容易复现。

## Roadmap

当前方向很明确：先把学校报告类文档场景做实用，再逐步抽象成更通用的文档生成工具。

- 继续补常见实验报告、课程设计报告、实习报告模板和模板适配策略
- 强化教程链接、截图、正文、模板之间的自动串联能力
- 增强图片插入、图注和多图布局的配置能力
- 支持更多课程作业类和证据驱动类文档
- 逐步扩展到周报、月报、项目文档等可配置文档类型
- 继续完善样式 profile，让不同学校/模板的排版策略更容易切换

更完整的路线可以看 [ROADMAP.md](ROADMAP.md)。

## 仓库说明 / 协作

这个仓库已经包含开源协作所需的基础文件和流程，包括：

- [README.md](README.md)
- [LICENSE](LICENSE)
- [CONTRIBUTING.md](CONTRIBUTING.md)
- [CHANGELOG.md](CHANGELOG.md)
- [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md)
- [SECURITY.md](SECURITY.md)
- [SUPPORT.md](SUPPORT.md)
- `.github/` 下的 issue / PR 模板
- `.github/workflows/` 下的 CI 工作流

如果从 GitHub 公开项目的角度看，这一组文件基本已经构成了仓库当前的 `Repository Health` 基础面。

如果你想继续协作开发，先看 [CONTRIBUTING.md](CONTRIBUTING.md)；如果你只是想快速理解仓库目录和演示素材，可以先看 [demo/README.md](demo/README.md) 和 [examples](examples)。

## License

MIT
