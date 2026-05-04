# OpenClaw Experiment Report Skill

> 面向中文大学实验报告 / 课程设计报告的 OpenClaw skill + PowerShell 本地流水线。  
> 把“实验题目、教程链接、正文、学校模板、截图证据、版式检查”收成一条可复查、可复用、可交付的 `docx` 生成链路。

[![Quality Checks](https://github.com/lyf94697-droid/openclaw-experiment-report-skill/actions/workflows/quality.yml/badge.svg)](https://github.com/lyf94697-droid/openclaw-experiment-report-skill/actions/workflows/quality.yml)
[![Smoke Tests](https://github.com/lyf94697-droid/openclaw-experiment-report-skill/actions/workflows/smoke-tests.yml/badge.svg)](https://github.com/lyf94697-droid/openclaw-experiment-report-skill/actions/workflows/smoke-tests.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

## 项目定位

这个仓库不是“万能文档生成器”，而是一个把中文大学实验报告场景做深、做稳的开源项目：

- 先生成或接收结构化中文报告正文
- 再把正文和基础信息填进 `docx` / Word / WPS 模板
- 再插入截图、生成图注、处理多图布局
- 最后做样式收尾和 layout check，得到可检查的成品

适合解决这些问题：

- 实验要求、教程、截图、代码和结果散落在不同地方
- 学校模板是空白 `docx`，字段和正文都需要手工填
- 截图不知道该插到哪一节、图注怎么写、几张图怎么排
- 交作业前还要手动修标题、空行、分页、外框和图注编号

## 项目亮点

- 默认主线聚焦 `experiment-report`，优先把常见中文实验报告做稳
- 显式支持 `course-design-report`，适合课程设计报告和学校固定模板
- 支持“已有正文直接填模板”和“从教程链接生成正文再填模板”两条路径
- 支持截图插入、图注生成、每行 2 图 / 2x2 图片块等常见布局
- 支持模板诊断、字段映射、布局检查和样式 profile
- 附带一键 demo、项目就绪检查、烟测、GitHub workflow、Issue / PR 模板和开源协作文档
- 仓库内置演示素材和 3 个典型案例，适合做 GitHub 展示、录屏或简历项目

## 演示预览

| 步骤截图 | 结果截图 |
| --- | --- |
| ![Network config preview](demo/assets/step-network-config.png) | ![Ping result preview](demo/assets/result-ping.png) |

完整演示素材、2x2 拼图预览和展示建议见 [demo/README.md](demo/README.md)。

<!-- project-readiness:usage-tutorial -->
## 中文使用教程

### 1. 安装 skill

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install-skill.ps1
```

安装完成后可以在 OpenClaw 中确认：

```powershell
openclaw skills list
```

### 2. 跑一键演示

这个演示不依赖在线生成正文，直接用仓库自带正文、模板和截图走完“填模板 + 插图 + 排版 + layout check”。

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-one-click-demo.ps1
```

默认会在 `tests-output/one-click-demo-时间戳/` 下生成最终 `docx`、字段映射、图片映射、图片放置计划、layout check 和摘要文件。

### 3. 检查项目展示材料

如果你要把仓库发到 GitHub、写进简历或拿给别人试用，先跑：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\check-project-readiness.ps1
```

它会检查 README、关键 docs、典型案例和案例 JSON 是否齐全。

### 4. 已有正文 + 学校模板

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report.ps1 `
  -TemplatePath ".\examples\report-templates\experiment-report-template.docx" `
  -ReportPath ".\examples\sample-report.txt" `
  -MetadataPath ".\examples\cases\network-dos\metadata.json" `
  -ImageSpecsPath ".\examples\cases\network-dos\image-specs.json" `
  -RequirementsPath ".\examples\cases\network-dos\requirements.json" `
  -OutputDir ".\tests-output\network-dos-case" `
  -StyleFinalDocx `
  -StyleProfile auto
```

### 5. 教程链接 + 截图 + 学校模板

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report-from-url.ps1 `
  -ReferenceUrls "https://blog.csdn.net/..." `
  -CourseName "计算机网络" `
  -ExperimentName "局域网搭建与常用 DOS 命令使用" `
  -TemplatePath "E:\reports\template.docx" `
  -StudentName "张三" `
  -StudentId "20260001" `
  -ClassName "计科 2201" `
  -ImagePaths "E:\reports\step-1.png","E:\reports\result-1.png" `
  -OutputDir "E:\reports\final-output"
```

### 6. 飞书 / 直聊场景

如果材料来自飞书、聊天窗口或本地附件路径，可以使用本地 wrapper，把正文、截图和生成产物归档到一个输出目录：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report-from-feishu.ps1 `
  -ReferenceUrls "https://blog.csdn.net/..." `
  -CourseName "计算机网络" `
  -ExperimentName "局域网搭建与常用 DOS 命令使用" `
  -TemplatePath "E:\reports\template.docx" `
  -StudentName "张三" `
  -StudentId "20260001" `
  -ClassName "计科 2201" `
  -ImagePaths "E:\reports\step-1.png","E:\reports\result-1.png" `
  -OutputDir "E:\reports\final-output"
```

课程设计报告请显式传：

```powershell
-ReportProfileName course-design-report
```

更完整的流程见 [docs/usage-flow.md](docs/usage-flow.md)。

<!-- project-readiness:input-output -->
## 输入与输出示例

### 典型输入

| 输入 | 说明 | 示例 |
| --- | --- | --- |
| `TemplatePath` | 学校模板或示例模板 | `examples/report-templates/experiment-report-template.docx` |
| `ReportPath` | 已有报告正文 | `examples/sample-report.txt` |
| `MetadataPath` | 姓名、学号、课程名、实验名等短字段 | `examples/cases/network-dos/metadata.json` |
| `RequirementsPath` | 章节、关键词和禁用词检查 | `examples/cases/network-dos/requirements.json` |
| `ImageSpecsPath` | 截图路径、图注、章节和布局 | `examples/cases/network-dos/image-specs.json` |
| `ReferenceUrls` | 教程或 CSDN 参考链接 | `https://blog.csdn.net/...` |

### 典型输出

```text
tests-output/network-dos-case/
├─ summary.json
├─ generated-field-map.json
├─ generated-image-map.json
├─ image-placement-plan.md
├─ layout-check.json
├─ report.cleaned.txt
└─ <最终报告>.docx
```

其中 `summary.json` 用来追踪本次生成链路，`layout-check.json` 用来检查图片、图注和版式，最终 `docx` 用来人工打开复核和交付。

<!-- project-readiness:scenarios -->
## 适用场景

- 中文大学实验报告：计算机网络、操作系统、数据库、软件测试、程序设计等课程
- 课程设计报告：需要需求分析、方案设计、运行结果和设计总结的小型项目
- 学校固定模板：老师给了 `.docx` 模板，需要保留字段、表格、外框和标题结构
- 截图证据型作业：报告必须嵌入命令输出、程序界面、流程图或测试结果截图
- 项目展示：需要一个能演示“提示词 + 本地脚本 + 文档自动化 + 质量检查”的简历项目

## 典型案例

- [network-dos](examples/cases/network-dos/README.md)：局域网搭建与 DOS 命令实验，带 4 张截图，可直接跑完整 `docx` 流程
- [os-process-scheduling](examples/cases/os-process-scheduling/README.md)：操作系统进程调度算法实验，适合看算法类报告正文约束
- [course-design-student-management](examples/cases/course-design-student-management/README.md)：学生成绩管理系统课程设计，演示 `course-design-report` profile

<!-- project-readiness:limitations -->
## 限制说明

- 不能保证任意复杂学校模板都零人工适配，复杂模板仍需要看字段映射 diagnostics
- 从 CSDN 或公开教程生成正文时，参考内容只提供背景和流程，结果必须以用户真实截图和数据为准
- 这个项目能降低照抄风险，但不能自动保证学术合规；提交前仍需使用者复核
- 图像插入依赖本地图片路径可访问，截图模糊或信息不足时不能凭空补细节
- `build-report-from-url.ps1` 需要 OpenClaw CLI 和浏览器 profile 可用
- Word / WPS GUI 自动操作不是主路径；项目优先生成可检查的 `docx` 文件

后续方向见 [ROADMAP.md](ROADMAP.md)。

## 仓库目录

```text
openclaw-experiment-report-skill/
├─ demo/                  GitHub / 小红书 / 抖音友好的演示素材
├─ docs/                  使用流程、模板机制、CSDN 改写、截图证据等文档
├─ examples/              典型案例、样例正文、JSON、Prompt、模板
├─ profiles/              报告 profile 定义
├─ references/            skill 运行时参考规则
├─ scripts/               主流程脚本、辅助脚本和检查脚本
├─ agents/                OpenClaw agent 配置
├─ SKILL.md               skill 主说明
└─ README.md              项目首页
```

第一次阅读建议顺序：

1. [docs/usage-flow.md](docs/usage-flow.md)
2. [examples/cases/README.md](examples/cases/README.md)
3. [docs/template-filling.md](docs/template-filling.md)
4. [docs/screenshot-evidence.md](docs/screenshot-evidence.md)
5. [docs/csdn-reference-policy.md](docs/csdn-reference-policy.md)

## 文档导航

- [docs/README.md](docs/README.md)：文档总导航
- [docs/usage-flow.md](docs/usage-flow.md)：完整使用流程
- [docs/template-filling.md](docs/template-filling.md)：模板填充机制
- [docs/csdn-reference-policy.md](docs/csdn-reference-policy.md)：CSDN 参考内容如何避免照抄
- [docs/screenshot-evidence.md](docs/screenshot-evidence.md)：截图证据如何嵌入报告
- [docs/one-click-demo.md](docs/one-click-demo.md)：一键演示流程
- [docs/course-design-fastline.md](docs/course-design-fastline.md)：课程设计报告快线
- [demo/README.md](demo/README.md)：演示素材和 2x2 布局预览
- [examples/README.md](examples/README.md)：示例文件总览

## 验证方式

项目展示材料检查：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\check-project-readiness.ps1
```

一键 demo：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-one-click-demo.ps1
```

完整烟测：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-smoke-tests.ps1
```

本地 OpenClaw 环境检查：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\self-check.ps1
```

## 开源协作

仓库已包含 [CONTRIBUTING.md](CONTRIBUTING.md)、[CHANGELOG.md](CHANGELOG.md)、[CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md)、[SECURITY.md](SECURITY.md)、[SUPPORT.md](SUPPORT.md)，以及 `.github/` 下的 issue / PR 模板和 workflow。

如果你想公开发布，先看 [docs/GITHUB_LAUNCH.md](docs/GITHUB_LAUNCH.md)；如果要同步做小红书 / 抖音内容，看 [docs/social-launch-kit.md](docs/social-launch-kit.md)。

## License

MIT
