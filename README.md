# OpenClaw Experiment Report Skill
> 面向中文大学实验报告 / 课程设计报告的 OpenClaw skill + PowerShell 本地流水线。  
> 把“实验题目、教程链接、正文、模板、截图、排版检查”收成一条可复查、可复用、可交付的 `docx` 生成链路。

[![Quality Checks](https://github.com/lyf94697-droid/openclaw-experiment-report-skill/actions/workflows/quality.yml/badge.svg)](https://github.com/lyf94697-droid/openclaw-experiment-report-skill/actions/workflows/quality.yml)
[![Smoke Tests](https://github.com/lyf94697-droid/openclaw-experiment-report-skill/actions/workflows/smoke-tests.yml/badge.svg)](https://github.com/lyf94697-droid/openclaw-experiment-report-skill/actions/workflows/smoke-tests.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

## 项目定位

这个仓库不是“万能文档生成器”，而是一个把中文大学实验报告场景做深、做稳的开源项目：

- 先生成结构化中文实验报告正文
- 再把正文塞进 `docx` / Word / WPS 模板
- 再插入截图、生成图注、处理多图布局
- 最后做样式收尾和 layout check，得到可检查的成品

如果你经常遇到这些痛点，这个项目就是为你准备的：

- 实验题目、教程、截图、代码、结果散落在不同地方
- 学校模板是空白 `docx`，填字段和填正文都麻烦
- 截图不知道该插到哪一节、图注怎么写、几张图怎么排
- 最后交作业前还要手动修标题、空行、分页和图注编号

## 项目亮点

- 默认主线聚焦 `experiment-report`，优先把常见中文实验报告做稳
- 显式支持 `course-design-report`，适合课程设计报告和学校固定模板
- 支持“已有正文直接填模板”以及“从教程链接生成正文再填模板”两种路径
- 支持截图插入、图注生成、每行 2 图 / 2x2 图片块等常见布局
- 支持模板诊断、字段映射、布局检查和样式 profile
- 附带自检、烟测、GitHub workflow、Issue / PR 模板和开源协作文档
- 仓库内置演示素材，适合做 GitHub README、短视频录屏或社媒展示

## 演示预览

| 步骤截图 | 结果截图 |
| --- | --- |
| ![Network config preview](demo/assets/step-network-config.png) | ![Ping result preview](demo/assets/result-ping.png) |

完整演示素材、2x2 拼图预览和展示建议见 [demo/README.md](demo/README.md)。

## 3 分钟上手

### 1. 安装 skill

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install-skill.ps1
```

安装完成后可以在 OpenClaw 中确认：

```powershell
openclaw skills list
```

### 2. 跑仓库自带的一键演示

这个演示不依赖 OpenClaw 在线生成正文，直接用仓库自带样例走完“填模板 + 插图 + 排版 + layout check”。

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-one-click-demo.ps1
```

默认会在 `tests-output/one-click-demo-时间戳/` 下生成：

- 最终 `docx`
- `generated-field-map.json`
- `image-placement-plan.md`
- `layout-check.json`
- `summary.json`

详细说明见 [docs/one-click-demo.md](docs/one-click-demo.md)。

### 3. 替换成你自己的材料

最常见的 3 个入口如下：

```powershell
# 已有正文 + 模板，直接走本地确定性流程
powershell -ExecutionPolicy Bypass -File .\scripts\build-report.ps1 `
  -TemplatePath "E:\reports\template.docx" `
  -ReportPath ".\examples\sample-report.txt" `
  -MetadataPath ".\examples\docx-report-metadata.json" `
  -ImageSpecsPath ".\examples\docx-image-specs-row.json" `
  -RequirementsPath ".\examples\e2e-sample-requirements.json" `
  -StyleFinalDocx `
  -StyleProfile auto
```

```powershell
# 从教程链接直接走到最终 docx
powershell -ExecutionPolicy Bypass -File .\scripts\build-report-from-url.ps1 `
  -ReferenceUrls "https://blog.csdn.net/..." `
  -CourseName "计算机网络" `
  -TemplatePath "E:\reports\template.docx" `
  -StudentName "张三" `
  -StudentId "20260001" `
  -ClassName "计科 2201"
```

```powershell
# 飞书 / 直聊场景的本地封装入口
powershell -ExecutionPolicy Bypass -File .\scripts\build-report-from-feishu.ps1 `
  -ReferenceUrls "https://blog.csdn.net/..." `
  -CourseName "计算机网络" `
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

对应说明见 [docs/course-design-fastline.md](docs/course-design-fastline.md)。

## 固定版式规范

这些规则是当前默认交付标准，后续对话接手本仓库时也应按此执行：

- 实验报告套原模板外框时，顶部信息表保留正常表格线；正文只放进一个外框区域，不在段落或小节之间额外加横线。
- 实验报告默认按老师优秀示例风格排：`云南师范大学信息学院` 居中在前，标题为 `实验报告`，正文 12pt 宋体，一级标题约 15pt 宋体加粗，报告标题约 22pt 宋体，正文密度接近优秀样例。
- 实验报告截图默认大图居中，宽度下限为 `15.8 cm`；不要自动把多张截图压成两列，除非输入里显式指定 row/2x2 布局。
- 课程设计报告的流程图 / 总体设计图默认单独放大，宽度下限为 `15.8 cm`，不要和其他截图自动并排压小。
- 课程设计流程图标题只保留居中文字，不加标题左右两侧的装饰横线。
- 生成成品后至少跑 layout check；关键交付建议导出 PDF 或打开页面检查外框、图宽和线条是否重叠。
- 不把这些规则扩展成通用绘图引擎，只服务课程设计 / 实验报告交付。

## 仓库目录

```text
openclaw-experiment-report-skill/
├─ demo/                  GitHub / 小红书 / 抖音友好的演示素材
├─ docs/                  使用说明、演示流程、发布清单
├─ examples/              样例正文、JSON、Prompt、模板
├─ profiles/              报告 profile 定义
├─ references/            skill 运行时参考规则
├─ scripts/               主流程脚本与辅助脚本
├─ agents/                OpenClaw agent 配置
├─ SKILL.md               skill 主说明
└─ README.md              项目首页
```

如果你是第一次看这个仓库，建议按这个顺序打开：

1. [README.md](README.md)
2. [docs/one-click-demo.md](docs/one-click-demo.md)
3. [examples/README.md](examples/README.md)
4. [scripts/README.md](scripts/README.md)

## 文档导航

- [docs/README.md](docs/README.md)：文档总导航
- [docs/one-click-demo.md](docs/one-click-demo.md)：一键演示流程
- [docs/course-design-fastline.md](docs/course-design-fastline.md)：课程设计报告快线
- [docs/social-launch-kit.md](docs/social-launch-kit.md)：GitHub / 小红书 / 抖音发布素材建议
- [demo/README.md](demo/README.md)：演示素材和 2x2 布局预览
- [examples/README.md](examples/README.md)：示例文件总览

## 当前范围与边界

当前稳定范围：

- 中文大学实验报告
- 显式 opt-in 的课程设计报告
- 基于本地脚本的模板填充、截图插入、样式收尾和 layout check

当前不承诺的范围：

- 任意复杂学校模板都能零人工适配
- GUI 自动点 Word / WPS 完成填写
- 完全脱离 OpenClaw 的“通用写作平台”

后续方向见 [ROADMAP.md](ROADMAP.md)。

## 验证方式

提交前建议至少跑一遍烟测：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-smoke-tests.ps1
```

如果你只想确认本地 OpenClaw 环境和浏览器 profile 是否正常：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\self-check.ps1
```

## 开源协作

仓库已包含：

- [CONTRIBUTING.md](CONTRIBUTING.md)
- [CHANGELOG.md](CHANGELOG.md)
- [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md)
- [SECURITY.md](SECURITY.md)
- [SUPPORT.md](SUPPORT.md)
- `.github/` 下的 issue / PR 模板与 workflow

如果你想把这个仓库公开发布，先看 [docs/GITHUB_LAUNCH.md](docs/GITHUB_LAUNCH.md)；如果你还希望同步做小红书 / 抖音内容，直接看 [docs/social-launch-kit.md](docs/social-launch-kit.md)。

## License

MIT
