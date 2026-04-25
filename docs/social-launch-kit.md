# 社媒发布素材包

这个文件用于把仓库包装成一个更容易传播的开源项目，重点覆盖 GitHub、小红书、抖音三种场景。

## GitHub 仓库简介

推荐仓库描述：

```text
OpenClaw skill + PowerShell pipeline for Chinese lab reports: generate report bodies, fill docx templates, insert screenshots, and run layout checks.
```

推荐 topics：

```text
openclaw
skill
powershell
docx
lab-report
template-filling
layout-check
openxml
university
```

## 小红书标题方向

- 我把实验报告写作最烦的 4 步做成了一个开源工具
- 这个开源项目可以把实验报告模板、截图、排版一次性收好
- 不想再手填 Word 实验报告了，我做了个能自动插图和排版的项目

## 小红书正文骨架

1. 先抛痛点：实验报告最烦的不是写字，是收截图、套模板、排版。
2. 再给结果：这个项目能把正文、模板、截图和 layout check 串起来。
3. 展示素材：放 `demo/README.md` 里的 2x2 预览图和最终 docx 截图。
4. 给行动指令：贴 GitHub 地址和一键演示命令。

## 抖音 / 视频号脚本骨架

15 到 30 秒即可：

1. 开场 3 秒
   画面：4 张截图 + 空白模板
   口播：实验报告最浪费时间的，其实是填模板和插图。

2. 中段 8 到 15 秒
   画面：运行 `scripts/run-one-click-demo.ps1`
   口播：这个开源项目把正文、模板、截图、排版检查串成了一条本地流水线。

3. 结尾 5 秒
   画面：最终生成的 `docx` + GitHub 仓库页
   口播：想自己试，仓库里直接带了一键演示。

## 配图建议

- README 顶部预览图
- `demo/assets/step-network-config.png`
- `demo/assets/result-ping.png`
- 最终 `styled.docx` 的首页截图
- `layout-check.json` 通过结果截图

## 评论区 / 结尾 CTA

- GitHub 已开源，仓库里自带模板、示例和一键演示
- 适合实验报告、课程设计、模板填充这类高重复文档场景
- 如果你也在折腾 Word / WPS 自动化，可以直接 fork
