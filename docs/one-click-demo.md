# 一键演示流程

这个演示的目标不是“调用模型生成正文”，而是让第一次接触仓库的人，不需要额外准备模板、截图和 JSON，就能把核心流水线跑通。

## 覆盖能力

演示会实际跑完这些步骤：

1. 读取仓库自带样例正文
2. 填充实验报告模板
3. 插入 4 张演示截图
4. 输出图片分配预案
5. 进行最终样式处理
6. 生成 `layout-check.json` 和 `summary.json`

## 前置条件

- Windows PowerShell
- 已 clone 本仓库
- 本地可运行 PowerShell 脚本

这个演示默认不要求 OpenClaw CLI 在线生成正文，因为正文、模板和截图都已经放在仓库里。

## 直接运行

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-one-click-demo.ps1
```

如果你想指定输出目录：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-one-click-demo.ps1 `
  -OutputDir "E:\reports\demo-output"
```

## 演示输入来自哪里

- 模板：`examples/report-templates/experiment-report-template.docx`
- 正文：`examples/demo-one-click/report.txt`
- metadata：`examples/demo-one-click/metadata.json`
- requirements：`examples/demo-one-click/requirements.json`
- 截图：`demo/assets/*.png`

## 运行后会产出什么

默认输出目录位于：

```text
tests-output/one-click-demo-时间戳/
```

关键文件包括：

- `*.styled.docx`：最终成品
- `generated-field-map.json`：模板字段映射
- `generated-image-map.json`：正式插图计划
- `image-placement-plan.md`：可读的插图预案
- `layout-check.json`：布局检查结果
- `summary.json`：整条流水线摘要

## 如何替换成你自己的材料

最简单的替换方式有两种：

1. 直接把 `run-one-click-demo.ps1` 里的输入路径换成你自己的模板、正文和截图。
2. 不再使用演示脚本，改用 [scripts/build-report.ps1](../scripts/build-report.ps1)、[scripts/build-report-from-url.ps1](../scripts/build-report-from-url.ps1) 或 [scripts/build-report-from-feishu.ps1](../scripts/build-report-from-feishu.ps1)。

如果你想做真正的“教程链接 -> 正文 -> 模板 -> 插图 -> 成品”完整链路，请改用：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report-from-url.ps1 ...
```

## 推荐阅读

- [examples/README.md](../examples/README.md)
- [scripts/README.md](../scripts/README.md)
- [demo/README.md](../demo/README.md)
