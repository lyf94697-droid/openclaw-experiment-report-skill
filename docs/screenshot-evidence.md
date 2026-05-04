# 截图证据如何嵌入报告

<!-- project-readiness:caption-policy -->

截图在实验报告里不是附件堆叠，而是证据。每张图都应该说明它证明了什么，并放在最接近相关文字的位置。

## 输入方式

最常用的是 `image-specs`：

```json
{
  "images": [
    {
      "path": ".\\demo\\assets\\result-ping.png",
      "section": "实验结果",
      "caption": "图1 ping 命令连通性测试结果",
      "widthCm": 15.8
    }
  ]
}
```

字段含义：

- `path`：本地图片路径，建议使用仓库根目录相对路径或绝对路径
- `section`：图片应插入的报告章节，例如 `实验步骤`、`实验结果`
- `caption`：中文图注，说明截图证明的事实
- `widthCm`：图片宽度，普通实验报告大图建议不低于 `15.8`
- `layout`：可选，指定并排、2x2 或分组布局

## 插图流程

日常使用可直接交给 `build-report.ps1`：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report.ps1 `
  -TemplatePath ".\examples\report-templates\experiment-report-template.docx" `
  -ReportPath ".\examples\sample-report.txt" `
  -MetadataPath ".\examples\cases\network-dos\metadata.json" `
  -ImageSpecsPath ".\examples\cases\network-dos\image-specs.json" `
  -RequirementsPath ".\examples\cases\network-dos\requirements.json" `
  -OutputDir ".\tests-output\network-dos-with-images" `
  -StyleFinalDocx
```

脚本会生成：

- `generated-image-map.json`：正式插图映射
- `image-placement-plan.md`：每张图的插入位置、图注和置信度
- `layout-check.json`：图片宽度、图注和布局检查结果

## 图注规则

好的图注应该短、具体、能回扣实验目的：

- `图1 主机 IP 地址配置结果`
- `图2 ping 命令连通性测试结果`
- `图3 进程调度算法运行结果`
- `图4 学生成绩管理系统查询界面`

不要写成：

- `截图1`
- `运行结果`
- `实验图片`
- `如下图所示`

图注不是装饰文字，它要告诉读者这张图能证明什么。

## 章节放置

推荐映射：

- 安装、配置、命令输入：放在 `实验步骤`
- 程序输出、命令结果、测试通过：放在 `实验结果`
- 报错、异常、对比输出：放在 `问题分析`
- 拓扑图、流程图、架构图：放在 `实验原理`、`方案设计与实现` 或 `总体设计`

报告正文中应在图片前后提到它，例如“连通性测试结果如图2所示”。不要把所有截图集中塞到最后。

## 并排布局

普通实验报告默认优先大图居中。只有明显成对的截图才建议并排：

```json
{
  "path": ".\\demo\\assets\\step-network-config.png",
  "section": "实验结果",
  "caption": "图1 主机 A 配置结果",
  "widthCm": 7.8,
  "layout": {
    "mode": "row",
    "columns": 2,
    "group": "network-results"
  }
}
```

同一个 `group` 的图片会尽量放在同一行。并排时要保证内容仍然清晰，不要为了省页数把命令行截图压到看不清。

## 课程设计图

课程设计报告里的总体设计图、流程图和模块关系图要按设计证据处理：

- 默认单独放大，不参与普通截图的两列布局
- 宽度建议不低于 `15.8 cm`
- 图注写清楚图的含义，例如 `图1 系统总体功能结构`
- 生成黑白流程图时，标题保持干净，不添加多余装饰线

## 交付前复核

插图后至少检查：

- 图片文件是否真实存在
- 图注是否连续编号
- 图片是否靠近相关段落
- 版式检查是否通过
- 打开最终 `docx` 后图片没有压缩到看不清，也没有覆盖文字或表格线
