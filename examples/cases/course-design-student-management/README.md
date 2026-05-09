# course-design-student-management：学生成绩管理系统课程设计

这个案例适合课程设计报告：题目不是单次验证实验，而是一个小型系统的需求分析、总体设计、模块实现、运行结果和总结。

## 输入文件

- `prompt.md`：课程设计报告生成约束
- `report.txt`：可直接进入课程设计模板填充的示例正文
- `metadata.json`：课程设计封面和基础字段
- `requirements.json`：课程设计报告校验规则

## 可运行命令

课程设计必须使用 `course-design-report` profile：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report.ps1 `
  -TemplatePath ".\examples\report-templates\course-design-report-template.docx" `
  -ReportPath ".\examples\cases\course-design-student-management\report.txt" `
  -MetadataPath ".\examples\cases\course-design-student-management\metadata.json" `
  -RequirementsPath ".\examples\cases\course-design-student-management\requirements.json" `
  -ReportProfileName course-design-report `
  -OutputDir ".\tests-output\course-design-student-management-case" `
  -StyleFinalDocx `
  -StyleProfile auto
```

## 适配建议

如果你有系统截图，建议把登录界面、主界面、成绩录入、查询统计和异常提示分别放到 `image-specs.json`。总体设计图或流程图应单独大图展示，不要自动并排压缩。
