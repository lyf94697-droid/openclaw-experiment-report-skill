# Web UI Demo

This example shows how to start the local Web UI and generate DOCX, PDF, and preview PNG artifacts from browser-uploaded materials.

## Start

Install the optional UI dependencies:

```powershell
python -m pip install -r requirements-web.txt
```

Start the UI:

```powershell
python web_ui.py
```

Open:

```text
http://127.0.0.1:7860
```

## Inputs

Fill these fields:

- 报告类型
- 生成方式
- 正文长度
- 课程名称
- 学生姓名
- 学号
- 班级
- 实验名称/题目名称
- 实验要求
- 参考链接或补充说明
- 对话式需求
- 本地截图文件夹/文件路径
- 本地代码文件夹/文件路径

Upload:

- an optional `.docx` or `.doc` template
- one or more result screenshots
- one or more code files

If no template is uploaded, experiment reports use `E:\实验报告\00-模板\实验报告模版1.docx` when available. Course-design reports use `E:\新建文件夹\课程设计-模板.doc` when available.

The chat-style box can accept text like:

```text
CSDN链接：https://example.com/article
课程名称：计算机网络
实验名称：根据教程链接填充
姓名：李亦非
学号：2444100198
班级：24C
截图材料："E:\实验报告\截图\计网实验六"
```

Manual fields take priority. Empty fields are filled from the chat-style text where possible.

## Output

Click `生成报告`. The page shows:

- generation status
- warnings or errors
- a DOCX download button
- a PDF download button
- a preview PNG download button
- an in-page preview image

Generated artifacts are copied to the selected output root, defaulting to:

```text
E:\实验报告\docx
E:\实验报告\pdf
E:\实验报告\预览图
```

The working files are also kept under:

```text
outputs/web-ui/
```

## Notes

- `智能长文（接近对话效果）` uses the local OpenClaw chat gateway when available.
- If the chat gateway is unavailable, the UI falls back to `快速本地草稿` and shows the reason in the warning box.
- PDF export requires WPS, Microsoft Word, or LibreOffice. Preview PNG rendering uses PyMuPDF.
