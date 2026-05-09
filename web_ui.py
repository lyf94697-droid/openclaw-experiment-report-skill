from __future__ import annotations

import html
import json
import os
import re
import shutil
import subprocess
import tempfile
import urllib.request
from datetime import datetime
from pathlib import Path
from typing import Any

try:
    import gradio as gr
except ImportError as exc:  # pragma: no cover - shown only when the optional UI dependency is missing.
    raise SystemExit(
        "Gradio is not installed. Install it with: python -m pip install -r requirements-web.txt"
    ) from exc


REPO_ROOT = Path(__file__).resolve().parent
WEB_OUTPUT_ROOT = REPO_ROOT / "outputs" / "web-ui"
DEFAULT_EXPERIMENT_TEMPLATE = Path(r"E:\实验报告\00-模板\实验报告模版1.docx")
DEFAULT_COURSE_DESIGN_TEMPLATE = Path(r"E:\新建文件夹\课程设计-模板.doc")
DEFAULT_DELIVERY_ROOT = Path(r"E:\实验报告")

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp"}
TEMPLATE_EXTENSIONS = {".docx", ".doc"}
CODE_EXTENSIONS = {
    ".java",
    ".py",
    ".c",
    ".cpp",
    ".h",
    ".hpp",
    ".cs",
    ".js",
    ".ts",
    ".html",
    ".css",
    ".sql",
    ".xml",
    ".json",
    ".yaml",
    ".yml",
    ".md",
    ".txt",
    ".ps1",
    ".bat",
    ".sh",
    ".properties",
    ".ini",
    ".conf",
}
MAX_UPLOAD_COUNT = 60
MAX_UPLOAD_BYTES = 35 * 1024 * 1024
REFERENCE_FETCH_TIMEOUT = 75
SMART_GENERATION_TIMEOUT = 420
LOCAL_GENERATION_TIMEOUT = 240
PDF_EXPORT_TIMEOUT = 180
ALLOWED_UPLOAD_ROOTS = (Path(tempfile.gettempdir()).resolve(),)
AUTO_TITLE_VALUES = {
    "",
    "根据教程链接填充",
    "根据链接填充",
    "自动填充",
    "自动识别",
    "按教程填写",
}


class UploadValidationError(ValueError):
    pass


def _is_relative_to(path: Path, root: Path) -> bool:
    try:
        path.relative_to(root)
        return True
    except ValueError:
        return False


def _coerce_upload_path(file_obj: Any) -> Path | None:
    if file_obj is None:
        return None
    if isinstance(file_obj, (str, Path)):
        return Path(file_obj)
    for attr in ("name", "path"):
        value = getattr(file_obj, attr, None)
        if value:
            return Path(value)
    if isinstance(file_obj, dict):
        value = file_obj.get("name") or file_obj.get("path")
        if value:
            return Path(value)
    return None


def normalize_upload_path(file_obj: Any, label: str = "上传文件") -> Path | None:
    path = _coerce_upload_path(file_obj)
    if path is None:
        return None

    try:
        resolved = path.expanduser().resolve(strict=True)
    except OSError as exc:
        raise UploadValidationError(f"{label}不存在或不可访问：{path}") from exc

    if not resolved.is_file():
        raise UploadValidationError(f"{label}必须是文件：{resolved}")

    allow_local_paths = os.environ.get("OPENCLAW_WEB_UI_ALLOW_LOCAL_PATHS") == "1"
    if not allow_local_paths and not any(
        resolved == root or _is_relative_to(resolved, root) for root in ALLOWED_UPLOAD_ROOTS
    ):
        raise UploadValidationError(
            f"{label}不是浏览器上传的临时文件，已拒绝读取本机路径：{resolved}。"
            "如果是在可信本机脚本中测试，请设置 OPENCLAW_WEB_UI_ALLOW_LOCAL_PATHS=1。"
        )

    size = resolved.stat().st_size
    if size > MAX_UPLOAD_BYTES:
        raise UploadValidationError(
            f"{label}过大：{resolved.name} 为 {size} 字节，单文件上限为 {MAX_UPLOAD_BYTES} 字节。"
        )

    return resolved


def normalize_upload_list(files: Any, label: str = "上传文件") -> list[Path]:
    if not files:
        return []
    if not isinstance(files, list):
        files = [files]
    if len(files) > MAX_UPLOAD_COUNT:
        raise UploadValidationError(f"{label}数量过多：{len(files)} 个，最多 {MAX_UPLOAD_COUNT} 个。")
    paths: list[Path] = []
    for file_obj in files:
        path = normalize_upload_path(file_obj, label=label)
        if path is not None:
            paths.append(path)
    return paths


def require_allowed_extensions(paths: list[Path], allowed: set[str], label: str) -> list[Path]:
    rejected = [path.name for path in paths if path.suffix.lower() not in allowed]
    if rejected:
        raise UploadValidationError(f"{label}包含不支持的文件类型：{', '.join(rejected[:5])}")
    return paths


def parse_local_path_lines(text: str) -> list[str]:
    paths: list[str] = []
    for raw_line in (text or "").splitlines():
        line = html.unescape(raw_line).strip().strip('"“”')
        if not line:
            continue
        for part in re.split(r"\s*[;；]\s*", line):
            part = part.strip().strip('"“”')
            if part:
                paths.append(part)
    return paths


def collect_local_files(path_text: str, allowed: set[str], label: str) -> list[Path]:
    collected: list[Path] = []
    for raw_path in parse_local_path_lines(path_text):
        source = Path(raw_path).expanduser()
        if not source.exists():
            raise UploadValidationError(f"{label}路径不存在：{source}")
        if source.is_file():
            candidates = [source]
        elif source.is_dir():
            candidates = sorted(path for path in source.rglob("*") if path.is_file())
        else:
            continue

        for candidate in candidates:
            if candidate.suffix.lower() not in allowed:
                continue
            if candidate.stat().st_size > MAX_UPLOAD_BYTES:
                raise UploadValidationError(
                    f"{label}文件过大：{candidate.name} 为 {candidate.stat().st_size} 字节。"
                )
            collected.append(candidate.resolve())

    if len(collected) > MAX_UPLOAD_COUNT:
        raise UploadValidationError(f"{label}数量过多：{len(collected)} 个，最多 {MAX_UPLOAD_COUNT} 个。")
    return collected


def safe_filename(value: str, fallback: str = "report") -> str:
    value = value.strip() or fallback
    value = re.sub(r'[\\/:*?"<>|\r\n\t]+', "-", value)
    value = re.sub(r"\s+", "", value)
    return value[:90] or fallback


def unique_path(directory: Path, filename: str) -> Path:
    directory.mkdir(parents=True, exist_ok=True)
    candidate = directory / filename
    if not candidate.exists():
        return candidate
    stem = candidate.stem
    suffix = candidate.suffix
    for index in range(2, 200):
        next_candidate = directory / f"{stem}-{index}{suffix}"
        if not next_candidate.exists():
            return next_candidate
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    return directory / f"{stem}-{timestamp}{suffix}"


def copy_uploads(paths: list[Path], target_dir: Path) -> list[Path]:
    target_dir.mkdir(parents=True, exist_ok=True)
    copied: list[Path] = []
    used_names: set[str] = set()
    for index, source in enumerate(paths, start=1):
        suffix = source.suffix
        stem = safe_filename(source.stem, f"file-{index}")
        name = f"{stem}{suffix}"
        if name.lower() in used_names:
            name = f"{stem}-{index}{suffix}"
        used_names.add(name.lower())
        target = target_dir / name
        shutil.copy2(source, target)
        copied.append(target)
    return copied


def resolve_powershell_executable() -> str:
    for candidate in ("pwsh", "powershell"):
        resolved = shutil.which(candidate)
        if resolved:
            return resolved
    raise RuntimeError("未找到 PowerShell。请安装 PowerShell 7 或确认 powershell 在 PATH 中。")


def write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8-sig")


def write_failure_log(output_dir: Path, command: list[str], stdout: str, stderr: str) -> Path:
    log_dir = output_dir / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / f"generation-error-{datetime.now().strftime('%H%M%S')}.log"
    log_text = "\n".join(
        [
            "COMMAND:",
            " ".join(command),
            "",
            "STDOUT:",
            stdout,
            "",
            "STDERR:",
            stderr,
        ]
    )
    log_path.write_text(log_text, encoding="utf-8")
    return log_path


def run_command(command: list[str], cwd: Path, timeout: int) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        command,
        cwd=cwd,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
        check=False,
    )


def read_text_file(path: Path, max_chars: int = 2200) -> str:
    for encoding in ("utf-8-sig", "utf-8", "gbk", "latin-1"):
        try:
            text = path.read_text(encoding=encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        return f"{path.name}：无法读取文本内容，仅记录文件名。"

    text = text.replace("\r\n", "\n").replace("\r", "\n").strip()
    if len(text) > max_chars:
        text = text[:max_chars].rstrip() + "\n..."
    return text


def parse_reference_input(reference_links: str) -> tuple[list[str], str]:
    urls: list[str] = []
    notes: list[str] = []
    url_pattern = re.compile(r"https?://[^\s<>'\"]+", re.IGNORECASE)
    for raw_line in (reference_links or "").splitlines():
        line = html.unescape(raw_line).strip()
        if not line:
            continue
        found = [match.rstrip("，。；;、)") for match in url_pattern.findall(line)]
        if found:
            urls.extend(found)
            residue = url_pattern.sub("", line).strip(" ：:-")
            if residue:
                notes.append(residue)
        else:
            notes.append(line)

    deduped_urls: list[str] = []
    seen: set[str] = set()
    for url in urls:
        if url not in seen:
            seen.add(url)
            deduped_urls.append(url)
    return deduped_urls, "\n".join(notes).strip()


def parse_chat_request(text: str) -> dict[str, Any]:
    if not text.strip():
        return {}

    def find_value(patterns: list[str]) -> str:
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
            if match:
                return html.unescape(match.group(1)).strip().strip('"“”')
        return ""

    screenshot_path = find_value(
        [
            r"截图材料\s*[:：]\s*([^\r\n]+)",
            r"截图文件夹\s*[:：]\s*([^\r\n]+)",
            r"截图路径\s*[:：]\s*([^\r\n]+)",
        ]
    )
    code_path = find_value(
        [
            r"代码文件夹\s*[:：]\s*([^\r\n]+)",
            r"代码路径\s*[:：]\s*([^\r\n]+)",
        ]
    )
    report_type = "课程设计报告" if "课程设计" in text and "实验报告" not in text[:120] else ""
    return {
        "report_type": report_type,
        "course_name": find_value([r"课程名称\s*[:：]\s*([^\r\n]+)", r"课程\s*[:：]\s*([^\r\n]+)"]),
        "experiment_name": find_value([r"实验名称\s*[:：]\s*([^\r\n]+)", r"题目名称\s*[:：]\s*([^\r\n]+)"]),
        "student_name": find_value([r"姓名\s*[:：]\s*([^\r\n]+)", r"学生姓名\s*[:：]\s*([^\r\n]+)"]),
        "student_id": find_value([r"学号\s*[:：]\s*([^\r\n]+)"]),
        "class_name": find_value([r"班级\s*[:：]\s*([^\r\n]+)"]),
        "screenshot_path": screenshot_path,
        "code_path": code_path,
        "text": text.strip(),
    }


def clean_request_text_for_body(text: str) -> str:
    cleaned_lines: list[str] = []
    metadata_pattern = re.compile(
        r"^\s*(课程名称|课程|实验名称|题目名称|姓名|学生姓名|学号|班级|截图材料|截图文件夹|截图路径|代码文件夹|代码路径|CSDN链接|参考链接)\s*[:：]\s*(.*)$",
        re.IGNORECASE,
    )
    requirement_pattern = re.compile(r"^\s*(实验要求|要求)\s*[:：]\s*(.*)$", re.IGNORECASE)
    for raw_line in (text or "").splitlines():
        line = raw_line.strip()
        if not line:
            cleaned_lines.append("")
            continue
        requirement_match = requirement_pattern.match(line)
        if requirement_match:
            if requirement_match.group(2).strip():
                cleaned_lines.append(requirement_match.group(2).strip())
            continue
        if metadata_pattern.match(line):
            continue
        cleaned_lines.append(line)

    return "\n".join(cleaned_lines).strip()


def first_nonempty(*values: str) -> str:
    for value in values:
        if str(value or "").strip():
            return str(value).strip()
    return ""


def is_auto_title(value: str) -> bool:
    normalized = re.sub(r"\s+", "", html.unescape(value or "").strip().strip('"“”'))
    return normalized in AUTO_TITLE_VALUES or "根据教程链接填充" in normalized


def clean_inferred_title(title: str) -> str:
    value = html.unescape(title or "").strip()
    value = re.sub(r"\s+", " ", value)
    if re.match(
        r"^\s*(课程名称|课程|姓名|学生姓名|学号|班级|截图材料|截图文件夹|截图路径|代码文件夹|代码路径|实验要求)\s*[:：]",
        value,
    ):
        return ""
    value = re.sub(r"[-_|\s]*(CSDN博客|CSDN|博客园|知乎|哔哩哔哩|bilibili).*$", "", value, flags=re.IGNORECASE)
    value = re.sub(r"^\s*(TITLE|标题|实验名称|题目名称)\s*[:：]\s*", "", value, flags=re.IGNORECASE)
    value = value.strip(" -_｜|:：[]【】()（）\"'“”")
    if is_auto_title(value):
        return ""
    if len(value) < 3 or len(value) > 80:
        return ""
    if re.match(r"^https?://", value, re.IGNORECASE):
        return ""
    return value


def infer_experiment_name(reference_paths: list[Path], requirement_text: str, reference_notes: str) -> str:
    texts: list[str] = []
    for path in reference_paths[:4]:
        texts.append(read_text_file(path, max_chars=5000))
    if requirement_text.strip():
        texts.append(requirement_text.strip())
    if reference_notes.strip():
        texts.append(reference_notes.strip())

    combined = "\n".join(texts)
    patterns = [
        r"(?im)^\s*TITLE\s*[:：]\s*(.+?)\s*$",
        r"(?im)^\s*标题\s*[:：]\s*(.+?)\s*$",
        r"(?im)^\s*实验名称\s*[:：]\s*(.+?)\s*$",
        r"(?im)^\s*题目名称\s*[:：]\s*(.+?)\s*$",
    ]
    for pattern in patterns:
        match = re.search(pattern, combined)
        if match:
            title = clean_inferred_title(match.group(1))
            if title:
                return title

    for line in combined.splitlines():
        title = clean_inferred_title(line)
        if title and ("实验" in title or "配置" in title or "网络" in title or "设计" in title):
            return title
    return ""


def fallback_fetch_url_text(url: str, max_chars: int = 12000) -> str:
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        },
    )
    with urllib.request.urlopen(request, timeout=25) as response:
        raw = response.read(max_chars * 4)
        charset = response.headers.get_content_charset() or "utf-8"
    text = raw.decode(charset, errors="replace")
    title_match = re.search(r"(?is)<title[^>]*>(.*?)</title>", text)
    title = ""
    if title_match:
        title = re.sub(r"\s+", " ", html.unescape(title_match.group(1))).strip()
    text = re.sub(r"(?is)<script.*?</script>|<style.*?</style>", "\n", text)
    text = re.sub(r"(?s)<[^>]+>", "\n", text)
    text = html.unescape(text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text).strip()
    if title:
        text = f"TITLE: {title}\nURL: {url}\n\n{text}"
    return text[:max_chars]


def fetch_reference_texts(urls: list[str], reference_dir: Path, warnings: list[str]) -> list[Path]:
    if not urls:
        return []

    reference_dir.mkdir(parents=True, exist_ok=True)
    powershell = resolve_powershell_executable()
    reference_paths: list[Path] = []
    for index, url in enumerate(urls, start=1):
        target = reference_dir / f"reference-{index:02}.txt"
        command = [
            powershell,
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(REPO_ROOT / "scripts" / "fetch-web-article.ps1"),
            "-Url",
            url,
            "-MaxChars",
            "30000",
        ]
        fetched_text = ""
        try:
            result = run_command(command, cwd=REPO_ROOT, timeout=REFERENCE_FETCH_TIMEOUT)
            if result.returncode == 0 and result.stdout.strip():
                fetched_text = result.stdout.strip()
            else:
                warnings.append(f"参考链接自动提取失败，已尝试普通网页读取：{url}")
        except Exception as exc:
            warnings.append(f"参考链接浏览器提取超时或失败，已尝试普通网页读取：{url}（{exc}）")

        if not fetched_text:
            try:
                fetched_text = "URL: {0}\n\n{1}".format(url, fallback_fetch_url_text(url))
            except Exception as exc:
                fetched_text = f"URL: {url}\n\n未能自动提取网页正文，请以页面截图和手工填写的实验要求为准。错误：{exc}"
                warnings.append(f"参考链接内容不足：{url}")

        target.write_text(fetched_text + "\n", encoding="utf-8-sig")
        reference_paths.append(target)

    return reference_paths


def resolve_report_profile(report_type: str) -> tuple[str, Path | None]:
    if report_type == "课程设计报告":
        return "course-design-report", DEFAULT_COURSE_DESIGN_TEMPLATE if DEFAULT_COURSE_DESIGN_TEMPLATE.exists() else None
    return "experiment-report", DEFAULT_EXPERIMENT_TEMPLATE if DEFAULT_EXPERIMENT_TEMPLATE.exists() else None


def resolve_template_path(template_file: Any, report_type: str) -> Path:
    try:
        uploaded = normalize_upload_path(template_file, label="报告模板")
    except UploadValidationError:
        raise

    report_profile, default_template = resolve_report_profile(report_type)
    _ = report_profile
    template_path = uploaded or default_template
    if template_path is None or not template_path.exists():
        raise UploadValidationError(
            "未上传模板，也没有找到默认模板。实验报告默认模板应在 "
            f"{DEFAULT_EXPERIMENT_TEMPLATE}；课程设计默认模板应在 {DEFAULT_COURSE_DESIGN_TEMPLATE}。"
        )
    if template_path.suffix.lower() not in TEMPLATE_EXTENSIONS:
        raise UploadValidationError("模板只支持 .docx 或 .doc。")
    return template_path


def make_base_name(student_id: str, student_name: str, experiment_name: str) -> str:
    parts = [
        safe_filename(student_id, "").strip(),
        safe_filename(student_name, "").strip(),
        safe_filename(experiment_name, "实验报告").strip(),
    ]
    return "-".join(part for part in parts if part) or f"实验报告-{datetime.now().strftime('%Y%m%d-%H%M%S')}"


def write_reference_notes(notes: str, target_dir: Path) -> list[Path]:
    if not notes.strip():
        return []
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / "reference-notes.txt"
    target.write_text(notes.strip() + "\n", encoding="utf-8-sig")
    return [target]


def build_code_context(code_paths: list[Path]) -> str:
    if not code_paths:
        return "未上传代码文件。"

    blocks: list[str] = []
    for index, path in enumerate(code_paths[:6], start=1):
        if path.suffix.lower() in CODE_EXTENSIONS:
            blocks.append(f"{index}. {path.name}\n{read_text_file(path, max_chars=1800)}")
        else:
            blocks.append(f"{index}. {path.name}\n非文本代码文件，仅记录文件名。")
    if len(code_paths) > 6:
        blocks.append(f"其余 {len(code_paths) - 6} 个代码文件仅记录文件名。")
    return "\n\n".join(blocks)


def build_screenshot_context(screenshot_paths: list[Path]) -> str:
    if not screenshot_paths:
        return "未上传运行截图。"
    return "\n".join(f"{index}. {path.name}" for index, path in enumerate(screenshot_paths, start=1))


def build_prompt_text(
    report_type: str,
    course_name: str,
    experiment_name: str,
    requirement_text: str,
    reference_urls: list[str],
    reference_notes: str,
    screenshot_paths: list[Path],
    code_paths: list[Path],
    detail_level: str,
) -> str:
    length_hint = "正文写长一点，按优秀实验报告密度组织，约 2200 到 3500 字。" if detail_level == "full" else "正文保持完整但不要冗长，约 1200 到 1800 字。"
    screenshots = build_screenshot_context(screenshot_paths)
    code_context = build_code_context(code_paths)
    reference_block = "\n".join(reference_urls) if reference_urls else "未填写参考链接。"
    notes_block = reference_notes or "无额外参考说明。"

    return f"""请根据下面材料生成一份可直接进入模板排版的中文{report_type}正文。

基础信息：
课程名称：{course_name}
报告题目：{experiment_name}
详细程度：{length_hint}

写作要求：
1. 正文必须具体，按分点和小节写，不要空泛套话。
2. 截图要放在对应步骤或结果附近，正文中用“如图1所示”这类表述呼应截图。
3. 图注保持短句，放在图片下方，居中，无首行缩进。
4. 对参考教程只做学习整理，不照搬大段原文；以实验要求、截图、代码和运行结果为准。
5. 如果材料不足，明确写“根据已提供材料整理”，不要编造具体 IP、端口、账号、测试数值。
6. 实验报告优先包含：实验目的、实验环境、实验原理或任务要求、实验步骤、关键代码或配置、实验结果、问题分析、实验总结。
7. 课程设计报告优先包含：需求分析、系统设计、数据库或模块设计、详细实现、测试结果、总结。

实验要求：
{requirement_text.strip() or "未填写详细实验要求，请根据课程名称、题目、截图和代码整理完整报告。"}

参考链接：
{reference_block}

补充说明：
{notes_block}

运行截图文件：
{screenshots}

代码或配置文件摘录：
{code_context}
"""


def read_reference_excerpt(reference_paths: list[Path], max_chars: int = 4500) -> str:
    excerpts: list[str] = []
    budget = max_chars
    for path in reference_paths:
        if budget <= 0:
            break
        text = read_text_file(path, max_chars=min(1800, budget))
        excerpts.append(f"{path.name}\n{text}")
        budget -= len(text)
    return "\n\n".join(excerpts).strip() or "未提取到参考网页正文。"


def build_local_report_text(
    report_type: str,
    course_name: str,
    experiment_name: str,
    requirement_text: str,
    reference_urls: list[str],
    reference_notes: str,
    reference_paths: list[Path],
    screenshot_paths: list[Path],
    code_paths: list[Path],
    detail_level: str,
) -> str:
    screenshot_lines = build_screenshot_context(screenshot_paths)
    code_context = build_code_context(code_paths)
    reference_excerpt = read_reference_excerpt(reference_paths)
    references = "\n".join(reference_urls) if reference_urls else "未提供单独参考链接。"
    requirement = requirement_text.strip() or "根据课程实验要求完成环境准备、过程记录、结果验证和问题分析。"
    detail_tail = """
6. 对每一步操作都需要记录“为什么做、做了什么、看到什么结果”。如果某一步结果不符合预期，应回到环境、配置、命令参数和输入数据四个方向进行排查。
7. 在整理报告时，应让截图、代码和文字相互对应：截图证明运行结果，代码说明实现依据，文字解释实验逻辑。
""" if detail_level == "full" else ""

    if report_type == "课程设计报告":
        return f"""课程名称：{course_name}
题目名称：{experiment_name}

一、需求分析

本课程设计围绕“{experiment_name}”展开，目标是在给定课程背景下完成需求理解、方案设计、实现过程记录和结果验证。根据已提供材料，本报告重点整理系统应完成的主要功能、实现时涉及的关键文件、运行截图所体现的结果，以及设计过程中需要注意的问题。

1. 功能需求：系统需要围绕题目完成基础业务流程，保证输入、处理、输出三个环节清晰可追踪。
2. 数据需求：如果项目包含数据库、配置文件或本地数据，应说明数据来源、字段含义、存储位置和读写方式。
3. 运行需求：项目应能在指定环境中正常启动，并通过截图、控制台输出或页面结果证明主要功能已经运行。
4. 文档需求：报告需要把设计思路、实现过程、测试结果和问题总结整理成闭环，便于后续复查。

二、系统设计

1. 总体结构：系统可按输入层、业务处理层、数据存储层和结果展示层理解。输入层负责接收用户操作或测试数据，业务处理层完成核心逻辑，数据存储层负责保存或读取数据，结果展示层通过页面、控制台或文件输出最终结果。
2. 模块划分：根据代码文件和运行材料，可将项目拆分为初始化模块、核心功能模块、辅助工具模块和测试验证模块。每个模块都应在报告中说明作用，避免只罗列代码。
3. 关键流程：系统运行时通常先完成环境加载和参数初始化，再进入核心处理流程，最后输出结果并进行校验。截图应插入到对应流程说明之后。

三、详细实现

实验或设计要求如下：

{requirement}

参考资料如下：

{references}

自动提取的参考内容摘要：

{reference_excerpt}

关键代码或配置摘录：

{code_context}

在实现过程中，需要重点说明核心代码为什么这样组织。例如，入口文件负责启动流程，配置文件保存运行参数，业务函数或类承担数据处理逻辑，测试代码用于验证功能是否符合预期。如果代码中包含异常处理，应说明异常出现时系统如何提示和恢复。

四、测试与运行结果

运行截图材料：

{screenshot_lines}

测试时应按照“环境准备、功能运行、结果观察、问题复核”的顺序进行。对于每张截图，正文应说明它对应的操作步骤和验证结论。例如，启动界面截图用于证明程序能正常运行，功能页面截图用于证明业务流程可用，控制台输出截图用于证明命令执行成功。

五、问题分析

1. 如果运行失败，优先检查依赖是否安装、路径是否正确、配置文件是否可读、端口或数据库服务是否启动。
2. 如果界面或输出结果与预期不一致，应检查输入数据、核心判断逻辑和结果展示代码是否对应。
3. 如果截图信息不足，应在提交前补充关键页面、关键命令输出和测试结果截图。

六、总结

通过本次课程设计整理，完成了从需求理解、结构设计、代码实现到结果验证的完整记录。报告将截图、代码和参考资料放到同一条逻辑链中，能够说明项目是如何运行、如何验证以及后续如何排查问题的。后续完善时，可继续补充更细的测试用例、异常场景和性能或可用性分析，使课程设计文档更加完整。
"""

    return f"""课程名称：{course_name}
实验名称：{experiment_name}

一、实验目的

1. 理解“{experiment_name}”涉及的核心概念、实验流程和结果验证方法。
2. 能够根据实验要求或参考教程完成环境准备、参数配置、代码运行、命令执行和结果记录。
3. 能够把运行截图、代码文件和实验步骤对应起来，说明每一项操作的目的和验证依据。
4. 通过报告整理形成完整实验记录，便于后续复盘、提交和问题排查。

二、实验环境

1. 实验平台：根据实际运行材料整理，可包含 Windows、Linux、虚拟机、浏览器、IDE、数据库、网络模拟器或命令行终端等。
2. 输入材料：页面填写的实验要求、参考链接、上传的模板、运行截图和代码文件。
3. 运行截图：共 {len(screenshot_paths)} 张。
{screenshot_lines}
4. 代码或配置文件：共 {len(code_paths)} 个。报告中记录其文件名并摘录关键文本内容，用于说明实现依据。

三、实验原理或任务要求

本次实验围绕“{experiment_name}”展开。实验要求如下：

{requirement}

参考链接如下：

{references}

自动提取的参考内容摘要：

{reference_excerpt}

整理报告时，应把参考教程作为过程依据，把截图作为结果证据，把代码或配置文件作为实现依据。正文不能只复述教程步骤，还需要说明每一步操作为什么必要、执行后应观察什么现象，以及观察结果如何证明实验目标已经完成。

四、实验步骤

1. 阅读实验要求和参考资料，明确实验目标、所需环境、关键命令或代码文件，以及最终需要验证的结果。
2. 准备实验环境。根据实验类型安装或打开所需软件，例如开发工具、数据库服务、浏览器、虚拟机、网络配置工具或命令行终端。
3. 检查项目或实验目录结构，确认代码文件、配置文件、截图材料和模板文件均可访问。对于代码类实验，应先确认依赖安装、编译或启动命令；对于网络类实验，应先确认主机、IP 地址、连通性和拓扑关系。
4. 按实验要求完成核心配置或代码编写。配置过程中要记录关键参数的含义，代码实现中要说明主要函数、类或命令承担的作用。
5. 运行程序或命令并观察输出。对关键步骤保留截图，例如环境配置完成界面、命令行输出、运行结果页面、测试通过结果或异常提示。
{detail_tail}8. 将上传的截图插入正文对应位置。截图应紧跟相关步骤或结果说明，图注放在图片下方，用简短语句说明该截图对应的实验现象。

五、关键代码或配置

以下为上传代码或配置文件的部分内容摘录，用于说明实验实现依据：

{code_context}

在报告中引用代码时，应重点说明代码作用、主要逻辑、输入输出关系和运行结果，而不是简单堆放完整代码。若代码文件较多，应按入口文件、配置文件、核心逻辑文件和测试文件分类描述。

六、实验结果

1. 已根据实验要求生成完整实验报告正文，并使用模板填充为 DOCX 文档。
2. 上传截图已按顺序插入到实验步骤或实验结果部分，图注用于说明截图对应的运行现象。
3. 上传代码文件已在报告中记录文件名，并摘录部分文本代码作为关键实现说明。
4. 如果实验材料中包含明确运行结果，应在最终提交前人工复核截图中的结果是否与正文描述一致。
5. 对于网络、数据库、Web、Java、安卓等课程实验，结果判断应以实际运行界面、命令输出、测试结果或日志信息为准。
6. 对于本次实验这类连通性验证任务，应重点核对主机地址、网关、子网掩码、目标主机地址和命令返回状态。若截图中显示测试通过，应在正文中说明该截图证明了哪一段配置或哪一个通信链路已经生效。
7. 如果上传了多张截图，应按“环境或拓扑、参数配置、运行验证”的顺序理解。这样能够让报告从准备过程自然过渡到最终结果，避免只有结果截图而缺少过程依据。

七、问题分析

1. 如果生成 DOCX 后缺少图片，通常是上传文件格式不属于常见图片格式，或浏览器上传过程中未成功保存临时文件。
2. 如果模板字段没有完全填充，可能是学校模板中的字段名称与当前映射规则差异较大，需要使用模板诊断脚本进一步适配。
3. 如果代码摘录显示乱码，通常是代码文件编码不是 UTF-8，可先转为 UTF-8 后重新上传。
4. 如果参考链接内容较长，本地 Web UI 会优先提取正文摘要，但最终仍应以截图、代码和实验要求中的事实为准。
5. 对于网络实验，如果连通性测试失败，应先检查 IP 地址是否处于同一网段，再检查防火墙、虚拟网卡、网关配置和目标主机是否在线。排查时应保留关键命令输出，便于在报告中说明问题定位过程。
6. 对于代码类实验，如果运行脚本与截图结果不一致，应检查代码中的目标地址、执行参数和运行环境是否与截图对应，不能只以代码文本推断实验结果。

八、实验总结

通过本次实验报告整理，完成了实验要求、参考资料、运行截图、代码文件和 DOCX 模板之间的整合。报告正文按实验目的、实验环境、任务要求、实验步骤、关键代码、实验结果、问题分析和实验总结组织，能够作为进一步修改和提交的基础版本。后续如需提高质量，应补充更具体的运行数据、错误处理过程、测试用例和结果分析，使报告内容更加贴合真实实验过程。整体来看，实验报告的重点不是简单罗列命令或截图，而是把操作目的、执行过程、观察结果和结论串联起来，使读者能够根据报告复现实验流程并判断结果是否可信。
"""


def build_metadata(
    student_name: str,
    student_id: str,
    class_name: str,
    course_name: str,
    experiment_name: str,
    report_type: str,
) -> dict[str, Any]:
    today = datetime.now().strftime("%Y-%m-%d")
    property_name = "课程设计" if report_type == "课程设计报告" else "验证性实验"
    return {
        "姓名": student_name.strip(),
        "学号": student_id.strip(),
        "班级": class_name.strip(),
        "课程名称": course_name.strip(),
        "实验名称": experiment_name.strip(),
        "题目名称": experiment_name.strip(),
        "实验性质": property_name,
        "日期": today,
        "实验时间": today,
        "Name": student_name.strip(),
        "StudentId": student_id.strip(),
        "ClassName": class_name.strip(),
        "CourseName": course_name.strip(),
        "ExperimentName": experiment_name.strip(),
        "ExperimentProperty": property_name,
        "ExperimentDate": today,
    }


def build_requirements(course_name: str, experiment_name: str, report_type: str, detail_level: str) -> dict[str, Any]:
    min_chars = 2200 if detail_level == "full" else 900
    if report_type == "课程设计报告":
        sections = [
            {"name": "需求分析", "aliases": ["需求分析"], "minChars": 120},
            {"name": "系统设计", "aliases": ["系统设计", "概要设计"], "minChars": 160},
            {"name": "详细实现", "aliases": ["详细实现", "系统实现"], "minChars": 220},
            {"name": "测试与运行结果", "aliases": ["测试与运行结果", "测试结果"], "minChars": 160},
            {"name": "总结", "aliases": ["总结", "课程设计总结"], "minChars": 100},
        ]
    else:
        sections = [
            {"name": "实验目的", "aliases": ["实验目的"], "minChars": 80},
            {"name": "实验环境", "aliases": ["实验环境"], "minChars": 80},
            {"name": "实验原理或任务要求", "aliases": ["实验原理或任务要求", "实验原理", "任务要求"], "minChars": 140},
            {"name": "实验步骤", "aliases": ["实验步骤", "实验过程"], "minChars": 260},
            {"name": "实验结果", "aliases": ["实验结果"], "minChars": 160},
            {"name": "问题分析", "aliases": ["问题分析", "结果分析"], "minChars": 100},
            {"name": "实验总结", "aliases": ["实验总结", "实验小结"], "minChars": 100},
        ]

    return {
        "courseName": course_name,
        "experimentName": experiment_name,
        "minChars": min_chars,
        "sections": sections,
        "requiredKeywords": [],
        "forbiddenPatterns": ["TODO", "待补充", "自行填写", "ChatGPT", "AI生成"],
    }


def build_image_specs(screenshot_paths: list[Path], report_type: str) -> dict[str, Any]:
    images = []
    for index, path in enumerate(screenshot_paths, start=1):
        section = "实验步骤" if index == 1 else "实验结果"
        if report_type == "课程设计报告":
            section = "详细实现" if index == 1 else "测试与运行结果"
        images.append(
            {
                "path": str(path),
                "section": section,
                "caption": f"图{index} {path.stem}",
                "widthCm": 15.8,
            }
        )
    return {"images": images}


def choose_final_docx(summary: dict[str, Any], fallback: Path | None) -> Path:
    for key in ("templateFrameDocxPath", "finalDocxPath"):
        value = summary.get(key)
        if value:
            path = Path(str(value))
            if path.exists():
                return path
    if fallback and fallback.exists():
        return fallback
    raise RuntimeError("生成脚本结束，但没有找到可用的 DOCX 文件。")


def run_smart_pipeline(
    output_dir: Path,
    template_path: Path,
    report_profile: str,
    detail_level: str,
    prompt_text: str,
    reference_paths: list[Path],
    screenshot_paths: list[Path],
    student_name: str,
    student_id: str,
    class_name: str,
    course_name: str,
    experiment_name: str,
) -> tuple[dict[str, Any], Path]:
    powershell = resolve_powershell_executable()
    pipeline_dir = output_dir / "pipeline-smart"
    final_docx_path = output_dir / f"{make_base_name(student_id, student_name, experiment_name)}-智能版.docx"
    command = [
        powershell,
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(REPO_ROOT / "scripts" / "build-report-from-url.ps1"),
        "-TemplatePath",
        str(template_path),
        "-PromptText",
        prompt_text,
        "-CourseName",
        course_name,
        "-ExperimentName",
        experiment_name,
        "-StudentName",
        student_name,
        "-StudentId",
        student_id,
        "-ClassName",
        class_name,
        "-ReportProfileName",
        report_profile,
        "-OutputDir",
        str(pipeline_dir),
        "-FinalDocxPath",
        str(final_docx_path),
        "-PipelineMode",
        "fast",
        "-StyleProfile",
        "excellent",
        "-DetailLevel",
        detail_level,
    ]
    if report_profile == "experiment-report":
        command.append("-CreateTemplateFrameDocx")
    if reference_paths:
        command.append("-ReferenceTextPaths")
        command.extend(str(path) for path in reference_paths)
    if screenshot_paths:
        command.append("-ImagePaths")
        command.extend(str(path) for path in screenshot_paths)

    result = run_command(command, cwd=REPO_ROOT, timeout=SMART_GENERATION_TIMEOUT)
    if result.returncode != 0:
        log_path = write_failure_log(output_dir, command, result.stdout, result.stderr)
        error_text = result.stderr.strip() or result.stdout.strip() or "未知错误"
        raise RuntimeError(f"智能长文生成失败：{error_text[-1800:]}\n日志：{log_path}")

    summary_path = pipeline_dir / "url-build-summary.json"
    if not summary_path.exists():
        raise RuntimeError(f"智能生成完成，但未找到摘要文件：{summary_path}")
    summary = json.loads(summary_path.read_text(encoding="utf-8-sig"))
    return summary, summary_path


def run_local_pipeline(
    output_dir: Path,
    template_path: Path,
    report_profile: str,
    report_type: str,
    detail_level: str,
    course_name: str,
    experiment_name: str,
    student_name: str,
    student_id: str,
    class_name: str,
    requirement_text: str,
    reference_urls: list[str],
    reference_notes: str,
    reference_paths: list[Path],
    screenshot_paths: list[Path],
    code_paths: list[Path],
) -> tuple[dict[str, Any], Path]:
    metadata_path = output_dir / "metadata.json"
    report_path = output_dir / "report.txt"
    requirements_path = output_dir / "requirements.json"
    image_specs_path = output_dir / "image-specs.json"
    pipeline_dir = output_dir / "pipeline-local"
    styled_docx_path = output_dir / f"{make_base_name(student_id, student_name, experiment_name)}-本地草稿.docx"

    report_text = build_local_report_text(
        report_type=report_type,
        course_name=course_name,
        experiment_name=experiment_name,
        requirement_text=requirement_text,
        reference_urls=reference_urls,
        reference_notes=reference_notes,
        reference_paths=reference_paths,
        screenshot_paths=screenshot_paths,
        code_paths=code_paths,
        detail_level=detail_level,
    )
    write_json(metadata_path, build_metadata(student_name, student_id, class_name, course_name, experiment_name, report_type))
    report_path.write_text(report_text, encoding="utf-8-sig")
    write_json(requirements_path, build_requirements(course_name, experiment_name, report_type, detail_level))
    write_json(image_specs_path, build_image_specs(screenshot_paths, report_type))

    powershell = resolve_powershell_executable()
    command = [
        powershell,
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(REPO_ROOT / "scripts" / "build-report.ps1"),
        "-TemplatePath",
        str(template_path),
        "-ReportPath",
        str(report_path),
        "-MetadataPath",
        str(metadata_path),
        "-RequirementsPath",
        str(requirements_path),
        "-OutputDir",
        str(pipeline_dir),
        "-StyledDocxOutPath",
        str(styled_docx_path),
        "-StyleFinalDocx",
        "-ReportProfileName",
        report_profile,
        "-StyleProfile",
        "excellent",
        "-PipelineMode",
        "fast",
    ]
    if report_profile == "experiment-report":
        command.append("-CreateTemplateFrameDocx")
    if screenshot_paths:
        command.extend(["-ImageSpecsPath", str(image_specs_path)])

    result = run_command(command, cwd=REPO_ROOT, timeout=LOCAL_GENERATION_TIMEOUT)
    if result.returncode != 0:
        log_path = write_failure_log(output_dir, command, result.stdout, result.stderr)
        error_text = result.stderr.strip() or result.stdout.strip() or "未知错误"
        raise RuntimeError(f"本地草稿生成失败：{error_text[-2200:]}\n日志：{log_path}")

    summary_path = pipeline_dir / "summary.json"
    if not summary_path.exists():
        raise RuntimeError(f"本地草稿生成完成，但未找到摘要文件：{summary_path}")
    summary = json.loads(summary_path.read_text(encoding="utf-8-sig"))
    return summary, summary_path


def export_docx_to_pdf(docx_path: Path, pdf_path: Path) -> None:
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    powershell = resolve_powershell_executable()
    script = r'''
param(
  [Parameter(Mandatory = $true)]
  [string]$DocxPath,

  [Parameter(Mandatory = $true)]
  [string]$PdfPath
)

$ErrorActionPreference = "Stop"
$docx = [System.IO.Path]::GetFullPath($DocxPath)
$pdf = [System.IO.Path]::GetFullPath($PdfPath)
$parent = Split-Path -Parent $pdf
if (-not [string]::IsNullOrWhiteSpace($parent)) {
  New-Item -ItemType Directory -Path $parent -Force | Out-Null
}
$errors = New-Object System.Collections.Generic.List[string]
foreach ($progId in @("KWPS.Application", "Word.Application")) {
  $app = $null
  $doc = $null
  try {
    $app = New-Object -ComObject $progId
    $app.Visible = $false
    $doc = $app.Documents.Open($docx)
    $doc.ExportAsFixedFormat($pdf, 17)
    $doc.Close($false)
    $doc = $null
    $app.Quit()
    $app = $null
    if (Test-Path -LiteralPath $pdf -PathType Leaf) {
      exit 0
    }
  } catch {
    [void]$errors.Add(("{0}: {1}" -f $progId, $_.Exception.Message))
  } finally {
    if ($null -ne $doc) {
      try { $doc.Close($false) | Out-Null } catch {}
      try { [Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null } catch {}
    }
    if ($null -ne $app) {
      try { $app.Quit() | Out-Null } catch {}
      try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}
$soffice = Get-Command soffice -ErrorAction SilentlyContinue
if ($null -ne $soffice -and -not [string]::IsNullOrWhiteSpace($soffice.Source)) {
  & $soffice.Source --headless --convert-to pdf --outdir $parent $docx | Out-Null
  $converted = Join-Path $parent (([System.IO.Path]::GetFileNameWithoutExtension($docx)) + ".pdf")
  if (Test-Path -LiteralPath $converted -PathType Leaf) {
    if (-not [string]::Equals($converted, $pdf, [System.StringComparison]::OrdinalIgnoreCase)) {
      Move-Item -LiteralPath $converted -Destination $pdf -Force
    }
    exit 0
  }
}
throw ("DOCX 转 PDF 失败：" + ($errors -join " | "))
'''
    temp_script = Path(tempfile.gettempdir()) / f"openclaw-docx-to-pdf-{os.getpid()}-{datetime.now().strftime('%H%M%S%f')}.ps1"
    temp_script.write_text(script, encoding="utf-8-sig")
    try:
        result = run_command(
            [
                powershell,
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(temp_script),
                "-DocxPath",
                str(docx_path),
                "-PdfPath",
                str(pdf_path),
            ],
            cwd=REPO_ROOT,
            timeout=PDF_EXPORT_TIMEOUT,
        )
        if result.returncode != 0 or not pdf_path.exists():
            error_text = result.stderr.strip() or result.stdout.strip() or "未知错误"
            raise RuntimeError(error_text[-1800:])
    finally:
        try:
            temp_script.unlink(missing_ok=True)
        except OSError:
            pass


def render_pdf_preview(pdf_path: Path, preview_path: Path) -> None:
    preview_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        import fitz  # type: ignore
    except ImportError as exc:
        raise RuntimeError("缺少 PyMuPDF，无法渲染 PDF 预览图。请运行 python -m pip install -r requirements-web.txt") from exc

    document = fitz.open(str(pdf_path))
    if document.page_count == 0:
        raise RuntimeError("PDF 没有页面，无法渲染预览图。")

    try:
        from PIL import Image
    except ImportError:
        page = document.load_page(0)
        pixmap = page.get_pixmap(matrix=fitz.Matrix(1.4, 1.4), alpha=False)
        pixmap.save(str(preview_path))
        return

    page_images = []
    for page_index in range(document.page_count):
        page = document.load_page(page_index)
        pixmap = page.get_pixmap(matrix=fitz.Matrix(1.3, 1.3), alpha=False)
        image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
        page_images.append(image)

    gap = 28
    margin = 22
    width = max(image.width for image in page_images) + margin * 2
    height = sum(image.height for image in page_images) + gap * (len(page_images) - 1) + margin * 2
    canvas = Image.new("RGB", (width, height), (238, 238, 238))
    y = margin
    for image in page_images:
        x = (width - image.width) // 2
        canvas.paste(image, (x, y))
        y += image.height + gap
    canvas.save(preview_path)


def format_check_value(value: Any) -> str:
    if value is True:
        return "通过"
    if value is False:
        return "未通过"
    return "未运行"


def copy_final_artifacts(
    output_root_text: str,
    base_name: str,
    source_docx: Path,
    export_pdf: bool,
    render_preview: bool,
    warnings: list[str],
) -> tuple[Path, Path | None, Path | None]:
    output_root = Path(output_root_text.strip()) if output_root_text.strip() else DEFAULT_DELIVERY_ROOT
    if not output_root.exists() and output_root == DEFAULT_DELIVERY_ROOT:
        output_root = WEB_OUTPUT_ROOT / "delivery"

    docx_target = unique_path(output_root / "docx", f"{base_name}.docx")
    shutil.copy2(source_docx, docx_target)

    pdf_target: Path | None = None
    preview_target: Path | None = None

    if export_pdf:
        pdf_target = unique_path(output_root / "pdf", f"{base_name}.pdf")
        try:
            export_docx_to_pdf(docx_target, pdf_target)
        except Exception as exc:
            warnings.append(f"PDF 导出失败：{exc}")
            pdf_target = None

    if render_preview:
        if pdf_target is None:
            warnings.append("未生成 PDF，预览图也无法渲染。")
        else:
            preview_target = unique_path(output_root / "预览图", f"{base_name}.png")
            try:
                render_pdf_preview(pdf_target, preview_target)
            except Exception as exc:
                warnings.append(f"PDF 预览图渲染失败：{exc}")
                preview_target = None

    return docx_target, pdf_target, preview_target


def generate_report(
    report_type: str,
    generation_mode: str,
    detail_level_label: str,
    export_pdf_checked: bool,
    render_preview_checked: bool,
    output_root: str,
    course_name: str,
    student_name: str,
    student_id: str,
    class_name: str,
    experiment_name: str,
    requirement_text: str,
    reference_links: str,
    chat_request_text: str,
    screenshot_path_text: str,
    code_path_text: str,
    template_file: Any,
    screenshot_files: Any,
    code_files: Any,
) -> tuple[str, str, str | None, str | None, str | None, str | None]:
    warnings: list[str] = []
    parsed_request = parse_chat_request(chat_request_text or "")
    report_type = first_nonempty(parsed_request.get("report_type", ""), report_type)
    course_name = first_nonempty(course_name, parsed_request.get("course_name", ""))
    student_name = first_nonempty(student_name, parsed_request.get("student_name", ""))
    student_id = first_nonempty(student_id, parsed_request.get("student_id", ""))
    class_name = first_nonempty(class_name, parsed_request.get("class_name", ""))
    experiment_name = first_nonempty(experiment_name, parsed_request.get("experiment_name", ""))
    screenshot_path_text = "\n".join(
        part for part in [screenshot_path_text.strip(), parsed_request.get("screenshot_path", "")] if part
    )
    code_path_text = "\n".join(
        part for part in [code_path_text.strip(), parsed_request.get("code_path", "")] if part
    )
    chat_reference_urls, _ = parse_reference_input(chat_request_text or "")
    reference_links = "\n".join(part for part in [reference_links.strip(), "\n".join(chat_reference_urls)] if part)
    effective_requirement_text = clean_request_text_for_body(requirement_text.strip())
    if chat_request_text.strip():
        chat_body_text = clean_request_text_for_body(chat_request_text.strip())
        effective_requirement_text = "\n\n".join(
            part for part in [effective_requirement_text, "对话式需求：\n" + chat_body_text if chat_body_text else ""] if part
        )

    required_fields = {
        "课程名称": course_name,
        "学生姓名": student_name,
        "学号": student_id,
        "班级": class_name,
    }
    missing = [label for label, value in required_fields.items() if not str(value or "").strip()]
    if missing:
        return "生成失败", "请先填写：" + "、".join(missing), None, None, None, None

    detail_level = "full" if "长" in detail_level_label or "完整" in detail_level_label else "standard"
    report_profile, _ = resolve_report_profile(report_type)
    reference_urls, reference_notes = parse_reference_input(reference_links)

    try:
        template_path = resolve_template_path(template_file, report_type)
        screenshots = require_allowed_extensions(
            normalize_upload_list(screenshot_files, label="运行截图"),
            IMAGE_EXTENSIONS,
            "运行截图",
        )
        screenshots.extend(collect_local_files(screenshot_path_text, IMAGE_EXTENSIONS, "截图文件夹"))
        code_paths = require_allowed_extensions(
            normalize_upload_list(code_files, label="代码文件"),
            CODE_EXTENSIONS,
            "代码文件",
        )
        code_paths.extend(collect_local_files(code_path_text, CODE_EXTENSIONS, "代码文件夹"))
    except UploadValidationError as exc:
        return "生成失败", str(exc), None, None, None, None

    if not effective_requirement_text.strip():
        warnings.append("未填写实验要求，正文会按题目和材料生成通用结构，具体性会弱一些。")
    if not screenshots:
        warnings.append("未上传运行截图，报告会缺少结果证据。")
    if not code_paths:
        warnings.append("未上传代码文件，关键代码部分会以过程说明为主。")

    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    title_needs_inference = is_auto_title(experiment_name)
    initial_experiment_name = experiment_name if not title_needs_inference else "待推断实验报告"
    base_name = make_base_name(student_id, student_name, initial_experiment_name)
    output_dir = WEB_OUTPUT_ROOT / f"{base_name}-{timestamp}"
    input_dir = output_dir / "inputs"
    output_dir.mkdir(parents=True, exist_ok=True)

    try:
        copied_template = copy_uploads([template_path], input_dir / "template")[0]
        copied_screenshots = copy_uploads(screenshots, input_dir / "screenshots")
        copied_code_files = copy_uploads(code_paths, input_dir / "code")
    except OSError as exc:
        return "生成失败", f"复制上传材料失败：{exc}", None, None, None, None

    reference_dir = input_dir / "references"
    try:
        reference_paths = fetch_reference_texts(reference_urls, reference_dir, warnings)
    except Exception as exc:
        return "生成失败", f"参考链接处理失败：{exc}", None, None, None, None
    reference_paths.extend(write_reference_notes(reference_notes, reference_dir))

    if title_needs_inference:
        inferred_name = infer_experiment_name(reference_paths, effective_requirement_text, reference_notes)
        experiment_name = inferred_name or f"{course_name.strip()}实验报告"
        warnings.append(f"实验名称已自动处理为：{experiment_name}")
    experiment_name = experiment_name.strip()
    base_name = make_base_name(student_id, student_name, experiment_name)

    prompt_text = build_prompt_text(
        report_type=report_type,
        course_name=course_name.strip(),
        experiment_name=experiment_name.strip(),
        requirement_text=effective_requirement_text,
        reference_urls=reference_urls,
        reference_notes=reference_notes,
        screenshot_paths=copied_screenshots,
        code_paths=copied_code_files,
        detail_level=detail_level,
    )
    (output_dir / "prompt-for-smart-generation.txt").write_text(prompt_text, encoding="utf-8-sig")

    summary: dict[str, Any]
    summary_path: Path
    generation_used = "快速本地草稿"

    if generation_mode.startswith("智能"):
        try:
            summary, summary_path = run_smart_pipeline(
                output_dir=output_dir,
                template_path=copied_template,
                report_profile=report_profile,
                detail_level=detail_level,
                prompt_text=prompt_text,
                reference_paths=reference_paths,
                screenshot_paths=copied_screenshots,
                student_name=student_name.strip(),
                student_id=student_id.strip(),
                class_name=class_name.strip(),
                course_name=course_name.strip(),
                experiment_name=experiment_name.strip(),
            )
            generation_used = "智能长文"
        except Exception as exc:
            warnings.append(str(exc))
            summary, summary_path = run_local_pipeline(
                output_dir=output_dir,
                template_path=copied_template,
                report_profile=report_profile,
                report_type=report_type,
                detail_level=detail_level,
                course_name=course_name.strip(),
                experiment_name=experiment_name.strip(),
                student_name=student_name.strip(),
                student_id=student_id.strip(),
                class_name=class_name.strip(),
                requirement_text=effective_requirement_text,
                reference_urls=reference_urls,
                reference_notes=reference_notes,
                reference_paths=reference_paths,
                screenshot_paths=copied_screenshots,
                code_paths=copied_code_files,
            )
            generation_used = "智能失败后回退本地草稿"
    else:
        try:
            summary, summary_path = run_local_pipeline(
                output_dir=output_dir,
                template_path=copied_template,
                report_profile=report_profile,
                report_type=report_type,
                detail_level=detail_level,
                course_name=course_name.strip(),
                experiment_name=experiment_name.strip(),
                student_name=student_name.strip(),
                student_id=student_id.strip(),
                class_name=class_name.strip(),
                requirement_text=effective_requirement_text,
                reference_urls=reference_urls,
                reference_notes=reference_notes,
                reference_paths=reference_paths,
                screenshot_paths=copied_screenshots,
                code_paths=copied_code_files,
            )
        except Exception as exc:
            return "生成失败", str(exc), None, None, None, None

    try:
        source_docx = choose_final_docx(summary, None)
    except Exception as exc:
        return "生成失败", str(exc), None, None, None, None

    final_docx, final_pdf, final_preview = copy_final_artifacts(
        output_root_text=output_root,
        base_name=f"{base_name}-{timestamp}",
        source_docx=source_docx,
        export_pdf=bool(export_pdf_checked),
        render_preview=bool(render_preview_checked),
        warnings=warnings,
    )

    layout_message = str(summary.get("layoutCheckMessage") or "").strip()
    if layout_message:
        warnings.append(layout_message)
    if summary.get("layoutCheckPassed") is False:
        warnings.append("版式检查未通过，请打开 DOCX/PDF 复核外框、图注和分页。")
    if summary.get("validationPassed") is False:
        warnings.append("正文校验未通过，请查看输出目录中的 validation.json。")

    status_lines = [
        "生成成功",
        f"生成方式：{generation_used}",
        f"工作目录：{output_dir}",
        f"摘要文件：{summary_path}",
        f"DOCX：{final_docx}",
        f"PDF：{final_pdf if final_pdf else '未生成'}",
        f"预览图：{final_preview if final_preview else '未生成'}",
        f"版式检查：{format_check_value(summary.get('layoutCheckPassed'))}",
        f"正文校验：{format_check_value(summary.get('validationPassed'))}",
        f"截图数量：{len(copied_screenshots)}",
        f"代码文件数量：{len(copied_code_files)}",
    ]
    warning_text = "\n".join(dict.fromkeys(line for line in warnings if line.strip()))
    preview_value = str(final_preview) if final_preview else None
    return "\n".join(status_lines), warning_text, str(final_docx), str(final_pdf) if final_pdf else None, preview_value, preview_value


def create_app() -> gr.Blocks:
    with gr.Blocks(title="实验报告生成 Web UI") as app:
        gr.Markdown("# 实验报告生成 Web UI")

        with gr.Row():
            report_type = gr.Radio(
                label="报告类型",
                choices=["实验报告", "课程设计报告"],
                value="实验报告",
            )
            generation_mode = gr.Radio(
                label="生成方式",
                choices=["智能长文（接近对话效果）", "快速本地草稿"],
                value="智能长文（接近对话效果）",
            )
            detail_level = gr.Radio(
                label="正文长度",
                choices=["长正文（推荐）", "标准正文"],
                value="长正文（推荐）",
            )

        with gr.Row():
            course_name = gr.Textbox(label="课程名称", placeholder="计算机网络")
            experiment_name = gr.Textbox(label="实验名称/题目名称", placeholder="根据教程链接填充")

        with gr.Row():
            student_name = gr.Textbox(label="学生姓名", placeholder="李亦非")
            student_id = gr.Textbox(label="学号", placeholder="2444100198")
            class_name = gr.Textbox(label="班级", placeholder="24C")

        requirement_text = gr.Textbox(label="实验要求", lines=6, placeholder="填写实验任务、步骤要求、验收标准等")
        reference_links = gr.Textbox(label="参考链接或补充说明", lines=3, placeholder="每行一个链接，也可以写补充说明")
        chat_request_text = gr.Textbox(
            label="对话式需求（可直接粘贴你平时发给我的整段要求）",
            lines=6,
            placeholder="例如：CSDN链接、课程名称、实验名称、姓名、学号、班级、截图材料路径等",
        )

        with gr.Row():
            screenshot_path_text = gr.Textbox(
                label="本地截图文件夹/文件路径",
                placeholder=r'E:\实验报告\截图\计网实验六',
            )
            code_path_text = gr.Textbox(
                label="本地代码文件夹/文件路径",
                placeholder=r'E:\某项目\src',
            )

        with gr.Row():
            template_file = gr.File(label="上传 docx/doc 模板（不传则使用默认模板）", file_types=[".docx", ".doc"])
            screenshot_files = gr.File(label="上传运行截图", file_count="multiple", file_types=sorted(IMAGE_EXTENSIONS))
            code_files = gr.File(label="上传代码文件", file_count="multiple")

        with gr.Row():
            output_root = gr.Textbox(label="输出根目录", value=str(DEFAULT_DELIVERY_ROOT), scale=2)
            export_pdf_checked = gr.Checkbox(label="导出 PDF", value=True)
            render_preview_checked = gr.Checkbox(label="生成预览图", value=True)

        generate_button = gr.Button("生成报告", variant="primary")

        status_output = gr.Textbox(label="生成状态", lines=10)
        error_output = gr.Textbox(label="警告/错误信息", lines=10)
        with gr.Row():
            docx_output = gr.File(label="下载 DOCX")
            pdf_output = gr.File(label="下载 PDF")
            preview_output = gr.File(label="下载预览图")
        preview_image = gr.Image(label="预览图", type="filepath")

        generate_button.click(
            fn=generate_report,
            inputs=[
                report_type,
                generation_mode,
                detail_level,
                export_pdf_checked,
                render_preview_checked,
                output_root,
                course_name,
                student_name,
                student_id,
                class_name,
                experiment_name,
                requirement_text,
                reference_links,
                chat_request_text,
                screenshot_path_text,
                code_path_text,
                template_file,
                screenshot_files,
                code_files,
            ],
            outputs=[status_output, error_output, docx_output, pdf_output, preview_output, preview_image],
        )

    return app


if __name__ == "__main__":
    create_app().launch(server_name="127.0.0.1", server_port=7860)
