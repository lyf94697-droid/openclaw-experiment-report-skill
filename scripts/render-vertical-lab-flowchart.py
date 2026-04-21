from __future__ import annotations

import argparse
import math
from pathlib import Path
from typing import Iterable

from PIL import Image, ImageDraw, ImageFont


def load_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = [
        Path(r"C:\Windows\Fonts\simsun.ttc"),
        Path(r"C:\Windows\Fonts\simhei.ttf"),
        Path(r"C:\Windows\Fonts\msyh.ttc"),
    ]
    for path in candidates:
        if path.exists():
            return ImageFont.truetype(str(path), size)
    return ImageFont.load_default()


def wrap_text(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, max_width: int) -> list[str]:
    if not text:
        return [""]

    if " " in text:
        words = text.split()
        lines: list[str] = []
        current = ""
        for word in words:
            candidate = (current + " " + word).strip()
            if not current or draw.textlength(candidate, font=font) <= max_width:
                current = candidate
            else:
                lines.append(current)
                current = word
        if current:
            lines.append(current)
        return lines

    lines: list[str] = []
    current = ""
    for char in text:
        candidate = current + char
        if not current or draw.textlength(candidate, font=font) <= max_width:
            current = candidate
        else:
            lines.append(current)
            current = char
    if current:
        lines.append(current)
    return lines


def draw_centered_text(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    text: str,
    font: ImageFont.ImageFont,
    fill: tuple[int, int, int],
    line_gap: int = 10,
) -> None:
    max_width = box[2] - box[0] - 56
    lines: list[str] = []
    for raw_line in text.splitlines() or [""]:
        lines.extend(wrap_text(draw, raw_line, font, max_width))

    bboxes = [draw.textbbox((0, 0), line, font=font) for line in lines]
    heights = [bbox[3] - bbox[1] for bbox in bboxes]
    total_height = sum(heights) + line_gap * max(0, len(lines) - 1)
    y = box[1] + (box[3] - box[1] - total_height) / 2

    for line, bbox, height in zip(lines, bboxes, heights):
        width = bbox[2] - bbox[0]
        x = box[0] + (box[2] - box[0] - width) / 2
        draw.text((x, y), line, font=font, fill=fill)
        y += height + line_gap


def draw_arrow(
    draw: ImageDraw.ImageDraw,
    start: tuple[int, int],
    end: tuple[int, int],
    color: tuple[int, int, int],
) -> None:
    draw.line((start, end), fill=color, width=3)
    draw_arrow_head(draw, start, end, color)


def draw_polyline_arrow(
    draw: ImageDraw.ImageDraw,
    points: list[tuple[int, int]],
    color: tuple[int, int, int],
) -> None:
    if len(points) < 2:
        return
    draw.line(points, fill=color, width=3)
    draw_arrow_head(draw, points[-2], points[-1], color)


def draw_arrow_head(
    draw: ImageDraw.ImageDraw,
    start: tuple[int, int],
    end: tuple[int, int],
    color: tuple[int, int, int],
) -> None:
    angle = math.atan2(end[1] - start[1], end[0] - start[0])
    size = 14
    wing = math.radians(26)
    p1 = (
        end[0] - size * math.cos(angle - wing),
        end[1] - size * math.sin(angle - wing),
    )
    p2 = (
        end[0] - size * math.cos(angle + wing),
        end[1] - size * math.sin(angle + wing),
    )
    draw.polygon([end, p1, p2], fill=color)


def draw_badge(
    draw: ImageDraw.ImageDraw,
    center: tuple[int, int],
    text: str,
    font: ImageFont.ImageFont,
    fill: tuple[int, int, int],
    ink: tuple[int, int, int],
) -> None:
    radius = 28
    x, y = center
    draw.ellipse((x - radius, y - radius, x + radius, y + radius), fill=fill, outline=ink, width=3)
    bbox = draw.textbbox((0, 0), text, font=font)
    draw.text((x - (bbox[2] - bbox[0]) / 2, y - (bbox[3] - bbox[1]) / 2 - 2), text, font=font, fill=ink)


def draw_process(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    text: str,
    font: ImageFont.ImageFont,
    ink: tuple[int, int, int],
) -> None:
    draw.rectangle(box, fill=(255, 255, 255), outline=ink, width=2)
    draw_centered_text(draw, box, text, font, ink, line_gap=5)


def draw_terminator(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    text: str,
    font: ImageFont.ImageFont,
    ink: tuple[int, int, int],
) -> None:
    draw.rounded_rectangle(box, radius=10, fill=(255, 255, 255), outline=ink, width=2)
    draw_centered_text(draw, box, text, font, ink, line_gap=5)


def draw_decision(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    text: str,
    font: ImageFont.ImageFont,
    ink: tuple[int, int, int],
) -> None:
    x1, y1, x2, y2 = box
    points = [
        ((x1 + x2) // 2, y1),
        (x2, (y1 + y2) // 2),
        ((x1 + x2) // 2, y2),
        (x1, (y1 + y2) // 2),
    ]
    draw.polygon(points, fill=(255, 255, 255), outline=ink)
    draw.line(points + [points[0]], fill=ink, width=2)
    draw_centered_text(draw, box, text, font, ink, line_gap=5)


def draw_connector_label(
    draw: ImageDraw.ImageDraw,
    xy: tuple[int, int],
    text: str,
    font: ImageFont.ImageFont,
    ink: tuple[int, int, int],
) -> None:
    bbox = draw.textbbox((0, 0), text, font=font)
    pad = 5
    rect = (xy[0] - pad, xy[1] - pad, xy[0] + (bbox[2] - bbox[0]) + pad, xy[1] + (bbox[3] - bbox[1]) + pad)
    draw.rectangle(rect, fill=(255, 255, 255))
    draw.text(xy, text, font=font, fill=ink)


def crop_to_ink(img: Image.Image, pad: int = 28) -> Image.Image:
    mask = Image.eval(img.convert("L"), lambda px: 255 if px < 250 else 0)
    bbox = mask.getbbox()
    if bbox is None:
        return img

    left = max(0, bbox[0] - pad)
    top = max(0, bbox[1] - pad)
    right = min(img.width, bbox[2] + pad)
    bottom = min(img.height, bbox[3] + pad)
    return img.crop((left, top, right, bottom))


def render_branched_flowchart(title: str, steps: list[str], out_path: Path) -> None:
    width = 1300
    height = 1560
    img = Image.new("RGB", (width, height), (255, 255, 255))
    draw = ImageDraw.Draw(img)
    ink = (20, 20, 20)

    title_font = load_font(28)
    node_font = load_font(24)
    label_font = load_font(19)

    center_x = width // 2
    top = 38
    if title.strip():
        title_width = draw.textlength(title, font=title_font)
        draw.text(((width - title_width) / 2, 16), title, font=title_font, fill=ink)
        top = 70

    node_w = 420
    node_h = 58
    term_w = 120
    term_h = 42
    gap = 36
    decision_w = 390
    decision_h = 122

    decision_index = next((i for i, step in enumerate(steps) if "是否" in step or "?" in step or "？" in step), -1)
    if decision_index < 1 or len(steps) < decision_index + 3:
        render_linear_flowchart(title, steps, out_path)
        return

    def centered_box(y: int, w: int = node_w, h: int = node_h) -> tuple[int, int, int, int]:
        return (center_x - w // 2, y, center_x + w // 2, y + h)

    y = top
    start_box = centered_box(y, term_w, term_h)
    draw_terminator(draw, start_box, "开始", node_font, ink)
    previous_bottom = ((start_box[0] + start_box[2]) // 2, start_box[3])
    y = start_box[3] + gap

    search_box: tuple[int, int, int, int] | None = None
    for idx, step in enumerate(steps[:decision_index]):
        box = centered_box(y)
        draw_arrow(draw, previous_bottom, (center_x, box[1]), ink)
        draw_process(draw, box, step, node_font, ink)
        if idx == decision_index - 1:
            search_box = box
        previous_bottom = (center_x, box[3])
        y = box[3] + gap

    decision_box = centered_box(y, decision_w, decision_h)
    draw_arrow(draw, previous_bottom, (center_x, decision_box[1]), ink)
    draw_decision(draw, decision_box, steps[decision_index], node_font, ink)
    decision_center_y = (decision_box[1] + decision_box[3]) // 2

    no_box = (900, decision_center_y - 29, 1215, decision_center_y + 29)
    draw_polyline_arrow(draw, [(decision_box[2], decision_center_y), (no_box[0], decision_center_y)], ink)
    draw_connector_label(draw, (decision_box[2] + 24, decision_center_y - 34), "否", label_font, ink)
    draw_process(draw, no_box, "提示无结果", node_font, ink)

    retry_box = (900, decision_center_y + 88, 1215, decision_center_y + 146)
    draw_arrow(draw, ((no_box[0] + no_box[2]) // 2, no_box[3]), ((retry_box[0] + retry_box[2]) // 2, retry_box[1]), ink)
    draw_process(draw, retry_box, "调整关键词重新查询", node_font, ink)
    if search_box is not None:
        draw_polyline_arrow(
            draw,
            [
                (retry_box[2], (retry_box[1] + retry_box[3]) // 2),
                (1240, (retry_box[1] + retry_box[3]) // 2),
                (1240, (search_box[1] + search_box[3]) // 2),
                (search_box[2], (search_box[1] + search_box[3]) // 2),
            ],
            ink,
        )

    y = decision_box[3] + gap
    draw_connector_label(draw, (center_x + 18, decision_box[3] + 6), "是", label_font, ink)
    after = steps[decision_index + 1 :]
    end_text = after[-1]
    body = after[:-1]

    first_after = body[0] if body else "执行下一步操作"
    first_box = centered_box(y)
    draw_arrow(draw, (center_x, decision_box[3]), (center_x, first_box[1]), ink)
    draw_process(draw, first_box, first_after, node_font, ink)
    y = first_box[3] + 56

    remaining = body[1:]
    merge_start = (center_x, first_box[3])
    if len(remaining) >= 2:
        left_box = (70, y, 490, y + node_h)
        right_box = (760, y, 1180, y + node_h)
        split_y = y - 24
        draw_polyline_arrow(draw, [merge_start, (center_x, split_y), ((left_box[0] + left_box[2]) // 2, split_y), ((left_box[0] + left_box[2]) // 2, left_box[1])], ink)
        draw_polyline_arrow(draw, [(center_x, split_y), ((right_box[0] + right_box[2]) // 2, split_y), ((right_box[0] + right_box[2]) // 2, right_box[1])], ink)
        draw_process(draw, left_box, remaining[0], node_font, ink)
        draw_process(draw, right_box, remaining[1], node_font, ink)
        merge_y = left_box[3] + 52
        join_left = (center_x - 52, merge_y)
        join_right = (center_x + 52, merge_y)
        draw_polyline_arrow(draw, [((left_box[0] + left_box[2]) // 2, left_box[3]), ((left_box[0] + left_box[2]) // 2, merge_y), join_left], ink)
        draw_polyline_arrow(draw, [((right_box[0] + right_box[2]) // 2, right_box[3]), ((right_box[0] + right_box[2]) // 2, merge_y), join_right], ink)
        draw.line((join_left, join_right), fill=ink, width=3)
        draw_arrow(draw, (center_x, merge_y), (center_x, merge_y + 26), ink)
        y = merge_y + 34
        previous_bottom = (center_x, merge_y + 26)
        for step in remaining[2:]:
            box = centered_box(y)
            draw_arrow(draw, previous_bottom, (center_x, box[1]), ink)
            draw_process(draw, box, step, node_font, ink)
            previous_bottom = (center_x, box[3])
            y = box[3] + gap
    else:
        previous_bottom = (center_x, first_box[3])
        for step in remaining:
            box = centered_box(y)
            draw_arrow(draw, previous_bottom, (center_x, box[1]), ink)
            draw_process(draw, box, step, node_font, ink)
            previous_bottom = (center_x, box[3])
            y = box[3] + gap

    end_box = centered_box(y, term_w, term_h)
    draw_arrow(draw, previous_bottom, (center_x, end_box[1]), ink)
    draw_terminator(draw, end_box, "结束", node_font, ink)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    crop_to_ink(img.crop((0, 0, width, min(height, end_box[3] + 46)))).save(out_path)


def render_linear_flowchart(title: str, steps: list[str], out_path: Path) -> None:
    width = 1100
    node_w = 420
    node_h = 58
    gap = 36
    top = 68 if title.strip() else 30
    height = top + 58 + len(steps) * (node_h + gap) + 100

    img = Image.new("RGB", (width, height), (255, 255, 255))
    draw = ImageDraw.Draw(img)
    ink = (20, 20, 20)
    title_font = load_font(28)
    node_font = load_font(24)
    center_x = width // 2

    if title.strip():
        title_width = draw.textlength(title, font=title_font)
        draw.text(((width - title_width) / 2, 14), title, font=title_font, fill=ink)

    y = top
    start_box = (center_x - 90, y, center_x + 90, y + 56)
    draw_terminator(draw, start_box, "开始", node_font, ink)
    previous_bottom = (center_x, start_box[3])
    y = start_box[3] + gap
    for step in steps:
        box = (center_x - node_w // 2, y, center_x + node_w // 2, y + node_h)
        draw_arrow(draw, previous_bottom, (center_x, box[1]), ink)
        draw_process(draw, box, step, node_font, ink)
        previous_bottom = (center_x, box[3])
        y = box[3] + gap
    end_box = (center_x - 90, y, center_x + 90, y + 56)
    draw_arrow(draw, previous_bottom, (center_x, end_box[1]), ink)
    draw_terminator(draw, end_box, "结束", node_font, ink)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    crop_to_ink(img.crop((0, 0, width, min(height, end_box[3] + 46)))).save(out_path)


def render_flowchart(title: str, steps: Iterable[str], out_path: Path) -> None:
    steps = list(steps)
    if len(steps) < 2:
        raise ValueError("At least two flowchart steps are required.")

    render_branched_flowchart(title, steps, out_path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Render a portrait vertical lab flowchart PNG.")
    parser.add_argument("--out", required=True, help="Output PNG path.")
    parser.add_argument("--title", default="", help="Optional flowchart title.")
    parser.add_argument(
        "--steps-file",
        required=True,
        help="UTF-8 text file containing one step per line. Blank lines are ignored.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    steps_path = Path(args.steps_file)
    steps = [line.strip() for line in steps_path.read_text(encoding="utf-8").splitlines() if line.strip()]
    render_flowchart(args.title, steps, Path(args.out))


if __name__ == "__main__":
    main()
