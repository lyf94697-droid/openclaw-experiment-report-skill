from __future__ import annotations

import argparse
import math
import re
from pathlib import Path
from typing import Iterable

from PIL import Image, ImageDraw, ImageFont


INK = (18, 18, 18)
WHITE = (255, 255, 255)


def load_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = [
        Path(r"C:\Windows\Fonts\simhei.ttf"),
        Path(r"C:\Windows\Fonts\simsun.ttc"),
        Path(r"C:\Windows\Fonts\msyh.ttc"),
    ]
    for path in candidates:
        if path.exists():
            return ImageFont.truetype(str(path), size)
    return ImageFont.load_default()


def normalize_step_text(text: str) -> str:
    compact = (text or "").strip()
    compact = re.sub(r"^\s*[\(（]?[0-9一二三四五六七八九十]+[\)）.、]\s*", "", compact)
    compact = compact.replace("：", " ").replace(":", " ")
    compact = re.sub(r"\s+", " ", compact)
    return compact.strip()


def is_decision_step(text: str) -> bool:
    compact = normalize_step_text(text)
    return ("是否" in compact) or compact.endswith("?") or compact.endswith("？")


def wrap_text(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, max_width: int) -> list[str]:
    if not text:
        return [""]

    lines: list[str] = []
    current = ""
    for char in text:
        candidate = current + char
        if not current or draw.textlength(candidate, font=font) <= max_width:
            current = candidate
            continue
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
    fill: tuple[int, int, int] = INK,
    line_gap: int = 10,
    padding_x: int = 28,
) -> None:
    max_width = max(60, box[2] - box[0] - padding_x * 2)
    lines: list[str] = []
    for raw_line in text.splitlines() or [""]:
        lines.extend(wrap_text(draw, raw_line, font, max_width))

    boxes = [draw.textbbox((0, 0), line, font=font) for line in lines]
    heights = [bbox[3] - bbox[1] for bbox in boxes]
    total_height = sum(heights) + line_gap * max(0, len(lines) - 1)
    y = box[1] + (box[3] - box[1] - total_height) / 2

    for line, bbox, height in zip(lines, boxes, heights):
        width = bbox[2] - bbox[0]
        x = box[0] + (box[2] - box[0] - width) / 2
        draw.text((x, y), line, font=font, fill=fill)
        y += height + line_gap


def draw_title(draw: ImageDraw.ImageDraw, width: int, title: str, font: ImageFont.ImageFont) -> int:
    if not title.strip():
        return 48

    title = title.strip()
    top = 22
    baseline = 74
    text_box = draw.textbbox((0, 0), title, font=font)
    text_width = text_box[2] - text_box[0]
    text_height = text_box[3] - text_box[1]
    left = (width - text_width) / 2
    right = left + text_width

    draw.line((96, baseline, left - 32, baseline), fill=INK, width=3)
    draw.line((right + 32, baseline, width - 96, baseline), fill=INK, width=3)
    draw.text((left, baseline - text_height + 8), title, font=font, fill=INK)
    return 122


def draw_arrow_head(
    draw: ImageDraw.ImageDraw,
    start: tuple[int, int],
    end: tuple[int, int],
    color: tuple[int, int, int] = INK,
    width: int = 4,
) -> None:
    angle = math.atan2(end[1] - start[1], end[0] - start[0])
    size = 16
    wing = math.radians(26)
    p1 = (
        end[0] - size * math.cos(angle - wing),
        end[1] - size * math.sin(angle - wing),
    )
    p2 = (
        end[0] - size * math.cos(angle + wing),
        end[1] - size * math.sin(angle + wing),
    )
    draw.line((end, p1), fill=color, width=width)
    draw.line((end, p2), fill=color, width=width)


def draw_arrow(
    draw: ImageDraw.ImageDraw,
    start: tuple[int, int],
    end: tuple[int, int],
    color: tuple[int, int, int] = INK,
    width: int = 4,
) -> None:
    draw.line((start, end), fill=color, width=width)
    draw_arrow_head(draw, start, end, color=color, width=width)


def draw_polyline_arrow(
    draw: ImageDraw.ImageDraw,
    points: list[tuple[int, int]],
    color: tuple[int, int, int] = INK,
    width: int = 4,
) -> None:
    if len(points) < 2:
        return
    draw.line(points, fill=color, width=width)
    draw_arrow_head(draw, points[-2], points[-1], color=color, width=width)


def draw_badge(draw: ImageDraw.ImageDraw, center: tuple[int, int], text: str, font: ImageFont.ImageFont) -> None:
    radius = 26
    x, y = center
    draw.ellipse((x - radius, y - radius, x + radius, y + radius), outline=INK, width=3, fill=WHITE)
    bbox = draw.textbbox((0, 0), text, font=font)
    draw.text((x - (bbox[2] - bbox[0]) / 2, y - (bbox[3] - bbox[1]) / 2 - 2), text, font=font, fill=INK)


def draw_process(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    text: str,
    font: ImageFont.ImageFont,
) -> None:
    draw.rectangle(box, outline=INK, width=3, fill=WHITE)
    draw_centered_text(draw, box, text, font, line_gap=6)


def draw_terminator(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    text: str,
    font: ImageFont.ImageFont,
) -> None:
    draw.rounded_rectangle(box, radius=18, outline=INK, width=3, fill=WHITE)
    draw_centered_text(draw, box, text, font, line_gap=6)


def draw_decision(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    text: str,
    font: ImageFont.ImageFont,
) -> None:
    x1, y1, x2, y2 = box
    points = [
        ((x1 + x2) // 2, y1),
        (x2, (y1 + y2) // 2),
        ((x1 + x2) // 2, y2),
        (x1, (y1 + y2) // 2),
    ]
    draw.polygon(points, outline=INK, fill=WHITE)
    draw.line(points + [points[0]], fill=INK, width=3)
    draw_centered_text(draw, box, text, font, line_gap=6, padding_x=42)


def draw_connector_label(
    draw: ImageDraw.ImageDraw,
    xy: tuple[int, int],
    text: str,
    font: ImageFont.ImageFont,
) -> None:
    bbox = draw.textbbox((0, 0), text, font=font)
    pad_x = 10
    pad_y = 6
    rect = (
        xy[0] - pad_x,
        xy[1] - pad_y,
        xy[0] + (bbox[2] - bbox[0]) + pad_x,
        xy[1] + (bbox[3] - bbox[1]) + pad_y,
    )
    draw.rectangle(rect, outline=INK, fill=WHITE, width=2)
    draw.text(xy, text, font=font, fill=INK)


def crop_to_ink(img: Image.Image, pad: int = 34) -> Image.Image:
    mask = Image.eval(img.convert("L"), lambda px: 255 if px < 250 else 0)
    bbox = mask.getbbox()
    if bbox is None:
        return img

    left = max(0, bbox[0] - pad)
    top = max(0, bbox[1] - pad)
    right = min(img.width, bbox[2] + pad)
    bottom = min(img.height, bbox[3] + pad)
    return img.crop((left, top, right, bottom))


def render_linear_flowchart(title: str, steps: list[str], out_path: Path) -> None:
    width = 1360
    top_padding = 48
    node_w = 560
    node_h = 84
    term_w = 240
    term_h = 72
    gap = 62
    side_gap = 92

    title_font = load_font(30)
    node_font = load_font(26)
    badge_font = load_font(20)

    body_top = draw_title(ImageDraw.Draw(Image.new("RGB", (width, 160), WHITE)), width, title, title_font)
    height = body_top + term_h * 2 + len(steps) * (node_h + gap) + 140
    img = Image.new("RGB", (width, height), WHITE)
    draw = ImageDraw.Draw(img)
    top = draw_title(draw, width, title, title_font) if title.strip() else top_padding

    center_x = width // 2
    start_box = (center_x - term_w // 2, top, center_x + term_w // 2, top + term_h)
    draw_terminator(draw, start_box, "开始", node_font)

    previous_bottom = (center_x, start_box[3])
    y = start_box[3] + gap
    for index, raw_step in enumerate(steps, start=1):
        text = normalize_step_text(raw_step)
        box = (center_x - node_w // 2, y, center_x + node_w // 2, y + node_h)
        draw_arrow(draw, previous_bottom, (center_x, box[1]))
        draw_process(draw, box, text, node_font)
        draw_badge(draw, (box[0] - side_gap, (box[1] + box[3]) // 2), str(index), badge_font)
        previous_bottom = (center_x, box[3])
        y = box[3] + gap

    end_box = (center_x - term_w // 2, y, center_x + term_w // 2, y + term_h)
    draw_arrow(draw, previous_bottom, (center_x, end_box[1]))
    draw_terminator(draw, end_box, "结束", node_font)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    crop_to_ink(img).save(out_path)


def render_branched_flowchart(title: str, steps: list[str], out_path: Path) -> None:
    decision_index = next((index for index, step in enumerate(steps) if is_decision_step(step)), -1)
    if decision_index < 1 or len(steps) < decision_index + 3:
        render_linear_flowchart(title, steps, out_path)
        return

    pre_steps = [normalize_step_text(step) for step in steps[:decision_index]]
    decision_text = normalize_step_text(steps[decision_index])
    right_steps = [normalize_step_text(step) for step in steps[decision_index + 1 : -1]]
    result_text = normalize_step_text(steps[-1])

    width = 1780
    top_padding = 48
    left_center_x = 620
    right_box_left = 1120
    node_w = 520
    node_h = 84
    term_w = 240
    term_h = 72
    decision_w = 420
    decision_h = 168
    gap = 64
    badge_font = load_font(20)
    title_font = load_font(30)
    node_font = load_font(26)
    label_font = load_font(20)

    preview = ImageDraw.Draw(Image.new("RGB", (width, 2200), WHITE))
    top = draw_title(preview, width, title, title_font) if title.strip() else top_padding
    height = top + term_h * 2 + (len(pre_steps) + len(right_steps) + 2) * (node_h + gap) + 260
    img = Image.new("RGB", (width, height), WHITE)
    draw = ImageDraw.Draw(img)
    top = draw_title(draw, width, title, title_font) if title.strip() else top_padding

    def left_box(y: int, w: int = node_w, h: int = node_h) -> tuple[int, int, int, int]:
        return (left_center_x - w // 2, y, left_center_x + w // 2, y + h)

    def right_box(y: int) -> tuple[int, int, int, int]:
        return (right_box_left, y, right_box_left + node_w, y + node_h)

    start_box = left_box(top, term_w, term_h)
    draw_terminator(draw, start_box, "开始", node_font)
    previous_bottom = (left_center_x, start_box[3])
    y = start_box[3] + gap

    step_no = 1
    loop_target_box: tuple[int, int, int, int] | None = None
    for step in pre_steps:
        box = left_box(y)
        draw_arrow(draw, previous_bottom, (left_center_x, box[1]))
        draw_process(draw, box, step, node_font)
        draw_badge(draw, (box[0] - 88, (box[1] + box[3]) // 2), str(step_no), badge_font)
        step_no += 1
        previous_bottom = (left_center_x, box[3])
        loop_target_box = box
        y = box[3] + gap

    decision_box = left_box(y, decision_w, decision_h)
    draw_arrow(draw, previous_bottom, (left_center_x, decision_box[1]))
    draw_decision(draw, decision_box, decision_text, node_font)
    decision_center_y = (decision_box[1] + decision_box[3]) // 2

    result_box = left_box(decision_box[3] + gap)
    draw_connector_label(draw, (left_center_x + 22, decision_box[3] + 10), "是", label_font)
    draw_arrow(draw, (left_center_x, decision_box[3]), (left_center_x, result_box[1]))
    draw_process(draw, result_box, result_text, node_font)
    draw_badge(draw, (result_box[0] - 88, (result_box[1] + result_box[3]) // 2), str(step_no), badge_font)
    step_no += 1

    branch_start = (decision_box[2], decision_center_y)
    branch_y = decision_center_y - node_h // 2
    previous_right_bottom: tuple[int, int] | None = None
    right_boxes: list[tuple[int, int, int, int]] = []
    for idx, step in enumerate(right_steps, start=1):
        box = right_box(branch_y if idx == 1 else branch_y + (idx - 1) * (node_h + gap))
        right_boxes.append(box)
        if idx == 1:
            draw_connector_label(draw, (decision_box[2] + 18, decision_center_y - 38), "否", label_font)
            draw_polyline_arrow(draw, [branch_start, (box[0] - 42, decision_center_y), (box[0], (box[1] + box[3]) // 2)])
        else:
            assert previous_right_bottom is not None
            draw_arrow(draw, previous_right_bottom, ((box[0] + box[2]) // 2, box[1]))
        draw_process(draw, box, step, node_font)
        draw_badge(draw, (box[0] - 88, (box[1] + box[3]) // 2), str(step_no), badge_font)
        step_no += 1
        previous_right_bottom = ((box[0] + box[2]) // 2, box[3])

    if right_boxes and loop_target_box is not None:
        target_y = (loop_target_box[1] + loop_target_box[3]) // 2
        loop_start = previous_right_bottom or ((right_boxes[-1][0] + right_boxes[-1][2]) // 2, right_boxes[-1][3])
        loop_points = [
            loop_start,
            (loop_start[0], target_y + 120),
            (loop_target_box[0] - 60, target_y + 120),
            (loop_target_box[0] - 60, target_y),
            (loop_target_box[0], target_y),
        ]
        draw_polyline_arrow(draw, loop_points)
        draw_connector_label(draw, (loop_target_box[0] + 18, target_y + 82), "调整后继续验证", label_font)

    end_box = left_box(result_box[3] + gap, term_w, term_h)
    draw_arrow(draw, (left_center_x, result_box[3]), (left_center_x, end_box[1]))
    draw_terminator(draw, end_box, "结束", node_font)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    crop_to_ink(img).save(out_path)


def render_flowchart(title: str, steps: Iterable[str], out_path: Path) -> None:
    normalized_steps = [normalize_step_text(step) for step in steps if normalize_step_text(step)]
    if len(normalized_steps) < 2:
        raise ValueError("At least two flowchart steps are required.")

    if any(is_decision_step(step) for step in normalized_steps):
        render_branched_flowchart(title, normalized_steps, out_path)
        return

    render_linear_flowchart(title, normalized_steps, out_path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Render a black-and-white portrait flowchart PNG.")
    parser.add_argument("--out", required=True, help="Output PNG path.")
    parser.add_argument("--title", default="", help="Optional title.")
    parser.add_argument("--steps-file", required=True, help="UTF-8 text file containing one step per line.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    steps_path = Path(args.steps_file)
    steps = [line.strip() for line in steps_path.read_text(encoding="utf-8").splitlines() if line.strip()]
    render_flowchart(args.title, steps, Path(args.out))


if __name__ == "__main__":
    main()
