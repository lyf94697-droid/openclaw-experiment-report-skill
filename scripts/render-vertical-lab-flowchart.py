from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable

from PIL import Image, ImageDraw, ImageFont


def load_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = [
        Path(r"C:\Windows\Fonts\msyh.ttc"),
        Path(r"C:\Windows\Fonts\simhei.ttf"),
        Path(r"C:\Windows\Fonts\simsun.ttc"),
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
    draw.line((start, end), fill=color, width=5)
    tip_x, tip_y = end
    size = 18
    draw.polygon(
        [
            (tip_x, tip_y),
            (tip_x - size, tip_y - size),
            (tip_x + size, tip_y - size),
        ],
        fill=color,
    )


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


def render_flowchart(title: str, steps: Iterable[str], out_path: Path) -> None:
    steps = list(steps)
    if len(steps) < 2:
        raise ValueError("At least two flowchart steps are required.")

    width = 1300
    box_width = 820
    box_height = 136
    gap = 58
    header_height = 156
    bottom_margin = 90
    start_y = header_height + 48
    total_height = start_y + len(steps) * box_height + (len(steps) - 1) * gap + bottom_margin

    img = Image.new("RGB", (width, total_height), (248, 250, 252))
    draw = ImageDraw.Draw(img)

    ink = (31, 42, 55)
    muted = (86, 100, 116)
    line = (74, 92, 110)
    blue = (229, 240, 251)
    green = (226, 242, 232)
    yellow = (255, 244, 220)
    card_outline = (62, 78, 96)
    header_fill = (34, 58, 76)
    shadow = (220, 226, 232)
    white = (255, 255, 255)

    title_font = load_font(40)
    small_font = load_font(22)
    badge_font = load_font(24)
    box_font = load_font(29)

    draw.rectangle((0, 0, width, header_height), fill=header_fill)
    draw.rounded_rectangle((88, 44, 184, 88), radius=8, fill=(231, 238, 244))
    draw.text((111, 51), "流程", font=small_font, fill=header_fill)
    title_width = draw.textlength(title, font=title_font)
    draw.text(((width - title_width) / 2, 46), title, font=title_font, fill=white)
    subtitle = "按报告版式优化的纵向流程图"
    subtitle_width = draw.textlength(subtitle, font=small_font)
    draw.text(((width - subtitle_width) / 2, 102), subtitle, font=small_font, fill=(218, 227, 235))

    x1 = (width - box_width) // 2
    x2 = x1 + box_width
    boxes: list[tuple[int, int, int, int]] = []
    y = start_y
    for index, step in enumerate(steps):
        fill = green if index in (0, len(steps) - 1) else (yellow if "?" in step or "是否" in step else blue)
        radius = 48 if index in (0, len(steps) - 1) else 14
        box = (x1, y, x2, y + box_height)
        shadow_box = (box[0] + 8, box[1] + 8, box[2] + 8, box[3] + 8)
        draw.rounded_rectangle(shadow_box, radius=radius, fill=shadow)
        draw.rounded_rectangle(box, radius=radius, fill=fill, outline=card_outline, width=4)
        badge_fill = white if index not in (0, len(steps) - 1) else (241, 248, 244)
        draw_badge(draw, (x1 + 54, y + box_height // 2), str(index + 1), badge_font, badge_fill, card_outline)
        text_box = (x1 + 108, y, x2 - 32, y + box_height)
        draw_centered_text(draw, text_box, step, box_font, ink, line_gap=8)
        boxes.append(box)
        y += box_height + gap

    center_x = width // 2
    for first, second in zip(boxes, boxes[1:]):
        draw_arrow(draw, (center_x, first[3] + 4), (center_x, second[1] - 4), line)

    footer = "可直接插入课程设计或实验报告正文"
    footer_width = draw.textlength(footer, font=small_font)
    draw.text(((width - footer_width) / 2, total_height - 54), footer, font=small_font, fill=muted)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    img.save(out_path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Render a portrait vertical lab flowchart PNG.")
    parser.add_argument("--out", required=True, help="Output PNG path.")
    parser.add_argument("--title", required=True, help="Flowchart title.")
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
