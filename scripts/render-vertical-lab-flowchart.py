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
    draw.line((start, end), fill=color, width=4)
    tip_x, tip_y = end
    size = 16
    draw.polygon(
        [
            (tip_x, tip_y),
            (tip_x - size, tip_y - size),
            (tip_x + size, tip_y - size),
        ],
        fill=color,
    )


def render_flowchart(title: str, steps: Iterable[str], out_path: Path) -> None:
    steps = list(steps)
    if len(steps) < 2:
        raise ValueError("At least two flowchart steps are required.")

    width = 1100
    box_width = 760
    box_height = 128
    gap = 52
    title_top = 52
    title_gap = 60
    bottom_margin = 70
    start_y = title_top + 48 + title_gap
    total_height = start_y + len(steps) * box_height + (len(steps) - 1) * gap + bottom_margin

    img = Image.new("RGB", (width, total_height), "white")
    draw = ImageDraw.Draw(img)

    ink = (34, 45, 56)
    line = (88, 104, 122)
    blue = (233, 241, 251)
    green = (228, 243, 232)

    title_font = load_font(38)
    box_font = load_font(28)

    title_width = draw.textlength(title, font=title_font)
    draw.text(((width - title_width) / 2, title_top), title, font=title_font, fill=ink)

    x1 = (width - box_width) // 2
    x2 = x1 + box_width
    boxes: list[tuple[int, int, int, int]] = []
    y = start_y
    for index, step in enumerate(steps):
        fill = green if index in (0, len(steps) - 1) else blue
        radius = 64 if index in (0, len(steps) - 1) else 22
        box = (x1, y, x2, y + box_height)
        draw.rounded_rectangle(box, radius=radius, fill=fill, outline=line, width=4)
        draw_centered_text(draw, box, step, box_font, ink)
        boxes.append(box)
        y += box_height + gap

    center_x = width // 2
    for first, second in zip(boxes, boxes[1:]):
        draw_arrow(draw, (center_x, first[3]), (center_x, second[1]), line)

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
