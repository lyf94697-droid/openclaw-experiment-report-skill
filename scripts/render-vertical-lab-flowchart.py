from __future__ import annotations

import argparse
import math
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from PIL import Image, ImageDraw, ImageFont


INK = (24, 24, 24)
WHITE = (255, 255, 255)


@dataclass
class TreeGroup:
    name: str
    children: list[str]


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
    compact = re.sub(
        r"^\s*[\(\uff08]?[0-9\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]+[\)\uff09\.\u3001]\s*",
        "",
        compact,
    )
    compact = compact.replace("\uff1a", " ").replace(":", " ")
    compact = re.sub(r"\s+", " ", compact)
    return compact.strip()


def is_decision_step(text: str) -> bool:
    compact = normalize_step_text(text)
    return ("\u662f\u5426" in compact) or compact.endswith("?") or compact.endswith("\uff1f")


def wrap_text(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, max_width: int) -> list[str]:
    if not text:
        return [""]

    lines: list[str] = []
    current = ""
    for char in text:
        candidate = current + char
        if (not current) or draw.textlength(candidate, font=font) <= max_width:
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
    *,
    fill: tuple[int, int, int] = INK,
    line_gap: int = 8,
    padding_x: int = 24,
) -> None:
    max_width = max(40, box[2] - box[0] - padding_x * 2)
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


def draw_vertical_text(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    text: str,
    font: ImageFont.ImageFont,
    *,
    fill: tuple[int, int, int] = INK,
    char_gap: int = 4,
) -> None:
    chars = [char for char in (text or "").strip() if not char.isspace()]
    if not chars:
        chars = [""]

    boxes = [draw.textbbox((0, 0), char, font=font) for char in chars]
    heights = [bbox[3] - bbox[1] for bbox in boxes]
    total_height = sum(heights) + char_gap * max(0, len(chars) - 1)
    y = box[1] + (box[3] - box[1] - total_height) / 2

    for char, bbox, height in zip(chars, boxes, heights):
        width = bbox[2] - bbox[0]
        x = box[0] + (box[2] - box[0] - width) / 2
        draw.text((x, y), char, font=font, fill=fill)
        y += height + char_gap


def draw_title(draw: ImageDraw.ImageDraw, width: int, title: str, font: ImageFont.ImageFont) -> int:
    if not title.strip():
        return 44

    title = title.strip()
    baseline = 74
    text_box = draw.textbbox((0, 0), title, font=font)
    text_width = text_box[2] - text_box[0]
    text_height = text_box[3] - text_box[1]
    left = (width - text_width) / 2
    right = left + text_width

    draw.line((110, baseline, left - 28, baseline), fill=INK, width=3)
    draw.line((right + 28, baseline, width - 110, baseline), fill=INK, width=3)
    draw.text((left, baseline - text_height + 8), title, font=font, fill=INK)
    return 116


def draw_arrow_head(
    draw: ImageDraw.ImageDraw,
    start: tuple[int, int],
    end: tuple[int, int],
    *,
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
    *,
    color: tuple[int, int, int] = INK,
    width: int = 4,
) -> None:
    draw.line((start, end), fill=color, width=width)
    draw_arrow_head(draw, start, end, color=color, width=width)


def draw_polyline_arrow(
    draw: ImageDraw.ImageDraw,
    points: list[tuple[int, int]],
    *,
    color: tuple[int, int, int] = INK,
    width: int = 4,
) -> None:
    if len(points) < 2:
        return
    draw.line(points, fill=color, width=width)
    draw_arrow_head(draw, points[-2], points[-1], color=color, width=width)


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


def draw_box(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    *,
    radius: int = 0,
    border_width: int = 3,
) -> None:
    if radius > 0:
        draw.rounded_rectangle(box, radius=radius, outline=INK, width=border_width, fill=WHITE)
        return
    draw.rectangle(box, outline=INK, width=border_width, fill=WHITE)


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


def parse_tree_spec(lines: Iterable[str]) -> tuple[str, list[TreeGroup]] | None:
    root_label = ""
    groups: list[TreeGroup] = []
    for raw_line in lines:
        line = (raw_line or "").strip()
        if not line:
            continue

        upper = line.upper()
        if upper.startswith("@TREE"):
            root_label = line[5:].strip()
            continue
        if upper.startswith("@ROOT"):
            root_label = line[5:].strip()
            continue
        if upper.startswith("@GROUP"):
            payload = line[6:].strip()
            parts = [
                normalize_step_text(part)
                for part in re.split(r"[|｜]", payload)
                if normalize_step_text(part)
            ]
            if parts:
                groups.append(TreeGroup(name=parts[0], children=parts[1:]))

    if not groups:
        return None
    return (root_label or "\u7cfb\u7edf\u603b\u4f53\u8bbe\u8ba1"), groups


def render_tree_diagram(title: str, root_label: str, groups: list[TreeGroup], out_path: Path) -> None:
    root_w = 330
    root_h = 70
    group_w = 180
    group_h = 66
    child_w = 78
    child_h = 220
    child_gap = 24
    cluster_gap = 120
    margin_x = 120

    cluster_widths: list[int] = []
    for group in groups:
        child_count = max(1, len(group.children))
        child_total_width = child_count * child_w + max(0, child_count - 1) * child_gap
        cluster_widths.append(max(group_w, child_total_width))

    total_clusters_width = sum(cluster_widths) + max(0, len(cluster_widths) - 1) * cluster_gap
    width = max(1440, total_clusters_width + margin_x * 2)

    title_font = load_font(28)
    root_font = load_font(26)
    group_font = load_font(24)
    child_font = load_font(19)

    top = 38
    if title.strip():
        preview = ImageDraw.Draw(Image.new("RGB", (width, 180), WHITE))
        top = draw_title(preview, width, title, title_font)

    root_y = top
    group_y = root_y + 132
    child_y = group_y + 146
    height = child_y + child_h + 120

    img = Image.new("RGB", (width, height), WHITE)
    draw = ImageDraw.Draw(img)
    if title.strip():
        top = draw_title(draw, width, title, title_font)
        root_y = top
        group_y = root_y + 132
        child_y = group_y + 146

    root_box = (
        width // 2 - root_w // 2,
        root_y,
        width // 2 + root_w // 2,
        root_y + root_h,
    )
    draw_box(draw, root_box)
    draw_centered_text(draw, root_box, root_label, root_font, line_gap=5)

    cluster_left = (width - total_clusters_width) / 2
    group_centers: list[int] = []
    group_boxes: list[tuple[int, int, int, int]] = []
    child_boxes_per_group: list[list[tuple[int, int, int, int]]] = []
    for cluster_width, group in zip(cluster_widths, groups):
        cluster_center = cluster_left + cluster_width / 2
        group_box = (
            int(cluster_center - group_w / 2),
            group_y,
            int(cluster_center + group_w / 2),
            group_y + group_h,
        )
        group_boxes.append(group_box)
        group_centers.append(int(cluster_center))
        draw_box(draw, group_box)
        draw_centered_text(draw, group_box, group.name, group_font, line_gap=4)

        child_boxes: list[tuple[int, int, int, int]] = []
        child_count = max(1, len(group.children))
        child_total_width = child_count * child_w + max(0, child_count - 1) * child_gap
        child_left = cluster_center - child_total_width / 2
        for index, child in enumerate(group.children or ["\u6838\u5fc3\u6a21\u5757"]):
            left = int(child_left + index * (child_w + child_gap))
            box = (left, child_y, left + child_w, child_y + child_h)
            child_boxes.append(box)
            draw_box(draw, box)
            draw_vertical_text(draw, box, child, child_font)
        child_boxes_per_group.append(child_boxes)
        cluster_left += cluster_width + cluster_gap

    root_center_x = (root_box[0] + root_box[2]) // 2
    root_branch_y = root_box[3] + 42
    draw.line((root_center_x, root_box[3], root_center_x, root_branch_y), fill=INK, width=3)
    if group_centers:
        draw.line((group_centers[0], root_branch_y, group_centers[-1], root_branch_y), fill=INK, width=3)
        for group_center, group_box in zip(group_centers, group_boxes):
            draw.line((group_center, root_branch_y, group_center, group_box[1]), fill=INK, width=3)

    for group_box, child_boxes in zip(group_boxes, child_boxes_per_group):
        group_center_x = (group_box[0] + group_box[2]) // 2
        child_branch_y = group_box[3] + 44
        draw.line((group_center_x, group_box[3], group_center_x, child_branch_y), fill=INK, width=3)
        if child_boxes:
            child_centers = [int((box[0] + box[2]) / 2) for box in child_boxes]
            if len(child_centers) > 1:
                draw.line((child_centers[0], child_branch_y, child_centers[-1], child_branch_y), fill=INK, width=3)
            for child_center_x, child_box in zip(child_centers, child_boxes):
                draw.line((child_center_x, child_branch_y, child_center_x, child_box[1]), fill=INK, width=3)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    crop_to_ink(img).save(out_path)


def render_linear_flowchart(title: str, steps: list[str], out_path: Path) -> None:
    width = 1360
    top_padding = 48
    node_w = 560
    node_h = 84
    term_w = 240
    term_h = 72
    gap = 62

    title_font = load_font(30)
    node_font = load_font(26)

    preview = ImageDraw.Draw(Image.new("RGB", (width, 180), WHITE))
    top = draw_title(preview, width, title, title_font) if title.strip() else top_padding
    height = top + term_h * 2 + len(steps) * (node_h + gap) + 140
    img = Image.new("RGB", (width, height), WHITE)
    draw = ImageDraw.Draw(img)
    top = draw_title(draw, width, title, title_font) if title.strip() else top_padding

    center_x = width // 2
    start_box = (center_x - term_w // 2, top, center_x + term_w // 2, top + term_h)
    draw_box(draw, start_box, radius=18)
    draw_centered_text(draw, start_box, "\u5f00\u59cb", node_font)

    previous_bottom = (center_x, start_box[3])
    y = start_box[3] + gap
    for step in steps:
        box = (center_x - node_w // 2, y, center_x + node_w // 2, y + node_h)
        draw_arrow(draw, previous_bottom, (center_x, box[1]))
        draw_box(draw, box)
        draw_centered_text(draw, box, normalize_step_text(step), node_font, line_gap=6)
        previous_bottom = (center_x, box[3])
        y = box[3] + gap

    end_box = (center_x - term_w // 2, y, center_x + term_w // 2, y + term_h)
    draw_arrow(draw, previous_bottom, (center_x, end_box[1]))
    draw_box(draw, end_box, radius=18)
    draw_centered_text(draw, end_box, "\u7ed3\u675f", node_font)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    crop_to_ink(img).save(out_path)


def render_branched_flowchart(title: str, steps: list[str], out_path: Path) -> None:
    decision_index = next((index for index, step in enumerate(steps) if is_decision_step(step)), -1)
    if decision_index < 1 or len(steps) < decision_index + 3:
        render_linear_flowchart(title, steps, out_path)
        return

    pre_steps = [normalize_step_text(step) for step in steps[:decision_index]]
    decision_text = normalize_step_text(steps[decision_index])
    branch_steps = [normalize_step_text(step) for step in steps[decision_index + 1 : -1]]
    result_text = normalize_step_text(steps[-1])

    width = 1760
    top_padding = 48
    left_center_x = 610
    right_left = 1110
    node_w = 520
    node_h = 84
    term_w = 240
    term_h = 72
    decision_w = 420
    decision_h = 166
    gap = 64

    title_font = load_font(30)
    node_font = load_font(26)
    label_font = load_font(20)

    preview = ImageDraw.Draw(Image.new("RGB", (width, 2200), WHITE))
    top = draw_title(preview, width, title, title_font) if title.strip() else top_padding
    height = top + term_h * 2 + (len(pre_steps) + len(branch_steps) + 2) * (node_h + gap) + 260
    img = Image.new("RGB", (width, height), WHITE)
    draw = ImageDraw.Draw(img)
    top = draw_title(draw, width, title, title_font) if title.strip() else top_padding

    def left_box(y: int, w: int = node_w, h: int = node_h) -> tuple[int, int, int, int]:
        return (left_center_x - w // 2, y, left_center_x + w // 2, y + h)

    def right_box(y: int) -> tuple[int, int, int, int]:
        return (right_left, y, right_left + node_w, y + node_h)

    start_box = left_box(top, term_w, term_h)
    draw_box(draw, start_box, radius=18)
    draw_centered_text(draw, start_box, "\u5f00\u59cb", node_font)
    previous_bottom = (left_center_x, start_box[3])
    y = start_box[3] + gap

    loop_target_box: tuple[int, int, int, int] | None = None
    for step in pre_steps:
        box = left_box(y)
        draw_arrow(draw, previous_bottom, (left_center_x, box[1]))
        draw_box(draw, box)
        draw_centered_text(draw, box, step, node_font, line_gap=6)
        previous_bottom = (left_center_x, box[3])
        loop_target_box = box
        y = box[3] + gap

    decision_box = left_box(y, decision_w, decision_h)
    draw_arrow(draw, previous_bottom, (left_center_x, decision_box[1]))
    x1, y1, x2, y2 = decision_box
    diamond = [
        ((x1 + x2) // 2, y1),
        (x2, (y1 + y2) // 2),
        ((x1 + x2) // 2, y2),
        (x1, (y1 + y2) // 2),
    ]
    draw.polygon(diamond, outline=INK, fill=WHITE)
    draw.line(diamond + [diamond[0]], fill=INK, width=3)
    draw_centered_text(draw, decision_box, decision_text, node_font, line_gap=6, padding_x=42)
    decision_center_y = (decision_box[1] + decision_box[3]) // 2

    result_box = left_box(decision_box[3] + gap)
    draw_connector_label(draw, (left_center_x + 18, decision_box[3] + 8), "\u662f", label_font)
    draw_arrow(draw, (left_center_x, decision_box[3]), (left_center_x, result_box[1]))
    draw_box(draw, result_box)
    draw_centered_text(draw, result_box, result_text, node_font, line_gap=6)

    branch_start = (decision_box[2], decision_center_y)
    previous_right_bottom: tuple[int, int] | None = None
    right_boxes: list[tuple[int, int, int, int]] = []
    branch_y = decision_center_y - node_h // 2
    for index, step in enumerate(branch_steps, start=1):
        box = right_box(branch_y if index == 1 else branch_y + (index - 1) * (node_h + gap))
        right_boxes.append(box)
        if index == 1:
            draw_connector_label(draw, (decision_box[2] + 18, decision_center_y - 38), "\u5426", label_font)
            draw_polyline_arrow(draw, [branch_start, (box[0] - 42, decision_center_y), (box[0], (box[1] + box[3]) // 2)])
        else:
            assert previous_right_bottom is not None
            draw_arrow(draw, previous_right_bottom, ((box[0] + box[2]) // 2, box[1]))
        draw_box(draw, box)
        draw_centered_text(draw, box, step, node_font, line_gap=6)
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
        draw_connector_label(draw, (loop_target_box[0] + 18, target_y + 82), "\u8c03\u6574\u540e\u7ee7\u7eed\u9a8c\u8bc1", label_font)

    end_box = left_box(result_box[3] + gap, term_w, term_h)
    draw_arrow(draw, (left_center_x, result_box[3]), (left_center_x, end_box[1]))
    draw_box(draw, end_box, radius=18)
    draw_centered_text(draw, end_box, "\u7ed3\u675f", node_font)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    crop_to_ink(img).save(out_path)


def render_flowchart(title: str, steps: Iterable[str], out_path: Path) -> None:
    lines = [line.strip() for line in steps if (line or "").strip()]
    tree_spec = parse_tree_spec(lines)
    if tree_spec is not None:
        root_label, groups = tree_spec
        render_tree_diagram("", root_label, groups, out_path)
        return

    normalized_steps = [normalize_step_text(step) for step in lines if normalize_step_text(step)]
    if len(normalized_steps) < 2:
        raise ValueError("At least two flowchart steps are required.")

    if any(is_decision_step(step) for step in normalized_steps):
        render_branched_flowchart(title, normalized_steps, out_path)
        return

    render_linear_flowchart(title, normalized_steps, out_path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Render a black-and-white portrait process or structure diagram PNG.")
    parser.add_argument("--out", required=True, help="Output PNG path.")
    parser.add_argument("--title", default="", help="Optional title for process diagrams.")
    parser.add_argument("--steps-file", required=True, help="UTF-8 text file containing one step/spec line per row.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    steps_path = Path(args.steps_file)
    lines = [line.strip() for line in steps_path.read_text(encoding="utf-8").splitlines() if line.strip()]
    render_flowchart(args.title, lines, Path(args.out))


if __name__ == "__main__":
    main()
