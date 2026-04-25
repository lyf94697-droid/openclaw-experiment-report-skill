from __future__ import annotations

import argparse
import math
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from PIL import Image, ImageDraw, ImageFont


INK = (24, 24, 24)
MUTED = (110, 110, 110)
WHITE = (255, 255, 255)


@dataclass
class TreeGroup:
    name: str
    children: list[str]


@dataclass
class LayerRow:
    name: str
    items: list[str]


@dataclass
class SwimlaneRow:
    name: str
    steps: list[str]


def load_font(size: int, *, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = []
    if bold:
        candidates.extend(
            [
                Path(r"C:\Windows\Fonts\simhei.ttf"),
                Path(r"C:\Windows\Fonts\msyhbd.ttc"),
            ]
        )
    candidates.extend(
        [
            Path(r"C:\Windows\Fonts\simsun.ttc"),
            Path(r"C:\Windows\Fonts\msyh.ttc"),
        ]
    )
    for path in candidates:
        if path.exists():
            return ImageFont.truetype(str(path), size)
    return ImageFont.load_default()


def normalize_step_text(text: str) -> str:
    compact = (text or "").strip()
    compact = compact.lstrip("\ufeff")
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


def clean_spec_line(raw_line: str) -> str:
    return (raw_line or "").lstrip("\ufeff").strip()


def split_spec_payload(payload: str) -> list[str]:
    normalized_payload = (payload or "").replace("\uff5c", "|")
    return [normalize_step_text(part) for part in normalized_payload.split("|") if normalize_step_text(part)]

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
    char_gap: int = 5,
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
        return 42

    title = title.strip()
    baseline = 74
    text_box = draw.textbbox((0, 0), title, font=font)
    text_width = text_box[2] - text_box[0]
    text_height = text_box[3] - text_box[1]
    left = (width - text_width) / 2
    right = left + text_width

    draw.text((left, baseline - text_height + 8), title, font=font, fill=INK)
    return 114


def draw_box(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    *,
    radius: int = 0,
    border_width: int = 3,
    dashed: bool = False,
) -> None:
    if dashed:
        x1, y1, x2, y2 = box
        dash = 12
        gap = 8
        for start in range(x1, x2, dash + gap):
            draw.line((start, y1, min(start + dash, x2), y1), fill=INK, width=border_width)
            draw.line((start, y2, min(start + dash, x2), y2), fill=INK, width=border_width)
        for start in range(y1, y2, dash + gap):
            draw.line((x1, start, x1, min(start + dash, y2)), fill=INK, width=border_width)
            draw.line((x2, start, x2, min(start + dash, y2)), fill=INK, width=border_width)
        return
    if radius > 0:
        draw.rounded_rectangle(box, radius=radius, outline=INK, width=border_width, fill=WHITE)
        return
    draw.rectangle(box, outline=INK, width=border_width, fill=WHITE)


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


def draw_dashed_connector(
    draw: ImageDraw.ImageDraw,
    start: tuple[int, int],
    end: tuple[int, int],
    *,
    color: tuple[int, int, int] = MUTED,
    width: int = 2,
    dash: int = 10,
    gap: int = 8,
) -> None:
    total = math.dist(start, end)
    if total <= 0:
        return
    dx = (end[0] - start[0]) / total
    dy = (end[1] - start[1]) / total
    progress = 0.0
    while progress < total:
        dash_end = min(progress + dash, total)
        x1 = start[0] + dx * progress
        y1 = start[1] + dy * progress
        x2 = start[0] + dx * dash_end
        y2 = start[1] + dy * dash_end
        draw.line((x1, y1, x2, y2), fill=color, width=width)
        progress += dash + gap


def draw_connector_label(
    draw: ImageDraw.ImageDraw,
    xy: tuple[int, int],
    text: str,
    font: ImageFont.ImageFont,
) -> None:
    bbox = draw.textbbox((0, 0), text, font=font)
    pad_x = 8
    pad_y = 5
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


def extract_spec_title(lines: Iterable[str]) -> str:
    for raw_line in lines:
        line = clean_spec_line(raw_line)
        if line.upper().startswith("@TITLE"):
            return line[6:].strip()
    return ""


def parse_tree_spec(lines: Iterable[str]) -> tuple[str, list[TreeGroup]] | None:
    root_label = ""
    groups: list[TreeGroup] = []
    for raw_line in lines:
        line = clean_spec_line(raw_line)
        if not line:
            continue
        upper = line.upper()
        if upper.startswith("@TREE") or upper.startswith("@ROOT"):
            root_label = line[5:].strip()
            continue
        if upper.startswith("@GROUP"):
            parts = split_spec_payload(line[6:].strip())
            if parts:
                groups.append(TreeGroup(name=parts[0], children=parts[1:]))
    if not groups:
        return None
    return (root_label or "系统总体设计", groups)


def parse_layer_spec(lines: Iterable[str]) -> tuple[str, list[LayerRow]] | None:
    root_label = ""
    layers: list[LayerRow] = []
    for raw_line in lines:
        line = clean_spec_line(raw_line)
        if not line:
            continue
        upper = line.upper()
        if upper.startswith("@STACK"):
            root_label = line[6:].strip()
            continue
        if upper.startswith("@LAYER"):
            parts = split_spec_payload(line[6:].strip())
            if parts:
                layers.append(LayerRow(name=parts[0], items=parts[1:]))
    if not layers:
        return None
    return (root_label or "系统分层架构", layers)


def parse_swimlane_spec(lines: Iterable[str]) -> tuple[str, list[SwimlaneRow]] | None:
    title = extract_spec_title(lines)
    lanes: list[SwimlaneRow] = []
    for raw_line in lines:
        line = clean_spec_line(raw_line)
        if not line:
            continue
        if line.upper().startswith("@LANE"):
            parts = split_spec_payload(line[5:].strip())
            if parts:
                lanes.append(SwimlaneRow(name=parts[0], steps=parts[1:]))
    if not lanes:
        return None
    return (title or "业务协同泳道图", lanes)


def render_tree_diagram(title: str, root_label: str, groups: list[TreeGroup], out_path: Path) -> None:
    root_w = 360
    root_h = 78
    group_w = 180
    group_h = 66
    child_w = 82
    child_h = 220
    child_gap = 26
    cluster_gap = 120
    margin_x = 120

    cluster_widths: list[int] = []
    for group in groups:
        child_count = max(1, len(group.children))
        child_total_width = child_count * child_w + max(0, child_count - 1) * child_gap
        cluster_widths.append(max(group_w, child_total_width))

    total_width = sum(cluster_widths) + max(0, len(cluster_widths) - 1) * cluster_gap
    width = max(1460, total_width + margin_x * 2)

    title_font = load_font(28, bold=True)
    root_font = load_font(28, bold=True)
    group_font = load_font(23, bold=True)
    child_font = load_font(19)

    top = 34
    if title.strip():
        preview = ImageDraw.Draw(Image.new("RGB", (width, 180), WHITE))
        top = draw_title(preview, width, title, title_font)

    root_y = top
    group_y = root_y + 128
    child_y = group_y + 142
    height = child_y + child_h + 110

    img = Image.new("RGB", (width, height), WHITE)
    draw = ImageDraw.Draw(img)
    if title.strip():
        top = draw_title(draw, width, title, title_font)
        root_y = top
        group_y = root_y + 128
        child_y = group_y + 142

    root_box = (
        width // 2 - root_w // 2,
        root_y,
        width // 2 + root_w // 2,
        root_y + root_h,
    )
    draw_box(draw, root_box)
    draw_centered_text(draw, root_box, root_label, root_font, line_gap=5)

    cluster_left = (width - total_width) / 2
    group_centers: list[int] = []
    group_boxes: list[tuple[int, int, int, int]] = []
    child_box_groups: list[list[tuple[int, int, int, int]]] = []
    for cluster_width, group in zip(cluster_widths, groups):
        cluster_center = cluster_left + cluster_width / 2
        group_box = (
            int(cluster_center - group_w / 2),
            group_y,
            int(cluster_center + group_w / 2),
            group_y + group_h,
        )
        draw_box(draw, group_box)
        draw_centered_text(draw, group_box, group.name, group_font, line_gap=4)
        group_boxes.append(group_box)
        group_centers.append(int(cluster_center))

        children = group.children or ["核心模块"]
        child_total_width = len(children) * child_w + max(0, len(children) - 1) * child_gap
        child_left = cluster_center - child_total_width / 2
        boxes: list[tuple[int, int, int, int]] = []
        for index, child in enumerate(children):
            left = int(child_left + index * (child_w + child_gap))
            box = (left, child_y, left + child_w, child_y + child_h)
            draw_box(draw, box)
            draw_vertical_text(draw, box, child, child_font)
            boxes.append(box)
        child_box_groups.append(boxes)
        cluster_left += cluster_width + cluster_gap

    root_center_x = (root_box[0] + root_box[2]) // 2
    branch_y = root_box[3] + 40
    draw.line((root_center_x, root_box[3], root_center_x, branch_y), fill=INK, width=3)
    if group_centers:
        draw.line((group_centers[0], branch_y, group_centers[-1], branch_y), fill=INK, width=3)
        for center_x, group_box in zip(group_centers, group_boxes):
            draw.line((center_x, branch_y, center_x, group_box[1]), fill=INK, width=3)

    for group_box, child_boxes in zip(group_boxes, child_box_groups):
        group_center_x = (group_box[0] + group_box[2]) // 2
        child_branch_y = group_box[3] + 42
        draw.line((group_center_x, group_box[3], group_center_x, child_branch_y), fill=INK, width=3)
        if child_boxes:
            centers = [int((box[0] + box[2]) / 2) for box in child_boxes]
            if len(centers) > 1:
                draw.line((centers[0], child_branch_y, centers[-1], child_branch_y), fill=INK, width=3)
            for center_x, child_box in zip(centers, child_boxes):
                draw.line((center_x, child_branch_y, center_x, child_box[1]), fill=INK, width=3)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    crop_to_ink(img).save(out_path)


def render_layer_diagram(title: str, root_label: str, layers: list[LayerRow], out_path: Path) -> None:
    width = 1600
    title_font = load_font(28, bold=True)
    root_font = load_font(28, bold=True)
    layer_font = load_font(23, bold=True)
    item_font = load_font(19)

    top = 34
    if title.strip():
        preview = ImageDraw.Draw(Image.new("RGB", (width, 180), WHITE))
        top = draw_title(preview, width, title, title_font)

    root_box = (width // 2 - 220, top, width // 2 + 220, top + 76)
    band_y = root_box[3] + 56
    band_height = 118
    band_gap = 34
    height = band_y + len(layers) * (band_height + band_gap) + 80

    img = Image.new("RGB", (width, height), WHITE)
    draw = ImageDraw.Draw(img)
    if title.strip():
        top = draw_title(draw, width, title, title_font)
        root_box = (width // 2 - 220, top, width // 2 + 220, top + 76)
        band_y = root_box[3] + 56

    draw_box(draw, root_box)
    draw_centered_text(draw, root_box, root_label, root_font, line_gap=5)

    center_x = (root_box[0] + root_box[2]) // 2
    draw.line((center_x, root_box[3], center_x, band_y - 10), fill=INK, width=3)

    current_y = band_y
    previous_center_y = root_box[3]
    for index, layer in enumerate(layers):
        band_box = (120, current_y, width - 120, current_y + band_height)
        draw_box(draw, band_box)
        label_box = (band_box[0], band_box[1], band_box[0] + 220, band_box[3])
        draw_box(draw, label_box)
        draw_centered_text(draw, label_box, layer.name, layer_font, line_gap=4)

        items = layer.items or ["模块"]
        item_area_left = label_box[2] + 28
        item_area_right = band_box[2] - 28
        item_count = len(items)
        gap = 26
        item_width = min(210, int((item_area_right - item_area_left - gap * max(0, item_count - 1)) / item_count))
        total_items_width = item_count * item_width + gap * max(0, item_count - 1)
        start_x = int(item_area_left + ((item_area_right - item_area_left) - total_items_width) / 2)
        item_center_y = band_box[1] + 24
        for item_index, item in enumerate(items):
            x = start_x + item_index * (item_width + gap)
            item_box = (x, band_box[1] + 18, x + item_width, band_box[3] - 18)
            draw_box(draw, item_box)
            draw_centered_text(draw, item_box, item, item_font, line_gap=4, padding_x=14)

        if index > 0:
            draw_arrow(
                draw,
                (center_x, previous_center_y + band_height),
                (center_x, band_box[1]),
                color=MUTED,
                width=3,
            )
        previous_center_y = band_box[1]
        current_y = band_box[3] + band_gap

    out_path.parent.mkdir(parents=True, exist_ok=True)
    crop_to_ink(img).save(out_path)


def render_swimlane_diagram(title: str, lanes: list[SwimlaneRow], out_path: Path) -> None:
    lane_label_w = 150
    step_w = 185
    step_h = 70
    col_gap = 42
    lane_h = 130
    lane_gap = 28
    padding = 90

    max_columns = max(len(lane.steps) for lane in lanes)
    width = padding * 2 + lane_label_w + 36 + max_columns * step_w + max(0, max_columns - 1) * col_gap

    title_font = load_font(28, bold=True)
    lane_font = load_font(22, bold=True)
    step_font = load_font(18)

    preview = ImageDraw.Draw(Image.new("RGB", (width, 180), WHITE))
    top = draw_title(preview, width, title, title_font) if title.strip() else 42
    start_y = top + 18
    height = start_y + len(lanes) * lane_h + max(0, len(lanes) - 1) * lane_gap + 80

    img = Image.new("RGB", (width, height), WHITE)
    draw = ImageDraw.Draw(img)
    if title.strip():
        top = draw_title(draw, width, title, title_font)
        start_y = top + 18

    grid_boxes: list[list[tuple[int, int, int, int]]] = []
    for lane_index, lane in enumerate(lanes):
        y = start_y + lane_index * (lane_h + lane_gap)
        lane_box = (padding, y, width - padding, y + lane_h)
        draw_box(draw, lane_box)
        label_box = (lane_box[0], lane_box[1], lane_box[0] + lane_label_w, lane_box[3])
        draw_box(draw, label_box)
        draw_centered_text(draw, label_box, lane.name, lane_font)

        boxes: list[tuple[int, int, int, int]] = []
        for step_index in range(max_columns):
            x = lane_box[0] + lane_label_w + 36 + step_index * (step_w + col_gap)
            step_box = (
                x,
                y + (lane_h - step_h) // 2,
                x + step_w,
                y + (lane_h - step_h) // 2 + step_h,
            )
            boxes.append(step_box)
            if step_index < len(lane.steps):
                draw_box(draw, step_box, radius=12)
                draw_centered_text(draw, step_box, lane.steps[step_index], step_font, line_gap=4, padding_x=16)
        grid_boxes.append(boxes)

    for lane_index, lane in enumerate(lanes):
        boxes = grid_boxes[lane_index]
        for step_index in range(len(lane.steps) - 1):
            start_box = boxes[step_index]
            end_box = boxes[step_index + 1]
            draw_arrow(
                draw,
                (start_box[2], (start_box[1] + start_box[3]) // 2),
                (end_box[0], (end_box[1] + end_box[3]) // 2),
                width=3,
            )

    for lane_index in range(len(lanes) - 1):
        upper_lane = lanes[lane_index]
        lower_lane = lanes[lane_index + 1]
        for step_index in range(min(len(upper_lane.steps), len(lower_lane.steps))):
            upper_box = grid_boxes[lane_index][step_index]
            lower_box = grid_boxes[lane_index + 1][step_index]
            draw_dashed_connector(
                draw,
                ((upper_box[0] + upper_box[2]) // 2, upper_box[3]),
                ((lower_box[0] + lower_box[2]) // 2, lower_box[1]),
            )

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

    title_font = load_font(30, bold=True)
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
    draw_centered_text(draw, start_box, "开始", node_font)

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
    draw_centered_text(draw, end_box, "结束", node_font)

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

    width = 1780
    top_padding = 48
    left_center_x = 620
    right_left = 1140
    node_w = 530
    node_h = 84
    term_w = 240
    term_h = 72
    decision_w = 430
    decision_h = 170
    gap = 64

    title_font = load_font(30, bold=True)
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
    draw_centered_text(draw, start_box, "开始", node_font)
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
    draw_centered_text(draw, decision_box, decision_text, node_font, line_gap=6, padding_x=48)
    decision_center_y = (decision_box[1] + decision_box[3]) // 2

    result_box = left_box(decision_box[3] + gap)
    draw_connector_label(draw, (left_center_x + 20, decision_box[3] + 8), "是", label_font)
    draw_arrow(draw, (left_center_x, decision_box[3]), (left_center_x, result_box[1]))
    draw_box(draw, result_box)
    draw_centered_text(draw, result_box, result_text, node_font, line_gap=6)

    branch_start = (decision_box[2], decision_center_y)
    branch_y = decision_center_y - node_h // 2
    previous_right_bottom: tuple[int, int] | None = None
    right_boxes: list[tuple[int, int, int, int]] = []
    for index, step in enumerate(branch_steps, start=1):
        box = right_box(branch_y if index == 1 else branch_y + (index - 1) * (node_h + gap))
        if index == 1:
            draw_connector_label(draw, (decision_box[2] + 18, decision_center_y - 38), "否", label_font)
            draw_polyline_arrow(draw, [branch_start, (box[0] - 46, decision_center_y), (box[0], (box[1] + box[3]) // 2)])
        else:
            assert previous_right_bottom is not None
            draw_arrow(draw, previous_right_bottom, ((box[0] + box[2]) // 2, box[1]))
        draw_box(draw, box)
        draw_centered_text(draw, box, step, node_font, line_gap=6)
        previous_right_bottom = ((box[0] + box[2]) // 2, box[3])
        right_boxes.append(box)

    if right_boxes and loop_target_box is not None:
        target_y = (loop_target_box[1] + loop_target_box[3]) // 2
        last_branch_box = right_boxes[-1]
        loop_start = (last_branch_box[0], (last_branch_box[1] + last_branch_box[3]) // 2)
        loop_drop_x = loop_start[0] - 72
        loop_join_x = loop_target_box[0] - 72
        future_end_bottom = result_box[3] + gap + term_h
        loop_y = future_end_bottom + 36
        loop_points = [
            loop_start,
            (loop_drop_x, loop_start[1]),
            (loop_drop_x, loop_y),
            (loop_join_x, loop_y),
            (loop_join_x, target_y),
            (loop_target_box[0], target_y),
        ]
        draw_polyline_arrow(draw, loop_points, color=MUTED, width=3)
        draw_connector_label(draw, (loop_drop_x - 136, loop_y - 42), "修正后继续验证", label_font)

    end_box = left_box(result_box[3] + gap, term_w, term_h)
    draw_arrow(draw, (left_center_x, result_box[3]), (left_center_x, end_box[1]))
    draw_box(draw, end_box, radius=18)
    draw_centered_text(draw, end_box, "结束", node_font)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    crop_to_ink(img).save(out_path)


def render_flowchart(title_arg: str, steps: Iterable[str], out_path: Path) -> None:
    lines = [clean_spec_line(line) for line in steps if clean_spec_line(line)]

    layer_spec = parse_layer_spec(lines)
    if layer_spec is not None:
        title = extract_spec_title(lines) or title_arg
        root_label, layers = layer_spec
        render_layer_diagram(title, root_label, layers, out_path)
        return

    swimlane_spec = parse_swimlane_spec(lines)
    if swimlane_spec is not None:
        title, lanes = swimlane_spec
        render_swimlane_diagram(title or title_arg, lanes, out_path)
        return

    tree_spec = parse_tree_spec(lines)
    if tree_spec is not None:
        title = extract_spec_title(lines)
        root_label, groups = tree_spec
        render_tree_diagram(title, root_label, groups, out_path)
        return

    normalized_steps = [normalize_step_text(step) for step in lines if normalize_step_text(step) and not step.upper().startswith("@TITLE")]
    if len(normalized_steps) < 2:
        raise ValueError("At least two flowchart steps are required.")

    title = extract_spec_title(lines) or title_arg
    if any(is_decision_step(step) for step in normalized_steps):
        render_branched_flowchart(title, normalized_steps, out_path)
        return
    render_linear_flowchart(title, normalized_steps, out_path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Render black-and-white process, structure, swimlane, or layered diagrams.")
    parser.add_argument("--out", required=True, help="Output PNG path.")
    parser.add_argument("--title", default="", help="Optional title for regular process diagrams.")
    parser.add_argument("--steps-file", required=True, help="UTF-8 text file containing one step/spec line per row.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    steps_path = Path(args.steps_file)
    lines = [line.rstrip("\n") for line in steps_path.read_text(encoding="utf-8-sig").splitlines()]
    render_flowchart(args.title, lines, Path(args.out))


if __name__ == "__main__":
    main()
