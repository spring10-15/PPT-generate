#!/usr/bin/env python3
from __future__ import annotations

import json
import posixpath
import sys
import zipfile
from collections import defaultdict
from pathlib import Path

from lxml import etree


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}

EMU_PER_PT = 12700.0
DEFAULT_FONT_PT = 14.0


def as_float(value: str | None, fallback: float = 0.0) -> float:
    if value is None:
        return fallback
    try:
        return float(value)
    except ValueError:
        return fallback


def emu_to_pt(value: str | None) -> float:
    return round(as_float(value) / EMU_PER_PT, 2)


def extract_font_size_pt(node: etree._Element, fallback: float | None = DEFAULT_FONT_PT) -> float | None:
    for xpath in (".//a:rPr", ".//a:defRPr", ".//a:endParaRPr"):
        for item in node.xpath(xpath, namespaces=NS):
            size = item.get("sz")
            if size:
                return round(as_float(size) / 100.0, 2)
    return fallback


def extract_shape_name(node: etree._Element) -> tuple[str, str]:
    c_nv_pr = node.xpath("./p:nvSpPr/p:cNvPr | ./p:nvPicPr/p:cNvPr | ./p:nvGraphicFramePr/p:cNvPr", namespaces=NS)
    if not c_nv_pr:
        return "unknown", "0"
    element = c_nv_pr[0]
    return element.get("name", "unknown"), element.get("id", "0")


def extract_placeholder_kind(node: etree._Element) -> str:
    if node.tag.endswith("pic"):
        return "PICTURE"

    if node.tag.endswith("graphicFrame"):
        if node.xpath(".//a:tbl", namespaces=NS):
            return "TABLE"
        if node.xpath(".//c:chart", namespaces=NS):
            return "CHART"
        return "TEXT"

    ph = node.xpath("./p:nvSpPr/p:nvPr/p:ph", namespaces=NS)
    if ph:
        ph_type = (ph[0].get("type") or "").lower()
        if ph_type in {"pic", "media"}:
            return "PICTURE"
        if ph_type == "tbl":
            return "TABLE"
        if ph_type == "chart":
            return "CHART"

    if node.xpath(".//p:txBody", namespaces=NS):
        return "TEXT"

    return "TEXT"


def extract_placeholder_ref(node: etree._Element) -> tuple[str, str] | None:
    ph = node.xpath(
        "./p:nvSpPr/p:nvPr/p:ph | ./p:nvPicPr/p:nvPr/p:ph | ./p:nvGraphicFramePr/p:nvPr/p:ph",
        namespaces=NS,
    )
    if not ph:
        return None

    item = ph[0]
    ph_type = (item.get("type") or "body").lower()
    ph_idx = item.get("idx") or "0"
    return ph_type, ph_idx


def extract_text(node: etree._Element) -> str:
    parts = [text.strip() for text in node.xpath(".//a:t/text()", namespaces=NS)]
    return "".join(part for part in parts if part)


def estimate_text_capacity(width_pt: float, height_pt: float, font_pt: float) -> tuple[int, int]:
    safe_font = max(font_pt, 8.0)
    chars_per_line = max(1, int((width_pt - safe_font * 0.6) / max(safe_font * 0.92, 1.0)))
    line_height = max(safe_font * 1.45, 10.0)
    max_lines = max(1, int((height_pt - safe_font * 0.6) / line_height))
    return chars_per_line * max_lines, max_lines


def extract_geometry(node: etree._Element) -> tuple[float, float, float, float]:
    xfrm = node.xpath("./p:spPr/a:xfrm | ./p:spPr/p:xfrm | ./p:xfrm | ./p:pic/p:spPr/a:xfrm | ./p:graphicFrame/p:xfrm", namespaces=NS)
    if not xfrm:
        xfrm = node.xpath(".//a:xfrm | .//p:xfrm", namespaces=NS)
    if not xfrm:
        return 0.0, 0.0, 0.0, 0.0

    xfrm_node = xfrm[0]
    off = xfrm_node.xpath("./a:off | ./p:off", namespaces=NS)
    ext = xfrm_node.xpath("./a:ext | ./p:ext", namespaces=NS)
    x = emu_to_pt(off[0].get("x") if off else None)
    y = emu_to_pt(off[0].get("y") if off else None)
    w = emu_to_pt(ext[0].get("cx") if ext else None)
    h = emu_to_pt(ext[0].get("cy") if ext else None)
    return x, y, w, h


def iter_shape_nodes(root: etree._Element) -> list[etree._Element]:
    nodes: list[etree._Element] = []
    for xpath in ("./p:cSld/p:spTree/p:sp", "./p:cSld/p:spTree/p:pic", "./p:cSld/p:spTree/p:graphicFrame"):
        nodes.extend(root.xpath(xpath, namespaces=NS))
    return nodes


def build_placeholder_lookup(root: etree._Element | None) -> dict[tuple[str, str], etree._Element]:
    if root is None:
        return {}

    lookup: dict[tuple[str, str], etree._Element] = {}
    for node in iter_shape_nodes(root):
        placeholder_ref = extract_placeholder_ref(node)
        if placeholder_ref and placeholder_ref not in lookup:
            lookup[placeholder_ref] = node
    return lookup


def merge_geometry(
    node: etree._Element,
    inherited_nodes: list[tuple[etree._Element, str]],
) -> tuple[float, float, float, float, str | None]:
    x, y, width_pt, height_pt = extract_geometry(node)
    if width_pt > 0 and height_pt > 0:
        return x, y, width_pt, height_pt, None

    for inherited_node, source in inherited_nodes:
        inherited_x, inherited_y, inherited_width, inherited_height = extract_geometry(inherited_node)
        if inherited_width > 0 and inherited_height > 0:
            return inherited_x, inherited_y, inherited_width, inherited_height, source

    return x, y, width_pt, height_pt, None


def merge_font_size(node: etree._Element, inherited_nodes: list[tuple[etree._Element, str]]) -> tuple[float, str | None]:
    local_font = extract_font_size_pt(node, None)
    if local_font:
        return local_font, None

    for inherited_node, source in inherited_nodes:
        inherited_font = extract_font_size_pt(inherited_node, None)
        if inherited_font:
            return inherited_font, source

    return DEFAULT_FONT_PT, None


def build_placeholder(
    node: etree._Element,
    occurrences: dict[str, int],
    layout_lookup: dict[tuple[str, str], etree._Element],
    master_lookup: dict[tuple[str, str], etree._Element],
) -> dict:
    name, shape_id = extract_shape_name(node)
    occurrences[name] += 1
    occurrence = occurrences[name]
    kind = extract_placeholder_kind(node)
    placeholder_ref = extract_placeholder_ref(node)
    inherited_nodes: list[tuple[etree._Element, str]] = []
    if placeholder_ref:
        layout_match = layout_lookup.get(placeholder_ref)
        master_match = master_lookup.get(placeholder_ref)
        if layout_match is not None:
            inherited_nodes.append((layout_match, "layout"))
        if master_match is not None:
            inherited_nodes.append((master_match, "master"))

    x, y, width_pt, height_pt, geometry_source = merge_geometry(node, inherited_nodes)
    font_pt, font_source = merge_font_size(node, inherited_nodes)
    sample_text = extract_text(node)
    max_chars = None
    max_lines = None

    if kind == "TEXT":
        max_chars, max_lines = estimate_text_capacity(width_pt, height_pt, font_pt)

    return {
        "id": shape_id,
        "name": name,
        "occurrence": occurrence,
        "kind": kind,
        "sampleText": sample_text,
        "xPt": x,
        "yPt": y,
        "widthPt": width_pt,
        "heightPt": height_pt,
        "fontSizePt": font_pt,
        "maxChars": max_chars,
        "maxLines": max_lines,
        "placeholderType": placeholder_ref[0] if placeholder_ref else None,
        "placeholderIndex": placeholder_ref[1] if placeholder_ref else None,
        "inheritedGeometryFrom": geometry_source,
        "inheritedFontFrom": font_source,
    }


def resolve_part(base_part: str, target: str) -> str:
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def read_part_xml(archive: zipfile.ZipFile, part_name: str) -> etree._Element | None:
    try:
        return etree.fromstring(archive.read(part_name))
    except KeyError:
        return None


def load_relationship_targets(archive: zipfile.ZipFile, source_part: str) -> dict[str, str]:
    rels_part = posixpath.join(
        posixpath.dirname(source_part),
        "_rels",
        f"{posixpath.basename(source_part)}.rels",
    )
    try:
        root = etree.fromstring(archive.read(rels_part))
    except KeyError:
        return {}

    targets: dict[str, str] = {}
    for relationship in root.xpath("./pr:Relationship", namespaces=NS):
        rel_id = relationship.get("Id")
        target = relationship.get("Target")
        if rel_id and target:
            targets[rel_id] = resolve_part(source_part, target)
    return targets


def load_inheritance_roots(
    archive: zipfile.ZipFile,
    slide_part: str,
) -> tuple[etree._Element | None, etree._Element | None]:
    slide_relationships = load_relationship_targets(archive, slide_part)
    layout_part = next((target for target in slide_relationships.values() if "slideLayout" in target), None)
    if not layout_part:
        return None, None

    layout_root = read_part_xml(archive, layout_part)
    layout_relationships = load_relationship_targets(archive, layout_part)
    master_part = next((target for target in layout_relationships.values() if "slideMaster" in target), None)
    master_root = read_part_xml(archive, master_part) if master_part else None
    return layout_root, master_root


def scan_slide(archive: zipfile.ZipFile, part_name: str, xml_bytes: bytes, slide_number: int) -> dict:
    root = etree.fromstring(xml_bytes)
    layout_root, master_root = load_inheritance_roots(archive, part_name)
    layout_lookup = build_placeholder_lookup(layout_root)
    master_lookup = build_placeholder_lookup(master_root)
    occurrences: dict[str, int] = defaultdict(int)
    placeholders = []

    for node in iter_shape_nodes(root):
        placeholders.append(build_placeholder(node, occurrences, layout_lookup, master_lookup))

    text_count = sum(1 for item in placeholders if item["kind"] == "TEXT")
    picture_slots = sum(1 for item in placeholders if item["kind"] == "PICTURE")
    table_slots = sum(1 for item in placeholders if item["kind"] == "TABLE")
    chart_slots = sum(1 for item in placeholders if item["kind"] == "CHART")

    return {
        "sourceSlide": slide_number,
        "placeholderCount": len(placeholders),
        "layoutInherited": layout_root is not None,
        "masterInherited": master_root is not None,
        "placeholders": placeholders,
        "capacities": {
            "textCount": text_count,
            "pictureSlots": picture_slots,
            "tableSlots": table_slots,
            "chartSlots": chart_slots,
        },
    }


def scan_template(path: Path) -> dict:
    with zipfile.ZipFile(path) as archive:
        slide_files = sorted(
            [name for name in archive.namelist() if name.startswith("ppt/slides/slide") and name.endswith(".xml")],
            key=lambda item: int(Path(item).stem.replace("slide", "")),
        )

        slides = []
        for file_name in slide_files:
            slide_number = int(Path(file_name).stem.replace("slide", ""))
            slides.append(scan_slide(archive, file_name, archive.read(file_name), slide_number))

    return {
        "templatePath": str(path),
        "slideCount": len(slides),
        "slides": slides,
    }


def main() -> int:
    if len(sys.argv) != 2:
        print("Usage: scan_template_schema.py <template.pptx>", file=sys.stderr)
        return 1

    path = Path(sys.argv[1]).expanduser()
    if not path.exists():
        print(json.dumps({"error": f"Template not found: {path}"}))
        return 2

    result = scan_template(path)
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
