"""Build certificate PPTX files by editing the supplied OOXML template in place.

No office converter is involved: slide geometry, text-run formatting and all decorative
objects remain byte-for-byte based on the user's template.
"""

import json
import re
import sys
import zipfile
from copy import deepcopy
from pathlib import Path
from xml.etree import ElementTree as ET

from PIL import ImageFont
from pptx import Presentation


P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
PERSIAN_DIGITS = str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")
EMU_PER_PIXEL_96_DPI = 9525
DEFAULT_TEXT_INSET_EMU = 91425
MAIN_TEXT_LINE_GAP_EMU = 120000
UIGHUR_FONT_PATH = Path(__file__).resolve().parent / "assets" / "fonts" / "MSUIGHUR.TTF"

for prefix, uri in (("p", P), ("a", A), ("r", R)):
    ET.register_namespace(prefix, uri)
ET.register_namespace("", PR)


def xml_bytes(root):
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def next_rid(root):
    values = []
    for relation in root.findall(f"{{{PR}}}Relationship"):
        match = re.fullmatch(r"rId(\d+)", relation.get("Id", ""))
        if match:
            values.append(int(match.group(1)))
    return f"rId{max(values, default=0) + 1}"


def shape_text(shape):
    return "".join(node.text or "" for node in shape.iter(f"{{{A}}}t"))


def wrapped_line_count(text, width_emu, font_size_pt=24):
    """Estimate PowerPoint wrapping using the exact certificate font at 96 DPI."""
    usable_pixels = max(1, (width_emu - 2 * DEFAULT_TEXT_INSET_EMU) / EMU_PER_PIXEL_96_DPI)
    font_pixels = max(1, round(font_size_pt * 96 / 72))
    font = ImageFont.truetype(str(UIGHUR_FONT_PATH), font_pixels)
    lines = 1
    current_width = 0.0
    for word in str(text).split():
        word_width = font.getlength(word, direction="rtl")
        space_width = font.getlength(" ", direction="rtl") if current_width else 0.0
        if current_width and current_width + space_width + word_width > usable_pixels:
            lines += 1
            current_width = word_width
        else:
            current_width += space_width + word_width
    return max(1, lines)


def position_main_separator(root, main_shape):
    """Keep the pale gray separator a consistent gap below the last body line."""
    transform = main_shape.find(f"{{{P}}}spPr/{{{A}}}xfrm")
    offset = transform.find(f"{{{A}}}off") if transform is not None else None
    extent = transform.find(f"{{{A}}}ext") if transform is not None else None
    if offset is None or extent is None:
        return
    line_count = wrapped_line_count(shape_text(main_shape), int(extent.get("cx", "0")))
    font = ImageFont.truetype(str(UIGHUR_FONT_PATH), round(24 * 96 / 72))
    ascent, descent = font.getmetrics()
    line_height_emu = round((ascent + descent) * 1.08 * EMU_PER_PIXEL_96_DPI)
    separator_y = (
        int(offset.get("y", "0")) + DEFAULT_TEXT_INSET_EMU
        + line_count * line_height_emu + MAIN_TEXT_LINE_GAP_EMU
    )
    for connector in root.iter(f"{{{P}}}cxnSp"):
        gray = connector.find(f".//{{{A}}}solidFill/{{{A}}}schemeClr[@val='bg2']")
        connector_offset = connector.find(f"{{{P}}}spPr/{{{A}}}xfrm/{{{A}}}off")
        connector_extent = connector.find(f"{{{P}}}spPr/{{{A}}}xfrm/{{{A}}}ext")
        if gray is not None and connector_offset is not None and connector_extent is not None \
                and connector_extent.get("cy") == "0":
            connector_offset.set("y", str(separator_y))
            break


def replace_slide_tokens(slide_bytes, student):
    replacements = {
        "{{HONORIFIC}}": student["honorific"],
        "{{STUDENT_NAME}}": student["persian_name"],
        "{{NATIONAL_ID}}": student["national_id"],
        "{{COURSE_TITLE}}": student["course_title"],
        "{{DURATION}}": student["duration"],
        "{{INSTRUCTOR}}": student["instructor"],
        "{{VENUE}}": student["venue"],
        "{{ORGANIZATION}}": student["organization"],
    }
    root = ET.fromstring(slide_bytes)
    for shape_tree in root.iter(f"{{{P}}}spTree"):
        for picture in list(shape_tree.findall(f"{{{P}}}pic")):
            shape_tree.remove(picture)
    main_shape = None
    centered_organization_shape = None
    for shape in root.iter(f"{{{P}}}sp"):
        original_text = shape_text(shape)
        if "{{NATIONAL_ID}}" in original_text and "{{COURSE_TITLE}}" in original_text:
            main_shape = shape
            text_nodes = list(shape.iter(f"{{{A}}}t"))
            for index, node in enumerate(text_nodes[:-1]):
                if "{{NATIONAL_ID}}" in (node.text or "") and not (text_nodes[index + 1].text or "").strip():
                    # A non-breaking space survives RTL run boundaries in PowerPoint.
                    text_nodes[index + 1].text = "\u00a0"
        if original_text.strip() == "{{ORGANIZATION}}":
            centered_organization_shape = shape

    found = set()
    for node in root.iter(f"{{{A}}}t"):
        value = node.text or ""
        for token, replacement in replacements.items():
            if token in value:
                value = value.replace(token, str(replacement))
                found.add(token)
        # Convert both inserted values and any fixed template numbers.
        node.text = value.translate(PERSIAN_DIGITS)
    for paragraph in root.iter(f"{{{A}}}p"):
        paragraph_properties = paragraph.find(f"{{{A}}}pPr")
        if paragraph_properties is None:
            paragraph_properties = ET.Element(f"{{{A}}}pPr")
            paragraph.insert(0, paragraph_properties)
        paragraph_properties.set("algn", "r")
        paragraph_properties.set("rtl", "1")
    if centered_organization_shape is not None:
        for paragraph in centered_organization_shape.iter(f"{{{A}}}p"):
            paragraph_properties = paragraph.find(f"{{{A}}}pPr")
            if paragraph_properties is not None:
                paragraph_properties.set("algn", "ctr")
                paragraph_properties.set("rtl", "1")
    for body_properties in root.iter(f"{{{A}}}bodyPr"):
        body_properties.set("rtlCol", "1")
    missing = set(replacements) - found
    if missing:
        raise RuntimeError("Certificate template token(s) missing: " + ", ".join(sorted(missing)))
    if main_shape is not None:
        position_main_separator(root, main_shape)
    return xml_bytes(root)


def slide_relationships(source_bytes, include_notes=True):
    root = ET.fromstring(source_bytes)
    if not include_notes:
        for relation in list(root.findall(f"{{{PR}}}Relationship")):
            if relation.get("Type", "").endswith("/notesSlide"):
                root.remove(relation)
    for relation in list(root.findall(f"{{{PR}}}Relationship")):
        if relation.get("Type", "").endswith("/image"):
            root.remove(relation)
    return xml_bytes(root)


def build(config_path):
    config = json.loads(Path(config_path).read_text(encoding="utf-8"))
    students = config["students"]
    if not students:
        raise ValueError("At least one student is required")

    template_path = Path(config["template_path"])
    output_path = Path(config["output_path"])
    with zipfile.ZipFile(template_path, "r") as source:
        files = {name: source.read(name) for name in source.namelist()}

    slide_source = files["ppt/slides/slide1.xml"]
    rel_source = files["ppt/slides/_rels/slide1.xml.rels"]
    presentation = ET.fromstring(files["ppt/presentation.xml"])
    presentation_rels = ET.fromstring(files["ppt/_rels/presentation.xml.rels"])
    content_types = ET.fromstring(files["[Content_Types].xml"])
    slide_list = presentation.find(f"{{{P}}}sldIdLst")
    if slide_list is None or not list(slide_list):
        raise RuntimeError("Certificate template has no source slide")

    source_slide_id = list(slide_list)[0]
    for extra in list(slide_list)[1:]:
        slide_list.remove(extra)
    existing_ids = [int(item.get("id", "255")) for item in slide_list]
    source_rid = source_slide_id.get(f"{{{R}}}id")

    for index, student in enumerate(students, start=1):
        slide_name = f"ppt/slides/slide{index}.xml"
        rel_name = f"ppt/slides/_rels/slide{index}.xml.rels"
        files[slide_name] = replace_slide_tokens(slide_source, student)
        files[rel_name] = slide_relationships(rel_source, include_notes=index == 1)
        if index == 1:
            continue
        relation_id = next_rid(presentation_rels)
        ET.SubElement(presentation_rels, f"{{{PR}}}Relationship", {
            "Id": relation_id,
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
            "Target": f"slides/slide{index}.xml",
        })
        slide_id = deepcopy(source_slide_id)
        slide_id.set("id", str(max(existing_ids) + 1))
        slide_id.set(f"{{{R}}}id", relation_id)
        existing_ids.append(int(slide_id.get("id")))
        slide_list.append(slide_id)
        ET.SubElement(content_types, f"{{{CT}}}Override", {
            "PartName": f"/ppt/slides/slide{index}.xml",
            "ContentType": "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
        })

    # Keep the source slide relationship stable; only newly duplicated slides get new rIds.
    if not source_rid:
        raise RuntimeError("Source slide relationship is missing")
    files["ppt/presentation.xml"] = xml_bytes(presentation)
    files["ppt/_rels/presentation.xml.rels"] = xml_bytes(presentation_rels)
    files["[Content_Types].xml"] = xml_bytes(content_types)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as target:
        for name, payload in files.items():
            target.writestr(name, payload)
    # Normalize OPC relationship targets/content types for maximum PowerPoint compatibility.
    # python-pptx preserves slide geometry and run formatting while repairing package metadata.
    Presentation(output_path).save(output_path)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("Usage: certificate_worker.py CONFIG.json")
    build(sys.argv[1])
