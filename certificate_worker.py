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

from pptx import Presentation


P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"

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
    found = set()
    for node in root.iter(f"{{{A}}}t"):
        value = node.text or ""
        for token, replacement in replacements.items():
            if token in value:
                node.text = value.replace(token, str(replacement))
                found.add(token)
    missing = set(replacements) - found
    if missing:
        raise RuntimeError("Certificate template token(s) missing: " + ", ".join(sorted(missing)))
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
