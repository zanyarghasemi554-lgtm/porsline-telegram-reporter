import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from certificate_worker import A, P, replace_slide_tokens


PROJECT_DIR = Path(__file__).resolve().parents[1]


def student(course_title="کمک‌های اولیه"):
    return {
        "honorific": "آقای", "persian_name": "علی 2", "national_id": "00123",
        "course_title": course_title, "duration": "12 ساعت", "instructor": "مدرس 5",
        "venue": "سالن 6", "organization": "گروه آموزشی 7",
    }


def test_slide_text_is_rtl_right_aligned_and_digits_are_persian():
    tokens = (
        "{{HONORIFIC}}|{{STUDENT_NAME}}|{{NATIONAL_ID}}|{{COURSE_TITLE}}|"
        "{{DURATION}}|{{INSTRUCTOR}}|{{VENUE}}|{{ORGANIZATION}}"
    )
    slide = (
        f'<p:sld xmlns:p="{P}" xmlns:a="{A}"><p:cSld><p:spTree><p:sp>'
        f'<p:txBody><a:bodyPr/><a:p><a:r><a:t>{tokens}</a:t></a:r></a:p>'
        f'</p:txBody></p:sp></p:spTree></p:cSld></p:sld>'
    ).encode()
    root = ET.fromstring(replace_slide_tokens(slide, student("دوره 4")))
    text = "".join(node.text or "" for node in root.iter(f"{{{A}}}t"))
    assert "00123" not in text
    assert "۰۰۱۲۳" in text
    paragraph_properties = next(root.iter(f"{{{A}}}pPr"))
    assert paragraph_properties.get("algn") == "r"
    assert paragraph_properties.get("rtl") == "1"
    assert next(root.iter(f"{{{A}}}bodyPr")).get("rtlCol") == "1"


def test_template_spacing_centered_organization_and_dynamic_separator():
    with zipfile.ZipFile(PROJECT_DIR / "assets" / "certificate-template.pptx") as archive:
        source = archive.read("ppt/slides/slide1.xml")
    short_root = ET.fromstring(replace_slide_tokens(source, student()))
    long_root = ET.fromstring(replace_slide_tokens(source, student("آموزش تخصصی " * 28)))

    main_text = next(
        "".join(node.text or "" for node in shape.iter(f"{{{A}}}t"))
        for shape in short_root.iter(f"{{{P}}}sp") if "۰۰۱۲۳" in "".join(
            node.text or "" for node in shape.iter(f"{{{A}}}t")
        )
    )
    assert "۰۰۱۲۳\u00a0دوره" in main_text

    organization_paragraph = next(
        shape.find(f".//{{{A}}}p") for shape in short_root.iter(f"{{{P}}}sp")
        if "".join(node.text or "" for node in shape.iter(f"{{{A}}}t")).strip() == "گروه آموزشی ۷"
    )
    properties = organization_paragraph.find(f"{{{A}}}pPr")
    assert properties.get("algn") == "ctr"
    assert properties.get("rtl") == "1"

    def gray_line_y(root):
        connector = next(
            item for item in root.iter(f"{{{P}}}cxnSp")
            if item.find(f".//{{{A}}}solidFill/{{{A}}}schemeClr[@val='bg2']") is not None
        )
        return int(connector.find(f"{{{P}}}spPr/{{{A}}}xfrm/{{{A}}}off").get("y"))

    assert gray_line_y(long_root) > gray_line_y(short_root)
