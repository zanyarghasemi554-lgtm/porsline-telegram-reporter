from xml.etree import ElementTree as ET

from certificate_worker import A, P, replace_slide_tokens


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
    student = {
        "honorific": "آقای", "persian_name": "علی 2", "national_id": "00123",
        "course_title": "دوره 4", "duration": "12 ساعت", "instructor": "مدرس 5",
        "venue": "سالن 6", "organization": "گروه 7",
    }
    root = ET.fromstring(replace_slide_tokens(slide, student))
    text = "".join(node.text or "" for node in root.iter(f"{{{A}}}t"))
    assert "00123" not in text
    assert "۰۰۱۲۳" in text
    paragraph_properties = next(root.iter(f"{{{A}}}pPr"))
    assert paragraph_properties.get("algn") == "r"
    assert paragraph_properties.get("rtl") == "1"
    assert next(root.iter(f"{{{A}}}bodyPr")).get("rtlCol") == "1"
