import io
import os
import zipfile
import xml.etree.ElementTree as ET
from pptx import Presentation
from utils import remove_all_slides

_APP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
_VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

ET.register_namespace("ap", _APP_NS)
ET.register_namespace("vt", _VT_NS)


def _fix_app_xml(pptx_path: str, slides) -> None:
    """Rewrite docProps/app.xml so slide count metadata matches actual content.

    `slides` can be either:
    - a list of custom Slide dataclass objects (with .content.title), or
    - a pptx SlideParts collection (with .shapes.title.text).
    """
    slide_count = len(slides)
    slide_titles = []
    for s in slides:
        if hasattr(s, 'content'):
            slide_titles.append(s.content.title)
        else:
            # pptx Slide object
            slide_titles.append(s.shapes.title.text if s.shapes.title else '')

    with zipfile.ZipFile(pptx_path, "r") as zin:
        names = zin.namelist()
        files = {name: zin.read(name) for name in names}

    if "docProps/app.xml" not in files:
        return

    tree = ET.fromstring(files["docProps/app.xml"])
    ns = {"ap": _APP_NS, "vt": _VT_NS}

    def _set_text(tag, value):
        el = tree.find(f"ap:{tag}", ns)
        if el is not None:
            el.text = str(value)

    _set_text("Slides", slide_count)
    _set_text("Notes", 0)

    # Update HeadingPairs — find slide count variant and update it
    hp = tree.find("ap:HeadingPairs/vt:vector", ns)
    if hp is not None:
        variants = hp.findall("vt:variant", ns)
        for i in range(0, len(variants) - 1, 2):
            label_el = variants[i].find("vt:lpstr", ns)
            count_el = variants[i + 1].find("vt:i4", ns)
            if label_el is not None and label_el.text in ("Slide Titles", "Slides"):
                if count_el is not None:
                    count_el.text = str(slide_count)

    # Rebuild TitlesOfParts — keep non-slide entries, replace slide titles
    top = tree.find("ap:TitlesOfParts/vt:vector", ns)
    if top is not None:
        existing = top.findall("vt:lpstr", ns)

        # Count non-slide entries from HeadingPairs
        non_slide_count = 0
        if hp is not None:
            pairs = hp.findall("vt:variant", ns)
            for i in range(0, len(pairs) - 1, 2):
                label_el = pairs[i].find("vt:lpstr", ns)
                count_el = pairs[i + 1].find("vt:i4", ns)
                if label_el is not None and label_el.text not in ("Slide Titles", "Slides"):
                    if count_el is not None and count_el.text is not None:
                        non_slide_count += int(count_el.text)

        # Remove all lpstr children then re-add kept + new slide titles
        for el in existing:
            top.remove(el)
        for el in existing[:non_slide_count]:
            top.append(el)
        for title in slide_titles:
            lpstr = ET.SubElement(top, f"{{{_VT_NS}}}lpstr")
            lpstr.text = title
        top.set("size", str(non_slide_count + slide_count))

    xml_bytes = ET.tostring(tree, encoding="unicode").encode("utf-8")
    xml_bytes = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + xml_bytes
    files["docProps/app.xml"] = xml_bytes

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, files[name])

    with open(pptx_path, "wb") as f:
        f.write(buf.getvalue())


def generate_presentation(powerpoint_data, template_path: str, output_path: str):
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template file '{template_path}' does not exist.")

    prs = Presentation(template_path)
    remove_all_slides(prs)
    prs.core_properties.title = powerpoint_data.title

    for slide in powerpoint_data.slides:
        if slide.layout >= len(prs.slide_layouts):
            slide_layout = prs.slide_layouts[0]
        else:
            slide_layout = prs.slide_layouts[slide.layout]

        new_slide = prs.slides.add_slide(slide_layout)

        if new_slide.shapes.title:
            new_slide.shapes.title.text = slide.content.title

        for shape in new_slide.shapes:
            if shape.has_text_frame and not shape == new_slide.shapes.title:
                text_frame = shape.text_frame
                text_frame.clear()
                for point in slide.content.bullet_points:
                    p = text_frame.add_paragraph()
                    p.text = point
                    p.level = 0
                break

        if slide.content.image_path:
            image_full_path = os.path.join(os.getcwd(), slide.content.image_path)
            if os.path.exists(image_full_path):
                for shape in new_slide.placeholders:
                    if shape.placeholder_format.type == 18:
                        shape.insert_picture(image_full_path)
                        break

    prs.save(output_path)
    _fix_app_xml(output_path, powerpoint_data.slides)
    print(f"Presentation saved to '{output_path}'")
