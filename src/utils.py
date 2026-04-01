from pptx import Presentation
from pptx.oxml.ns import qn

def remove_all_slides(prs: Presentation):
    xml_slides = prs.slides._sldIdLst
    slide_ids = list(xml_slides)
    for slide_id in slide_ids:
        rId = slide_id.get(qn('r:id'))
        prs.part.drop_rel(rId)
        xml_slides.remove(slide_id)
    print("所有默认幻灯片已被移除。")
