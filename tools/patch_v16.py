from docx import Document
from docx.oxml.ns import qn
import os

doc = Document("output/Hengshan_Report_v15.docx")
body = doc.element.body

removed = 0
for p in list(body.findall(qn("w:p"))):
    texts = [t.text or "" for t in p.findall(".//" + qn("w:t"))]
    all_text = "".join(texts).strip()
    has_drawing = len(p.findall(".//" + qn("w:drawing"))) > 0
    if not all_text and not has_drawing:
        body.remove(p)
        removed += 1

prev_was_br = False
for p in list(body.findall(qn("w:p"))):
    has_br = len(p.findall(".//" + qn("w:br"))) > 0
    texts = [t.text or "" for t in p.findall(".//" + qn("w:t"))]
    all_text = "".join(texts).strip()
    if has_br and not all_text:
        if prev_was_br:
            body.remove(p)
            removed += 1
        else:
            prev_was_br = True
    elif all_text:
        prev_was_br = False

doc.save("output/Hengshan_Report_v16.docx")
print("Removed " + str(removed) + " blank paragraphs")
print("v16: " + str(os.path.getsize("output/Hengshan_Report_v16.docx") // 1024) + " KB")
