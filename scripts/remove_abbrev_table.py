# -*- coding: utf-8 -*-
"""
Remove the List of Abbreviations heading and table from Research_Proposal_EN.docx.
Inline first-use expansions (e.g. "Acute Febrile Illness (AFI)") are kept.
"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
from docx import Document
from docx.text.paragraph import Paragraph

FOLDER = r'G:\My Drive\Research\MORU'
DOCX   = os.path.join(FOLDER, 'Research_Proposal_EN.docx')
TMP    = os.path.join(FOLDER, 'Research_Proposal_EN_tmp.docx')

doc = Document(DOCX)
body = doc.element.body
children = list(body)

# Locate the abbreviations heading and table by scanning
to_remove = []
for i, child in enumerate(children):
    tag = child.tag.split('}')[-1]
    if tag == 'p':
        p = Paragraph(child, doc)
        txt = p.text.strip()
        if txt == 'List of Abbreviations':
            to_remove.append(child)
            # also remove up to 2 blank paragraphs immediately before it
            for j in range(i - 1, max(i - 3, -1), -1):
                prev = children[j]
                if prev.tag.split('}')[-1] == 'p':
                    prev_txt = Paragraph(prev, doc).text.strip()
                    if prev_txt == '':
                        to_remove.append(prev)
                    else:
                        break
                else:
                    break
    elif tag == 'tbl':
        from docx.table import Table
        t = Table(child, doc)
        if t.rows and t.rows[0].cells[0].text.strip() == 'Abbreviation':
            to_remove.append(child)

for elem in to_remove:
    body.remove(elem)
    print(f'Removed: {elem.tag.split("}")[-1]}'
          + (f' "{Paragraph(elem, doc).text.strip()[:50]}"'
             if elem.tag.split("}")[-1] == "p" else ' (abbreviations table)'))

doc.save(TMP)
os.replace(TMP, DOCX)
print('Done.')
