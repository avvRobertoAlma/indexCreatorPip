# coding=utf-8

import docx
import hyperlink
import helpers

doc = docx.Document('atto.docx')

paragraphs = doc.paragraphs

txt = 'doc. 1'

for p in paragraphs:
    runs = p.runs

    for run in runs:
        print run.text
        if txt in run.text:
            hyperlink.add(p, run, 'https://github.com')
            run.font.color.rgb = docx.shared.RGBColor(0, 0, 255)
            

doc.save('new-atto.docx')

    