import docx
import hyperlink

document = docx.Document('test.docx')

allegati = ['https://github.com', 'https://www.iusinaction.com', 'https://www.laroma24.it']

class Link(object):
    def __init__(self, txt, url):
        self.txt = txt
        self.url = url
    def __eq__(self, cmp):
        if self.txt == cmp:
            return self.url

links = []

for i, allegato in enumerate(allegati):
    i += 1
    text = 'doc. ' +str(i)
    links.append(Link(text, allegato))


for p in document.paragraphs:
    for run in p.runs:
        for link in links:
            found_link = link.__eq__(run.text)
            if(found_link):
                hyperlink.add(p, run, found_link)



document.save('new-text.docx')





    

