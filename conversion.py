# coding=utf-8
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

import mammoth

with open('atto.docx', 'rb') as docx_file:
    result = mammoth.convert_to_html(docx_file)
    html = result.value

obj = [
    {
        'key': 'doc. 1',
        'href': "<a href='https://www.github.com'>doc. 1</a>"
    },
    {
        'key': 'doc. 2',
        'href': "<a href='https://www.iusinaction.com'>doc. 2</a>"

    }
]

for index, object in enumerate(obj):
    if object['key'] in html:
        html = html.replace(object['key'], object['href'])

html_file = open('atto.html', 'w')
html_file.write(html)
html_file.close()