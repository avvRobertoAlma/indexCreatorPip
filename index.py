from os import path, listdir
from jinja2 import Template

pathDir = path.dirname(path.abspath(__file__))

listFiles = listdir(pathDir)

documents = []

class Doc(object):
    def __init__(self, url, txt):
        self.url = url
        self.txt = txt

def filterString(str):
    return str.split('.pdf')[0]

for file in listFiles:
    if file.endswith('.pdf'):
        name = filterString(file)
        documents.append(Doc(file, name))

template = Template("""<style>
.centered {
    text-align: center;
}
</style>
<div>
    <h1 class="centered">Indice atti e documenti</h1>
    <p></p>
    <p></p>
    <ol>
    {% for file in files %}
    <li><a href='./{{file.url}}'>{{file.txt}}</a></li>
    {% endfor %}
</div>""")

html = template.render(files=documents)

index = open('Indice.html', 'a+')

index.write(html)
