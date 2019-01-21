# coding=utf-8
# importo gli oggetti "path" e "listdir" #
# "path" serve per gestire i percorsi #
# "listdir" serve per elencare i file contenuti in una cartella #
from os import path, listdir

# importo l\'oggetto Template che sarà utilizzato per generare un file html #
from jinja2 import Template

# salvo il percorso relativo alla directory nella quale viene eseguito lo script #
pathDir = path.dirname(path.abspath(__file__))

# ottengo la lista dei file contenuti nella directory
listFiles = listdir(pathDir)

# ordino la lista
listFiles.sort()

documents = []

# creo una classe "DOC", il metodo "__init__" è la funzione costruttrice #
# ogni oggetto della classe "DOC" ha una proprietà url e una proprietà txt #
class Doc(object):
    def __init__(self, url, txt):
        self.url = url
        self.txt = txt

# funzione che riceve come parametro una stringa e rimuove l'estensione ".pdf" #
def filterString(str):
    return str.split('.pdf')[0]

# ciclo che per tutti i file della directory, verifica se l'estensione è .pdf o meno #
# se l'estensione è pdf, crea un'istanza della classe Doc e passa come parametro "url" #
#il nome del file come estensione, e come "txt" il solo nome, senza estensione #
for file in listFiles:
    if file.endswith('.pdf'):
        name = filterString(file)
        documents.append(Doc(file, name))

# creo il template #
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

# salvo in una variabile html il contenuto del template, passando come oggetto file il contenuto dell'array documents
html = template.render(files=documents)

# creo un puntatore al file indice.html, se non esiste viene creato
index = open('Indice.html', 'a+')

# salvo il contenuto di html all'interno del file "Indice.html"
index.write(html)
