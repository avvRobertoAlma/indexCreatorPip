# coding=utf-8
# importo gli oggetti "path" e "listdir" #
# "path" serve per gestire i percorsi #
# "listdir" serve per elencare i file contenuti in una cartella #
from os import path, listdir

# importo la libreria docx-python
import docx

# importo le funzioni contenute nel file helpers.py
from helpers import filterString, Doc, Link, createLinkArray

# importo la funzione hyperlink
import hyperlink

# salvo il percorso relativo alla directory nella quale viene eseguito lo script #
pathDir = path.dirname(path.abspath(__file__))

print('Benvenuti in linkGenerator versione 1.0.0\n Il programma è realizzato dall\' Avv. Roberto Alma\n, con il supporto dell\' Avv. Matteo Moscioni, dell\'Avv. Francesco Saverio Orlando e dell\' Avv. Antonio Cocco')

# inizializzo il puntatore al documento originale
atto = docx.Document('atto.docx')

# ottengo la lista dei file contenuti nella directory
listFiles = listdir(pathDir)

# ordino la lista
listFiles.sort()

# inizializzo l'array dei documenti
documents = []

# ciclo che per tutti i file della directory, verifica se l'estensione è .pdf o meno #
# se l'estensione è pdf, crea un'istanza della classe Doc e passa come parametro "url" #
#il nome del file come estensione, e come "txt" il solo nome, senza estensione #
for file in listFiles:
    if file.endswith('.pdf'):
        name = filterString(file)
        documents.append(Doc(file, name))

# informo l'utente che i documenti sono stati correttamente registrati
print('Ho analizzato i documenti pdf')

#ottengo l'array dei link
linksArray = createLinkArray(documents)

#creazione dell'indice dei documenti
print('Inizio creazione dell\'indice dei documenti')

indice = docx.Document()
indice.add_heading('Indice atti e documenti')

# ciclo che per ogni link salvato aggiunge un elemento all'indice e aggiunge il link ipertestuale
for link in linksArray:
    p = indice.add_paragraph()
    p.style = 'List Number 2'
    new_run = p.add_run(link.name)
    hyperlink.add(p, new_run, link.url)

indice.save('00_indice_atti_documenti.docx')


# informo l'utente che inizio a scansionare l'atto
print('Sto iniziando la scansione dell\'atto\n Attendere prego..')

# scansione dell'atto e aggiunta degli hyperlinks
for p in atto.paragraphs:
    for run in p.runs:
        for link in linksArray:
            found_link = link.__eq__(run.text)
            if(found_link):
                hyperlink.add(p, run, found_link)

# informo che la scansione è completata
print('Scansione completata\n Salvataggio in corso..')

atto.save('new-atto.docx')

# informo che il file è stato salvato
print('Il file completo dei link ipertestuali è stato salvato come new-atto.docx\n Grazie per aver utilizzato il programma!')




