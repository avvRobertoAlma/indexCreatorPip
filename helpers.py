# coding=utf-8

# verifica se siamo in produzione o in sviluppo
import sys
def checkProduction():
    if getattr(sys, 'frozen', False):
        return False
    else:
        return True


# funzione che riceve come parametro una stringa e rimuove l'estensione ".pdf" #
def filterString(str):
    return str.split('.pdf')[0]

# dichiarazione della classe "DOC", il metodo "__init__" è la funzione costruttrice #
# ogni oggetto della classe "DOC" ha una proprietà url e una proprietà txt #
class Doc(object):
    def __init__(self, url, txt):
        self.url = url
        self.txt = txt

# dichiarazione della classe Link
# la funzione costruttrice accetta tre parametri: il nome del file, l'url e l'index
# l'oggetto generato ha una proprietà nome che corrisponde al nome del file, l'url relativo, e il nome dell'occorrenza nell'atto
class Link(object):
    def __init__(self, name, url, index):
        self.name = name
        self.url = './' + url
        self.occurrence = 'doc. ' + str(index)
    def __eq__(self, cmp):
        if self.occurrence == cmp:
            return self.url

# funzione che restituisce un array di oggetti Link
def createLinkArray(docsArray):
    linkArray = []
    for index, doc in enumerate(docsArray):
        index += 1
        linkArray.append(Link(doc.txt, doc.url, index))
    return linkArray

