# Link Generator

## Descrizione

Il programma è stato realizzato dall'Avv. Roberto Alma, con il supporto dell'Avv. Matteo Moscioni, Francesco Saverio Orlando e Antonio Cocco.

Il programma analizza i file pdf contenuti nella cartella in cui viene eseguito e consente, in modo automatizzato:

1. la creazione dell'indice degli atti e dei documenti contenente la lista dei documenti pdf presenti nella cartella con impostati in automatico i collegamenti ipertestuali (ad esempio, il primo elemento dell'indice sarà costituito da 01_[nomeDoc].pdf e così via)
2. la modifica dell'atto processuale in modo da inserire i collegamenti ipertestuali in automatico ai documenti corretti (es. se il programma trova l'occorrenza '**doc. 1**' andrà ad applicare in automatico il collegamento ipertestuale al documento allegato 01_[nomeDoc].pdf)

> L'indice sarà salvato nella medesima cartella con il nome '**00_indice_atti_documenti.docx**'
> L'atto con i collegamenti ipertestuali sarà salvato nella medesima cartella con il nome '**new-atto.docx**'

## Prerequisiti

Per il corretto funzionamento del programma, è necessario osservare le seguenti semplici regole:

- i documenti devono essere salvati premettendo la numerazione progressiva '**01_**' al nome del file, in ossequio ai vigenti protocolli processuali (ciò è necessario per assicurare che i documenti siano ordinati in modo crescente)
- all'interno dell'atto le occorrenze dei documenti devono rispettare il seguente formato `doc. 1`;
- dopo che l'atto è stato finalizzato, è necessario salvare in grassetto ed applicare il colore blu a tutte le occorrenze contenenti `doc. 1` ecc. (è necessario per il buon funzionamento del ciclo di analisi dell'atto)
- è importante che l'atto sia salvato nella medesima cartella nella quale sono salvati gli allegati e sia rinominato come '**atto.docx**';
- è importante che nel nome degli allegati non ci siano lettere accentate (es. 18_attivit**à**.pdf) restituirà un errore, rinominare allegato ed eseguire di nuovo il programma;
- tra un'occorrenza (es. **doc.1**) e un'altra (es. **doc.2**) è necessario che si vada a capo (altrimenti **doc.2** viene spostato a fianco a **doc.1** e bisogna rispostarlo a mano).

## Esecuzione del programma

Prerequisiti:

- python 2.7+
- docx-python, installabile con il comando `pip install docx-python`

Per eseguire il programma digitare `python linkGenerator.py`

> N.B., è necessario che tutti i file, compresi quelli degli allegati e il file atto.docx siano copiati nella cartella dove ci sono tutti i file del programma.

> In alternativa è necessario installare `pyinstaller` e creare un eseguibile per il proprio sistema operativo con `pyinstaller --onefile linkGenerator.py` quindi copiare il file eseguibile nella cartella di lavoro dove sono contenuti tutti i file.


