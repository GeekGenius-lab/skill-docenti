---
name: verifica-dal-modello
description: Prepara una verifica scritta in Word partendo dal modello di verifica della scuola (file .docx fornito dal docente), solo sugli argomenti che il docente dichiara svolti in classe, con uno stile semplice e umano, poi la controlla due volte prima di consegnarla. Usala quando un insegnante chiede di preparare, scrivere o impaginare una verifica, un compito in classe o una prova scritta "sul modello della scuola", "sul format d'istituto" o allegando il modello.
---

# Verifica dal modello della scuola

Prepari una verifica scritta per un insegnante. Il risultato è un file Word costruito sul modello della sua scuola, con le soluzioni per il docente in un file a parte.

## Regole che non si cambiano

1. **Il modello resta intatto.** Intestazione, loghi, tabelle, griglie di valutazione, rubriche e testi fissi non si toccano. Riempi solo gli spazi da compilare: segnaposto tra parentesi quadre, righe `____`, celle vuote da completare. L'unica eccezione sono i trattini lunghi e medi (— e –): li sostituisci in tutto il documento, modello compreso (vedi `riferimenti/stile.md`).
2. **Solo argomenti svolti in classe.** Scrivi esercizi solo sugli argomenti che il docente dichiara di aver fatto. Se non li ha detti, chiediglieli prima di scrivere qualsiasi cosa. Non aggiungere argomenti "vicini" o "di ripasso" che lui non ha nominato.
3. **Niente dati degli alunni.** Non chiedere nomi, diagnosi, certificazioni (BES, DSA, PEI, PDP), indirizzi o mail. Se il docente li scrive lo stesso, non riportarli nel file e non ripeterli nella risposta. Per le versioni adattate bastano misure anonime, per esempio "versione con meno esercizi" o "tempo più lungo del 30%". Il resto lo gestisce il docente sul suo computer.
4. **Nessun riferimento personale nel file**, se non quelli che il modello già contiene o che il docente chiede (nome dell'istituto, materia, classe).

## Passi

### 1. Raccogli le informazioni

Prima di scrivere ti servono queste cose. Chiedi solo quelle che mancano, tutte insieme in un unico messaggio:

- **il modello della scuola** (.docx). Se il docente non lo allega, guarda in `modelli/` di questa skill. Se anche lì non c'è, chiedi se vuole caricarlo o se preferisce una verifica senza modello (struttura semplice: intestazione, istruzioni, esercizi con punteggio, griglia di valutazione);
- **materia, classe e tipo di scuola** (per il livello delle domande);
- **gli argomenti svolti in classe**, uno per uno: è il punto più importante;
- **durata** della prova e **materiali ammessi** (calcolatrice, formulario, manuale);
- **numero e tipo di esercizi**, se il modello non li stabilisce già;
- **versioni adattate** che servono, solo come misure anonime (per esempio "una versione con obiettivi minimi");
- **istituto**, solo se il modello ha un campo da compilare e non è già scritto.

Se il docente dice "fai tu" su qualcosa che non sono gli argomenti, scegli tu in modo ragionevole e scrivilo nel riepilogo finale. Sugli argomenti non si sceglie al posto suo.

### 2. Leggi il modello

Esegui `scripts/segnaposto.py MODELLO.docx` (se puoi eseguire Python) per vedere la struttura: paragrafi, tabelle, celle, segnaposto da compilare e trattini da sostituire. Se non puoi eseguire script, apri il documento e fai lo stesso elenco a mano.

Capisci cosa chiede ogni sezione del modello: punteggi, tipi di domanda, livelli di competenza, rubriche. La verifica deve rispettare quella struttura e quei punteggi. Se una sezione non si adatta agli argomenti svolti (per esempio un "compito di realtà" su argomenti solo teorici), proponi al docente come riempirla. Non toglierla.

### 3. Scrivi le domande

Segui `riferimenti/stile.md`. In breve:

- frasi corte, una richiesta per frase, del tu ("Calcola", "Spiega", "Indica");
- esempi concreti, da banco o da laboratorio, legati alla vita reale degli studenti;
- dati numerici realistici, risultati che vengono "puliti" quando il calcolo è a mano;
- ogni domanda si risolve con quello che è stato fatto in classe;
- nessuna domanda suggerisce la risposta di un'altra;
- nelle risposte multiple, la risposta giusta non è sempre nella stessa posizione e non è sempre la più lunga;
- punteggi scritti accanto a ogni esercizio, con la somma uguale al totale del modello.

### 4. Compila il file

Lavora su una copia del modello, mai sull'originale. Se puoi eseguire Python, usa `python-docx`: sostituisci il testo dei segnaposto dentro i run esistenti, così il formato resta quello del modello. Per i nuovi paragrafi copia lo stile di un paragrafo vicino. Non ricreare il documento da zero e non cambiare margini, caratteri o tabelle.

Prepara due file:

- `Verifica_<materia>_<classe>_<argomento>.docx`, il file da stampare;
- `Verifica_<materia>_<classe>_<argomento>_SOLUZIONI.docx`, per il docente: soluzioni svolte, punteggio per ogni passaggio, errori tipici da guardare.

Per ogni versione adattata fai un file a parte, con lo stesso nome più `_versione_ridotta` (o il nome che dà il docente). La versione adattata deve sembrare una verifica normale: niente scritte come "per alunni DSA" sul foglio dello studente.

Se non puoi eseguire codice (per esempio su claude.ai senza strumenti di analisi), scrivi il testo completo, sezione per sezione, indicando in quale punto del modello va ogni parte, e spiega al docente che deve incollarlo nel modello.

### 5. Primo controllo: automatico

Esegui `scripts/controlla_docx.py --modello MODELLO.docx VERIFICA.docx [altri file]`. Lo script controlla che:

- il file si apra;
- non ci siano testi in caratteri simbolo (Wingdings, Symbol), caratteri rotti o invisibili;
- non ci siano trattini lunghi o medi;
- non siano rimasti segnaposto da compilare ("[Inserire ...]");
- le tabelle non abbiano celle spostate;
- non ci siano paragrafi ripetuti per errore;
- i testi fissi del modello siano ancora tutti presenti.

Correggi ogni problema e riesegui finché esce `OK`. Le caselle da spuntare del modello (☐, spesso in Wingdings) vanno bene: se lo script le segnala, controlla che il carattere simbolo sia solo sulle caselle e non sul testo.

Poi somma i punteggi a mano, sezione per sezione, e controlla che tornino con i totali del modello.

### 6. Secondo controllo: una ricontrollata separata

Rileggi la verifica **come se non l'avessi scritta tu**, con la lista di `riferimenti/controllo.md`.

- In Claude Code, o dove puoi avviare un sotto-agente, affida la ricontrollata a un agente separato. Dagli solo i file (modello, verifica, soluzioni) e l'elenco degli argomenti svolti, non le tue intenzioni. Chiedigli di risolvere la verifica da studente e di segnalare errori, ambiguità e domande fuori programma.
- Altrimenti fallo tu, in un passaggio a parte: risolvi ogni esercizio da zero senza guardare le soluzioni, poi confronta.

Correggi quello che trovi. Se hai corretto qualcosa, rifai il primo controllo.

### 7. Consegna

Consegna i file e un riepilogo breve, in italiano semplice:

- argomenti coperti, esercizio per esercizio;
- punteggi e durata;
- cosa hai scelto tu, dove il docente ha detto "fai tu";
- esito dei due controlli;
- cosa deve controllare il docente prima di stampare (per esempio una figura da aggiungere a mano).

Non dire che la verifica è "perfetta": il docente la rilegge comunque.
