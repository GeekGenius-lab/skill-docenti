# Skill per docenti

Skill per Claude che aiutano gli insegnanti nel lavoro di tutti i giorni. Per ora ce n'è una:

| Skill | Cosa fa |
|---|---|
| [verifica-dal-modello](verifica-dal-modello/SKILL.md) | Prepara una verifica scritta in Word sul modello della tua scuola, solo sugli argomenti che hai fatto in classe, e la controlla due volte prima di dartela. |

## verifica-dal-modello

### Cosa fa

1. Ti chiede quello che serve: il modello della scuola, la materia, la classe, gli **argomenti svolti in classe**, la durata e i materiali ammessi. Senza gli argomenti non scrive niente.
2. Riempie il modello senza cambiarlo. Tocca solo gli spazi da compilare e sostituisce i trattini lunghi con la punteggiatura normale.
3. Scrive domande semplici: frasi corte, del tu, esempi concreti.
4. Ti consegna la verifica e un file a parte con le soluzioni. Se le chiedi, prepara anche le versioni adattate (per esempio con meno esercizi).
5. Prima di consegnare fa due controlli. Il primo è automatico: caratteri, tabelle, testi del modello. Il secondo è una ricontrollata separata: risolve la verifica da studente e cerca errori e domande fuori programma.

### Cosa non fa

- Non chiede e non scrive dati degli studenti: nomi, diagnosi, PDP, PEI. Per le versioni adattate bastano indicazioni anonime, come "una versione con meno esercizi".
- Non inventa argomenti: usa solo quelli che gli dici.
- Non manda e non carica niente. I file li scarichi tu.

## Installazione

### Su claude.ai (sito o app)

Serve un account con le skill attive: oggi richiedono un piano a pagamento. Nei piani Team ed Enterprise potrebbe doverle attivare l'amministratore: da verificare con chi gestisce l'account della scuola.

1. Scarica questo repository: pulsante verde **Code**, poi **Download ZIP**. Estrai la cartella sul tuo computer.
2. Facoltativo: metti il modello di verifica della tua scuola (.docx) nella cartella `verifica-dal-modello/modelli/`. Così non devi allegarlo ogni volta.
3. Crea lo zip della skill. Se hai Python: `python strumenti/crea_zip.py` (su Mac `python3`, su Windows `py`). Senza Python: clic destro sulla cartella `verifica-dal-modello` e poi "Comprimi" (Windows: "Invia a > Cartella compressa"). Lo zip deve contenere la cartella `verifica-dal-modello` con dentro `SKILL.md`.
4. Su claude.ai apri **Impostazioni > Funzionalità** e attiva **Esecuzione del codice e creazione di file**. Serve per creare il file Word e per i controlli automatici.
5. Nella stessa pagina, alla voce **Skill**, scegli **Carica skill** e seleziona lo zip.
6. In una nuova chat scrivi per esempio: "Prepara una verifica di matematica per la 2B sul modello della scuola" e allega il modello. Claude userà la skill da solo.

I nomi delle voci di menu possono cambiare con gli aggiornamenti di claude.ai.

### Su Claude Code

1. Copia la cartella `verifica-dal-modello` nella cartella delle skill:
   - per tutti i tuoi progetti: `C:\Users\<tuo nome>\.claude\skills\` su Windows, `~/.claude/skills/` su Mac e Linux (è una cartella nascosta: nel Finder premi Cmd+Maiusc+punto per vederla);
   - solo per un progetto: `.claude/skills/` dentro la cartella del progetto.
2. Installa la libreria per i file Word: `python3 -m pip install python-docx` (su Windows `py -m pip install python-docx`). Se compare l'errore "externally-managed-environment", chiedi a Claude Code di creare un ambiente virtuale.
3. Chiudi e riapri Claude Code.

## Dove aggiornare

| Cosa vuoi cambiare | File |
|---|---|
| Il modello della scuola | metti il nuovo .docx in `verifica-dal-modello/modelli/` e togli quello vecchio |
| Lo stile delle domande (lunghezza delle frasi, punteggiatura, versioni adattate) | `verifica-dal-modello/riferimenti/stile.md` |
| La lista della ricontrollata finale | `verifica-dal-modello/riferimenti/controllo.md` |
| I passi della skill e le regole fisse | `verifica-dal-modello/SKILL.md` |
| I controlli automatici (caratteri, trattini, tabelle, testi del modello) | `verifica-dal-modello/scripts/controlla_docx.py` |

Su claude.ai, dopo ogni modifica ricrea lo zip e caricalo di nuovo. Prima togli la versione vecchia della skill.

Per sapere quale versione hai: apri `verifica-dal-modello/SKILL.md` e guarda in alto le righe `version` e `aggiornata` (data dell'ultima modifica). Confrontale con quelle del file nel repository.

## Gli script

Puoi usarli anche da soli, senza Claude:

```
python verifica-dal-modello/scripts/segnaposto.py modello.docx
python verifica-dal-modello/scripts/controlla_docx.py --modello modello.docx verifica.docx
```

- `segnaposto.py` mostra la struttura del modello: paragrafi, tabelle, spazi da compilare, trattini da sostituire.
- `controlla_docx.py` controlla la verifica finita: caratteri rotti o in font simbolo, trattini lunghi, segnaposto non compilati, tabelle spostate, testi ripetuti, testi del modello mancanti, nome dell'autore rimasto nelle proprietà del file. Esce con `OK` o con l'elenco dei problemi. Non controlla figure, formule e somma dei punteggi: quelli si guardano a occhio.

## Privacy

La skill lavora solo sul modello e sugli argomenti che le dai. Prima di mettere il modello in `modelli/` o di condividerlo, guarda le proprietà del file (in Word: File > Informazioni): autore e "ultima modifica di" possono contenere nomi di colleghi. La skill li svuota nelle verifiche che prepara. Non mettere nei modelli, negli esempi o nelle richieste nomi di studenti o informazioni sulla loro salute. Se ti serve una versione adattata, descrivi la misura ("meno esercizi", "tempo più lungo"), non lo studente.

## Licenza

Queste skill sono distribuite con licenza [Creative Commons Attribuzione 4.0 Internazionale (CC BY 4.0)](https://creativecommons.org/licenses/by/4.0/deed.it). Puoi usarle, copiarle e modificarle, anche per altre scuole, citando la fonte: "progetto skill-docenti", e dicendo se hai fatto modifiche. Il testo completo è nel file `LICENSE`, che trovi anche dentro la cartella `verifica-dal-modello` e quindi nello zip.
