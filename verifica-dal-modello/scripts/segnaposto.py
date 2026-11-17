"""Mostra la struttura di un modello di verifica .docx, nell'ordine del documento:
paragrafi, tabelle, istruzioni da sostituire, righe ____ e trattini lunghi.

Uso:  python segnaposto.py MODELLO.docx
Serve: pip install python-docx

Etichette:
  <ISTRUZIONE>  testo tra quadre o tra < > che spiega cosa scrivere ("[Inserire qui ...]"): va sostituito
  <RIGHE ____>  spazi vuoti: spesso li compila lo studente o il docente dopo la prova (nome, data, punteggio).
                Decidi caso per caso: si riempiono solo quelli che riguardano la preparazione della prova.
  <TRATTINO>    trattino lungo o medio da sostituire con altra punteggiatura
"""
import re
import sys

import docx
from docx.table import Table
from docx.text.paragraph import Paragraph

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
PAROLE = r"inserire|inserisci|scrivere|scrivi|indicare|indica|descrivere|completare|completa|aggiungere|aggiungi|riportare|qui"
ISTRUZIONE = re.compile(r"\[[^\]]*\b(?:%s)\b[^\]]*\]|<[^<>]*\b(?:%s)\b[^<>]*>|\bX{3,}\b" % (PAROLE, PAROLE), re.I)
RIGHE = re.compile(r"_{4,}|\.{6,}")
TRATTINI = re.compile("[–—]")


def segna(testo):
    note = [n for n, rx in (("ISTRUZIONE", ISTRUZIONE), ("RIGHE ____", RIGHE), ("TRATTINO", TRATTINI)) if rx.search(testo)]
    return ("  <" + ", ".join(note) + ">") if note else ""


def blocchi(contenitore, elemento):
    """Paragrafi e tabelle nell'ordine in cui compaiono."""
    for figlio in elemento.iterchildren():
        if figlio.tag == W + "p":
            yield Paragraph(figlio, contenitore)
        elif figlio.tag == W + "tbl":
            yield Table(figlio, contenitore)


def mostra(contenitore, elemento, rientro=""):
    t_n = 0
    for b in blocchi(contenitore, elemento):
        if isinstance(b, Paragraph):
            testo = b.text.strip()
            if testo:
                stile = b.style.name if b.style is not None else ""
                print("%s[%s] %s%s" % (rientro, stile, testo.replace("\n", "\n" + rientro + "   "), segna(testo)))
        else:
            t_n += 1
            print("%s== Tabella %d (%d righe x %d colonne)" % (rientro, t_n, len(b.rows), len(b.columns)))
            for r_i, r in enumerate(b.rows, 1):
                for c_i, tc in enumerate(r._tr.tc_lst, 1):        # celle reali, una volta sola anche se unite
                    print("%s  R%d C%d:" % (rientro, r_i, c_i))
                    mostra(b, tc, rientro + "      ")


def principale(percorso):
    d = docx.Document(percorso)
    print("== Corpo")
    mostra(d, d.element.body)
    for s_i, s in enumerate(d.sections, 1):
        for nome, parte in (("intestazione", s.header), ("piè di pagina", s.footer),
                            ("intestazione prima pagina", s.first_page_header), ("piè di pagina prima pagina", s.first_page_footer),
                            ("intestazione pagine pari", s.even_page_header), ("piè di pagina pagine pari", s.even_page_footer)):
            if not parte.is_linked_to_previous and (any(p.text.strip() for p in parte.paragraphs) or parte.tables):
                print("== Sezione %d, %s" % (s_i, nome))
                mostra(parte, parte._element, "  ")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    try:
        principale(sys.argv[1])
    except Exception as e:
        sys.exit("Non riesco a leggere %s: %s" % (sys.argv[1], e))
