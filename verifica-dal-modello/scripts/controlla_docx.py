"""Primo controllo automatico delle verifiche .docx, prima di consegnarle.

Uso:  python controlla_docx.py [--modello MODELLO.docx] FILE.docx [altri file]
      --modello  controlla anche che i testi fissi del modello siano ancora nel file
La somma dei punteggi non si controlla qui (totali di sezione e sottopunti si mescolano): va fatta a mano.
Esce con 1 se trova problemi. Serve: pip install python-docx
"""
import argparse
import re
import sys
import zipfile
from collections import Counter

import docx
from docx.oxml.ns import qn

ROTTI = re.compile("Ã|Â|â€|�|ï¿½")
FONT_SIMBOLO = re.compile(r'w:(?:ascii|hAnsi|cs|eastAsia)="(Wingdings[^"]*|Webdings|Symbol|Marlett|MT Extra)"')
SYM = re.compile(r'<w:sym [^>]*w:font="([^"]+)"')
TRATTINI = re.compile("[–—]")
INVISIBILI = re.compile("[​‌‍⁠﻿­]")
# testo tra quadre con un'istruzione per chi compila ("[Inserire qui ...]"), righe ____ e puntini
ISTRUZIONE = r"\[[^\]]*\b(?:[Ii]nserire|[Ii]nserisci|[Ss]crivere|[Ss]crivi|[Ii]ndicare|[Dd]escrivere|qui)\b[^\]]*\]"
SEGNAPOSTO = re.compile(ISTRUZIONE + r"|_{4,}|\.{6,}")


def paragrafi(d):
    """Tutti i paragrafi del corpo, delle tabelle (anche annidate), di intestazioni e piè di pagina."""
    def da(contenitore):
        for p in contenitore.paragraphs:
            yield p
        for t in contenitore.tables:
            for r in t.rows:
                visti = set()
                for c in r.cells:
                    if id(c._tc) not in visti:
                        visti.add(id(c._tc))
                        yield from da(c)
    yield from da(d)
    for s in d.sections:
        for parte in (s.header, s.footer):
            if not parte.is_linked_to_previous:
                yield from da(parte)


def testi_fissi(modello):
    """Paragrafi del modello senza spazi da compilare: devono restare (trattini a parte)."""
    fissi = []
    for p in paragrafi(docx.Document(modello)):
        t = p.text.strip()
        if len(t) > 15 and not SEGNAPOSTO.search(t):
            fissi.append(normalizza(t))
    return fissi


def normalizza(t):
    """Solo lettere e cifre: i trattini del modello si possono sostituire con altra punteggiatura."""
    return "".join(ch for ch in t.lower() if ch.isalnum())


def controlla(percorso, fissi=None):
    problemi = []
    try:
        d = docx.Document(percorso)
    except Exception as e:
        return ["non si apre: %s" % e]
    with zipfile.ZipFile(percorso) as z:
        for nome in z.namelist():
            if nome == "word/document.xml" or nome.startswith(("word/header", "word/footer")):
                xml = z.read(nome).decode("utf-8", "replace")
                for f in set(FONT_SIMBOLO.findall(xml)):
                    problemi.append("%s: testo in carattere simbolo %s (controlla che sia solo una casella da spuntare)" % (nome, f))
                for f in set(SYM.findall(xml)):
                    problemi.append("%s: simbolo inserito con il carattere %s" % (nome, f))

    testi = [p.text for p in paragrafi(d)]
    tutto = "\n".join(testi)
    for nome, rx in (("carattere rotto", ROTTI), ("carattere invisibile", INVISIBILI), ("trattino lungo o medio", TRATTINI)):
        for m in rx.finditer(tutto):
            problemi.append("%s %r: ...%s..." % (nome, m.group(), tutto[max(0, m.start() - 30):m.end() + 30].replace("\n", " | ")))
    for m in re.finditer(ISTRUZIONE, tutto):
        problemi.append("segnaposto non compilato: %s" % m.group()[:80])

    for i, t in enumerate(d.tables, 1):
        griglia = len(t._tbl.tblGrid.findall(qn("w:gridCol")))
        for j, r in enumerate(t.rows, 1):
            span = sum(int(c.tcPr.gridSpan.val) if c.tcPr is not None and c.tcPr.gridSpan is not None else 1 for c in r._tr.tc_lst)
            if span != griglia:
                problemi.append("tabella %d riga %d: %d colonne su una griglia di %d" % (i, j, span, griglia))

    corpo = [p.text.strip() for p in d.paragraphs if p.text.strip().strip("_. ")]
    for t, n in Counter(x for x in corpo if len(x) > 40).items():
        if n > 1:
            problemi.append("paragrafo ripetuto %d volte: %s" % (n, t[:80]))
    for a, b in zip(corpo, corpo[1:]):
        if a == b and len(a) <= 40:
            problemi.append("riga doppia di fila: %s" % a)

    if fissi:
        presenti = normalizza(tutto.replace("\n", " "))
        for f in fissi:
            if f not in presenti:
                problemi.append("testo del modello mancante o cambiato: %s" % f[:100])

    return problemi


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    a = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    a.add_argument("--modello")
    a.add_argument("file", nargs="+")
    args = a.parse_args()
    fissi = testi_fissi(args.modello) if args.modello else None
    errori = 0
    for f in args.file:
        p = controlla(f, fissi)
        print(("OK   " if not p else "KO   ") + f)
        for x in p:
            print("     - " + x)
        errori += bool(p)
    sys.exit(1 if errori else 0)
