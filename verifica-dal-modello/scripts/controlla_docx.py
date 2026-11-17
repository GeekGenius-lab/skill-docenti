"""Primo controllo automatico delle verifiche .docx, prima di consegnarle.

Uso:  python controlla_docx.py [--modello MODELLO.docx] FILE.docx [altri file]
      --modello  controlla anche che i testi fissi del modello siano ancora nel file.
                 Usalo per la verifica e le versioni adattate, non per le soluzioni se sono un documento semplice.
La somma dei punteggi non si controlla qui (totali di sezione e sottopunti si mescolano): va fatta a mano.
Esce con 1 se trova problemi. Serve: pip install python-docx
"""
import argparse
import re
import sys
import zipfile
from collections import Counter

import docx
from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
ROTTI = re.compile("Ã|Â|â€|�|ï¿½")
FONT_SIMBOLO = re.compile(r'w:(?:ascii|hAnsi|cs|eastAsia)="(Wingdings[^"]*|Webdings|Symbol|Marlett|MT Extra)"')
SYM = re.compile(r'<w:sym [^>]*w:font="([^"]+)"')
TRATTINI = re.compile("[–—]")
INVISIBILI = re.compile("[​‌‍⁠﻿­]")
# istruzioni per chi compila: "[Inserire qui ...]", "<inserire materia>", "XXX"
PAROLE = r"inserire|inserisci|scrivere|scrivi|indicare|indica|descrivere|completare|completa|aggiungere|aggiungi|riportare|qui"
ISTRUZIONE = re.compile(r"\[[^\]]*\b(?:%s)\b[^\]]*\]|<[^<>]*\b(?:%s)\b[^<>]*>|\bX{3,}\b" % (PAROLE, PAROLE), re.I)
RIGHE = re.compile(r"_{4,}|\.{6,}")
# parti del file con testo visibile: corpo (con caselle di testo), tutte le intestazioni e i piè di pagina, note
PARTI = re.compile(r"word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$")


def paragrafi(percorso):
    """Testo di ogni paragrafo del file, nell'ordine, comprese tabelle, caselle di testo, intestazioni e note."""
    testi = []
    with zipfile.ZipFile(percorso) as z:
        for nome in z.namelist():
            if PARTI.match(nome):
                radice = etree.fromstring(z.read(nome))
                for p in radice.iter(W + "p"):
                    if p.find(".//" + W + "p") is not None:        # contiene una casella di testo: si legge il contenuto
                        continue
                    parti = []
                    for e in p.iter(W + "t", W + "tab", W + "br", W + "cr"):
                        parti.append((e.text or "") if e.tag == W + "t" else ("\t" if e.tag == W + "tab" else "\n"))
                    testi.append("".join(parti))
    return testi


def pezzi_fissi(testo):
    """Pezzi di testo che non sono spazi da compilare (si confrontano solo lettere e cifre).
    Restituisce (pezzo normalizzato, riga_intera, testo originale): una riga senza spazi da compilare deve ritrovarsi identica."""
    for riga in testo.split("\n"):
        pezzi = re.split(ISTRUZIONE.pattern + "|" + RIGHE.pattern, riga, flags=re.I)
        for pezzo in pezzi:
            n = normalizza(pezzo or "")
            if len(n) >= 4 and not n.isdigit():
                yield n, len(pezzi) == 1, re.sub(r"\s+", " ", pezzo).strip()


def testi_fissi(modello):
    visti = {}
    for t in paragrafi(modello):
        for n, intera, originale in pezzi_fissi(t):
            visti.setdefault((n, intera), originale)
    return [(n, intera, o) for (n, intera), o in visti.items()]


def normalizza(t):
    """Solo lettere e cifre: i trattini del modello si possono sostituire con altra punteggiatura."""
    return "".join(ch for ch in t.lower() if ch.isalnum())


def solo_righe(t):
    """Riga di risposta: quasi tutta trattini bassi o puntini."""
    return len(re.sub(r"[_.\s]", "", t)) <= max(12, len(t) // 4)


def controlla(percorso, fissi=None):
    problemi = []
    try:
        d = docx.Document(percorso)
        testi = paragrafi(percorso)
    except Exception as e:
        return ["non si apre: %s" % e]
    with zipfile.ZipFile(percorso) as z:
        for nome in z.namelist():
            if PARTI.match(nome) or nome == "word/styles.xml":
                xml = z.read(nome).decode("utf-8", "replace")
                for f in sorted(set(FONT_SIMBOLO.findall(xml))):
                    problemi.append("%s: carattere simbolo %s (va bene solo sulle caselle da spuntare, non sul testo)" % (nome, f))
                for f in sorted(set(SYM.findall(xml))):
                    problemi.append("%s: simbolo inserito con il carattere %s (va bene solo sulle caselle da spuntare)" % (nome, f))

    tutto = "\n".join(testi)
    for nome, rx in (("carattere rotto", ROTTI), ("carattere invisibile", INVISIBILI), ("trattino lungo o medio", TRATTINI)):
        for m in rx.finditer(tutto):
            problemi.append("%s %r: ...%s..." % (nome, m.group(), tutto[max(0, m.start() - 30):m.end() + 30].replace("\n", " | ")))
    for m in ISTRUZIONE.finditer(tutto):
        problemi.append("segnaposto non compilato: %s" % m.group()[:80])

    for i, t in enumerate(d.tables, 1):
        griglia = len(t._tbl.tblGrid.findall(W + "gridCol"))
        for j, r in enumerate(t.rows, 1):
            span = sum(int(c.tcPr.gridSpan.val) if c.tcPr is not None and c.tcPr.gridSpan is not None else 1 for c in r._tr.tc_lst)
            trpr = r._tr.trPr
            for extra in ("gridBefore", "gridAfter"):
                e = trpr.find(W + extra) if trpr is not None else None
                span += int(e.get(W + "val")) if e is not None else 0
            if span != griglia:
                problemi.append("tabella %d riga %d: %d colonne su una griglia di %d" % (i, j, span, griglia))

    righe = [r.strip() for t in testi for r in t.split("\n") if r.strip() and not solo_righe(r.strip())]
    for t, n in Counter(x for x in righe if len(x) > 40).items():
        if n > 1:
            problemi.append("testo ripetuto %d volte: %s" % (n, t[:80]))
    for a, b in zip(righe, righe[1:]):
        if a == b and len(a) <= 40:
            problemi.append("riga doppia di fila: %s" % a)

    if fissi:
        righe_n = [normalizza(r) for t in testi for r in t.split("\n")]
        intere, unite = set(righe_n), "|".join(righe_n)     # "|": niente corrispondenze a cavallo di due righe
        for f, riga_intera, originale in fissi:
            if (f not in intere) if riga_intera else (f not in unite):
                problemi.append("testo del modello mancante o cambiato: %s" % originale[:100])

    p = d.core_properties
    for campo, valore in (("autore", p.author), ("ultima modifica di", p.last_modified_by)):
        if valore and valore.strip() and valore.strip() != "python-docx":
            problemi.append("proprietà del file '%s' = %r: svuotala prima di consegnare" % (campo, valore))
    return problemi


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    a = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    a.add_argument("--modello")
    a.add_argument("file", nargs="+")
    args = a.parse_args()
    fissi = None
    if args.modello:
        try:
            fissi = testi_fissi(args.modello)
        except Exception as e:
            sys.exit("Non riesco ad aprire il modello %s: %s" % (args.modello, e))
    errori = 0
    for f in args.file:
        p = controlla(f, fissi)
        print(("OK   " if not p else "KO   ") + f)
        for x in p:
            print("     - " + x)
        errori += bool(p)
    sys.exit(1 if errori else 0)
