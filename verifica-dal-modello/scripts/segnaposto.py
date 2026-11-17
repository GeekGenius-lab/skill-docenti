"""Mostra la struttura di un modello di verifica .docx: paragrafi, tabelle, spazi da compilare e trattini da sostituire.

Uso:  python segnaposto.py MODELLO.docx
Serve: pip install python-docx
"""
import re
import sys

import docx

# testo tra quadre con un'istruzione per chi compila ("[Inserire qui ...]"), righe ____ e puntini
ISTRUZIONE = r"\[[^\]]*\b(?:[Ii]nserire|[Ii]nserisci|[Ss]crivere|[Ss]crivi|[Ii]ndicare|[Dd]escrivere|qui)\b[^\]]*\]"
SEGNAPOSTO = re.compile(ISTRUZIONE + r"|_{4,}|\.{6,}")
TRATTINI = re.compile("[–—]")


def segna(testo):
    note = []
    if SEGNAPOSTO.search(testo):
        note.append("DA COMPILARE")
    if TRATTINI.search(testo):
        note.append("TRATTINO")
    return ("  <" + ", ".join(note) + ">") if note else ""


def mostra(percorso):
    d = docx.Document(percorso)
    print("== Corpo")
    for i, p in enumerate(d.paragraphs):
        if p.text.strip():
            print("P%-4d [%s] %s%s" % (i, p.style.name, p.text.strip()[:160], segna(p.text)))
    for t_i, t in enumerate(d.tables):
        print("== Tabella %d (%d righe x %d colonne)" % (t_i, len(t.rows), len(t.columns)))
        for r_i, r in enumerate(t.rows):
            visti = set()
            for c_i, c in enumerate(r.cells):
                if id(c._tc) in visti:          # celle unite: python-docx le ripete
                    continue
                visti.add(id(c._tc))
                testo = c.text.strip().replace("\n", " | ")
                print("  T%d R%d C%d: %s%s" % (t_i, r_i, c_i, testo[:160] or "(vuota)", segna(c.text)))
    for s_i, s in enumerate(d.sections):
        for nome, parte in (("intestazione", s.header), ("piè di pagina", s.footer)):
            testo = " | ".join(p.text.strip() for p in parte.paragraphs if p.text.strip())
            if testo:
                print("== Sezione %d, %s: %s%s" % (s_i, nome, testo[:160], segna(testo)))


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    mostra(sys.argv[1])
