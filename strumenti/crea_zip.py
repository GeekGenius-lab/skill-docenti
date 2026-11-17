"""Crea verifica-dal-modello.zip da caricare su claude.ai.

Uso: python strumenti/crea_zip.py   (dalla cartella principale del repository; funziona su Windows, Mac e Linux)
Lo zip contiene la cartella verifica-dal-modello con SKILL.md, riferimenti, script e i modelli che hai messo in modelli/.
"""
import zipfile
from pathlib import Path

RADICE = Path(__file__).resolve().parent.parent
SKILL = RADICE / "verifica-dal-modello"
ZIP = RADICE / "verifica-dal-modello.zip"

with zipfile.ZipFile(ZIP, "w", zipfile.ZIP_DEFLATED) as z:
    for f in sorted(SKILL.rglob("*")):
        if f.is_file() and "__pycache__" not in f.parts and f.suffix != ".pyc":
            z.write(f, f.relative_to(RADICE))
            print("  " + str(f.relative_to(RADICE)))
print("Creato " + str(ZIP))
