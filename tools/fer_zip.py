"""Fa dist/Mudae-tauler.zip per compartir o penjar com a release: el programa, sense cap dada personal.

Executa: python tools/fer_zip.py
"""
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "dist" / "Mudae-tauler.zip"
FILES = ["mudae.py", "i18n.py", "oc.html", "Actualitza.bat", "Configura.bat", "README.md", "README.ca.md", "README.es.md", "LICENSE",
         "icones/tauler.ico", "assets/icona-256.png", "assets/tauler.png", "assets/col-leccions.png", "assets/cercador.png"]
PERSONAL = (".xlsx", "dashboard.html", "harem.txt", "wishlist.txt", "config.json", "historial.csv", "actualitzat.json",
            "series-sobrants.json", "mudae.ico", "hypixel")

OUT.parent.mkdir(exist_ok=True)
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    for f in FILES:
        z.write(ROOT / f, "Mudae/" + f)
    for f in sorted((ROOT / "lib").rglob("*")):  # openpyxl i et_xmlfile (MIT), amb les seves llicències
        if f.is_file() and "__pycache__" not in f.parts:
            z.write(f, "Mudae/" + f.relative_to(ROOT).as_posix())
    z.writestr("Mudae/dades/series/", "")  # carpeta buida: la resta la crea Configura.bat

names = zipfile.ZipFile(OUT).namelist()
leaks = [n for n in names if any(p in n for p in PERSONAL)]
assert not leaks, f"s'hi ha colat una dada personal: {leaks}"
print(f"{OUT.relative_to(ROOT)}: {OUT.stat().st_size // 1024} KB, {len(names)} fitxers, cap dada personal")
