"""Actualitza l'Excel i el tauler de Mudae a partir del que enganxes a la carpeta dades/.

Ús: python mudae.py   (o doble clic a Actualitza.bat)

- dades/series/<Sèrie>.txt  -> el que respon `$ima <sèrie>` (una sèrie per fitxer)
- dades/harem.txt           -> el teu `$mm`
- dades/wishlist.txt        -> el teu `$wl`
Pots enganxar-ho tal qual de Discord: les capçaleres i la numeració s'ignoren.
El que escrius a mà a l'Excel (cel·les grogues) es conserva.
"""
import csv
import base64
import datetime as dt
import html
import json
import os
import re
import sys
import time
from pathlib import Path

# openpyxl va dins de lib/: així funciona es llanci des d'on es llanci (l'app de Claude i Windows no veuen la mateixa carpeta d'usuari)
sys.path.insert(0, str(Path(__file__).resolve().parent / "lib"))
from openpyxl import Workbook, load_workbook  # noqa: E402
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.datavalidation import DataValidation

DIR = Path(__file__).resolve().parent
DADES = DIR / "dades"
XLSX = DIR / "Mudae - estrategia.xlsx"
DASH = DIR / "dashboard.html"
HIST = DADES / "historial.csv"
CONFIG_FILE = DADES / "config.json"
try:
    CONFIG = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
except (OSError, ValueError):
    CONFIG = {}
ME = CONFIG.get("usuari", "")  # el teu nom a Mudae (el que surt a «Pertenece a …»)
DISCORD_ID = CONFIG.get("discord_id", "")  # opcional, de $myid: les ordres copiades et mencionen bé
LANG = CONFIG.get("idioma", "en")  # ca, es o en (el tria Configura.bat)
from i18n import STRINGS  # noqa: E402  (castellà i anglès; el text en català fa de clau)


def t(text, **kw):
    """Text en l'idioma de config.json. Les dades internes (Meu, Alta, Pendent…) es guarden sempre en català."""
    s = STRINGS.get(LANG, {}).get(text, text)
    return s.format(**kw) if kw else s


CANON = {STRINGS[lang].get(k, k): k for lang in STRINGS for k in
         ("Alta", "Mitjana", "Baixa", "Pendent", "Enviat", "Acceptat", "Rebutjat", "Guardar", "Per intercanviar", "Divorciar")}


def canon(v):
    """Valor escrit a l'Excel (en qualsevol idioma) -> la clau interna en català."""
    return CANON.get(v, v)


TITLE_TEXT = f"Mudae · {CONFIG['servidor']}" if CONFIG.get("servidor") else t("Tauler de Mudae")
ALTRES = "Altres (no la sé: mira-la amb $im)"
# Personatges que vols que el tauler et proposi primer quan s'allibera un lloc a la wishlist (després, la resta de prioritat Alta)
PREFERITS = CONFIG.get("preferits", [])

CLAIM_HOURS = CONFIG.get("hores_reclam", [1, 4, 7, 10, 13, 16, 19, 22])  # reclams cada 3 h al minut 02
GRUPS = {
    "dia": "EN COMENÇAR EL DIA",
    "hora": "CADA HORA (minut 02)",
    "reclam": "MOMENT BO: pots reclamar ({hours})",
    "aconseguit": "UN COP PER SETMANA (o quan aconsegueixis algú)",
    "temps": "QUAN TINGUIS TEMPS",
}
RT_HORES = CONFIG.get("rt_hores", 50)  # $rt: cada 50 h, -10 h per cada insígnia Esmeralda II, III i IV
# (grup, què, ordre, com, quan es torna a poder fer): surt al full Rutina i al tauler.
# «hora» = al minut 02 · «reclam» = al pròxim reclam · «utc» = a mitjanit UTC · un número = hores des que la fas
RUTINA = [
    ("dia", "Mirar què tens", "$tu", "Resum de tot el que tens disponible", "hora"),
    ("dia", "Votar", "$vote", "Et guarda un reinici de tirades: el gastes amb $rolls. Cada 12 h", 12),
    ("dia", "Reinici de tirades", "$daily", "Et torna les tirades. Cada 20 h: guarda'l per quan te'n quedis sense", 20),
    ("dia", "Esferes", "$oh (×4)", "Primer blaves (destapen 3), als últims clics les millors. 2 minuts!", "utc"),
    ("dia", "Trobar la vermella", "$oc", "Si en tens. Fes servir l'ajudant: et diu on clicar", "utc"),
    ("hora", "Tirades", "$m", "Fes-les totes: les que no fas es perden", "hora"),
    ("hora", "Kakera", "clic al botó", "Cada reacció gasta 40 % de poder (20 % amb personatges de 10+ claus)", "hora"),
    ("reclam", "1. Tirades", "$m", "Totes", "reclam"),
    ("reclam", "2. Reclamar", "clic al cor", "El millor: wishlist > les teves sèries > molta kakera", "reclam"),
    ("reclam", "3. Més tirades", "$us · $rolls", "Tirades guardades ($us) i el reinici del vot ($rolls)", "reclam"),
    ("reclam", "4. Un altre de bo?", "$rt", "Et deixa tornar a reclamar. Cada {rt} h: només per a la wishlist o les teves sèries", RT_HORES),
    ("reclam", "5. Sense poder?", "$dk", "Recarrega el poder i dona kakera. Cada 20 h", 20),
    ("aconseguit", "Treure'l de la wishlist", "$wishremove <nom>", "I posa-hi el primer de «Per omplir» (full Wishlist)", None),
    ("aconseguit", "Posar al dia les dades", "$mms · $wl · $imaso <sèrie>", "Per privat i sense pàgines. Enganxa-ho a «Posar al dia» del tauler", None),
    ("aconseguit", "Obrir el tauler", "Mudae", "Accés directe de l'escriptori: posa al dia l'Excel i obre el tauler", None),
    ("temps", "Intercanvis", "$trade @persona", "Full Intercanvis: a qui demanar què. Ofereix dels Sobrants", None),
    ("temps", "Kakera estalviada", "$build 8", "Quan arribis als 35.000: pis [8] de la torre (o l'[11])", None),
]

SKIP = re.compile(r"Harem de|Wishlist|^Page\b|\bAPP\b|—\s*\d{1,2}:\d{2}|^\$|^Mudae$|^\d+\s*/\s*\d+$", re.I)
HEART_ONLY = re.compile(r"^(.*?)\s*(?:💞|💖|💗|❤️|❤|♥️|♥)\s*$")  # «Nom 💞» del $imas (sense la «o»): reclamat, propietari desconegut
CLAIM = re.compile(r"^(.*?)\s*(?:💞|💖|💗|❤️|❤|♥️|♥)?\s*=>\s*(\S+)\s*$")
IMASO_TITLE = re.compile(r"^(?:\d+\.\s+)?\S.*\s\d+\s*/\s*\d+$")  # «1. Harry Potter 49/94»: el títol del $imaso


def drop_header(text):
    """El $imaso comença amb el títol de la sèrie («1. Harry Potter 49/94») i un bloc amb els àlies: ho traiem.

    El bloc d'àlies acaba amb «… $ima? …» o amb la línia separadora de Mudae (un espai). Si no hi és, només traiem
    el títol: més val que surti algun àlies com a personatge que perdre personatges de veritat.
    """
    rows = text.splitlines()
    for i, raw in enumerate(rows[:6]):
        if IMASO_TITLE.match(raw.strip()):
            k = next((j for j in range(i + 1, len(rows)) if rows[j].strip()), len(rows))   # primera línia del bloc
            end = next((j for j in range(k, len(rows)) if not rows[j].strip()), len(rows))  # línia buida que el tanca
            block = rows[k:end]
            aliases = block and not any("=>" in r for r in block) and (
                "$ima?" in block[-1] or (end < len(rows) and rows[end] != ""))
            return "\n".join(rows[:i] + rows[end + 1 if aliases else k:])
    return text


def lines(path):
    """Línies útils d'un fitxer enganxat de Discord, amb la marca ⭐ si n'hi havia."""
    out = []
    if not path.exists():
        return out
    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        star = "⭐" in raw
        line = re.sub(r"^\d+\.\s+", "", raw.replace("⭐", "").strip())  # numeració de Discord: "2. Remus Lupin"
        if line and not SKIP.search(line):
            out.append((line.strip(), star))
    return out


def specials():
    """Col·leccions especials: grups de personatges triats per tu (p. ex. els Eevee de la Pokédex), no sèries de Mudae."""
    f = DADES / "especials.json"
    return json.loads(f.read_text(encoding="utf-8")) if f.exists() else []


def parse_owners(text):
    """Nom -> propietari d'un text de Mudae ($imaso, $imas o una llista): "" lliure, None «💞» sense saber de qui."""
    out = {}
    for raw in text.splitlines():
        line = re.sub(r"^\d+\.\s+", "", raw.replace("⭐", "").strip())
        if not line or SKIP.search(line):
            continue
        m, h = CLAIM.match(line), HEART_ONLY.match(line)
        name, owner = (m.group(1), m.group(2)) if m else (h.group(1), None) if h else (line, "")
        out.setdefault(name.strip(), owner)
    return out


def read_series():
    rows = []
    spec = set(specials())
    harem = {n for n, _ in lines(DADES / "harem.txt")} if spec else set()
    for f in sorted((DADES / "series").glob("*.txt")):
        seen = set()
        for line, _ in lines(f):
            m = CLAIM.match(line)
            h = HEART_ONLY.match(line)
            name, owner = (m.group(1).strip(), m.group(2)) if m else (h.group(1).strip(), "algú") if h else (line, "")
            if owner.casefold() == ME.casefold():  # Mudae escriu els noms en minúscula («vlarac»), tu pots haver posat «Vlarac»
                owner = ME
            if f.stem in spec and name in harem:  # a les especials, el que és al teu harem és teu (no hi ha cap $imaso que ho digui)
                owner = ME
            if name not in seen:
                seen.add(name)
                rows.append({"serie": f.stem, "name": name, "owner": owner})
    return rows


def old_inputs():
    """Llegeix el que hi ha escrit a mà a l'Excel anterior, per no perdre-ho."""
    old = {"col": {}, "col_name": {}, "trade": {}, "sob": {}}
    if not XLSX.exists():
        return old
    wb = load_workbook(XLSX)
    ws = sheet(wb, "Col·lecció")
    for r in ws.iter_rows(min_row=2, values_only=True) if ws else ():
        if r[1]:
            old["col"][(r[0], r[1])] = old["col_name"][r[1]] = {"owner": r[2] or "", "prio": canon(r[4]), "notes": r[6]}
    ws = sheet(wb, "Intercanvis")
    for r in ws.iter_rows(min_row=5, values_only=True) if ws else ():
        if r[0] and r[1]:
            old["trade"][(r[0], r[1])] = (r[4], canon(r[5]))
    ws = sheet(wb, "Sobrants")
    for r in ws.iter_rows(min_row=6, values_only=True) if ws else ():
        if r[1]:
            old["sob"][r[1]] = canon(r[2])
    return old


def state(owner):
    return "Meu" if owner == ME else ("Lliure" if not owner else "D'algú altre")


def build_data():
    rows = read_series()
    old = old_inputs()
    changes = []
    for r in rows:
        prev = old["col"].get((r["serie"], r["name"])) or old["col_name"].get(r["name"])  # per nom: si la sèrie s'ha reanomenat
        r["prio"] = (prev or {}).get("prio") or "Mitjana"
        r["notes"] = (prev or {}).get("notes")
        r["state"] = state(r["owner"])
        if prev is None:
            if old["col"]:
                changes.append(("nou", r, ""))
        elif prev["owner"] != r["owner"]:
            changes.append(("canvi", r, prev["owner"]))

    by_name = {r["name"]: r for r in rows}
    wl_raw = lines(DADES / "wishlist.txt")
    wishlist = []
    for name, star in wl_raw:
        r = by_name.get(name)
        st = r["state"] if r else "?"
        if st == "Meu":
            todo = t("Ja és teu: treu-lo")
        elif st == "D'algú altre":
            todo = t("L'ha reclamat {owner}: treu-lo (intercanvi)", owner=r["owner"])
        elif st == "Lliure":
            todo = t("Bé")
        else:
            todo = t("No és a cap sèrie teva: mira-ho amb $im")
        wishlist.append({"name": name, "star": star, "serie": r["serie"] if r else "", "state": st, "ok": st == "Lliure", "todo": todo})
    wl_names = {w["name"] for w in wishlist}
    free_alta = [r for r in rows if r["state"] == "Lliure" and r["prio"] == "Alta" and r["name"] not in wl_names]
    rank = {n: i for i, n in enumerate(PREFERITS)}
    free_alta.sort(key=lambda r: (rank.get(r["name"], len(rank)), r["serie"], r["name"]))

    trades, seen = [], set()
    for r in sorted((r for r in rows if r["state"] == "D'algú altre"), key=lambda r: (r["owner"], r["serie"], r["name"])):
        if r["name"] in seen:
            continue
        seen.add(r["name"])
        offer, st = old["trade"].get((r["owner"], r["name"]), (None, "Pendent"))
        trades.append({**r, "offer": offer, "status": st or "Pendent"})

    sob_file = DADES / "series-sobrants.json"
    sob_map = json.loads(sob_file.read_text(encoding="utf-8")) if sob_file.exists() else {}
    sob_serie = {n: s for s, ns in sob_map.items() for n in ns}
    mine = {r["name"] for r in rows if r["state"] == "Meu"}
    harem = [n for n, _ in lines(DADES / "harem.txt")]
    sobrants = [{"name": h, "serie": sob_serie.get(h, ALTRES), "todo": old["sob"].get(h)} for h in harem if h not in mine]
    size = {}
    for s in sobrants:
        size[s["serie"]] = size.get(s["serie"], 0) + 1
    sobrants.sort(key=lambda s: (s["serie"] == ALTRES, -size[s["serie"]], s["serie"], s["name"]))
    missing = sorted(mine - set(harem))  # teus segons $ima però no al $mm: sol voler dir que el harem és vell

    series = {}
    for r in rows:
        s = series.setdefault(r["serie"], {"name": r["serie"], "total": 0, "Meu": 0, "D'algú altre": 0, "Lliure": 0})
        s["total"] += 1
        s[r["state"]] += 1
    series = sorted(series.values(), key=lambda s: (-s["Meu"], s["name"]))
    return {"rows": rows, "series": series, "wishlist": wishlist, "free_alta": free_alta, "trades": trades,
            "sobrants": sobrants, "sob_size": size, "changes": changes, "harem": harem, "missing": missing}


# ---------------------------------------------------------------- Excel
F = "Arial"
H_FILL, H_FONT = PatternFill("solid", fgColor="2F3B52"), Font(name=F, bold=True, color="FFFFFF")
TITLE, BOLD = Font(name=F, bold=True, size=14), Font(name=F, bold=True)
INPUT = PatternFill("solid", fgColor="FFF2CC")
GREEN, RED, BLUE = (PatternFill("solid", fgColor=c) for c in ("D9EAD3", "F4CCCC", "CFE2F3"))
WRAP = Alignment(wrap_text=True, vertical="top")


def put(ws, row, col, value):
    """Cel·la amb un text que ve de Discord (nom, sèrie, jugador): sempre text, mai fórmula (=, +, -, @)."""
    cell = ws.cell(row, col, value)
    if isinstance(value, str) and value[:1] in "=+-@":
        cell.data_type = "s"
    return cell


def header(ws, row, cols):
    for i, c in enumerate(cols, 1):
        cell = ws.cell(row=row, column=i, value=c)
        cell.fill, cell.font = H_FILL, H_FONT


def widths(ws, w):
    for col, v in w.items():
        ws.column_dimensions[col].width = v


def sheet(wb, name):
    """Full de l'Excel anterior pel seu nom en qualsevol idioma (si has canviat d'idioma, no es perd res)."""
    for n in (name, *(STRINGS[lang].get(name, name) for lang in STRINGS)):
        if n in wb.sheetnames:
            return wb[n]
    return None


def write_excel(d):
    col = t("Col·lecció")
    MEU, LLIURE, ALTRE, ALTA = t("Meu"), t("Lliure"), t("D'algú altre"), t("Alta")
    q = lambda v: str(v).replace('"', '""')  # cometes dins d'una fórmula
    options = lambda *vals: '"' + ",".join(t(v) for v in vals) + '"'
    wb = Workbook()
    # Resum
    ws = wb.active
    ws.title = t("Resum")
    ws["A1"], ws["A1"].font = t("{title} · com vas de col·lecció", title=TITLE_TEXT), TITLE
    ws["A2"] = t("Actualitzat el {date} amb mudae.py. Es calcula sol a partir de «{sheet}».",
                 date=f"{dt.datetime.now():%d/%m/%Y %H:%M}", sheet=col)
    header(ws, 4, [t("Sèrie"), t("Total"), t("Meus"), t("Lliures"), t("D'algú altre"), t("% completat")])
    for i, s in enumerate(d["series"]):
        r = 5 + i
        put(ws, r, 1, s["name"])
        ws.cell(r, 2, f"=COUNTIF('{col}'!$A:$A,A{r})")
        for c, st in (("C", MEU), ("D", LLIURE), ("E", ALTRE)):
            ws[f"{c}{r}"] = f"=COUNTIFS('{col}'!$A:$A,A{r},'{col}'!$D:$D,\"{q(st)}\")"
        ws.cell(r, 6, f"=IF(B{r}=0,0,C{r}/B{r})").number_format = "0%"
    tr = 5 + len(d["series"])
    ws.cell(tr, 1, t("Total")).font = BOLD
    for c in "BCDE":
        ws[f"{c}{tr}"] = f"=SUM({c}5:{c}{tr - 1})"
        ws[f"{c}{tr}"].font = BOLD
    ws[f"F{tr}"] = f"=IF(B{tr}=0,0,C{tr}/B{tr})"
    ws[f"F{tr}"].number_format, ws[f"F{tr}"].font = "0%", BOLD
    ws.cell(tr + 2, 1, t("Per posar-ho al dia: obre el tauler (accés directe Mudae) i fes servir «Posar al dia»."))
    widths(ws, {"A": 24, "B": 9, "C": 9, "D": 9, "E": 14, "F": 13})

    # Col·lecció
    ws = wb.create_sheet(col)
    header(ws, 1, [t("Sèrie"), t("Personatge"), t("Propietari"), t("Estat"), t("Prioritat"), t("Què fer"), t("Notes")])
    for i, r in enumerate(d["rows"], start=2):
        put(ws, i, 1, r["serie"])
        put(ws, i, 2, r["name"])
        put(ws, i, 3, r["owner"])
        ws.cell(i, 4, f'=IF(C{i}="","{q(LLIURE)}",IF(C{i}="{q(ME)}","{q(MEU)}","{q(ALTRE)}"))')
        ws.cell(i, 5, t(r["prio"])).fill = INPUT
        ws.cell(i, 6, f'=IF(D{i}="{q(MEU)}","—",IF(D{i}="{q(ALTRE)}","{q(t("Intercanvi amb "))}"&C{i},'
                      f'IF(E{i}="{q(ALTA)}","Wishlist","{q(t("Tirades (si surt, reclama-ho)"))}")))')
        put(ws, i, 7, r["notes"]).fill = INPUT
    last = max(len(d["rows"]) + 1, 2)  # buit (primera vegada): openpyxl no accepta rangs com A2:G1
    rng = f"A2:G{last}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'$D2="{q(MEU)}"'], fill=GREEN))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'$D2="{q(ALTRE)}"'], fill=RED))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND($D2="{q(LLIURE)}",$E2="{q(ALTA)}")'], fill=BLUE))
    dv = DataValidation(type="list", formula1=options("Alta", "Mitjana", "Baixa"), allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"E2:E{last}")
    ws.auto_filter.ref, ws.freeze_panes = f"A1:G{last}", "A2"
    widths(ws, {"A": 20, "B": 30, "C": 32, "D": 12, "E": 10, "F": 38, "G": 36})

    # Sobrants
    ws = wb.create_sheet(t("Sobrants"))
    ws["A1"], ws["A1"].font = t("El que tens i no col·lecciones"), TITLE
    ws["A2"] = t("Personatges del teu harem que no són de les sèries que col·lecciones. Serveixen per oferir en "
                 "intercanvis o, si no t'importen, per divorciar-te'n i treure'n kakera ($divorce).")
    ws["A3"] = t("La sèrie pot tenir errors: $im <nom> et diu la de debò, i la pots corregir al tauler.")
    header(ws, 5, [t("Sèrie"), t("Personatge"), t("Què en faig")])
    for i, s in enumerate(d["sobrants"], start=6):
        put(ws, i, 1, t(s["serie"]))
        put(ws, i, 2, s["name"])
        ws.cell(i, 3, t(s["todo"]) if s["todo"] else None).fill = INPUT
    last = max(5 + len(d["sobrants"]), 6)
    dv = DataValidation(type="list", formula1=options("Guardar", "Per intercanviar", "Divorciar"), allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"C6:C{last}")
    ws.auto_filter.ref, ws.freeze_panes = f"A5:C{last}", "A6"
    widths(ws, {"A": 34, "B": 32, "C": 40})

    # Wishlist
    ws = wb.create_sheet(t("Wishlist"))
    ws["A1"], ws["A1"].font = t("Wishlist"), TITLE
    ws["A2"] = t("Norma: només personatges LLIURES. Si ja són d'algú, no els pots reclamar: van per intercanvi.")
    header(ws, 4, [t("La teva wishlist ($wl)"), t("Sèrie"), t("Estat"), t("Què fer")])
    for i, w in enumerate(d["wishlist"], start=5):
        put(ws, i, 1, w["name"] + (" ⭐" if w["star"] else ""))
        put(ws, i, 2, w["serie"])
        ws.cell(i, 3, t(w["state"]))
        put(ws, i, 4, w["todo"]).fill = GREEN if w["ok"] else RED
    r0 = 5 + len(d["wishlist"]) + 1
    ws.cell(r0, 1, t("Per omplir llocs lliures (lliures i de prioritat Alta, per ordre)")).font = BOLD
    header(ws, r0 + 1, [t("Personatge"), t("Sèrie")])
    for k, r in enumerate(d["free_alta"], start=r0 + 2):
        put(ws, k, 1, r["name"])
        put(ws, k, 2, r["serie"])
    k = r0 + 3 + len(d["free_alta"])
    ws.cell(k, 1, t("Ordres")).font = BOLD
    for j, line in enumerate([t("Afegir: $wish <nom> (diversos: separa'ls amb $)   ·   Treure: $wishremove <nom>"),
                              t("El desig estrella ($sw) també compta dins dels 15.")], start=k + 1):
        ws.cell(j, 1, "• " + line)
    widths(ws, {"A": 34, "B": 22, "C": 12, "D": 44})

    # Intercanvis
    ws = wb.create_sheet(t("Intercanvis"))
    ws["A1"], ws["A1"].font = t("Intercanvis: qui té el que busques"), TITLE
    ws["A2"] = t("Omple les cel·les grogues. Quan un intercanvi es faci, marca'l al tauler i la fila desapareixerà sola.")
    header(ws, 4, [t("Propietari"), t("Personatge"), t("Sèrie"), t("Prioritat"), t("Què li ofereixo"), t("Estat")])
    for i, tr_ in enumerate(d["trades"], start=5):
        for j, v in enumerate((tr_["owner"], tr_["name"], tr_["serie"], t(tr_["prio"])), start=1):
            put(ws, i, j, v)
        put(ws, i, 5, tr_["offer"]).fill = INPUT
        ws.cell(i, 6, t(tr_["status"])).fill = INPUT
    last = 4 + len(d["trades"])
    dv = DataValidation(type="list", formula1=options("Pendent", "Enviat", "Acceptat", "Rebutjat"), allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"F5:F{max(last, 5)}")
    ws.auto_filter.ref = f"A4:F{max(last, 5)}"
    widths(ws, {"A": 34, "B": 28, "C": 20, "D": 10, "E": 34, "F": 12})

    # Rutina (text fix)
    ws = wb.create_sheet(t("Rutina"))
    ws["A1"], ws["A1"].font = t("Rutina de farmeig"), TITLE
    header(ws, 3, [t("Quan"), t("Què"), t("Ordre"), t("Per què / com")])
    prev = None
    for i, (grup, que, ordre, com, _torna) in enumerate(RUTINA, start=4):
        ws.cell(i, 1, group_title(grup) if grup != prev else None).font = BOLD
        prev = grup
        for j, v in enumerate((t(que), cmd_text(ordre), t(com, rt=RT_HORES)), start=2):
            ws.cell(i, j, v).alignment = WRAP
    widths(ws, {"A": 30, "B": 24, "C": 24, "D": 66})

    for s in wb.worksheets:
        for row in s.iter_rows():
            for c in row:
                if c.font.name != F:
                    c.font = Font(name=F, bold=c.font.bold, color=c.font.color, size=c.font.size)
    wb.calculation = CalcProperties(fullCalcOnLoad=True)  # openpyxl no desa resultats: Excel els calcula en obrir
    try:
        wb.save(XLSX)
    except PermissionError:
        raise RuntimeError(t("No puc desar l'Excel: el tens obert. Tanca'l i torna-ho a provar."))


# ---------------------------------------------------------------- historial
def update_history(d):
    today = dt.date.today().isoformat()
    rows = []
    if HIST.exists():
        with HIST.open(encoding="utf-8", newline="") as f:
            rows = [r for r in csv.DictReader(f) if r["data"] != today]
    rows += [{"data": today, "serie": s["name"], "meus": s["Meu"], "total": s["total"]} for s in d["series"]]
    with HIST.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["data", "serie", "meus", "total"])
        w.writeheader()
        w.writerows(rows)
    totals = {}
    for r in rows:
        t = totals.setdefault(r["data"], [0, 0])
        t[0] += int(r["meus"])
        t[1] += int(r["total"])
    return sorted(totals.items())


# ---------------------------------------------------------------- tauler
def e(x):
    return html.escape(str(x))


JS_HEAD = """
// Textos en l'idioma triat (el català fa de clau) i escapament per posar text a l'HTML
const I18N = __I18N__;
const T = (s, v) => { let r = I18N[s] ?? s; for (const k in v || {}) r = r.split('{' + k + '}').join(v[k]); return r; };
const esc = t => String(t).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
const OFF_MSG = () => '<span class="tag bad">' + esc(T('Botó apagat')) + '</span> ' + esc(T("Obre Mudae a l'escriptori."));
const NET_MSG = () => '<span class="tag bad">' + esc(T('Sense connexió')) + '</span> ' + esc(T("Obre Mudae a l'escriptori."));
const ERR_MSG = e => '<span class="tag bad">' + esc(T("No s'ha desat")) + '</span> ' + esc(e);
const postJSON = (path, body) => fetch(path, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body || {}) }).then(r => r.json());
"""


JS_TODAY = """
const CLAIM = __CLAIM__, clock = document.getElementById('clock');
const pad = n => String(n).padStart(2, '0');
const fmt = m => m >= 60 ? `${Math.floor(m / 60)} h ${m % 60} min` : `${m} min`;
const hm = d => `${pad(d.getHours())}:${pad(d.getMinutes())}`;
// Quan es pot tornar a fer una tasca marcada a l'hora `t` (ms): al minut 02, al pròxim reclam, a mitjanit UTC o al cap de N hores
function until(torna, t) {
  const d = new Date(t);
  if (torna === 'hora') { d.setMinutes(2, 0, 0); if (d <= t) d.setHours(d.getHours() + 1); return +d; }
  if (torna === 'utc') { d.setUTCHours(24, 0, 0, 0); return +d; }
  if (torna === 'reclam') {
    for (let k = 0; k < 2; k++) for (const h of CLAIM) {
      const c = new Date(t); c.setDate(c.getDate() + k); c.setHours(h, 2, 0, 0); if (c > t) return +c;
    }
  }
  return t + Number(torna) * 3600e3;
}
function load() { try { return JSON.parse(localStorage.getItem('mudae-fet') || '{}'); } catch (e) { return {}; } }
function save(s) { try { localStorage.setItem('mudae-fet', JSON.stringify(s)); } catch (e) {} }
function tick() {
  const now = new Date(), s = load(), m = now.getHours() * 60 + now.getMinutes();
  document.querySelectorAll('input[data-torna]').forEach(i => {
    const t = s[i.dataset.id], back = typeof t === 'number' ? until(i.dataset.torna, t) : 0;
    i.checked = back > +now;
    const when = i.closest('.task').querySelector('.when');
    if (when) when.textContent = i.checked ? T('Torna a les {h} ({rest})', { h: hm(new Date(back)), rest: fmt(Math.ceil((back - now) / 60000)) }) : '';
  });
  const nr = (now.getMinutes() < 2 ? now.getHours() : now.getHours() + 1) % 24;
  const toRoll = (nr * 60 + 2 - m + 1440) % 1440;
  const nc = CLAIM.find(h => h * 60 + 2 > m) ?? CLAIM[0] + 24;
  const last = [...CLAIM].reverse().find(h => h * 60 + 2 <= m);
  const fresh = last !== undefined && m - (last * 60 + 2) < 60;
  document.querySelector('.grp[data-period="reclam"]')?.classList.toggle('now', fresh);
  clock.innerHTML = T('Ara <b>{now}</b> · tirades noves a les <b>{roll}</b> ({rollIn}) · pròxim reclam a les <b>{claim}</b> ({claimIn})',
    { now: hm(now), roll: pad(nr) + ':02', rollIn: fmt(toRoll), claim: pad(nc % 24) + ':02', claimIn: fmt(nc * 60 + 2 - m) })
    + (fresh ? '<br><span class="claim-now">' + esc(T('Pots reclamar des de les {h}: fes els passos de «Moment bo»', { h: pad(last) + ':02' })) + '</span>' : '');
}
document.addEventListener('change', ev => {
  const i = ev.target; if (!i.dataset || !i.dataset.torna) return;
  const s = load();
  if (i.checked) s[i.dataset.id] = Date.now(); else delete s[i.dataset.id];
  save(s); tick();
});
tick(); setInterval(tick, 30000);

// mentre la pestanya és oberta, diu al servidor que continua aquí; si es tanca, el servidor plega sol
if (location.protocol.startsWith('http')) setInterval(() => fetch('/ping', { cache: 'no-store' }).catch(() => {}), 20000);
"""


JS_COLL = """
const COLL = __DATA__;
const chips = document.getElementById('chips'), hide = document.getElementById('hideMine'), title = document.getElementById('albumTitle');
const ORDER = { wish: 0, alta: 1, other: 2, free: 3, mine: 4 };  // primer el que et falta
const kind = c => c.st === 'Meu' ? 'mine' : c.st !== 'Lliure' ? 'other' : c.wl ? 'wish' : c.prio === 'Alta' ? 'alta' : 'free';
let cur = Object.keys(COLL)[0];
if (!cur) document.getElementById('coll').querySelector('.album').innerHTML =
  '<p class="muted" style="margin:0">' + esc(T("Encara no col·lecciones cap sèrie. Afegeix-ne una a «Posar al dia»: tria «Una sèrie» i «Sèrie nova…».")) + '</p>';
try { const saved = localStorage.getItem('mudae-coll'); if (saved in COLL) cur = saved; hide.checked = localStorage.getItem('mudae-hide') === '1'; } catch (e) {}
function draw() {
  if (!cur) return;
  const list = COLL[cur], mine = list.filter(c => c.st === 'Meu').length;
  title.textContent = `${cur} · ` + T('{a} de {b}', { a: mine, b: list.length });
  document.querySelectorAll('.row[data-serie]').forEach(r => r.classList.toggle('on', r.dataset.serie === cur));
  chips.innerHTML = [...list].sort((a, b) => ORDER[kind(a)] - ORDER[kind(b)] || a.n.localeCompare(b.n))
    .filter(c => !(hide.checked && c.st === 'Meu')).map(c => {
      const k = kind(c);
      const tip = { mine: T('Teu'), other: T('De {o}: per intercanviar', { o: c.o }), wish: T('Lliure i a la teva wishlist'),
                    alta: T('Lliure · prioritat alta'), free: T('Lliure · prioritat {p}', { p: T(c.prio) }) }[k];
      const inner = `${k === 'wish' ? '<span class="star">★</span> ' : ''}${esc(c.n)}${k === 'other' ? `<small>${esc(c.o)}</small>` : ''}`;
      if (k === 'mine') return `<span class="ch mine" data-tip="${esc(T('Teu'))}">${inner}</span>`;
      // els que no tens: clic per seleccionar-los i marcar-los com a aconseguits
      const sel = got.has(c.n);
      return `<button type="button" class="ch ${k} pick${sel ? ' sel' : ''}" data-pick="${esc(c.n)}" aria-pressed="${sel}" data-tip="${esc(tip + ' · ' + T('clica si ja el tens'))}">${sel ? '✓ ' : ''}${inner}</button>`;
    }).join('') || '<span class="muted">' + esc(T('Ja els tens tots.')) + '</span>';
  gotBar.hidden = !got.size;
  gotCount.textContent = got.size === 1 ? T('1 personatge seleccionat') : T('{n} personatges seleccionats', { n: got.size });
}
const got = new Set();
const gotBar = document.getElementById('gotBar'), gotCount = document.getElementById('gotCount'), gotMsg = document.getElementById('gotMsg');
chips.addEventListener('click', ev => {
  const b = ev.target.closest('[data-pick]'); if (!b) return;
  got.has(b.dataset.pick) ? got.delete(b.dataset.pick) : got.add(b.dataset.pick); draw();
});
document.getElementById('gotClear').onclick = () => { got.clear(); gotMsg.textContent = ''; draw(); };
document.getElementById('gotBtn').onclick = async ev => {
  const btn = ev.target;
  if (location.protocol === 'file:') { gotMsg.innerHTML = OFF_MSG(); return; }
  btn.disabled = true; gotMsg.textContent = T('Desant…');
  try {
    const j = await postJSON('/rebut', { names: [...got] });
    if (j.ok) location.reload(); else { gotMsg.innerHTML = ERR_MSG(j.error); btn.disabled = false; }
  } catch (err) { gotMsg.innerHTML = NET_MSG(); btn.disabled = false; }
};
function show(s) { cur = s; try { localStorage.setItem('mudae-coll', s); } catch (e) {} draw(); }
hide.onchange = () => { try { localStorage.setItem('mudae-hide', hide.checked ? '1' : '0'); } catch (e) {} draw(); };
document.querySelectorAll('.row[data-serie]').forEach(r => {
  r.onclick = () => show(r.dataset.serie);
  r.onkeydown = ev => { if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); show(r.dataset.serie); } };
});
draw();
"""


STYLE = """
/* Fosc grafit, targetes de vidre i un sol accent blau.
   El color només marca estat: blau = teu, taronja = d'algú altre, puntejat = lliure, verd/vermell = avisos */
:root {
  --bg:#141417; --bg-grad:linear-gradient(160deg,#1b1b1f 0%,#141417 50%,#111114 100%);
  --glow1:rgba(88,166,255,.08); --glow2:rgba(255,255,255,.035);
  --glass:rgba(255,255,255,.05); --glass-b:rgba(255,255,255,.09); --line:rgba(255,255,255,.08);
  --surface:#1c1c20; --sunk:rgba(255,255,255,.06);
  --ink:#f4f4f5; --ink-2:#c8c8d0; --muted:#a1a1ab;
  --accent:#58a6ff; --acc-rgb:88,166,255; --accent-soft:rgba(88,166,255,.13); --accent-ink:#a5d0ff;
  --acc-fort:#2563eb; --acc-fort2:#1d4ed8; --on-accent:#ffffff;
  --warm:#f7a84a; --warm-soft:rgba(247,168,74,.13); --warm-ink:#fbc98a;
  --good:#3ecf8e; --good-soft:rgba(62,207,142,.13); --bad:#ff9aa6; --bad-soft:rgba(255,154,166,.13);
  --sh:0 4px 28px rgba(0,0,0,.3);
  --sans:system-ui,-apple-system,"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif;
  --mono:ui-monospace,"SF Mono",Menlo,Consolas,"Liberation Mono",monospace;
  color-scheme:dark;
}
* { box-sizing:border-box; }
[hidden] { display:none !important; }
html { background:var(--bg); -webkit-tap-highlight-color:transparent; }
body { margin:0; min-height:100vh; color:var(--ink); font:14.5px/1.5 var(--sans); background:var(--bg-grad) fixed;
  overflow-x:hidden; -webkit-font-smoothing:antialiased; }
/* fons: dues llums suaus i una trama de punts que s'esvaeix cap avall */
body::before { content:''; position:fixed; top:-140px; left:-100px; width:520px; height:520px; pointer-events:none;
  background:radial-gradient(circle,var(--glow1) 0%,transparent 68%); }
body::after { content:''; position:fixed; bottom:-160px; right:-120px; width:600px; height:600px; pointer-events:none;
  background:radial-gradient(circle,var(--glow2) 0%,transparent 68%); }
main { position:relative; max-width:1120px; margin:0 auto; padding-block:24px 40px;
  padding-left:max(20px, env(safe-area-inset-left, 0px)); padding-right:max(20px, env(safe-area-inset-right, 0px));
  display:grid; gap:14px; }
main::before { content:''; position:fixed; inset:0; pointer-events:none;
  background-image:radial-gradient(rgba(255,255,255,.06) 1px,transparent 1px); background-size:24px 24px;
  -webkit-mask-image:linear-gradient(180deg,#000,transparent 65%); mask-image:linear-gradient(180deg,#000,transparent 65%); }
h1, h2 { text-wrap:balance; margin:0; }
h2 { font-size:15px; font-weight:650; letter-spacing:-.2px; }
a { color:var(--accent-ink); }
button { font:inherit; color:inherit; }
:focus-visible { outline:2px solid var(--accent); outline-offset:3px; border-radius:6px; }
.muted { color:var(--muted); }
code { font:12.5px var(--mono); background:var(--sunk); color:var(--ink-2); padding:1px 6px; border-radius:5px; white-space:nowrap; }
.glass, .panel { background:var(--glass); backdrop-filter:blur(22px); -webkit-backdrop-filter:blur(22px);
  border:1px solid var(--glass-b); border-radius:18px; box-shadow:var(--sh), inset 0 1px 0 rgba(255,255,255,.08); }
.panel { padding:20px 22px; min-width:0; display:grid; gap:14px; align-content:start; }
.panel-head { display:flex; flex-wrap:wrap; align-items:baseline; justify-content:space-between; gap:8px; }
.fu { animation:fu .45s ease both; }
@keyframes fu { from { opacity:0; transform:translateY(8px); } to { opacity:1; transform:none; } }

/* capçalera */
.hdr { display:flex; align-items:center; justify-content:space-between; gap:12px; flex-wrap:wrap; position:relative; z-index:5; }
.hdr-l { display:flex; align-items:center; gap:13px; min-width:0; }
.logo { width:42px; height:42px; border-radius:11px; flex-shrink:0; display:block; box-shadow:0 4px 16px rgba(var(--acc-rgb),.3); }
.logo.round { border-radius:50%; object-fit:cover; }
.logo-name { font-size:17px; font-weight:650; letter-spacing:-.3px; margin:0; }
.logo-sub { font-size:12.5px; color:var(--muted); }
.hdr-r { display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
.pill { display:flex; align-items:center; gap:7px; padding:7px 13px; border-radius:50px; font-size:12px; color:var(--ink-2);
  background:var(--glass); border:1px solid var(--glass-b); }
button.pill { cursor:pointer; transition:color .2s, border-color .2s; }
button.pill:hover:not(:disabled) { color:var(--ink); border-color:rgba(255,255,255,.2); }
button.pill:disabled { opacity:.45; cursor:default; }
.dot { width:8px; height:8px; border-radius:50%; flex-shrink:0; background:var(--muted); }
.dot.ok { background:var(--good); } .dot.err { background:var(--bad); }
#toolMsg:empty { display:none; }

/* resum: salutació, rellotge i avisos; a la dreta, les xifres */
.resum { display:grid; grid-template-columns:minmax(0,1.3fr) minmax(0,1fr); gap:22px; align-items:center; padding:26px 28px 24px; }
.hola { font-size:30px; font-weight:650; letter-spacing:-.8px; line-height:1.1; }
.clock { margin:8px 0 0; color:var(--ink-2); font-variant-numeric:tabular-nums; max-width:62ch; }
.clock b { color:var(--ink); }
.claim-now { display:inline-block; margin-top:8px; padding:3px 11px; border-radius:999px; background:var(--good-soft); color:var(--good); font-size:13px; font-weight:650; }
.alerts { list-style:none; padding:16px 0 0; margin:16px 0 0; border-top:1px solid var(--line); display:grid; gap:6px; max-height:260px; overflow:auto; }
.alerts li { display:flex; gap:10px; align-items:baseline; padding:8px 12px; border-radius:10px; background:rgba(255,255,255,.04); border:1px solid var(--line); font-size:13.5px; }
.fets { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:8px; align-content:center; }
.fet { padding:12px 14px; border-radius:12px; background:rgba(255,255,255,.05); border:1px solid var(--line); min-width:0; }
.fet-l { font-size:11.5px; color:var(--ink-2); }
.fet-v { font-size:24px; font-weight:650; letter-spacing:-.6px; margin-top:2px; font-variant-numeric:tabular-nums; }
.fet-s { font-size:12px; color:var(--muted); }
.tag { font-size:11px; font-weight:650; padding:2px 9px; border-radius:20px; white-space:nowrap; flex:none; background:rgba(255,255,255,.08); color:var(--ink-2); }
.tag.good { background:var(--good-soft); color:var(--good); }
.tag.bad { background:var(--bad-soft); color:var(--bad); }
.tag.accent { background:var(--accent-soft); color:var(--accent); }

/* pestanyes: una vista cada vegada */
.nav { position:sticky; top:10px; z-index:4; display:flex; gap:4px; padding:5px; border-radius:50px; width:fit-content; max-width:100%; overflow-x:auto;
  background:rgba(28,28,32,.75); backdrop-filter:blur(22px); -webkit-backdrop-filter:blur(22px); border:1px solid var(--glass-b); box-shadow:var(--sh); }
.nav button { padding:8px 16px; border-radius:50px; border:0; background:transparent; color:var(--muted); font-size:13.5px; font-weight:600; cursor:pointer; white-space:nowrap;
  transition:background .2s, color .2s; }
.nav button:hover { color:var(--ink-2); background:rgba(255,255,255,.06); }
.nav button[aria-selected="true"] { background:rgba(var(--acc-rgb),.14); color:var(--accent); }
.nav .n { font-size:11px; margin-left:6px; padding:1px 7px; border-radius:20px; background:var(--bad-soft); color:var(--bad); }
.view { display:grid; gap:14px; animation:fu .3s ease both; }
.g2 { display:grid; grid-template-columns:repeat(auto-fit, minmax(340px, 1fr)); gap:14px; align-items:start; }

/* avui: llista + $oc */
.grp { display:grid; }
.grp h3, .og h3 { margin:10px 0 2px; font-size:11.5px; font-weight:650; letter-spacing:1.2px; text-transform:uppercase; color:var(--muted); }
.grp.now h3 { color:var(--good); }
.task { display:flex; gap:12px; align-items:flex-start; padding:9px 0; border-bottom:1px solid var(--line); cursor:pointer; }
.task:last-child { border-bottom:0; }
.task input { width:17px; height:17px; margin-top:3px; accent-color:var(--accent); flex:none; }
.task > span { min-width:0; }
.task:has(input:checked) > span > :not(.when) { opacity:.45; text-decoration:line-through; }
.task .desc { display:block; color:var(--muted); font-size:13px; }
.task .when { display:block; font-size:12.5px; font-weight:650; color:var(--good); }
.bullet { width:17px; text-align:center; color:var(--muted); flex:none; }
details summary { cursor:pointer; color:var(--ink-2); font-size:13.5px; margin-top:6px; }
details[open] summary { margin-bottom:6px; }
.ocbox iframe { width:100%; height:880px; border:0; display:block; }

/* col·leccions: barres + àlbum */
.legend { display:flex; flex-wrap:wrap; gap:6px 16px; font-size:12.5px; color:var(--ink-2); align-items:center; }
.legend i { display:inline-block; width:10px; height:10px; border-radius:3px; margin-right:6px; vertical-align:-1px; }
.bars { display:grid; gap:2px; }
.row { display:grid; grid-template-columns:180px 1fr 110px; gap:14px; align-items:center; padding:7px 10px; margin:0 -10px; border-radius:10px; cursor:pointer; transition:background .2s; }
.row:hover { background:rgba(255,255,255,.05); }
.row.on { background:rgba(var(--acc-rgb),.10); }
.lbl { overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.val { text-align:right; font-variant-numeric:tabular-nums; color:var(--ink-2); font-size:13.5px; }
.val b { color:var(--ink); }
.track { display:flex; gap:2px; height:6px; background:rgba(255,255,255,.08); border-radius:99px; overflow:hidden; }
.seg.mine { background:var(--accent); } .seg.other { background:var(--warm); }
.album { border-top:1px solid var(--line); padding-top:14px; display:grid; gap:12px; }
.album-head { display:flex; flex-wrap:wrap; justify-content:space-between; align-items:baseline; gap:8px; }
.album-head h3 { margin:0; font-size:17px; font-weight:650; letter-spacing:-.3px; }
.chips { display:flex; flex-wrap:wrap; gap:6px; }
.ch { font-size:12.5px; padding:4px 11px; border-radius:20px; border:1px solid transparent; white-space:nowrap; }
.ch small { opacity:.75; margin-left:4px; }
.ch.wish { background:rgba(var(--acc-rgb),.06); border-color:var(--accent); color:var(--ink); font-weight:650; }
.ch.alta { background:rgba(255,255,255,.06); border-color:rgba(255,255,255,.18); color:var(--ink); font-weight:650; }
.ch.other { background:var(--warm-soft); color:var(--warm-ink); }
.ch.free { background:transparent; border:1px dashed rgba(255,255,255,.28); color:var(--ink-2); }
.ch.mine { background:var(--accent-soft); color:var(--accent-ink); }
.star { color:var(--accent); }
button.ch { font:inherit; font-size:12.5px; cursor:pointer; background:transparent; }
button.ch.pick:hover { border-color:var(--good); }
button.ch.sel { background:var(--good-soft); color:var(--good); border:1px solid var(--good); font-weight:650; }
.tab { font-size:12.5px; font-weight:600; padding:6px 13px; border-radius:20px; border:0; background:rgba(255,255,255,.07); color:var(--ink-2); cursor:pointer; transition:background .2s; }
.tab:hover { background:rgba(255,255,255,.12); color:var(--ink); }
.toggle { font-size:13px; color:var(--ink-2); display:flex; gap:6px; align-items:center; cursor:pointer; }
.toggle input { accent-color:var(--accent); }
textarea, select, input[type=text] { font:14px var(--sans); color:var(--ink); background:rgba(0,0,0,.22); border:1px solid var(--glass-b); border-radius:10px; }
textarea { padding:10px 12px; resize:vertical; width:100%; }
select, input[type=text] { padding:6px 10px; border-radius:8px; }
select option { background:var(--surface); }
textarea:focus, input[type=text]:focus, select:focus { outline:2px solid var(--accent); outline-offset:0; border-color:transparent; }
#mineOut { display:grid; gap:14px; }
.mg { display:grid; gap:6px; }
.mg-head { display:flex; flex-wrap:wrap; gap:8px; align-items:center; font-size:14px; }
.mg-head label { display:flex; gap:8px; align-items:center; cursor:pointer; font-weight:650; }
.mg-head input { accent-color:var(--accent); }
.mg.off .chips { opacity:.4; }
.give { display:grid; gap:8px; padding:14px; border-radius:12px; background:rgba(255,255,255,.04); border:1px solid var(--line); }
.give-row { display:flex; flex-wrap:wrap; gap:8px; align-items:center; font-size:13.5px; color:var(--ink-2); }
.give input[type=text] { font:13px var(--mono); width:160px; }
.give code { white-space:normal; overflow-wrap:anywhere; display:block; padding:8px 10px; background:rgba(0,0,0,.22); }
.serieslist { display:flex; flex-wrap:wrap; gap:6px; }
.serieslist button { font-size:12.5px; padding:3px 11px; border-radius:20px; border:1px solid var(--glass-b); background:rgba(255,255,255,.04); color:var(--ink-2); cursor:pointer; }
.serieslist button:hover { border-color:var(--accent); color:var(--accent-ink); }
.serieslist span { color:var(--muted); font-variant-numeric:tabular-nums; }
.unk { list-style:none; margin:0; padding:0; display:flex; flex-wrap:wrap; gap:6px; }
#cls input[type=text] { min-width:0; width:220px; }
#wlBuild { display:grid; gap:10px; }
#wlBuild .ch { cursor:pointer; }
.ogrid { display:grid; grid-template-columns:repeat(auto-fit, minmax(300px, 1fr)); gap:10px 28px; }
.oline { display:flex; gap:10px; align-items:baseline; padding:4px 0; font-size:13.5px; color:var(--ink-2); }
.oline .cmd { flex:none; }

/* wishlist, sobrants, intercanvis */
table { width:100%; border-collapse:collapse; font-size:13.5px; }
td { padding:8px 0; border-bottom:1px solid var(--line); vertical-align:baseline; }
tr:last-child td { border-bottom:0; }
td + td { padding-left:12px; }
td.num { text-align:right; font-variant-numeric:tabular-nums; color:var(--ink-2); }
td.sub { color:var(--muted); }
.owners { display:grid; grid-template-columns:repeat(auto-fit, minmax(240px, 1fr)); gap:12px; }
.owner { padding:14px 16px; border-radius:14px; background:rgba(255,255,255,.04); border:1px solid var(--line); min-width:0; }
.owner h3 { margin:0 0 6px; font-size:14px; font-weight:650; overflow-wrap:anywhere; }
.owner ul { list-style:none; margin:0; padding:0; display:grid; gap:2px; }
.owner li { display:flex; gap:8px; align-items:baseline; justify-content:space-between; padding:4px 0; font-size:13.5px; }
.owner label { display:flex; gap:8px; align-items:baseline; cursor:pointer; min-width:0; }
.owner input { accent-color:var(--accent); }
.status { font-size:11.5px; color:var(--muted); white-space:nowrap; }
.cmd, .copybtn { font:12.5px var(--mono); background:rgba(var(--acc-rgb),.08); color:var(--accent-ink); padding:2px 8px; border-radius:7px;
  border:1px solid rgba(var(--acc-rgb),.2); cursor:copy; white-space:nowrap; transition:border-color .2s, background .2s; }
.cmd:hover, .copybtn:hover { border-color:var(--accent); background:rgba(var(--acc-rgb),.14); }
.copybtn { font:650 11.5px var(--sans); margin-left:8px; padding:2px 10px; border-radius:20px; vertical-align:1px; }
.cmd.done, .copybtn.done { background:var(--good-soft); color:var(--good); border-color:transparent; }
.recvbar { display:flex; flex-wrap:wrap; gap:10px; align-items:center; color:var(--ink-2); font-size:13.5px; margin:0; }
.foot { padding:12px 18px; text-align:center; font-size:12px; color:var(--muted); }
.panel > div:empty { display:none; }
.foot a { color:var(--ink-2); font-weight:650; text-decoration:none; } .foot a:hover { color:var(--accent); }
button.primary { font-size:13px; font-weight:650; padding:9px 16px; border-radius:10px; border:0; cursor:pointer; color:var(--on-accent);
  background:linear-gradient(135deg,var(--acc-fort),var(--acc-fort2)); box-shadow:0 2px 10px rgba(var(--acc-rgb),.22); transition:filter .2s, transform .1s; }
button.primary:hover:not(:disabled) { filter:brightness(1.12); }
button.primary:active:not(:disabled), .tab:active, .nav button:active { transform:scale(.97); }
button.primary:disabled { opacity:.35; cursor:default; box-shadow:none; }
button, a, label, .row, .task { touch-action:manipulation; }
#tip { position:fixed; pointer-events:none; padding:5px 10px; border-radius:8px; font-size:12px; color:var(--ink-2); background:var(--surface);
  border:1px solid var(--glass-b); box-shadow:0 4px 14px rgba(0,0,0,.35); display:none; max-width:280px; z-index:9; }
/* llum que segueix el ratolí per sobre de les targetes (no al mòbil) */
@media (hover:hover) and (pointer:fine) {
  .panel, .resum { position:relative; }
  .panel::after, .resum::after { content:''; position:absolute; inset:0; border-radius:inherit; pointer-events:none; opacity:0; transition:opacity .3s;
    background:radial-gradient(280px circle at var(--mx,50%) var(--my,50%),rgba(var(--acc-rgb),.07),transparent 70%); }
  .panel:hover::after, .resum:hover::after { opacity:1; }
}
@media (max-width:900px) { .resum { grid-template-columns:1fr; } }
@media (max-width:640px) {
  main { padding-left:max(16px, env(safe-area-inset-left, 0px)); padding-right:max(16px, env(safe-area-inset-right, 0px)); }
  .resum { padding:20px 18px; } .hola { font-size:26px; } .panel { padding:16px; }
  .g2 { grid-template-columns:1fr; }
  .row { grid-template-columns:1fr auto; } .row .track { grid-column:1 / -1; grid-row:2; }
}
@media (prefers-reduced-motion:reduce) { *, *::before, *::after { animation:none !important; transition:none !important; } }
"""


ORDRES = [  # (grup, [(ordre, què fa)]): requadre «Ordres d'ús habitual» del tauler
    ("Tirades", [("$m", "Tirada: personatge qualsevol"), ("$us", "Gasta tirades guardades"), ("$rolls", "Reinici de tirades del vot"),
                 ("$daily", "Reinici de tirades (cada 20 h)"), ("$vote", "Vota i guanya un reinici")]),
    ("Reclamar i kakera", [("$tu", "Tot el que tens disponible"), ("$rt", "Torna a poder reclamar"), ("$dk", "Recarrega el poder i dona kakera"),
                           ("$mu", "Quan podràs reclamar"), ("$k", "La teva kakera i com guanyar-ne")]),
    ("La teva col·lecció", [("$mms", "El teu harem sencer, per privat"), ("$mmbs", "El harem amb la sèrie de cada un, per privat"),
                            ("$imaso <sèrie>", "Una sèrie sencera amb propietaris, per privat"), ("$im <personatge>", "Fitxa d'un personatge"),
                            ("$wl", "La teva wishlist"), ("$wish <nom>", "Afegeix a la wishlist (diversos: separa'ls amb $)"),
                            ("$wishremove <nom>", "Treu de la wishlist"), ("$sw <nom>", "Desig estrella")]),
    ("Intercanvis", [("$give @persona", "Regala personatges (separa'ls amb $)"), ("$trade @persona", "Intercanvi amb confirmació dels dos")]),
    ("Esferes", [("$oh", "Esferes (4 al dia)"), ("$oc", "Troba la vermella: fes servir l'ajudant"), ("$oq", "Joc d'esferes"), ("$ot", "Joc d'esferes")]),
    ("Torre (kakeratower)", [("$kt", "Pisos, beneficis i preu del següent"), ("$towervalue", "Quant costa cada pis"),
                             ("$build 8", "Pis recomanat: reaccions a kakera −4 % de poder"), ("$build 11", "Pis: +1 tirada per hora"),
                             ("$build <número>", "Construeix el pis que triïs"), ("$destroy", "Desfà la torre i et torna la kakera")]),
    ("Kakeraloots", [("$kl", "Un kakeraloot (500 de kakera)"), ("$kl <número>", "Diversos alhora"), ("$lk", "Les teves estadístiques i inventari"),
                     ("$infokl", "Què donen els kakeraloots"), ("$quantity <número>", "Puja nivells de quantitat"), ("$quality <número>", "Puja nivells de qualitat")]),
    ("Kakera, insígnies i claus", [("$kakerareward", "Recompenses de la kakera"), ("$badges", "Les teves insígnies"), ("$bronze", "Insígnia de bronze (i $silver, $gold…)"),
                                   ("$bku", "Bonus de les claus daurades"), ("$infokeys", "Com funcionen les claus"), ("$givekakera @persona <número>", "Regala kakera")]),
    ("Admin del servidor", [("$forcedivorce <nom>", "Allibera un personatge"), ("$restorelist", "Personatges alliberats per admins"),
                            ("$restore <nom>", "Torna un personatge alliberat")]),
]


JS_MINE = r"""
// Cercador (text cru a Python: les barres invertides són del JavaScript): troba els teus personatges de les sèries que enganxis (tolera faltes) i prepara el $give
const MINE = __MINE__, COLLS = new Set(__COLLS__), ALTRES = __ALTRES__, BYNAME = '\u0000byname';
const q = document.getElementById('mineQ'), out = document.getElementById('mineOut');
const norm = t => t.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
function lev(a, b) {
  const d = Array.from({ length: a.length + 1 }, (_, i) => [i]);
  for (let j = 1; j <= b.length; j++) d[0][j] = j;
  for (let i = 1; i <= a.length; i++) for (let j = 1; j <= b.length; j++)
    d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
  return d[a.length][b.length];
}
const SERIES = [...new Set(MINE.map(c => c.s))];
const label = s => s === ALTRES ? T(ALTRES) : s === BYNAME ? T('Per nom') : s;
// una sèrie es pot dir de diverses maneres: "Naruto / Boruto", "Touhou Koumakyou ~ the Embodiment of Scarlet Devil"
const parts = s => [s, ...s.split(/[\/()~:]/)].map(norm).filter(p => p.length >= 2);
function matches(term, serie) {
  const t = norm(term); if (t.length < 3 || serie === ALTRES) return false;
  return parts(serie).some(p => p.includes(t) || (p.length >= 5 && t.includes(p)) || lev(t, p) <= Math.max(1, Math.floor(p.length * 0.2)));
}
const off = new Set();  // sèries desmarcades; les teves col·leccions comencen desmarcades
const isOff = s => COLLS.has(s) ? !off.has('+' + s) : off.has(s);
let lastChosen = [], lastWho = '';
function render() {
  const terms = q.value.split(/[\/,;\n]+/).map(t => t.trim()).filter(Boolean);
  if (!terms.length) {
    const counts = SERIES.map(s => [s, MINE.filter(c => c.s === s).length]).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
    out.innerHTML = '<p class="eyebrow" style="margin:0">' + esc(T('Les teves sèries · clica per afegir-la')) + '</p><div class="serieslist">'
      + counts.map(([s, n]) => `<button type="button" data-add="${esc(s)}">${esc(label(s))} <span>${n}</span></button>`).join('') + '</div>';
    return;
  }
  const groups = new Map(), byName = [], missing = [];
  for (const t of terms) {
    const ss = SERIES.filter(s => matches(t, s));
    ss.forEach(s => groups.set(s, MINE.filter(c => c.s === s)));
    const names = ss.length || norm(t).length < 3 ? [] : MINE.filter(c => norm(c.n).includes(norm(t)));
    names.forEach(c => byName.includes(c) || byName.push(c));
    if (!ss.length && !names.length) missing.push(t);
  }
  if (byName.length) groups.set(BYNAME, byName);
  const chosen = [];
  let html = '';
  for (const [s, list] of groups) {
    const o = isOff(s);
    if (!o) list.forEach(c => chosen.includes(c.n) || chosen.push(c.n));
    html += `<div class="mg${o ? ' off' : ''}"><div class="mg-head"><label><input type="checkbox" data-grp="${esc(s)}" ${o ? '' : 'checked'}>${esc(label(s))} <span class="muted">· ${list.length}</span></label>`
      + (COLLS.has(s) ? '<span class="tag accent">' + esc(T('la col·lecciones tu')) + '</span>' : '')
      + `</div><div class="chips">${list.map(c => `<span class="ch mine">${esc(c.n)}</span>`).join('')}</div></div>`;
  }
  if (!groups.size) html += '<p class="muted" style="margin:0">' + esc(T("Cap dels teus personatges és d'aquestes sèries.")) + '</p>';
  if (missing.length) html += '<p class="muted" style="margin:0;font-size:14px">' + esc(T('Sense cap personatge teu: {list}', { list: missing.join(' · ') })) + '</p>';
  if (chosen.length) {
    let who = '@' + T('persona'); try { who = localStorage.getItem('mudae-give-to') || who; } catch (e) {}
    const cmd = `$give ${who} ${chosen.join(' $ ')}`;
    lastChosen = chosen; lastWho = who;
    html += `<div class="give"><div class="give-row"><b>${esc(T('Ordre per regalar-los ({n})', { n: chosen.length }))}</b> · ${esc(T('per a'))} <input type="text" id="giveTo" spellcheck="false" autocomplete="off" value="${esc(who)}" aria-label="${esc(T('Per a qui'))}">`
      + `<button class="copybtn" data-copy="${esc(cmd)}">${esc(T('Copiar ordre'))}</button></div><code>${esc(cmd)}</code>`
      + '<span class="muted" style="font-size:13px">' + esc(T("Un regal no es pot desfer: revisa la llista abans de confirmar a Mudae. Si en són molts, potser l'hauràs de dividir.")) + '</span>'
      + '<div class="give-row"><span>' + esc(T("Quan Mudae t'hagi confirmat el regal:")) + '</span><button class="primary" id="gaveBtn">' + esc(T('Ja els he donat')) + '</button><span id="gaveMsg" role="status"></span></div></div>';
  }
  out.innerHTML = html;
}
q.addEventListener('input', render);
out.addEventListener('change', ev => {
  const g = ev.target.dataset.grp;
  if (g !== undefined) {
    const key = COLLS.has(g) ? '+' + g : g, wantOff = !ev.target.checked;
    if (COLLS.has(g) ? !wantOff : wantOff) off.add(key); else off.delete(key);
    render();
  }
  if (ev.target.id === 'giveTo') { try { localStorage.setItem('mudae-give-to', ev.target.value.trim()); } catch (e) {} render(); }
});
out.addEventListener('click', async ev => {
  if (ev.target.id === 'gaveBtn') {  // treu-los del teu harem (i, si eren d'una col·lecció teva, posa'ls a nom de qui els rep)
    const btn = ev.target, msg = document.getElementById('gaveMsg');
    if (location.protocol === 'file:') { msg.innerHTML = OFF_MSG(); return; }
    if (!confirm(T('Marcar {n} personatges com a regalats a {who}? Sortiran del teu harem.', { n: lastChosen.length, who: lastWho }))) return;
    btn.disabled = true; msg.textContent = T('Desant…');
    try {
      const j = await postJSON('/donat', { names: lastChosen, to: lastWho });
      if (j.ok) { q.value = ''; location.reload(); } else { msg.innerHTML = ERR_MSG(j.error); btn.disabled = false; }
    } catch (err) { msg.innerHTML = NET_MSG(); btn.disabled = false; }
    return;
  }
  const b = ev.target.closest('[data-add]'); if (!b) return;
  q.value = (q.value.trim() ? q.value.trim() + ' / ' : '') + b.dataset.add; render(); q.focus();
});
render();
"""


JS_UPD = r"""
// «Posar al dia»: enganxes el que diu Mudae i el servidor ho desa a dades/ i refà l'Excel i el tauler
const META = __META__, SPECIALS = __SPECIALS__;
const kindSel = document.getElementById('updKind'), serieBox = document.getElementById('updSerieBox'), serieSel = document.getElementById('updSerie');
const newIn = document.getElementById('updNew'), txt = document.getElementById('updText'), updBtn = document.getElementById('updBtn');
const updMsg = document.getElementById('updMsg'), updCmd = document.getElementById('updCmd'), age = document.getElementById('updAge');
const days = iso => iso ? Math.floor((Date.now() - new Date(iso + 'T00:00:00')) / 864e5) : null;
const fa = n => n === null ? T('mai') : n === 0 ? T('avui') : n === 1 ? T('ahir') : T('fa {n} dies', { n });
(function () {
  const h = days(META.harem), w = days(META.wishlist);
  const ser = Object.entries(META.series || {}).sort((a, b) => a[1].localeCompare(b[1]))[0];
  age.textContent = T('Harem: {h} · Wishlist: {w}', { h: fa(h), w: fa(w) }) + (ser ? ' · ' + T('Sèrie més antiga: {s}, {d}', { s: ser[0], d: fa(days(ser[1])) }) : '');
  if (h === null || h >= 7) age.innerHTML = '<span class="tag bad">' + esc(T('Toca posar al dia el harem: $mms')) + '</span> ' + esc(age.textContent);
})();
const serieName = () => serieSel.value === '__new__' ? newIn.value.trim() : serieSel.value;
function syncUpd() {
  serieBox.hidden = kindSel.value !== 'serie';
  newIn.hidden = serieSel.value !== '__new__';
  const esp = kindSel.value === 'serie' && SPECIALS.includes(serieName());
  document.getElementById('updEsp').hidden = !esp;
  const cmd = kindSel.value === 'harem' ? '$mms' : kindSel.value === 'wishlist' ? '$wl' : esp ? '$imaso ' : '$imaso ' + (serieName() || '<' + T('sèrie') + '>');
  updCmd.textContent = cmd; updCmd.dataset.copy = cmd;
  updBtn.disabled = location.protocol === 'file:' || !txt.value.trim() || (kindSel.value === 'serie' && !serieName());
}
[kindSel, serieSel, newIn, txt].forEach(el => el.addEventListener('input', syncUpd));
if (location.protocol === 'file:') updMsg.innerHTML = OFF_MSG();
updBtn.onclick = async () => {
  updBtn.disabled = true; updMsg.textContent = T('Desant…');
  try {
    const j = await postJSON('/dades', { kind: kindSel.value, serie: serieName(), text: txt.value });
    if (j.ok) location.reload(); else { updMsg.innerHTML = ERR_MSG(j.error); syncUpd(); }
  } catch (err) { updMsg.innerHTML = NET_MSG(); syncUpd(); }
};
syncUpd();
"""


JS_ESP = r"""
// «Col·lecció especial»: un grup de personatges triat per tu; si escrius el nom d'una que ja tens, s'hi carreguen els seus
const espName = document.getElementById('espName'), espText = document.getElementById('espText');
const espBtn = document.getElementById('espBtn'), espMsg = document.getElementById('espMsg');
const espSync = () => { espBtn.disabled = location.protocol === 'file:' || !espName.value.trim() || !espText.value.trim(); };
espName.addEventListener('input', () => {
  const n = espName.value.trim();
  if (SPECIALS.includes(n) && COLL[n] && !espText.value.trim()) espText.value = COLL[n].map(c => c.n).join('\n');
  espSync();
});
espText.addEventListener('input', espSync);
if (location.protocol === 'file:') espMsg.innerHTML = OFF_MSG();
espBtn.onclick = async () => {
  espBtn.disabled = true; espMsg.textContent = T('Desant…');
  try {
    const j = await postJSON('/especial', { name: espName.value.trim(), text: espText.value });
    if (j.ok) { try { localStorage.setItem('mudae-coll', espName.value.trim()); } catch (e) {} location.reload(); }
    else { espMsg.innerHTML = ERR_MSG(j.error); espSync(); }
  } catch (err) { espMsg.innerHTML = NET_MSG(); espSync(); }
};
espSync();
"""


JS_CLS = r"""
// «Sèrie dels teus personatges»: desa la sèrie d'un personatge (servidor) i refà el tauler
const clsName = document.getElementById('clsName'), clsSerie = document.getElementById('clsSerie');
const clsBtn = document.getElementById('clsBtn'), clsMsg = document.getElementById('clsMsg');
const clsSync = () => { clsBtn.disabled = location.protocol === 'file:' || !clsName.value.trim() || !clsSerie.value.trim(); };
[clsName, clsSerie].forEach(el => el.addEventListener('input', clsSync));
document.addEventListener('click', ev => {
  const b = ev.target.closest('[data-fill]'); if (!b) return;
  clsName.value = b.dataset.fill; clsSync(); clsSerie.focus();
});
clsBtn.onclick = async () => {
  clsBtn.disabled = true; clsMsg.textContent = T('Desant…');
  try {
    const j = await postJSON('/classifica', { name: clsName.value.trim(), serie: clsSerie.value.trim() });
    if (j.ok) location.reload(); else { clsMsg.innerHTML = ERR_MSG(j.error); clsSync(); }
  } catch (err) { clsMsg.innerHTML = NET_MSG(); clsSync(); }
};
if (location.protocol === 'file:') clsMsg.innerHTML = OFF_MSG();
clsSync();
"""


JS_WL = r"""
// Muntador de wishlist: què treure (ja és teu o d'algú altre) i què posar-hi (lliures de prioritat alta, per ordre)
const WL = __WL__, MAX = 15;
const wlBox = document.getElementById('wlBuild');
const bad = WL.wl.filter(w => !w.ok), keep = WL.wl.filter(w => w.ok);
const free = MAX - keep.length;
const pick = new Set(WL.cand.slice(0, free).map(c => c.n));
const note = s => '<p class="muted" style="margin:0;font-size:14px">' + s + '</p>';
function drawWl() {
  let html = '';
  if (bad.length) {
    const cmd = '$wishremove ' + bad.map(w => w.n).join(' $ ');
    html += `<div class="give"><div class="give-row"><b>${esc(T('1. Treu-ne {n}', { n: bad.length }))}</b><button class="copybtn" data-copy="${esc(cmd)}">${esc(T('Copiar ordre'))}</button></div><code>${esc(cmd)}</code></div>`;
  }
  if (free > 0) {
    const chosen = WL.cand.filter(c => pick.has(c.n)).map(c => c.n);
    const cmd = '$wish ' + chosen.join(' $ ');
    html += `<div class="give"><div class="give-row"><b>${bad.length ? '2. ' : ''}${esc(T(free === 1 ? 'Omple {n} lloc' : 'Omple {n} llocs', { n: free }))}</b>`
      + `<span class="muted">${esc(T("tria'n {n}: {c} de {n}", { n: free, c: chosen.length }))}</span>`
      + (chosen.length ? `<button class="copybtn" data-copy="${esc(cmd)}">${esc(T('Copiar ordre'))}</button>` : '') + '</div>'
      + `<div class="chips">${WL.cand.slice(0, 40).map(c => `<label class="ch ${pick.has(c.n) ? 'wish' : 'free'}" data-tip="${esc(c.s)}">`
        + `<input type="checkbox" data-wlc="${esc(c.n)}" ${pick.has(c.n) ? 'checked' : ''} hidden>${esc(c.n)}</label>`).join('')}</div>`
      + (chosen.length ? `<code>${esc(cmd)}</code>` : '')
      + (chosen.length > free ? `<span class="tag bad">${esc(T("N'has triat {n} de més", { n: chosen.length - free }))}</span>` : '')
      + '</div>';
  }
  if (!WL.wl.some(w => w.star && w.ok))
    html += note(esc(T('No tens desig estrella: posa-hi el que vulguis més amb')) + ` <button class="cmd" data-copy="$sw ">$sw &lt;${esc(T('nom'))}&gt;</button>`);
  html += note(esc(bad.length || free > 0 ? T('Quan ho hagis fet, enganxa el teu $wl a «Posar al dia».') : T('La wishlist està plena i tot està lliure: res a fer.')));
  wlBox.innerHTML = html;
}
wlBox.addEventListener('change', ev => {
  const n = ev.target.dataset.wlc; if (n === undefined) return;
  ev.target.checked ? pick.add(n) : pick.delete(n); drawWl();
});
drawWl();
"""


PLACEHOLDER = re.compile(r"<([^>]+)>|@persona")


def cmd_text(ordre):
    """Ordre per mostrar: tradueix «<nom>», «@persona» i el que no és una ordre («clic al cor»)."""
    return " · ".join(
        PLACEHOLDER.sub(lambda m: f"<{t(m.group(1))}>" if m.group(1) else "@" + t("persona"), p) if p.startswith("$") else t(p)
        for p in ordre.split(" · "))


def cmd_html(ordre):
    """`$us · $rolls` -> un botó de copiar per ordre; el que no és ordre queda com a text."""
    out = []
    for part, shown in zip(ordre.split(" · "), cmd_text(ordre).split(" · ")):
        if part.startswith("$"):
            text = re.sub(r"\s*\(.*?\)", "", part)            # "$oh (×4)" -> "$oh"
            hole = re.search(r"<[^>]*>|@\S*", text)               # "$wishremove <nom>" -> "$wishremove " (hi escrius el nom)
            text = re.sub(r"<[^>]*>|@\S*", "", text).rstrip() + (" " if hole else "")
            out.append(f'<button class="cmd" data-copy="{e(text)}" data-tip="{e(t("Copiar"))}">{e(shown)}</button>')
        else:
            out.append(f"<code>{e(shown)}</code>")
    return " ".join(out)


def group_title(g):
    if g == "reclam":
        return t("MOMENT BO: pots reclamar ({hours})", hours=", ".join(f"{h:02d}:02" for h in CLAIM_HOURS))
    return t(GRUPS[g])


def give_cmd(ts):
    me = f"<@{DISCORD_ID}>" if DISCORD_ID else f"@{ME}"
    return f"$give {me} " + " $ ".join(tr["name"] for tr in ts)


def write_dashboard(d, hist):
    rows, series = d["rows"], d["series"]
    owned = sum(s["Meu"] for s in series)
    total = sum(s["total"] for s in series)
    pct = round(100 * owned / total) if total else 0
    wl_bad = [w for w in d["wishlist"] if not w["ok"]]
    pend = sum(tr["status"] in ("Pendent", "Enviat") for tr in d["trades"])

    def alert(tag, cls, text):
        return f'<li><span class="tag {cls}">{e(tag)}</span><span>{text}</span></li>'
    alerts = [alert(t("Wishlist"), "bad", f'<b>{e(w["name"])}</b>: {e(w["todo"])}') for w in wl_bad]
    if len(d["wishlist"]) < 15:
        alerts.append(alert(t("Wishlist"), "accent", e(t("Tens {n} llocs lliures. Mira «Per omplir» a la wishlist.", n=15 - len(d["wishlist"])))))
    for kind, r, prev in d["changes"]:
        if kind == "nou":
            continue
        if r["owner"] == ME:
            alerts.append(alert(t("Aconseguit"), "good", f'<b>{e(r["name"])}</b> · {e(r["serie"])}'))
        elif not r["owner"]:
            alerts.append(alert(t("Lliure"), "accent", e(t("{name} ha quedat lliure (era de {prev}): posa'l a la wishlist", name=r["name"], prev=prev))))
        else:
            alerts.append(alert(t("Canvi"), "", e(t("{name} ara és de {owner} (abans: {prev})", name=r["name"], owner=r["owner"], prev=prev or t("lliure")))))
    news = sum(k == "nou" for k, _, _ in d["changes"])
    if news:
        alerts.append(alert(t("Nous"), "", e(t("{n} personatges nous a les teves sèries", n=news))))
    if d["missing"]:
        alerts.append(alert(t("Harem"), "", e(t("{n} personatges teus no surten al harem: torna a enganxar el $mms a «Posar al dia»", n=len(d["missing"])))))

    spec = specials()
    esp_tag = f' <span class="tag accent">{e(t("especial"))}</span>'
    bars = []
    for s in series:
        p = s["Meu"] / s["total"] * 100
        o = s["D'algú altre"] / s["total"] * 100
        tip = t("{serie}: {mine} teus · {other} d'algú altre · {free} lliures", serie=s["name"], mine=s["Meu"], other=s["D'algú altre"], free=s["Lliure"])
        bars.append(
            f'<div class="row" data-serie="{e(s["name"])}" data-tip="{e(tip)}" role="button" tabindex="0">'
            f'<div class="lbl">{e(s["name"])}{esp_tag if s["name"] in spec else ""}</div>'
            f'<div class="track"><div class="seg mine" style="width:{p:.2f}%"></div>'
            f'<div class="seg other" style="width:{o:.2f}%"></div></div>'
            f'<div class="val">{s["Meu"]}/{s["total"]} · <b>{round(p)}%</b></div></div>')

    wl = "".join(
        f'<tr><td>{"<span class=star>★</span> " if w["star"] else ""}{e(w["name"])}</td><td class="sub">{e(w["serie"])}</td>'
        f'<td style="text-align:right"><span class="tag {"good" if w["ok"] else "bad"}">'
        f'{e(t("Lliure")) if w["ok"] else e(w["todo"])}</span></td></tr>'
        for w in d["wishlist"])
    wl_json = json.dumps({"wl": [{"n": w["name"], "star": w["star"], "ok": w["ok"]} for w in d["wishlist"]],
                          "cand": [{"n": r["name"], "s": r["serie"]} for r in d["free_alta"]]}, ensure_ascii=False).replace("</", r"<\/")

    by_owner = {}
    for tr in d["trades"]:
        by_owner.setdefault(tr["owner"], []).append(tr)
    trades = "".join(
        f'<div class="owner"><h3>{e(o)} <span class="muted">· {len(ts)}</span>'
        f'<button class="copybtn" data-copy="{e(give_cmd(ts))}" data-tip="{e(t("Ordre perquè {o} te la passi", o=o))}">{e(t("Copiar ordre"))}</button></h3><ul>'
        + "".join(f'<li><label><input type="checkbox" class="recv" value="{e(tr["name"])}">'
                  f'<span>{e(tr["name"])} <span class="muted">· {e(tr["serie"])}</span></span></label>'
                  f'<span class="status">{e(t(tr["status"]))}</span></li>' for tr in ts)
        + "</ul></div>"
        for o, ts in sorted(by_owner.items(), key=lambda kv: -len(kv[1])))

    sizes = sorted(d["sob_size"].items(), key=lambda kv: (kv[0] == ALTRES, -kv[1], kv[0]))
    sob = "".join(f"<tr><td>{e(t(s))}</td><td class='num'>{n}</td></tr>" for s, n in sizes[:12])
    if len(sizes) > 12:
        sob += f"<tr><td class='sub'>{e(t('i {n} sèries més, al full «Sobrants» de l’Excel', n=len(sizes) - 12))}</td><td></td></tr>"

    # Llista d'avui: cada tasca marcada queda marcada fins que es pot tornar a fer (data-torna)
    periodic, other, prev = [], [], None
    for i, (g, que, ordre, com, torna) in enumerate(RUTINA):
        out = periodic if g in ("dia", "hora", "reclam") else other
        if g != prev:
            out.append(f'</div><div class="grp" data-period="{g}"><h3>{e(group_title(g))}</h3>')
            prev = g
        box = (f'<input type="checkbox" data-id="r{i}" data-torna="{torna}">' if torna is not None
               else '<span class="bullet">·</span>')
        out.append(f'<label class="task">{box}<span><b>{e(t(que))}</b> {cmd_html(ordre)}'
                   f'<span class="desc">{e(t(com, rt=RT_HORES))}</span><span class="when"></span></span></label>')
    checklist = ("<div>" + "".join(periodic) + "</div>"
                 + f'<details><summary>{e(t("Altres moments"))}</summary><div>' + "".join(other) + "</div></details>")

    wl_names = {w["name"] for w in d["wishlist"]}
    coll = {s["name"]: [] for s in series}
    for r in rows:
        coll[r["serie"]].append({"n": r["name"], "o": r["owner"], "st": r["state"], "prio": r["prio"], "wl": r["name"] in wl_names})
    js = lambda v: json.dumps(v, ensure_ascii=False).replace("</", r"<\/")  # que un nom no pugui tancar el <script>
    meus = {r["name"]: r["serie"] for r in rows if r["state"] == "Meu"}
    serie_of = {**meus, **{x["name"]: x["serie"] for x in d["sobrants"]}}
    mine = [{"n": n, "s": serie_of.get(n, ALTRES)} for n in d["harem"]]
    unknown = [x["name"] for x in d["sobrants"] if x["serie"] == ALTRES]
    all_series = sorted({x["serie"] for x in d["sobrants"] if x["serie"] != ALTRES})
    cls_rows = "".join(
        f'<li><button class="cmd" data-copy="$im {e(n)}" data-fill="{e(n)}" data-tip="{e(t("Copia l’ordre i posa el nom al formulari"))}">$im {e(n)}</button></li>'
        for n in unknown)
    cls_html = (f'<p class="muted" style="margin:0;font-size:14px">{e(t("{n} personatges teus sense sèrie. Copia el $im, mira a Discord de quina sèrie és i desa-la aquí sota.", n=len(unknown)))}</p>'
                f'<ul class="unk">{cls_rows}</ul>' if unknown
                else f'<p class="muted" style="margin:0;font-size:14px">{e(t("Tots els teus personatges tenen sèrie."))}</p>')
    datalists = ('<datalist id="dlNames">' + "".join(f'<option value="{e(x["name"])}">' for x in d["sobrants"]) + '</datalist>'
                 + '<datalist id="dlSeries">' + "".join(f'<option value="{e(x)}">' for x in all_series) + '</datalist>')
    serie_opts = "".join(f'<option value="{e(x["name"])}">{e(x["name"])}</option>' for x in series)
    ordres = "".join(
        f'<div class="og"><h3>{e(t(g))}</h3>' + "".join(f'<div class="oline">{cmd_html(c)}<span>{e(t(desc))}</span></div>' for c, desc in items) + "</div>"
        for g, items in ORDRES)
    oc = f"oc.html?lang={LANG}"
    avatar = DIR / "icones" / "mudae_rodo.png"  # no va al repo ni al .zip (no és nostre): qui no el té veu el logo propi
    if avatar.exists():
        src = "data:image/png;base64," + base64.b64encode(avatar.read_bytes()).decode()
        logo = f'<img class="logo round" src="{src}" alt="" width="42" height="42">'
        favicon = f'<link rel="icon" href="{src}">'
    else:
        logo = ('<svg class="logo" viewBox="0 0 42 42" aria-hidden="true"><defs><linearGradient id="lg" x1="0" y1="0" x2="1" y2="1">'
                '<stop offset="0" stop-color="#58a6ff"/><stop offset="1" stop-color="#1d4ed8"/></linearGradient></defs>'
                '<rect width="42" height="42" rx="11" fill="url(#lg)"/><path d="M21 8l10 10-10 16-10-16z" fill="#fff" fill-opacity=".92"/>'
                '<path d="M11 18h20M21 8l-3.5 10L21 34l3.5-16z" fill="none" stroke="#1d4ed8" stroke-opacity=".35"/></svg>')
        favicon = ""
    T_ = lambda s, **kw: e(t(s, **kw))  # text traduït i escapat, per a l'HTML

    page = f"""<!doctype html>
<html lang="{LANG}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="color-scheme" content="dark"><meta name="theme-color" content="#141417">
<title>{e(TITLE_TEXT)}</title>{favicon}
<style>{STYLE}</style></head><body><main>
<header class="hdr fu">
 <div class="hdr-l">
  {logo}
  <div><h1 class="logo-name">{e(TITLE_TEXT)}</h1><div class="logo-sub">{T_("Dades del {date}", date=f"{dt.datetime.now():%d/%m/%Y %H:%M}")}</div></div>
 </div>
 <div class="hdr-r">
  <span class="pill" id="toolMsg" role="status"></span>
  <button type="button" class="pill" id="lnkBtn">{T_("Crear accés directe a l'escriptori")}</button>
  <button type="button" class="pill" id="stopBtn"><span class="dot ok"></span>{T_("Tancar el servidor")}</button>
 </div>
</header>
<section class="resum glass fu" aria-labelledby="hola">
 <div>
  <h2 class="hola" id="hola">{T_("Hola, {name}", name=ME) if ME else T_("Hola")}</h2>
  <p id="clock" class="clock"></p>
  {f'<ul class="alerts" aria-label="{T_("Avisos")}">{"".join(alerts)}</ul>' if alerts else ""}
 </div>
 <div class="fets">
  <div class="fet"><div class="fet-l">{T_("Col·lecció")}</div><div class="fet-v">{pct}%</div><div class="fet-s">{T_("{a} de {b}", a=owned, b=total)}</div></div>
  <div class="fet"><div class="fet-l">{T_("Harem")}</div><div class="fet-v">{len(d["harem"])}</div><div class="fet-s">{T_("{n} sobrants", n=len(d["sobrants"]))}</div></div>
  <div class="fet"><div class="fet-l">{T_("Wishlist")}</div><div class="fet-v">{len(d["wishlist"])}/15</div><div class="fet-s">{T_("tot bé") if not wl_bad else T_("{n} per arreglar", n=len(wl_bad))}</div></div>
  <div class="fet"><div class="fet-l">{T_("Intercanvis")}</div><div class="fet-v">{len(d["trades"])}</div><div class="fet-s">{T_("{n} oberts", n=pend)}</div></div>
 </div>
</section>
<nav class="nav" role="tablist" aria-label="{T_("Seccions")}">
 <button type="button" role="tab" data-go="avui" aria-controls="v-avui">{T_("Avui")}</button>
 <button type="button" role="tab" data-go="coll" aria-controls="v-coll">{T_("Col·leccions")}</button>
 <button type="button" role="tab" data-go="wl" aria-controls="v-wl">{T_("Wishlist")}{f'<span class="n">{len(wl_bad)}</span>' if wl_bad else ""}</button>
 <button type="button" role="tab" data-go="trades" aria-controls="v-trades">{T_("Intercanvis")}</button>
 <button type="button" role="tab" data-go="eines" aria-controls="v-eines">{T_("Eines")}</button>
</nav>
<div class="view g2" id="v-avui" role="tabpanel">
 <section class="panel"><div class="panel-head"><h2>{T_("Què toca ara")}</h2></div>{checklist}</section>
 <section class="panel ocbox"><div class="panel-head"><h2>{T_("Ajudant del $oc")}</h2><a href="{oc}" style="font-size:13px">{T_("Obrir a part")}</a></div>
 <iframe src="{oc}" title="{T_("Ajudant del $oc")}"></iframe></section>
</div>
<div class="view" id="v-coll" role="tabpanel" hidden>
<section class="panel" id="coll">
 <div class="panel-head"><h2>{T_("Col·leccions")}</h2>
  <div class="legend"><span><i style="background:var(--accent)"></i>{T_("Teus")}</span><span><i style="background:var(--warm)"></i>{T_("D'algú altre")}</span><span><i style="background:var(--sunk)"></i>{T_("Lliures")}</span></div></div>
 <div class="bars">{"".join(bars)}</div>
 <div class="album">
  <div class="album-head"><h3 id="albumTitle"></h3>
   <label class="toggle"><input type="checkbox" id="hideMine"> {T_("Amaga els que ja tinc")}</label></div>
  <div class="legend">
   <span class="ch wish"><span class="star">★</span> {T_("A la wishlist")}</span><span class="ch alta">{T_("Lliure · prioritat alta")}</span>
   <span class="ch other">{T_("D'algú altre")}<small>{T_("qui el té")}</small></span><span class="ch free">{T_("Lliure")}</span><span class="ch mine">{T_("Teu")}</span></div>
  <div id="chips" class="chips"></div>
  <div class="give-row" id="gotBar" hidden><b id="gotCount"></b><button class="primary" id="gotBtn">{T_("Ja són meus")}</button>
   <button type="button" class="tab" id="gotClear">{T_("Desmarca")}</button><span id="gotMsg" role="status"></span></div>
 </div>
</section>
<section class="panel" id="esp">
 <div class="panel-head"><h2>{T_("Col·lecció especial")}</h2><span class="muted" style="font-size:13px">{T_("Un grup triat per tu: per exemple, els Eevee de la Pokédex")}</span></div>
 <div class="give-row"><input type="text" id="espName" list="dlEsp" spellcheck="false" autocomplete="off" placeholder="{T_("Nom del grup")}" aria-label="{T_("Nom del grup")}">
  <datalist id="dlEsp">{"".join(f'<option value="{e(x)}">' for x in spec)}</datalist></div>
 <textarea id="espText" rows="5" spellcheck="false" aria-label="{T_("Personatges")}" placeholder="{T_("Un personatge per línia, com surt a Mudae")}"></textarea>
 <div class="give-row"><button class="primary" id="espBtn" disabled>{T_("Desar la col·lecció")}</button><span id="espMsg" role="status"></span></div>
 <p class="muted" style="margin:0;font-size:13px">{T_("Els teus es marquen sols. Per saber de qui són els altres, ves a Eines → «Posar al dia», tria la col·lecció i enganxa el $imaso de la sèrie d'on surten. Per canviar-la, escriu el nom d'una que ja tens: s'hi carreguen els seus personatges.")}</p>
</section>
</div>
<div class="view g2" id="v-wl" role="tabpanel" hidden>
 <section class="panel" id="wlp"><div class="panel-head"><h2>{T_("Wishlist")}</h2><span class="muted" style="font-size:13px">{T_("{n} de 15 llocs (l'estrella també compta)", n=len(d["wishlist"]))}</span></div>
  <table>{wl}</table>
  <div id="wlBuild"></div>
 </section>
 <section class="panel"><h2>{T_("Sobrants per sèrie")}</h2><table>{sob}</table></section>
</div>
<div class="view" id="v-trades" role="tabpanel" hidden>
<section class="panel">
 <div class="panel-head"><h2>{T_("Intercanvis per persona")}</h2>
  <p class="recvbar"><button id="recvBtn" class="primary" disabled>{T_("Ja són meus")}</button><span id="recvMsg" role="status">{T_("Marca els que t'han donat.")}</span></p></div>
 <div class="owners">{trades}</div>
</section>
<section class="panel" id="mine">
 <div class="panel-head"><h2>{T_("Buscar entre els meus")}</h2><span class="muted" style="font-size:13px">{T_("{n} personatges al teu harem", n=len(d["harem"]))}</span></div>
 <label for="mineQ" class="muted" style="font-size:14px">{T_("Enganxa sèries o noms, separats per / , comes o línies (com la llista que t'envia un amic). No cal escriure-les bé del tot.")}</label>
 <textarea id="mineQ" rows="3" spellcheck="false" placeholder="Naruto / South Park / full metal alchimist / Rezero"></textarea>
 <div id="mineOut"></div>
</section>
</div>
<div class="view" id="v-eines" role="tabpanel" hidden>
<section class="panel" id="upd">
 <div class="panel-head"><h2>{T_("Posar al dia")}</h2><span id="updAge" class="muted" style="font-size:13px"></span></div>
 <div class="give-row">{T_("Què enganxes:")}
  <select id="updKind" aria-label="{T_("Què enganxes:")}"><option value="harem">{T_("El teu harem ($mms)")}</option><option value="wishlist">{T_("La wishlist ($wl)")}</option><option value="serie">{T_("Una sèrie ($imaso)")}</option></select>
  <span id="updSerieBox" hidden>{T_("Sèrie:")} <select id="updSerie" aria-label="{T_("Sèrie:")}">{serie_opts}<option value="__new__">{T_("Sèrie nova…")}</option></select>
  <input type="text" id="updNew" spellcheck="false" autocomplete="off" aria-label="{T_("Nom de la sèrie, com surt a Mudae")}" placeholder="{T_("Nom de la sèrie, com surt a Mudae")}" hidden></span>
  <span>{T_("Ordre:")} <button class="cmd" id="updCmd" data-copy="$mms">$mms</button></span></div>
 <p class="muted" id="updEsp" style="margin:0;font-size:13px" hidden>{T_("Col·lecció especial: enganxa el $imaso de la sèrie d'on surten els personatges (totes les pàgines). Només se n'agafaran els d'aquesta col·lecció.")}</p>
 <textarea id="updText" rows="4" spellcheck="false" aria-label="{T_("Posar al dia")}" placeholder="{T_("Enganxa aquí el missatge sencer que t'ha enviat Mudae per privat")}"></textarea>
 <div class="give-row"><button class="primary" id="updBtn" disabled>{T_("Desar i actualitzar")}</button><span id="updMsg" role="status"></span></div>
</section>
<section class="panel" id="cls">
 <div class="panel-head"><h2>{T_("Sèrie dels teus personatges")}</h2><span class="muted" style="font-size:13px">{T_("dels que no col·lecciones")}</span></div>
 {cls_html}
 <div class="give-row"><input type="text" id="clsName" list="dlNames" spellcheck="false" autocomplete="off" placeholder="{T_("Personatge")}" aria-label="{T_("Personatge")}">
  <span>{T_("és de")}</span><input type="text" id="clsSerie" list="dlSeries" spellcheck="false" autocomplete="off" placeholder="{T_("Sèrie")}" aria-label="{T_("Sèrie")}">
  <button class="primary" id="clsBtn" disabled>{T_("Desar")}</button><span id="clsMsg" role="status"></span></div>
 {datalists}
</section>
<section class="panel"><h2>{T_("Ordres d'ús habitual")}</h2><div class="ogrid">{ordres}</div></section>
</div>
<footer class="foot glass">Developed by <a href="https://github.com/Vlarac03">Vlarac</a> · {T_("Tauler de Mudae")} · {T_("Projecte de fans, no afiliat amb Mudae")}</footer>
</main><div id="tip"></div>
<script>{JS_HEAD.replace("__I18N__", js(STRINGS.get(LANG, {})))}
// Botó «Ja són meus» dels intercanvis: només funciona amb el servidor local engegat
const recvBtn = document.getElementById('recvBtn'), recvMsg = document.getElementById('recvMsg');
const picked = () => [...document.querySelectorAll('.recv:checked')].map(i => i.value);
if (location.protocol === 'file:') {{
  // obert com a fitxer: si el servidor està engegat, hi saltem; si no, avisem
  document.querySelectorAll('.recv').forEach(i => i.disabled = true);
  fetch('http://127.0.0.1:{PORT}/ping').then(r => r.json()).then(() => location.replace('http://127.0.0.1:{PORT}/dashboard.html'))
    .catch(() => {{ recvMsg.innerHTML = OFF_MSG(); }});
}} else {{
  document.addEventListener('change', ev => {{ if (ev.target.classList.contains('recv')) recvBtn.disabled = !picked().length; }});
  recvBtn.onclick = async () => {{
    recvBtn.disabled = true; recvMsg.textContent = T('Desant…');
    try {{
      const j = await postJSON('/rebut', {{ names: picked() }});
      if (j.ok) location.reload(); else {{ recvMsg.innerHTML = ERR_MSG(j.error); recvBtn.disabled = false; }}
    }} catch (err) {{ recvMsg.innerHTML = NET_MSG(); recvBtn.disabled = false; }}
  }};
}}
// Botons de copiar: copien el text de data-copy i diuen «Copiat»
document.addEventListener('click', async ev => {{
  const b = ev.target.closest('[data-copy]'); if (!b) return;
  ev.preventDefault();
  const text = b.dataset.copy;
  try {{ await navigator.clipboard.writeText(text); }}
  catch (err) {{ const ta = document.createElement('textarea'); ta.value = text; document.body.appendChild(ta); ta.select(); document.execCommand('copy'); ta.remove(); }}
  const old = b.textContent; b.classList.add('done'); b.textContent = T('Copiat');
  setTimeout(() => {{ b.classList.remove('done'); b.textContent = old; }}, 1400);
}});
// Capçalera: accés directe i tancar el servidor
const toolMsg = document.getElementById('toolMsg'), lnkBtn = document.getElementById('lnkBtn'), stopBtn = document.getElementById('stopBtn');
if (location.protocol === 'file:') {{ lnkBtn.disabled = stopBtn.disabled = true; }}
lnkBtn.onclick = async () => {{
  lnkBtn.disabled = true; toolMsg.textContent = T('Creant…');
  try {{ const j = await postJSON('/dressera'); toolMsg.innerHTML = j.ok ? '<span class="tag good">' + esc(T("Fet: «Mudae» a l'escriptori")) + '</span>' : ERR_MSG(j.error); }}
  catch (err) {{ toolMsg.innerHTML = NET_MSG(); }}
  lnkBtn.disabled = false;
}};
stopBtn.onclick = async () => {{
  stopBtn.disabled = lnkBtn.disabled = true;
  try {{ await postJSON('/quit'); }} catch (err) {{}}
  toolMsg.innerHTML = '<span class="tag">' + esc(T('Servidor tancat')) + '</span> ' + esc(T("Ja pots tancar la pestanya. Per tornar-lo a obrir, fes servir l'accés directe Mudae."));
}};
const tip = document.getElementById('tip');
document.addEventListener('mousemove', ev => {{
  const el = ev.target.closest('[data-tip]');
  if (!el) {{ tip.style.display = 'none'; return; }}
  tip.textContent = el.dataset.tip; tip.style.display = 'block';
  tip.style.left = Math.min(ev.clientX + 12, innerWidth - tip.offsetWidth - 8) + 'px';
  tip.style.top = (ev.clientY + 14) + 'px';
  const g = ev.target.closest('.panel, .resum');  // la llum que segueix el ratolí
  if (g) {{ const r = g.getBoundingClientRect(); g.style.setProperty('--mx', ev.clientX - r.left + 'px'); g.style.setProperty('--my', ev.clientY - r.top + 'px'); }}
}});
// Pestanyes: una vista cada vegada; es recorda l'última (també després de desar, que recarrega la pàgina)
const tabs = [...document.querySelectorAll('.nav [data-go]')];
function go(v, focus) {{
  if (!document.getElementById('v-' + v)) v = 'avui';
  tabs.forEach(b => {{
    const on = b.dataset.go === v;
    b.setAttribute('aria-selected', on); b.tabIndex = on ? 0 : -1;
    document.getElementById('v-' + b.dataset.go).hidden = !on;
    if (on && focus) b.focus();
  }});
  try {{ localStorage.setItem('mudae-vista', v); }} catch (err) {{}}
}}
tabs.forEach((b, i) => {{
  b.onclick = () => go(b.dataset.go);
  b.onkeydown = ev => {{
    const k = {{ ArrowRight: 1, ArrowLeft: -1 }}[ev.key];
    if (k) {{ ev.preventDefault(); go(tabs[(i + k + tabs.length) % tabs.length].dataset.go, true); }}
  }};
}});
let first = 'avui'; try {{ first = localStorage.getItem('mudae-vista') || 'avui'; }} catch (err) {{}}
go(first);
</script><script>{JS_TODAY.replace("__CLAIM__", json.dumps(CLAIM_HOURS))}</script>
<script>{JS_COLL.replace("__DATA__", js(coll))}</script>
<script>{JS_UPD.replace("__META__", js(read_meta())).replace("__SPECIALS__", js(spec))}</script>
<script>{JS_ESP}</script>
<script>{JS_CLS}</script>
<script>{JS_WL.replace("__WL__", wl_json)}</script>
<script>{JS_MINE.replace("__MINE__", js(mine)).replace("__COLLS__", js([x["name"] for x in series])).replace("__ALTRES__", js(ALTRES))}</script></body></html>"""
    DASH.write_text(page, encoding="utf-8")


def run():
    d = build_data()
    write_excel(d)
    hist = update_history(d)
    write_dashboard(d, hist)
    owned = sum(s["Meu"] for s in d["series"])
    total = sum(s["total"] for s in d["series"])
    print(t("Fet: {owned}/{total} personatges de les teves sèries · {trades} per intercanviar · {sob} sobrants · {changes} canvis",
            owned=owned, total=total, trades=len(d["trades"]), sob=len(d["sobrants"]), changes=sum(k != "nou" for k, _, _ in d["changes"])))
    for w in d["wishlist"]:
        if not w["ok"]:
            print(f"  Wishlist: {w['name']} -> {w['todo']}")




def read_meta():
    try:
        return json.loads((DADES / "actualitzat.json").read_text(encoding="utf-8"))  # quan vas enganxar cada cosa (avís setmanal)
    except (OSError, ValueError):
        return {}


def merge_owners(text, old):
    """El $imas (sense la «o») diu «Nom 💞» però no de qui és: hi posem el propietari que ja teníem apuntat."""
    prev = {}
    for raw in old.splitlines():
        m = CLAIM.match(re.sub(r"^\d+\.\s+", "", raw.strip()))
        if m:
            prev[m.group(1).strip()] = m.group(2)
    harem = {n for n, _ in lines(DADES / "harem.txt")}
    out = []
    for raw in text.splitlines():
        line = re.sub(r"^\d+\.\s+", "", raw.strip())
        h = HEART_ONLY.match(line)
        if h and not CLAIM.match(line):
            name = h.group(1).strip()
            out.append(f"{name} 💞 => {prev.get(name) or (ME if name in harem else 'algú')}")  # algú = reclamat, no sabem per qui
        else:
            out.append(raw)
    return "\n".join(out)


def series_path(serie):
    if not serie or re.search(r'[\\/:*?"<>|]', serie) or serie.startswith("."):
        raise ValueError(t("Posa el nom de la sèrie (sense / \\ : * ? \" < > |)."))
    return DADES / "series" / f"{serie}.txt"


def touch_meta(serie):
    meta = read_meta()
    meta.setdefault("series", {})[serie] = dt.date.today().isoformat()
    (DADES / "actualitzat.json").write_text(json.dumps(meta, ensure_ascii=False, indent=1), encoding="utf-8")


def owner_line(name, owner):
    return f"{name} 💞 => {owner}" if owner else name


def save_special(name, text):
    """Crea o canvia una col·lecció especial: el nom del grup i els seus personatges, un per línia (o separats per $).

    Els teus surten del harem; dels altres, es manté el propietari que ja sabíem (d'aquesta o d'una altra col·lecció).
    """
    name = name.strip()
    path = series_path(name)
    if path.exists() and name not in specials():
        raise ValueError(t("Ja tens una sèrie que es diu «{name}»: tria un altre nom.", name=name))
    names = list(parse_owners("\n".join(re.split(r"\s+\$\s+", text))))
    if not names:
        raise ValueError(t("Escriu almenys un personatge, un per línia."))
    known = {}
    for r in read_series():
        if r["owner"] and (r["serie"] == name or r["name"] not in known):
            known[r["name"]] = r["owner"]  # el d'aquesta mateixa col·lecció mana
    harem = {n for n, _ in lines(DADES / "harem.txt")}
    rows = [owner_line(n, ME if n in harem else "" if known.get(n) == ME else known.get(n, "")) for n in names]
    path.write_text("\n".join(rows) + "\n", encoding="utf-8")
    spec = specials()
    if name not in spec:
        (DADES / "especials.json").write_text(json.dumps(spec + [name], ensure_ascii=False, indent=1), encoding="utf-8")
    touch_meta(name)


def update_special(name, text):
    """«Posar al dia» d'una especial: enganxes el $imaso de la sèrie d'on surten i només se n'agafen els seus personatges."""
    found = {n.casefold(): o for n, o in parse_owners(drop_header(text)).items()}
    path = series_path(name)
    harem = {n for n, _ in lines(DADES / "harem.txt")}
    rows, hits = [], 0
    for n, old in parse_owners(path.read_text(encoding="utf-8-sig")).items():
        new = found.get(n.casefold(), old)
        if n.casefold() in found:
            hits += 1
        if new is None:  # «Nom 💞» sense propietari: el que sabíem, o teu si és al harem
            new = old or (ME if n in harem else "algú")
        rows.append(owner_line(n, new))
    if not hits:
        raise ValueError(t("Cap personatge de «{name}» al text enganxat. Enganxa el $imaso de la sèrie d'on surten.", name=name))
    path.write_text("\n".join(rows) + "\n", encoding="utf-8")
    touch_meta(name)


def save_paste(kind, serie, text):
    """Desa el que has enganxat al tauler al fitxer de dades/ que toca. Si no s'hi troba cap nom, no toca res."""
    if kind == "harem":
        path = DADES / "harem.txt"
    elif kind == "wishlist":
        path = DADES / "wishlist.txt"
    elif kind == "serie":
        serie = serie.strip()
        path = series_path(serie)
        if serie in specials():
            return update_special(serie, text)
    else:
        raise ValueError(t("No sé què és «{kind}».", kind=kind))
    old = path.read_text(encoding="utf-8-sig") if path.exists() else None
    if kind == "serie":
        text = drop_header(text)
    if kind == "serie" and any(HEART_ONLY.match(l.strip()) for l in text.splitlines()):
        text = merge_owners(text, old or "")
    path.write_text(text.strip() + "\n", encoding="utf-8")
    if not lines(path):
        if old is None:
            path.unlink()
        else:
            path.write_text(old, encoding="utf-8")
        raise ValueError(t("No he trobat cap nom al text enganxat. Copia el missatge de Mudae sencer."))
    meta = read_meta()
    today = dt.date.today().isoformat()
    if kind == "serie":
        meta.setdefault("series", {})[serie] = today
    else:
        meta[kind] = today
    (DADES / "actualitzat.json").write_text(json.dumps(meta, ensure_ascii=False, indent=1), encoding="utf-8")


def classify(name, serie):
    """Posa `name` (un personatge teu que no col·lecciones) a la sèrie `serie` de dades/series-sobrants.json."""
    name, serie = name.strip(), serie.strip()
    if not name or not serie:
        raise ValueError(t("Falta el nom del personatge o la sèrie."))
    if name not in [n for n, _ in lines(DADES / "harem.txt")]:
        raise ValueError(t("«{name}» no és al teu harem. Revisa com està escrit.", name=name))
    path = DADES / "series-sobrants.json"
    m = json.loads(path.read_text(encoding="utf-8"))
    for k in list(m):
        m[k] = [n for n in m[k] if n != name]
        if not m[k]:
            del m[k]
    m.setdefault(serie, []).append(name)
    path.write_text(json.dumps(m, ensure_ascii=False, indent=1), encoding="utf-8")


def mark_given(names, to=""):
    """Has regalat `names`: fora del harem; si algun era d'una col·lecció teva, passa a nom de qui l'ha rebut."""
    names = [n.strip() for n in names if n.strip()]
    if not names:
        raise ValueError(t("No hi ha cap personatge per marcar."))
    to = to.strip().lstrip("@") or "algú"
    harem = DADES / "harem.txt"
    have = [n for n, _ in lines(harem)]
    missing = [n for n in names if n not in have]
    if missing:
        raise ValueError(t("No són al teu harem: {names}", names=", ".join(missing)))
    kept = [raw for raw in harem.read_text(encoding="utf-8-sig").splitlines()
            if re.sub(r"^\d+\.\s+", "", raw.replace("⭐", "").strip()) not in names]
    harem.write_text("\n".join(kept) + "\n", encoding="utf-8")
    for f in (DADES / "series").glob("*.txt"):
        text = f.read_text(encoding="utf-8-sig").splitlines()
        changed = False
        for k, raw in enumerate(text):
            m = CLAIM.match(re.sub(r"^\d+\.\s+", "", raw.strip()))
            if m and m.group(1).strip() in names and m.group(2).casefold() == ME.casefold():
                text[k] = f"{m.group(1).strip()} 💞 => {to}"
                changed = True
        if changed:
            f.write_text("\n".join(text) + "\n", encoding="utf-8")


def mark_received(name):
    """Posa `name` com a teu al fitxer de la seva sèrie (i al harem), com si ho hagués dit el $imaso."""
    found = False
    for f in sorted((DADES / "series").glob("*.txt")):
        text = f.read_text(encoding="utf-8-sig").splitlines()
        for k, raw in enumerate(text):
            line = re.sub(r"^\d+\.\s+", "", raw.strip())
            m = CLAIM.match(line)
            if (m.group(1).strip() if m else line) == name:
                text[k] = f"{name} 💞 => {ME}"
                f.write_text("\n".join(text) + "\n", encoding="utf-8")
                found = True
    if found:
        harem = DADES / "harem.txt"
        if name not in [n for n, _ in lines(harem)]:
            harem.write_text(harem.read_text(encoding="utf-8-sig").rstrip("\n") + f"\n{name}\n", encoding="utf-8")
        return
    raise ValueError(t("No trobo «{name}» a cap de les teves sèries.", name=name))


PUBLIC = {"/dashboard.html", "/oc.html"}  # l'únic que el servidor deixa veure
PORT = 8766  # (el 8765 l'ocupaven servidors vells que no es podien tancar sols)
VERSION = str(Path(__file__).stat().st_mtime_ns)  # canvia cada vegada que es modifica aquest fitxer
IDLE_SECONDS = 150  # el tauler avisa cada 20 s; amb la pestanya en segon pla el navegador pot espaiar-ho fins a 1 min


def configure(user, server="", lang=""):
    """Primera vegada: desa l'usuari, prepara dades/ i posa l'accés directe «Mudae» a l'escriptori."""
    global LANG
    LANG = {"1": "en", "2": "es", "3": "ca"}.get(lang.strip(), lang.strip()) if lang.strip() in ("1", "2", "3", "ca", "es", "en") else "en"
    user = user.strip().lstrip("@")
    if not user:
        sys.exit(t("Cal el teu nom d'usuari de Mudae."))
    (DADES / "series").mkdir(parents=True, exist_ok=True)
    for f in ("harem.txt", "wishlist.txt"):
        (DADES / f).touch()
    if not (DADES / "series-sobrants.json").exists():
        (DADES / "series-sobrants.json").write_text("{}", encoding="utf-8")
    CONFIG_FILE.write_text(json.dumps({**CONFIG, "usuari": user, "idioma": LANG, **({"servidor": server.strip()} if server.strip() else {})},
                                      ensure_ascii=False, indent=1), encoding="utf-8")
    print(t("Usuari desat: {user}", user=user))
    try:
        make_shortcut()
        print(t("Accés directe «Mudae» creat a l'escriptori."))
    except RuntimeError:
        print(t("No he pogut crear l'accés directe: obre Actualitza.bat."))


def make_shortcut():
    """Accés directe «Mudae» a l'escriptori: obre Actualitza.bat (posa al dia i obre el tauler)."""
    import subprocess
    icon = DIR / "icones" / ("mudae.ico" if (DIR / "icones" / "mudae.ico").exists() else "tauler.ico")  # mudae.ico no va al repo: no és nostra
    q = lambda v: "'" + str(v).replace("'", "''") + "'"  # cometes simples: PowerShell no hi expandeix $ ni `
    ps = ('$s = (New-Object -ComObject WScript.Shell).CreateShortcut([Environment]::GetFolderPath("Desktop") + "\\Mudae.lnk");'
          f'$s.TargetPath = {q(DIR / "Actualitza.bat")}; $s.WorkingDirectory = {q(DIR)};'
          f'$s.IconLocation = {q(f"{icon},0")}; $s.WindowStyle = 7; $s.Save()')
    r = subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True, text=True,
                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    if r.returncode != 0:
        raise RuntimeError(t("No he pogut crear l'accés directe a l'escriptori."))


def do_action(action, body):
    """El que fan els botons del tauler. S'executa en un procés nou, sempre amb la versió actual del codi."""
    if action == "rebut":
        for n in body["names"]:
            mark_received(str(n))
    elif action == "donat":
        mark_given([str(n) for n in body["names"]], str(body.get("to") or ""))
    elif action == "classifica":
        classify(str(body["name"]), str(body["serie"]))
    elif action == "dressera":
        make_shortcut()
    elif action == "dades":
        save_paste(str(body["kind"]), str(body.get("serie") or ""), str(body["text"]))
    elif action == "especial":
        save_special(str(body["name"]), str(body["text"]))
    else:
        raise ValueError(t("No sé fer «{action}».", action=action))


def serve(port=PORT):
    """Servidor del tauler a http://127.0.0.1:<port> per als botons. Es queda engegat en segon pla.
    Només n'hi pot haver un; si n'hi ha un de vell (codi anterior), el tanca i en posa un de nou."""
    import socket
    import subprocess
    import threading
    import urllib.request
    import webbrowser
    from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

    class Server(ThreadingHTTPServer):
        allow_reuse_address = False  # a Windows deixaria engegar-ne diversos al mateix port
        daemon_threads = True

        def server_bind(self):
            if hasattr(socket, "SO_EXCLUSIVEADDRUSE"):
                self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
            super().server_bind()

    class Handler(SimpleHTTPRequestHandler):
        def __init__(self, *a, **k):
            super().__init__(*a, directory=str(DIR), **k)

        def log_message(self, *a):
            pass

        def end_headers(self):
            self.send_header("Cache-Control", "no-store")  # sempre la darrera versió del tauler
            super().end_headers()

        def reply(self, res, cors=False):
            body = json.dumps(res).encode()
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            if cors:
                self.send_header("Access-Control-Allow-Origin", "*")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def host_ok(self):
            """Una web de fora que fes «DNS rebinding» arribaria aquí amb el seu nom al Host: la rebutgem."""
            if self.headers.get("Host") in (f"127.0.0.1:{port}", f"localhost:{port}"):
                return True
            self.send_error(403)
            return False

        def do_OPTIONS(self):  # Chrome pregunta abans de deixar que un fitxer local parli amb el servidor
            if not self.host_ok():
                return
            self.send_response(204)
            self.send_header("Access-Control-Allow-Origin", "*")
            self.send_header("Access-Control-Allow-Methods", "GET")
            self.send_header("Access-Control-Allow-Private-Network", "true")
            self.end_headers()

        def do_GET(self):
            if not self.host_ok():
                return
            last_seen[0] = time.monotonic()
            path = self.path.split("?")[0]
            if path == "/ping":  # el tauler obert (cada 20 s), o un servidor nou, pregunta si n'hi ha un
                return self.reply({"ok": True, "v": VERSION}, cors=True)
            if path not in PUBLIC:  # res de dades/, l'Excel ni el codi: només el tauler i l'ajudant
                return self.send_error(404)
            super().do_GET()

        def do_POST(self):
            if not self.host_ok():
                return
            # Només JSON: així una web de fora no pot fer clic als botons per tu (el navegador ho bloqueja)
            if not (self.headers.get("Content-Type") or "").startswith("application/json"):
                return self.send_error(415)
            raw = self.rfile.read(int(self.headers.get("Content-Length") or 0)).decode("utf-8")
            if self.path == "/quit":  # un servidor més nou vol el port
                self.reply({"ok": True})
                return threading.Thread(target=httpd.shutdown, daemon=True).start()
            try:
                py = Path(sys.executable).with_name("python.exe")
                r = subprocess.run([str(py if py.exists() else sys.executable), str(DIR / "mudae.py"), "accio", self.path.strip("/")],
                                   input=raw, capture_output=True, text=True, encoding="utf-8",
                                   env={**os.environ, "PYTHONIOENCODING": "utf-8"}, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
                err = (r.stderr.strip().splitlines() or ["Error desconegut"])[-1]
                res = {"ok": True} if r.returncode == 0 else {"ok": False, "error": err}
            except Exception as ex:
                res = {"ok": False, "error": str(ex)}
            self.reply(res)

    base = f"http://127.0.0.1:{port}"
    last_seen = [time.monotonic()]
    httpd = None
    for _ in range(2):
        try:
            httpd = Server(("127.0.0.1", port), Handler)  # només aquest ordinador
            break
        except OSError:  # ja n'hi ha un: si és d'una versió anterior, li demanem que plegui
            try:
                v = json.loads(urllib.request.urlopen(base + "/ping", timeout=3).read()).get("v")
                if v == VERSION:
                    break
                urllib.request.urlopen(urllib.request.Request(base + "/quit", data=b"{}", headers={"Content-Type": "application/json"}), timeout=3)
                time.sleep(1.5)
            except OSError:
                time.sleep(1.5)
    webbrowser.open(base + "/dashboard.html")
    if httpd:
        def watchdog():  # sense notícies del tauler (pestanya tancada), el servidor plega sol
            while True:
                time.sleep(10)
                if time.monotonic() - last_seen[0] > IDLE_SECONDS:
                    return httpd.shutdown()
        threading.Thread(target=watchdog, daemon=True).start()
        httpd.serve_forever()
        httpd.server_close()


def main():
    if sys.stdout:  # amb pythonw no hi ha consola
        sys.stdout.reconfigure(encoding="utf-8")
    args = sys.argv[1:]
    if args[:1] == ["configura"]:  # Configura.bat: python mudae.py configura <usuari> [servidor] [idioma]
        return configure(*(args[1:4] or [""]))
    if args[:1] == ["tauler"]:  # només el servidor (Actualitza.bat el llança en segon pla)
        return serve()
    try:
        if args[:1] == ["accio"]:  # el servidor: python mudae.py accio <nom> < JSON
            if sys.stdin and hasattr(sys.stdin, "reconfigure"):
                sys.stdin.reconfigure(encoding="utf-8")
            do_action(args[1], json.loads(sys.stdin.read() or "{}"))
        if args[:1] == ["rebut"]:  # python mudae.py rebut "Nom 1" "Nom 2"
            for n in args[1:]:
                mark_received(n)
        if args[:1] == ["donat"]:  # python mudae.py donat "@persona" "Nom 1" "Nom 2"
            mark_given(args[2:], args[1])
        if args[:1] == ["classifica"]:  # python mudae.py classifica "Personatge" "Sèrie"
            classify(args[1], args[2])
        run()
    except (RuntimeError, ValueError, KeyError) as ex:
        sys.exit(str(ex))
    if args[:1] == ["serve"]:
        serve()


if __name__ == "__main__":
    main()
