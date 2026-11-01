"""Actualitza l'Excel i el tauler de Mudae a partir del que enganxes a la carpeta dades/.

Ús: python mudae.py   (o doble clic a Actualitza.bat)

- dades/series/<Sèrie>.txt  -> el que respon `$ima <sèrie>` (una sèrie per fitxer)
- dades/harem.txt           -> el teu harem (el desa el `$mma+s` de «Posar al dia»)
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

GRUPS = {
    "reclam": "POTS RECLAMAR",
    "hora": "CADA HORA: no les perdis",
    "dia": "UN COP AL DIA",
    "comodi": "COMODINS: guarda'ls per quan calguin",
    "aconseguit": "UN COP PER SETMANA (o quan aconsegueixis algú)",
    "temps": "QUAN TINGUIS TEMPS",
}
# (grup, què, ordre, com, quan es torna a poder fer): surt al full Rutina i al tauler, per ordre d'importància.
# «hora» = al reinici de tirades · «reclam» = al pròxim reclam · «rt» = les hores del $rt · «utc» = a mitjanit UTC · un número = hores des que la fas
# Les hores que depenen del servidor o de tu (reinicis, $rt) surten del que enganxes ($tu, $settings, $bonus), mai d'aquí.
RUTINA = [
    ("reclam", "Reclamar", "clic al cor", "El millor que surti: wishlist > les teves sèries > molta kakera. Si no surt res bo, espera't: a l'última hora abans del reinici, reclama el millor que hi hagi", "reclam"),
    ("hora", "Tirades", "$m", "Fes-les totes: les que no fas es perden", "hora"),
    ("hora", "Kakera", "clic al botó", "Cada reacció gasta poder: enganxa el $tu per saber quant", "hora"),
    ("hora", "Pokéslot", "$p", "Tirada de Pokémon: cada 2 h, per a tots els servidors", 2),
    ("hora", "Mirar què tens", "$tu", "Enganxa'l a la casella de dalt: marca sol el que ja has fet", "hora"),
    ("dia", "Votar", "$vote", "Et guarda un reinici de tirades: el gastes amb $rolls. Cada 12 h", 12),
    ("dia", "Reinici de tirades", "$daily", "Et guarda un reinici de tirades, com el $vote: el gastes amb $rolls. Cada 20 h (cal tenir 5 personatges al $ll)", 20),
    ("dia", "Esferes", "$oh (×4)", "Primer blaves (destapen 3), als últims clics les millors. 2 minuts!", "utc"),
    ("dia", "Trobar la vermella", "$oc", "Si en tens. Fes servir l'ajudant: et diu on clicar", "utc"),
    ("comodi", "Un altre reclam", "$rt", "Ja has reclamat i en surt un de la wishlist o de les teves sèries? Torna a reclamar", "rt"),
    ("comodi", "Recarregar el poder", "$dk", "Quan no et quedi poder per a la kakera: el torna tot i dona kakera. Cada 20 h", 20),
    ("aconseguit", "Treure'l de la wishlist", "$wishremove <nom>", "I posa-hi el primer de «Per omplir» (full Wishlist)", None),
    ("aconseguit", "Posar al dia les dades", "$mma+s · $wl · $imaso <sèrie>", "Per privat i sense pàgines. Enganxa-ho a la casella d'Avui", None),
    ("aconseguit", "Obrir el tauler", "Mudae", "Accés directe de l'escriptori: posa al dia l'Excel i obre el tauler", None),
    ("temps", "Intercanvis", "$trade @persona", "Full Intercanvis: a qui demanar què. Ofereix dels Sobrants", None),
    ("temps", "Kakera estalviada", "$kt", "A la pestanya Millores et diu en què gastar-la: insígnies, torre o kakeraloots", None),
]

SKIP = re.compile(r"Harem de|Wishlist|^Page\b|\bAPP\b|—\s*\d{1,2}:\d{2}|^\$|^Mudae$|^\d+\s*/\s*\d+$", re.I)
HEART_ONLY = re.compile(r"^(.*?)\s*(?:💞|💖|💗|❤️|❤|♥️|♥)\s*$")  # «Nom 💞» del $imas (sense la «o»): reclamat, propietari desconegut
CLAIM = re.compile(r"^(.*?)\s*(?:💞|💖|💗|❤️|❤|♥️|♥)?\s*=>\s*(\S+)\s*$")
IMASO_TITLE = re.compile(r"^(?:\d+\.\s+)?\S.*\s\d+\s*/\s*\d+$")  # «1. Harry Potter 49/94»: el títol del $imaso


def drop_header(text, check=True):
    """El $imaso comença amb el títol de la sèrie («1. Harry Potter 49/94») i un bloc amb els àlies: ho traiem.

    El bloc d'àlies acaba amb «… $ima? …», amb la línia separadora de Mudae (un espai) o, si no, és àlies quan sense
    ell queden just els personatges que diu el títol. Si al final no en queden tants, no es desa res.
    """
    rows = text.splitlines()
    count = lambda rs: len(parse_owners("\n".join(rs)))
    for i, raw in enumerate(rows[:6]):
        if IMASO_TITLE.match(raw.strip()):
            total = int(re.search(r"(\d+)\s*$", raw).group(1))
            k = next((j for j in range(i + 1, len(rows)) if rows[j].strip()), len(rows))   # primera línia del bloc
            end = next((j for j in range(k, len(rows)) if not rows[j].strip()), len(rows))  # línia buida que el tanca
            block = rows[k:end]
            aliases = block and not any("=>" in r for r in block) and (
                "$ima?" in block[-1] or (end < len(rows) and rows[end] != "") or count(rows[end + 1:]) == total)
            body = rows[end + 1 if aliases else k:]
            if check and count(body) != total:  # a les especials n'hi ha prou amb un tros (els Eevee dels 1.025 Pokémon)
                raise ValueError(t("Mudae diu que la sèrie té {total} personatges i n'he llegit {n}. Copia el $imaso sencer, per privat i sense pàgines.",
                                   total=total, n=count(body)))
            return "\n".join(rows[:i] + body)
    return text


def lines(path):
    """Línies útils d'un fitxer enganxat de Discord, amb la marca ⭐ si n'hi havia."""
    out = []
    if not path.exists():
        return out
    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        star = "⭐" in raw
        line = re.sub(r"^\d+\.\s+", "", raw.replace("⭐", "").strip())  # numeració de Discord: "2. Remus Lupin"
        line = re.sub(r"\s+\|\s.*$", "", line)  # nota de Mudae ($note): "Ginger Torte | RANDOM"
        line = re.sub(r"(?:\s*(?:✅|❌|✔️?|☑️?)|\s+\+\d+\s*%)+\s*$", "", line)  # marques del $wl: "Albedo ✅ +100%"
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
        seen, serie = set(), f.stem.translate(FROM_FILE)
        for line, _ in lines(f):
            m = CLAIM.match(line)
            h = HEART_ONLY.match(line)
            name, owner = (m.group(1).strip(), m.group(2)) if m else (h.group(1).strip(), "algú") if h else (line, "")
            if owner.casefold() == ME.casefold():  # Mudae escriu els noms en minúscula («pep»), tu pots haver posat «Pep»
                owner = ME
            if serie in spec and name in harem:  # a les especials, el que és al teu harem és teu (no hi ha cap $imaso que ho digui)
                owner = ME
            if name not in seen:
                seen.add(name)
                rows.append({"serie": serie, "name": name, "owner": owner})
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
                              t("El desig estrella ($sw) també ocupa un lloc de la wishlist.")], start=k + 1):
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
        for j, v in enumerate((t(que), cmd_text(ordre), t(com)), start=2):
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
// Plegables (details[data-fold]): cada un es queda obert o tancat com el vas deixar, també d'un dia per l'altre (en aquest navegador)
// si no l'has tocat mai, com ve (data-open: obert d'entrada)
const foldOpen = (k, def = false) => { try { const v = localStorage.getItem('mudae-fold-' + k); return v == null ? def : v === '1'; } catch (e) { return def; } };
document.querySelectorAll('details[data-fold]').forEach(d => { d.open = foldOpen(d.dataset.fold, d.hasAttribute('data-open')); });
document.addEventListener('toggle', ev => {  // toggle no puja: l'escoltem en la fase de captura
  const d = ev.target; if (!d.dataset || !d.dataset.fold) return;
  if (d.closest('#ordres') && document.getElementById('ordQ').value.trim()) return;  // el cercador els obre: no és teu
  try { localStorage.setItem('mudae-fold-' + d.dataset.fold, d.open ? '1' : '0'); } catch (e) {}
}, true);
"""


# Tasques de RUTINA amb hora: el seu nom al tauler (data-id) i la línia del $tu que la marca (les mateixes claus que readTu)
TU_KEY = {"$tu": "tu", "$vote": "vote", "$daily": "daily", "$oh (×4)": "oh", "$oc": "oc", "$m": "rolls",
          "clic al botó": "kakera", "clic al cor": "claim", "$rt": "rt", "$dk": "dk", "$p": "p"}

JS_TODAY = """
// Els reinicis del teu servidor: res de suposat. Del $tu, el minut de les tirades (cada hora) i el pròxim reclam;
// del $settings, cada quant hi ha reclam ($setclaim). Sense aquestes dades, el tauler et diu què has d'enganxar.
let RS = {}; try { RS = JSON.parse(localStorage.getItem('mudae-reinici')) || {}; } catch (e) {}
if (RS.c) RS = {};  // el format d'abans (hores fixes de config.json): es torna a aprendre del $tu
const MIN = 60000, HOUR = 3600000;
const clock = document.getElementById('clock');
const pad = n => String(n).padStart(2, '0');
const fmt = m => m >= 60 ? `${Math.floor(m / 60)} h ${m % 60} min` : `${m} min`;
const hm = d => `${pad(d.getHours())}:${pad(d.getMinutes())}`;
// El pròxim reinici de tirades i de reclam després de `t` (ms), o null si encara no ho sabem
function nextRoll(t) { if (RS.m == null) return null; const d = new Date(t); d.setMinutes(RS.m, 0, 0); return +d > t ? +d : +d + HOUR; }
function nextClaim(t) {
  if (!RS.next) return null;
  if (!RS.iv) return RS.next > t ? RS.next : null;
  const iv = RS.iv * MIN; return RS.next + (Math.floor((t - RS.next) / iv) + 1) * iv;
}
const lastClaim = t => { const n = nextClaim(t); return n == null ? (RS.next && RS.next <= t ? RS.next : null) : RS.iv ? n - RS.iv * MIN : null; };
// Quan es pot tornar a fer una tasca marcada a l'hora `t` (ms): al reinici de cada hora, al pròxim reclam, a mitjanit UTC,
// al cap de les hores del $rt (del $bonus) o al cap de N hores. Infinity: no ho sabem (falta enganxar alguna cosa).
function until(torna, t) {
  if (torna === 'hora') return nextRoll(t) ?? Infinity;
  if (torna === 'reclam') return nextClaim(t) ?? Infinity;
  if (torna === 'rt') return window.RT_H ? t + window.RT_H * HOUR : Infinity;
  if (torna === 'utc') { const d = new Date(t); d.setUTCHours(24, 0, 0, 0); return +d; }
  return t + Number(torna) * HOUR;
}
// Del $tu, d'aquí quants minuts es reinicien el reclam i les tirades (0 = no ho diu)
function readResets(text) {
  const ls = text.split(/\\r?\\n/), isClaim = l => /reclamar|\\bclaim/i.test(l);
  const c = ls.find(isClaim), r = ls.find(l => !isClaim(l) && !/\\$daily/i.test(l) && /reinicio|reset/i.test(l));
  return { claim: c ? tuMins(c) : 0, roll: r ? tuMins(r) : 0 };
}
// Aprèn el minut de les tirades i l'hora del pròxim reclam del $tu enganxat a l'hora `now` (Mudae diu minuts sencers: el reinici cau a mig minut)
function learnResets(R, r, now) {
  const at = n => { const d = new Date(now + (n + 0.5) * MIN); d.setSeconds(0, 0); return d; };
  const out = { ...R };
  if (r.roll) out.m = at(r.roll).getMinutes();
  if (r.claim) out.next = +at(r.claim);
  return out;
}
// Del $settings: cada quant hi ha reclam en minuts (iv, de la línia del $setclaim, entre 60 i 600) i el minut dels reinicis (m, del $setinterval).
// «· Reinicio de reclamos: Cada 180 min. ($setclaim)» i «· Minuto exacto de cada reinicio: xx:02 ($setinterval)» (de Victor, 01/11/2026). null si no hi és.
function readSettings(text) {
  const ls = text.split(/\\r?\\n/), line = re => ls.find(l => re.test(l));
  const c = line(/\\$setclaim/i) || line(/reclam|claim/i), i = line(/\\$setinterval/i);
  const m = c && c.match(/(\\d{2,3})\\s*min/i), h = c && c.match(/(\\d+(?:[.,]\\d+)?)\\s*h/i);
  const iv = m ? +m[1] : h ? Math.round(parseFloat(h[1].replace(',', '.')) * 60) : null;
  const mm = i && i.match(/:\\s*(\\d{1,2})\\b/);
  return iv >= 60 && iv <= 600 ? { iv, m: mm && +mm[1] < 60 ? +mm[1] : null } : null;
}
// El $tu enganxat: per a cada tasca, quants minuts falten perquè la puguis tornar a fer (0 = ja la pots fer).
// Fem servir els minuts que diu Mudae, no les hores fixes: cada servidor reinicia tirades i reclams a una hora diferent.
// true = no la pots fer però el $tu no diu quan: torna com si l'haguessis marcada a mà.
const tuMins = l => { const m = l.match(/(\\d+)\\s*h(?:\\s*(\\d+)\\s*min)?|(\\d+)\\s*min/i); return m ? (m[1] ? m[1] * 60 + (+m[2] || 0) : +m[3]) : 0; };
function readTu(text) {
  const r = {}, set = (k, v) => { if (!(k in r)) r[k] = v; }, mins = tuMins;  // la primera línia que en parla mana
  const no = l => /\\bno puedes\\b|\\bno pots\\b|can't|cannot|can not/i.test(l);
  const ready = l => /disponible|available|ready|listo|lista|a punt/i.test(l) && !/\\bnot?\\b/i.test(l);  // «no está disponible», «not ready»
  const wait = l => ready(l) ? 0 : mins(l) || true;
  const ls = text.split(/\\r?\\n/);
  ls.forEach((l, k) => {
    const low = l.toLowerCase(), n = re => { const m = l.match(re); return m ? +m[1] : null; };
    if (/reclamar|\\bclaim/.test(low)) set('claim', no(l) ? mins(l) || true : 0);
    else if (n(/(\\d+)\\s*rolls?\\b/i) !== null) { set('rollsN', n(/(\\d+)\\s*rolls?\\b/i)); if (mins(l)) set('rollsIn', mins(l)); }
    else if (/\\$daily/.test(low)) set('daily', wait(l));
    else if (/reinicio|reset/.test(low)) set('rollsIn', mins(l));  // «El siguiente reinicio será en 18 min.»: les tirades
    else if (/kakera/.test(low) && /reacc|react/.test(low)) set('kakera', no(l) ? mins(l) || true : 0);
    else if (/\\$rt\\b/.test(low)) set('rt', wait(l));
    else if (/\\$dk\\b/.test(low)) set('dk', wait(l));
    else if (/\\$p\\b|pok[eé]mon/.test(low)) set('p', wait(l));  // shortcut: format suposat (no tenim cap $tu amb la línia del Pokéslot); mirar-ho amb un de real
    else if (/\\bvot|\\bvote/.test(low)) set('vote', mins(l));
    else if (/\\$oh\\b/.test(low)) {
      const next = mins(l) || mins(ls.slice(k + 1).find(x => x.trim()) || '') || true;  // «14h 16 min hasta la recarga.»
      const oh = n(/(\\d+)\\s*\\$oh\\b/i), oc = n(/(\\d+)\\s*\\$oc\\b/i);
      if (oh !== null) set('oh', oh ? 0 : next);
      if (oc !== null) set('oc', oc ? 0 : next);
    }
  });
  if ('rollsN' in r) r.rolls = r.rollsN ? 0 : r.rollsIn || true;
  delete r.rollsN; delete r.rollsIn;
  if (Object.keys(r).length) r.tu = true;  // el $tu mateix: acabes de fer-lo
  return r;
}
// Marca les tasques segons el $tu: si en sabem els minuts, tornen just aleshores ({b: quan}); si no, com si les marquessis a mà
function applyTu(r, now) {
  const s = load();
  document.querySelectorAll('input[data-tu]').forEach(i => {
    const v = r[i.dataset.tu]; if (v === undefined) return;
    if (!v) delete s[i.dataset.id];
    else s[i.dataset.id] = v === true ? now : { b: now + v * 60000 };
  });
  save(s); tick();
}
// El poder de kakera del $tu («Poder: 110%», «consume 36% de su poder», «10+ llaves … (18%)»): {p, c, h}, o null si no hi surt
function readPower(text) {
  const p = text.match(/(?:^|\\n)\\s*(?:poder|power)\\s*:\\s*(\\d+)\\s*%/i), c = text.match(/(\\d+)\\s*%[^\\n]*(?:poder|power)/i);
  const h = text.match(/10\\+[^\\n]*?(\\d+)\\s*%/);
  return p && c ? { p: +p[1], c: +c[1], h: h ? +h[1] : Math.round(c[1] / 2) } : null;
}
// Què gasta cada reacció a kakera: canvia amb les millores, així que surt del darrer $tu enganxat
function powerDraw() {
  let P = null; try { P = JSON.parse(localStorage.getItem('mudae-poder')); } catch (e) {}
  const d = document.querySelector('input[data-tu="kakera"]')?.closest('.task').querySelector('.desc');
  if (P && d) d.textContent = T('Cada reacció gasta {c} % de poder ({h} % amb personatges de 10+ claus). Al $tu de les {t} tenies {p} %.',
    { c: P.c, h: P.h ?? Math.round(P.c / 2), t: hm(new Date(P.t)), p: P.p });
}
// fins a l'01/11/2026 les tasques es deien r0, r1… (la seva posició): les que tenies marcades passen al nom nou
const OLD = { r0: 'tu', r1: 'vote', r2: 'daily', r3: 'oh', r4: 'oc', r5: 'rolls', r6: 'kakera', r8: 'claim', r10: 'rt', r11: 'dk', r17: 'p' };
function load() {
  let s; try { s = JSON.parse(localStorage.getItem('mudae-fet') || '{}'); } catch (e) { return {}; }
  for (const k in s) if (/^r[0-9]/.test(k)) { if (OLD[k] && !(OLD[k] in s)) s[OLD[k]] = s[k]; delete s[k]; }
  return s;
}
function save(s) { try { localStorage.setItem('mudae-fet', JSON.stringify(s)); } catch (e) {} }
const backOf = (i, s) => { const t = s[i.dataset.id]; return typeof t === 'number' ? until(i.dataset.torna, t) : (t && t.b) || 0; };  // {b}: l'hora que deia el $tu
// Tasques que avisen quan tornen soles. El $oh i el $oc no: tornen a mitjanit UTC (la 1 de la matinada).
const NOTI = ['rolls', 'claim', 'daily', 'vote', 'dk', 'rt', 'p'], was = {};
// La cua: el que pots fer, al seu grup d'«Ara» (per l'ordre de la llista); el que has fet s'amaga a «Fets», pel que torna abans
const TASKS = [...document.querySelectorAll('.task[data-home]')], later = document.getElementById('later');
const byBack = (a, b) => a.back === b.back ? 0 : a.back < b.back ? -1 : 1;  // Infinity (no sabem quan torna) al final
function place(s) {
  const f = document.activeElement, done = [];
  TASKS.forEach(t => {
    const i = t.querySelector('input');
    if (i.checked) done.push({ t, back: backOf(i, s) });
    else document.querySelector(`.grp[data-period="${t.dataset.home}"] .tl`).append(t);
  });
  done.sort(byBack).forEach(x => later.append(x.t));
  document.querySelectorAll('.now-col .grp').forEach(g => { g.hidden = !g.querySelector('.task'); });
  document.getElementById('nowNone').hidden = done.length < TASKS.length;
  document.getElementById('laterNone').hidden = done.length > 0;
  later.closest('details').hidden = !done.length; document.getElementById('laterN').textContent = done.length;
  if (f && f !== document.activeElement && f.isConnected) f.focus({ preventScroll: true });  // en moure'l, no perdis on eres
}
// «Estat»: sí o no de cada cosa que va i ve, i fins quan o quan torna. La kakera, amb les reaccions que et donava el poder del darrer $tu.
const STATE = ['claim', 'rolls', 'kakera', 'p', 'rt', 'dk'];
function stateDraw(now) {
  let P = null; try { P = JSON.parse(localStorage.getItem('mudae-poder')); } catch (e) {}
  const nc = nextClaim(+now);
  document.getElementById('state').innerHTML = STATE.map(k => {
    const i = document.querySelector(`input[data-id="${k}"]`); if (!i) return '';
    const t = i.closest('.task'), cmd = t.querySelector('.cmd, code');
    const more = i.checked ? t.querySelector('.when').textContent
      : k === 'kakera' && P ? T('{n} reaccions: tenies {p} % de poder al $tu de les {t}', { n: Math.floor(P.p / P.c), p: P.p, t: hm(new Date(P.t)) })
      : k === 'claim' && nc != null ? T('fins a les {h} ({in})', { h: hm(new Date(nc)), in: fmt(Math.ceil((nc - now) / 60000)) }) : '';
    return `<div class="st"><span><b>${esc(t.querySelector('b').textContent)}</b> ${cmd ? `<code>${esc(cmd.textContent)}</code>` : ''}</span>`
      + `<span class="tag ${i.checked ? 'bad' : 'good'}">${esc(i.checked ? T('Encara no') : T('Ara pots'))}</span>` + (more ? `<span class="muted">${esc(more)}</span>` : '') + '</div>';
  }).join('');
}
function tick() {
  const now = new Date(), s = load();
  document.querySelectorAll('input[data-torna]').forEach(i => {
    const back = backOf(i, s);
    i.checked = back > +now;
    const when = i.closest('.task').querySelector('.when');
    if (when) when.textContent = !i.checked ? '' : isFinite(back) ? T('Torna a les {h} ({rest})', { h: hm(new Date(back)), rest: fmt(Math.ceil((back - now) / 60000)) })
      : T('No sé quan torna: enganxa el {c}', { c: i.dataset.torna === 'rt' ? '$bonus' : '$tu' });
  });
  place(s); stateDraw(now);
  // a la barra de pestanyes (es veu des de totes): quan tornen les tirades i el reclam, segons el que has enganxat
  const nr = nextRoll(+now), nc = nextClaim(+now), lc = lastClaim(+now), mins = x => Math.ceil((x - now) / 60000);
  // la primera hora després del reinici del reclam, si no has reclamat ja (el $tu ho diu: «no puedes reclamar hasta…»)
  const fresh = lc != null && now - lc < HOUR && !document.querySelector('input[data-tu="claim"]')?.checked;
  const at = x => ({ h: hm(new Date(x)), in: fmt(mins(x)) }), cs = document.querySelector('.grp[data-period="reclam"] h3 small');
  if (cs) cs.textContent = nc == null ? T('enganxa el $tu') : T('fins a les {h} ({in})', at(nc));  // si no reclames abans del reinici, el perds
  clock.innerHTML = (nr == null ? esc(T('Tirades: enganxa el $tu')) : T('Tirades a les <b>{h}</b> ({in})', at(nr))) + ' · '
    + (nc == null ? esc(RS.next ? T('reclam: enganxa el $tu o el $settings') : T('reclam: enganxa el $tu')) : T('reclam a les <b>{h}</b> ({in})', at(nc))) + (fresh ? ' <span class="tag good">' + esc(T('Pots reclamar')) + '</span>' : '');
  // tasques que acaben de tornar soles (no les has desmarcat tu, que n'esborra l'hora): avisa
  const back = [];
  NOTI.forEach(k => document.querySelectorAll(`input[data-tu="${k}"]`).forEach(i => {
    if (was[i.dataset.id] && !i.checked && s[i.dataset.id]) back.push(i);
    was[i.dataset.id] = i.checked;
  }));
  if (back.length && notiOn()) notify(back);
}
// Notificacions: només mentre el tauler és obert. Diuen que tornes a tenir tirades i com tens el reclam i la kakera.
const notiBtn = document.getElementById('notiBtn'), HAS_NOTI = 'Notification' in window;
const notiOn = () => { try { return HAS_NOTI && Notification.permission === 'granted' && localStorage.getItem('mudae-noti') === '1'; } catch (e) { return false; } };
function notiBody() {
  const s = load(), box = k => document.querySelector(`input[data-tu="${k}"]`), at = i => { const b = backOf(i, s); return isFinite(b) ? hm(new Date(b)) : '?'; };
  const c = box('claim'), k = box('kakera'), out = [];
  if (c) out.push(c.checked ? T('Reclam: a les {h}', { h: at(c) }) : T('Pots reclamar personatge.'));
  if (k) out.push(k.checked ? T('Kakera: a les {h}', { h: at(k) }) : T('Pots reaccionar a kakera.'));
  let P = null; try { P = JSON.parse(localStorage.getItem('mudae-poder')); } catch (e) {}
  if (P) out.push(T('Poder al $tu de les {h}: {p}% ({n} botons de kakera)', { h: hm(new Date(P.t)), p: P.p, n: Math.floor(P.p / P.c) }));
  return out.join('\\n');
}
// Un so curt al tauler: Windows sovint no en fa. El navegador només deixa sonar si ja has clicat la pàgina.
let audio;
const audioOn = () => { try { audio = audio || new AudioContext(); audio.resume(); } catch (e) {} };
document.addEventListener('pointerdown', audioOn, { once: true });
function beep() {
  audioOn(); if (!audio) return;
  try {
    [0, 0.3].forEach(d => {
      const o = audio.createOscillator(), g = audio.createGain(), t = audio.currentTime + d;
      o.frequency.value = 880; g.gain.setValueAtTime(0.25, t); g.gain.exponentialRampToValueAtTime(0.001, t + 0.25);
      o.connect(g).connect(audio.destination); o.start(t); o.stop(t + 0.25);
    });
  } catch (e) {}
}
const TITLE = document.title;  // la pestanya avisa fins que hi tornes
document.addEventListener('visibilitychange', () => { if (!document.hidden) document.title = TITLE; });
// Les tirades, amb el reclam i la kakera; la resta de tasques que han tornat, per la seva ordre («$daily, $vote»)
function notify(back) {
  beep();
  const rolls = back.some(i => i.dataset.tu === 'rolls');
  const names = back.filter(i => i.dataset.tu !== 'rolls' && !(rolls && i.dataset.tu === 'claim'))  // el reclam ja surt al text de les tirades
    .map(i => { const t = i.closest('.task'); return (t.querySelector('.cmd') || t.querySelector('b')).textContent.replace(/^\\d+\\.\\s*/, ''); }).join(', ');  // l'ordre ($daily); si no en té, el nom (Reclamar)
  const title = rolls ? T('Ja pots fer $m') : T('Ja pots: {list}', { list: names });
  const body = rolls ? notiBody() + (names ? '\\n' + T('També: {list}', { list: names }) : '') : '';
  if (document.hidden) document.title = '🔔 ' + title;
  // renotify: amb la mateixa etiqueta, si no, la segona surt sense avisar; requireInteraction: es queda fins que la tanques
  try { const n = new Notification(title, { body, tag: 'mudae-m', renotify: true, requireInteraction: true }); n.onclick = () => { window.focus(); n.close(); }; } catch (e) {}
}
function notiDraw() {
  const on = notiOn();
  notiBtn.setAttribute('aria-pressed', on); notiBtn.querySelector('.dot').className = 'dot' + (on ? ' ok' : '');
  notiBtn.lastChild.textContent = on ? T('Notificacions: sí') : T('Notificacions: no');
}
if (notiBtn && HAS_NOTI) {
  notiBtn.hidden = false; notiDraw();
  notiBtn.onclick = async () => {
    if (notiOn()) { try { localStorage.removeItem('mudae-noti'); } catch (e) {} return notiDraw(); }
    if (await Notification.requestPermission() !== 'granted') {
      toolMsg.innerHTML = '<span class="tag bad">' + esc(T('Notificacions bloquejades')) + '</span> ' + esc(T("Permet-les al navegador: el cadenat de la barra d'adreces."));
      return notiDraw();
    }
    try { localStorage.setItem('mudae-noti', '1'); } catch (e) {}
    notiDraw(); beep();
    try { new Notification(T('Notificacions activades'), { body: T("T'avisaré quan tornis a tenir tirades, el reclam, el $daily, el $vote, el $dk, el $rt o el $p, mentre el tauler sigui obert.") }); } catch (e) {}
  };
}
document.addEventListener('change', ev => {
  const i = ev.target; if (!i.dataset || !i.dataset.torna) return;
  const s = load();
  if (i.checked) s[i.dataset.id] = Date.now(); else delete s[i.dataset.id];
  save(s); tick();
});
// a «Cada hora», el minut de les tirades; a «Estat», les hores dels reclams del teu servidor (o què falta)
function resetDraw() {
  const s = document.querySelector('.grp[data-period="hora"] h3 small'), info = document.getElementById('laterInfo'), nc = nextClaim(Date.now());
  if (s) s.textContent = RS.m == null ? T('enganxa el $tu') : T('minut {m}', { m: pad(RS.m) });
  if (info) info.textContent = nc == null ? '' : !RS.iv ? T('Pròxim reclam a les {h}', { h: hm(new Date(nc)) })
    : T('Reclams a les {list}', { list: Array.from({ length: Math.min(24, Math.floor(1440 / RS.iv)) }, (_, k) => new Date(nc + k * RS.iv * MIN)).sort((a, b) => a.getHours() * 60 + a.getMinutes() - b.getHours() * 60 - b.getMinutes()).map(hm).join(', ') });
}
powerDraw(); resetDraw(); tick(); setInterval(tick, 30000);

// mentre la pestanya és oberta, diu al servidor que continua aquí; si es tanca, el servidor plega sol
if (location.protocol.startsWith('http')) setInterval(() => fetch('/ping', { cache: 'no-store' }).catch(() => {}), 20000);
"""


JS_COLL = """
const COLL = __DATA__;
const album = document.querySelector('#coll .album'), chips = document.getElementById('chips'), hide = document.getElementById('hideMine'), title = document.getElementById('albumTitle');
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
  document.querySelectorAll('.row[data-serie]').forEach(r => {
    const on = r.dataset.serie === cur;
    r.classList.toggle('on', on && !album.hidden); r.setAttribute('aria-expanded', String(on && !album.hidden));
    if (on) r.after(album);  // l'àlbum s'obre just sota la barra que has clicat, no a baix de tot
  });
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
document.getElementById('collDel').onclick = async ev => {  // ja no la fas: fora de les col·leccions (el harem no es toca)
  if (location.protocol === 'file:') { gotMsg.innerHTML = OFF_MSG(); gotBar.hidden = false; return; }
  if (!confirm(T("Deixes de fer «{s}»? S'esborra de les col·leccions; el teu harem no es toca.", { s: cur }))) return;
  ev.target.disabled = true;
  try {
    const j = await postJSON('/esborra', { name: cur });
    if (j.ok) { try { localStorage.removeItem('mudae-coll'); } catch (e) {} location.reload(); }
    else { gotBar.hidden = false; gotMsg.innerHTML = ERR_MSG(j.error); ev.target.disabled = false; }
  } catch (err) { gotBar.hidden = false; gotMsg.innerHTML = NET_MSG(); ev.target.disabled = false; }
};
document.getElementById('gotBtn').onclick = async ev => {
  const btn = ev.target;
  if (location.protocol === 'file:') { gotMsg.innerHTML = OFF_MSG(); return; }
  btn.disabled = true; gotMsg.textContent = T('Desant…');
  try {
    const j = await postJSON('/rebut', { names: [...got] });
    if (j.ok) location.reload(); else { gotMsg.innerHTML = ERR_MSG(j.error); btn.disabled = false; }
  } catch (err) { gotMsg.innerHTML = NET_MSG(); btn.disabled = false; }
};
function show(s) {  // clicar la sèrie oberta la tanca
  album.hidden = s === cur && !album.hidden; cur = s;
  try { localStorage.setItem('mudae-coll', s); } catch (e) {}
  draw();
  if (!album.hidden) album.scrollIntoView({ block: 'nearest', behavior: matchMedia('(prefers-reduced-motion: reduce)').matches ? 'auto' : 'smooth' });
}
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
  --bg:#0e0f15; --bg-grad:linear-gradient(160deg,#161a28 0%,#0f1017 45%,#0c0d12 100%);
  --glow1:rgba(88,166,255,.16); --glow2:rgba(167,139,250,.11);
  --glass:rgba(255,255,255,.045); --glass-b:rgba(255,255,255,.085); --line:rgba(255,255,255,.075);
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
html { background:var(--bg); -webkit-tap-highlight-color:transparent; scroll-padding-top:80px; }  /* que la barra fixa no tapi el camp on ets */
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
h2 { font-size:16.5px; font-weight:650; letter-spacing:-.3px; }
a { color:var(--accent-ink); }
button { font:inherit; color:inherit; }
:focus-visible { outline:2px solid var(--accent); outline-offset:3px; border-radius:6px; }
.muted { color:var(--muted); }
code { font:12.5px var(--mono); background:var(--sunk); color:var(--ink-2); padding:1px 6px; border-radius:5px; white-space:nowrap; }
/* sense backdrop-filter a les targetes: a sota només hi ha el fons, i Chrome hi pintava franges en repintar (en obrir el $oc o passar-hi el ratolí) */
.glass, .panel { background:var(--glass);
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
/* menú «···»: les opcions que es fan servir poc */
.more { position:relative; }
.more > summary { list-style:none; cursor:pointer; letter-spacing:2px; font-weight:700; }
.more > summary::-webkit-details-marker { display:none; }
.more > summary:hover, .more[open] > summary { color:var(--ink); border-color:rgba(255,255,255,.2); }
.menu { position:absolute; right:0; top:calc(100% + 6px); z-index:10; display:grid; gap:2px; padding:6px; min-width:250px; border-radius:14px;
  background:rgba(28,28,32,.96); backdrop-filter:blur(22px); -webkit-backdrop-filter:blur(22px); border:1px solid var(--glass-b); box-shadow:var(--sh); }
.menu button { display:flex; align-items:center; gap:8px; text-align:left; padding:9px 12px; border:0; border-radius:9px; background:transparent;
  color:var(--ink-2); font-size:13px; cursor:pointer; }
.menu button:hover:not(:disabled) { background:rgba(255,255,255,.07); color:var(--ink); }
.menu button:disabled { opacity:.45; cursor:default; }

/* barra fixa: pestanyes i, a la dreta, quan tornen les tirades i el reclam (es veu des de totes les pestanyes) */
.bar { position:sticky; top:10px; z-index:4; display:flex; flex-wrap:wrap; align-items:center; justify-content:space-between; gap:8px; }
.clock { margin:0; padding:8px 14px; border-radius:50px; font-size:13px; color:var(--ink-2); font-variant-numeric:tabular-nums;
  background:rgba(28,28,32,.75); backdrop-filter:blur(22px); -webkit-backdrop-filter:blur(22px); border:1px solid var(--glass-b); box-shadow:var(--sh); }
.clock:empty { display:none; }
.clock b { color:var(--ink); }

/* resum (només a Avui): salutació i quatre xifres que porten a la seva pestanya */
.resum { display:grid; grid-template-columns:auto minmax(0,1fr); gap:28px; align-items:center; padding:20px 26px;
  background:linear-gradient(115deg,rgba(88,166,255,.13),rgba(167,139,250,.08) 55%,rgba(255,255,255,.03)); }
.hola { font-size:26px; font-weight:650; letter-spacing:-.7px; line-height:1.1; }
.alerts { list-style:none; padding:0; margin:0; display:grid; gap:6px; max-height:340px; overflow:auto; }
.alerts li { display:flex; gap:10px; align-items:baseline; padding:8px 12px; border-radius:10px; background:rgba(255,255,255,.04); border:1px solid var(--line); font-size:13.5px; }
.alerts li > span:nth-child(2) { flex:1; min-width:0; }
.go { flex:none; border:0; background:transparent; color:var(--muted); font-size:12px; font-weight:600; cursor:pointer; padding:0; white-space:nowrap; }
.go:hover { color:var(--accent-ink); }
.fets { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); }
.fet { display:grid; text-align:left; padding:4px 18px; border:0; border-left:1px solid var(--line); border-radius:0; background:transparent; min-width:0; cursor:pointer; }
.fet-l { font-size:12px; color:var(--ink-2); }
.fet-v { font-size:24px; font-weight:650; letter-spacing:-.6px; font-variant-numeric:tabular-nums; transition:color .2s; }
.fet:hover .fet-v { color:var(--accent-ink); }
.fet-s { font-size:12px; color:var(--muted); }
.col { display:grid; gap:14px; align-content:start; min-width:0; }
/* targetes plegables: el que es fa servir poc s'obre amb un clic */
details.panel { display:block; }
.fold > summary { list-style:none; cursor:pointer; display:flex; flex-wrap:wrap; align-items:baseline; gap:4px 10px; }
.fold > summary::-webkit-details-marker { display:none; }
.fold > summary::before { content:'›'; display:inline-block; width:10px; color:var(--muted); font-weight:700; transition:transform .2s; }
.fold[open] > summary::before { transform:rotate(90deg); }
.fold > summary:hover h2 { color:var(--accent-ink); }
.fold-body { display:grid; gap:14px; margin-top:14px; }
.fold > summary .og-h { margin:0; }
.fold > summary .muted { font-size:12px; }
.serieslist.fold-body { display:flex; gap:6px; margin-top:10px; }
.tag { font-size:11px; font-weight:650; padding:2px 9px; border-radius:20px; white-space:nowrap; flex:none; background:rgba(255,255,255,.08); color:var(--ink-2); }
.tag.good { background:var(--good-soft); color:var(--good); }
.tag.bad { background:var(--bad-soft); color:var(--bad); }
.tag.accent { background:var(--accent-soft); color:var(--accent); }

/* pestanyes: una vista cada vegada */
.nav { display:flex; gap:4px; padding:5px; border-radius:50px; width:fit-content; max-width:100%; overflow-x:auto;
  background:rgba(28,28,32,.75); backdrop-filter:blur(22px); -webkit-backdrop-filter:blur(22px); border:1px solid var(--glass-b); box-shadow:var(--sh); }
.nav button { padding:8px 16px; border-radius:50px; border:0; background:transparent; color:var(--muted); font-size:13.5px; font-weight:600; cursor:pointer; white-space:nowrap;
  transition:background .2s, color .2s; }
.nav button:hover { color:var(--ink-2); background:rgba(255,255,255,.06); }
.nav button[aria-selected="true"] { background:rgba(var(--acc-rgb),.14); color:var(--accent); }
.nav .n { font-size:11px; margin-left:6px; padding:1px 7px; border-radius:20px; background:var(--bad-soft); color:var(--bad); }
.view { display:grid; gap:14px; animation:fu .3s ease both; }
.g2 { display:grid; grid-template-columns:repeat(auto-fit, minmax(340px, 1fr)); gap:14px; align-items:start; }

/* avui: llista + $oc */
.today { gap:16px; }
.today-head { display:flex; flex-wrap:wrap; align-items:center; gap:10px 16px; }
.paste { grid-column:1 / -1; display:flex; flex-wrap:wrap; align-items:center; gap:8px 12px; }
.paste textarea { flex:1; min-width:240px; height:40px; min-height:40px; padding:9px 12px; resize:none; overflow:hidden; }
#tuMsg:empty, #pdMsg:empty { display:none; }
/* els botons de copiar, dins d'un requadre igual que la casella de text */
.pend { grid-column:1 / -1; padding:10px 14px; border-radius:12px; background:var(--accent-soft); border:1px solid rgba(var(--acc-rgb),.3); font-size:13.5px; }
.pend > b { color:var(--accent-ink); }
.pend ul { margin:6px 0 0; padding:0; list-style:none; display:grid; gap:4px; }
.paste-cmds { flex-basis:100%; display:flex; flex-wrap:wrap; align-items:center; gap:6px; min-height:40px; padding:6px 12px; font-size:12.5px;
  background:rgba(0,0,0,.22); border:1px solid var(--glass-b); border-radius:10px; }
#pdSum { display:grid; gap:8px; }
.pdrow { display:grid; grid-template-columns:minmax(0,200px) 1fr 120px; gap:14px; align-items:center; font-size:13.5px; }
.pdrow .val { font-size:13px; }
.groups { display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:12px; align-items:start; }
.grp { display:grid; padding:14px 16px 4px; border-radius:14px; background:rgba(255,255,255,.035); border:1px solid var(--line); transition:border-color .3s, background .3s; }
.grp h3 { margin:0 0 4px; font-size:14px; font-weight:650; display:grid; gap:1px; }
.grp h3 small { font-size:11.5px; font-weight:500; color:var(--muted); }
.plan { display:grid; grid-template-columns:minmax(0,1fr) minmax(0,1fr); gap:12px; align-items:start; }
.now-col { display:grid; gap:12px; }
.grp.claim { border-color:rgba(62,207,142,.45); background:rgba(62,207,142,.06); }
.grp.claim h3 { color:var(--good); }
.grp .tip { margin:0 0 12px; font-size:13px; color:var(--muted); }
.plan .none { margin:2px 0 12px; font-size:13.5px; }
.st { display:grid; grid-template-columns:minmax(0,1fr) auto; align-items:center; gap:2px 10px; padding:7px 0; border-bottom:1px solid var(--line); font-size:13.5px; }
.st:last-child { border-bottom:0; }
.st > .muted { grid-column:1 / -1; font-size:12.5px; }
.done-fold { margin:6px 0 10px; }
#later { gap:0; margin-top:6px; }
#later .task { padding:6px 0; }
#later .task > span { flex:1; display:flex; flex-wrap:wrap; align-items:baseline; gap:2px 8px; }
#later .when { display:inline; margin-left:auto; font-weight:500; }
#later .desc { display:none; }
#later .task:has(input:checked) > span > :not(.when) { opacity:.7; text-decoration:none; }
.today .fold > summary h3 { margin:0; font-size:13.5px; font-weight:600; color:var(--ink-2); }
.og h3 { margin:10px 0 2px; font-size:11.5px; font-weight:650; letter-spacing:1.2px; text-transform:uppercase; color:var(--muted); }
.task { display:flex; gap:12px; align-items:flex-start; padding:9px 0; border-bottom:1px solid var(--line); cursor:pointer; }
.task:last-child { border-bottom:0; }
.task input { width:17px; height:17px; margin-top:3px; accent-color:var(--accent); flex:none; }
.task > span { min-width:0; }
.task:has(input:checked) > span > :not(.when) { opacity:.45; text-decoration:line-through; }
.task .desc { display:block; color:var(--muted); font-size:13px; }
.task .when { display:block; font-size:12.5px; font-weight:650; color:var(--good); }
.bullet { width:17px; text-align:center; color:var(--muted); flex:none; }
details:not(.fold, .more) summary { cursor:pointer; color:var(--ink-2); font-size:13.5px; margin-top:6px; }
details:not(.fold, .more)[open] summary { margin-bottom:6px; }
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
.bars .album { margin:4px 0 12px; padding:14px 0; border-bottom:1px solid var(--line); }
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
textarea, select, input[type=text], input[type=search] { font:14px var(--sans); color:var(--ink); background:rgba(0,0,0,.22); border:1px solid var(--glass-b); border-radius:10px; }
textarea { padding:10px 12px; resize:vertical; width:100%; }
select, input[type=text], input[type=search] { padding:6px 10px; border-radius:8px; }
select option { background:var(--surface); }
textarea:focus, input[type=text]:focus, input[type=search]:focus, select:focus { outline:2px solid var(--accent); outline-offset:0; border-color:transparent; }
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
.ogrid { display:grid; grid-template-columns:repeat(auto-fit, minmax(300px, 1fr)); gap:10px 28px; align-items:start; }
#ordres .og > summary { padding:6px 0; border-bottom:1px solid var(--line); }
#ordres .og > summary h3 { margin:0; }
#ordres .og > summary .muted { font-size:12px; }
#ordres .fold-body { gap:0; margin-top:6px; }
#ordQ { width:280px; max-width:100%; }
.oline { display:flex; gap:10px; align-items:baseline; padding:4px 0; font-size:13.5px; color:var(--ink-2); }
.oline .cmd { flex:none; }
.steps { list-style:none; margin:0; padding:0; display:grid; gap:6px; font-size:13.5px; color:var(--ink-2); }
.steps label { display:flex; gap:8px; align-items:baseline; cursor:pointer; }
.steps input { accent-color:var(--accent); }
.og-h { margin:4px 0 0; font-size:11.5px; font-weight:650; letter-spacing:1.2px; text-transform:uppercase; color:var(--muted); }

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
@media (max-width:900px) { .resum, .plan { grid-template-columns:1fr; gap:14px; } }
@media (max-width:640px) {
  main { padding-left:max(16px, env(safe-area-inset-left, 0px)); padding-right:max(16px, env(safe-area-inset-right, 0px)); }
  .resum { padding:18px; } .hola { font-size:22px; } .panel { padding:16px; }
  .fets { grid-template-columns:repeat(2,minmax(0,1fr)); }
  .g2 { grid-template-columns:1fr; }
  .row { grid-template-columns:1fr auto; } .row .track { grid-column:1 / -1; grid-row:2; }
}
@media (prefers-reduced-motion:reduce) { *, *::before, *::after { animation:none !important; transition:none !important; scroll-behavior:auto !important; } }
"""


# Pestanya Millores. Insígnies ($kakerareward): (nom, preu inicial, què dona). Cada nivell costa preu inicial × nivell.
BADGES = [
    ("Bronze", 1000, "+1 lloc a la wishlist per nivell"),
    ("Silver", 2000, "+25 % de sortir personatges de la wishlist per nivell"),
    ("Gold", 3000, "−10 % del poder que gastes en clicar kakera, per nivell"),
    ("Sapphire", 5000, "+1 tirada per nivell"),
    ("Ruby", 7000, "Wishlist, tirades i menys poder; al IV, totes les insígnies costen −25 %"),
    ("Emerald", 9000, "Desbloqueja el $rt (tornar a reclamar); cada nivell, 10 h menys d'espera"),
    ("Diamond", 12000, "+1 $oh cada dia; al IV, +1 $oc"),
]
# Ordre recomanat (wiki de Mudae, «Kakera Badges», 01/11/2026; l'ordre és proposta de Claude):
# 1. Gold fins al IV: cada nivell, −10 % del poder per reacció (més kakera de cada tirada), i el IV fa que el $dk et torni tot el poder.
# 2. Bronze II i Silver II (barats) desbloquegen el Ruby. 3. Ruby I-IV: −20 % de poder, +2 tirades, +2 llocs i +50 % de wishlist,
#    i el IV abaixa un 25 % totes les insígnies que queden (la ruta més barata de la wiki: B II, S II, G II i Ruby IV).
# 4. Sapphire (tirades), Emerald ($rt i, al IV, la kakera de cada personatge que reclames), Bronze i Silver al IV, Diamond (esferes).
# Respecta els desbloquejos: Sapphire, B, S i G a I; Ruby, a II; Emerald, a III o dues insígnies al IV.
BADGE_PLAN = [("Bronze", 1), ("Silver", 1), ("Gold", 1), ("Gold", 2), ("Gold", 3), ("Gold", 4), ("Bronze", 2), ("Silver", 2),
              *(("Ruby", n) for n in range(1, 5)), *(("Sapphire", n) for n in range(1, 5)), *(("Emerald", n) for n in range(1, 5)),
              ("Bronze", 3), ("Silver", 3), ("Bronze", 4), ("Silver", 4), *(("Diamond", n) for n in range(1, 5))]
# Pisos de la torre ($kt), de més a menys recomanat (proposta de Claude amb la wiki, «Kakera Tower/Perks»). Cada pis costa més que l'anterior,
# triïs el que triïs. El 8 i el 7 junts: amb el Gold IV i el Ruby III, el poder ple (110 %) dona 3 reaccions de 36 % en lloc de 2 de 40 %.
TOWER = [(8, "Clicar kakera gasta un 4 % menys de poder"), (7, "Poder màxim de kakera al 110 %"), (11, "+1 tirada per hora"),
         (6, "+30 esferes cada vegada que reclames"), (1, "+2 llocs a la wishlist"), (5, "El $oh et destapa 2 botons"),
         (4, "+10 % de segona clau si el personatge és a la wishlist"), (2, "El teu $starwish surt un 50 % més"),
         (9, "+500 de kakera cada 3 h a les claus daurades ($bku)"), (10, "La kakera de llum dona +1 kakera"),
         (3, "Menys límit de $limroul (depèn del servidor)"), (12, "Colors per al $mm i el $profile")]
# Millores d'esferes d'un personatge ($opp), pel número que té a Mudae, agrupades de més a menys recomanades
PERKS = {1: "Més probabilitat que surtin els personatges del costat a la teva wishlist", 2: "+20 de kakera base",
         3: "Probabilitat de +1 botó de kakera", 4: "Probabilitat de +1 clau",
         5: "Esferes per cada botó de kakera que hi cliques", 6: "Pot fer sortir un personatge de la teva wishlist",
         7: "Els botons de kakera poden passar a kakera del caos",
         8: "El primer cop del dia: 4 botons de kakera a meitat de poder, només per a tu",
         9: "El primer cop del dia: un botó d'esferes", 10: "El primer $oh del dia dona +20 esferes"}
PERK_GROUPS = [("Funciona cada dia, encara que no et surti el personatge, i se suma: +20 esferes per cada personatge que la té (fins a 100 personatges; després, menys). Es paga en 50 dies", [10]),
               ("Només quan et surt aquest personatge: posa-les al que et surt més (mira el $mmy)", [4, 5, 2, 3]),
               ("Només el primer cop del dia que te'l treus", [8, 9]),
               ("Les menys útils", [1, 6, 7])]
# Botiga d'esferes ($os): bonus per a tots els personatges, 10 nivells (4.000 esferes el primer, +4.000 cada nivell)
SHOP = {1: "Part del bonus d'aparició de la millora 1 també val per al personatge millorat",
        2: "Megaesferes: 1/50 en tirar qualsevol personatge teu (1 al dia), amb més premis per nivell",
        3: "El botó de més de la millora 3 pot no ser de kakera blava",
        4: "Les claus de la millora 4 poden ser claus Omega, per a qualsevol personatge ($ok)",
        5: "Les esferes de la millora 5 poden donar +1 $ot",
        6: "El desig de la millora 6 pot ser un de la teva wishlist sense reclamar, i reclamar-lo és gratis",
        7: "La kakera del caos de la millora 7 pot donar el doble",
        8: "Els botons de la millora 8 donen més kakera als personatges millorats del tot",
        9: "Més botons d'esferes de la millora 9 al dia, i donen més esferes",
        10: "El primer $oh del dia pot donar $ot per cada personatge millorat del tot"}
# Ordre recomanat (proposta de Claude): la 2 és l'única que no depèn d'haver millorat personatges; la 10 en demana de millorats del tot
SHOP_ORDER = [2, 4, 9, 5, 3, 1, 8, 6, 7, 10]


ORDRES = [  # (grup, [(ordre, què fa)]): requadre «Ordres d'ús habitual» del tauler
    # el primer és el resum (obert d'entrada): les de cada dia i les que entén el tauler, repetides del seu grup
    ("Les més útils", [("$tu", "Tot el que tens disponible"), ("$m", "Tirada: personatge qualsevol"),
                       ("$daily", "Guarda un reinici de tirades (cada 20 h)"), ("$vote", "Vota i guanya un reinici"),
                       ("$rt", "Torna a poder reclamar"), ("$dk", "Recarrega el poder i dona kakera"),
                       ("$oh", "Esferes (4 al dia)"), ("$oc", "Troba la vermella: fes servir l'ajudant"),
                       ("$p", "Tirada de Pokémon: cada 2 h, per a tots els servidors"), ("$k", "La teva kakera i com guanyar-ne"),
                       ("$kt", "Pisos, beneficis i preu del següent"), ("$opz", "Els teus personatges que tenen esferes"),
                       ("$opp <personatge>", "Millora un personatge amb esferes"), ("$or <personatge>", "Et torna les esferes d'un personatge (perds el 10 %)"),
                       ("$os", "Botiga d'esferes: millores per a tots"), ("$shoprefund", "Et torna les esferes de la botiga (cada 15 dies)"),
                       ("$wl", "La teva wishlist"), ("$mma+s", "El teu harem sencer amb la sèrie de cada un, per privat"),
                       ("$imaso <sèrie>", "Una sèrie sencera amb propietaris, per privat"), ("$bonus", "Els bonus que tens"),
                       ("$settings", "Com té configurat Mudae el servidor: cada quant hi ha reclam")]),
    ("Tirades", [("$m", "Tirada: personatge qualsevol"), ("$w", "Tirada: només noies"), ("$h", "Tirada: només nois"),
                 ("$ma", "Tirada: qualsevol, d'anime"), ("$wa", "Tirada: noies d'anime"), ("$ha", "Tirada: nois d'anime"),
                 ("$mg", "Tirada: qualsevol, de videojocs"), ("$wg", "Tirada: noies de videojocs"), ("$hg", "Tirada: nois de videojocs"),
                 ("$us", "Gasta tirades guardades"), ("$rolls", "Gasta un reinici guardat del $vote o el $daily"),
                 ("$daily", "Guarda un reinici de tirades (cada 20 h)"), ("$vote", "Vota i guanya un reinici")]),
    ("Reclamar i kakera", [("$tu", "Tot el que tens disponible"), ("$rt", "Torna a poder reclamar"), ("$dk", "Recarrega el poder i dona kakera"),
                           ("$mu", "Quan podràs reclamar"), ("$k", "La teva kakera i com guanyar-ne")]),
    ("La teva col·lecció", [("$mma+s", "El teu harem sencer amb la sèrie de cada un, per privat"), ("$mmbs", "El harem amb la sèrie de cada un, per privat"),
                            ("$imaso <sèrie>", "Una sèrie sencera amb propietaris, per privat"), ("$im <personatge>", "Fitxa d'un personatge"),
                            ("$wl", "La teva wishlist"), ("$wish <nom>", "Afegeix a la wishlist (diversos: separa'ls amb $)"),
                            ("$wishremove <nom>", "Treu de la wishlist"), ("$sw <nom>", "Desig estrella")]),
    # de la wiki de Mudae («List of Commands»), 31/10/2026
    ("Harem i perfil", [("$mm", "El teu harem"), ("$fm <personatge>", "El teu preferit: surt primer al $mm i al $pr"),
                        ("$sm <personatge>", "Ordena el harem: $sm Nom 1 $ Nom 2…"), ("$note <personatge>", "Una nota al harem: $note Nom $ text"),
                        ("$fn <text>", "Busca les teves notes"), ("$c <personatge>", "Tria quina imatge surt primer: $c Nom $ número"),
                        ("$a <personatge>", "Canvia el nom pel d'un àlies"), ("$pr", "El teu perfil"),
                        ("$divorce <personatge>", "Te'n divorcies i et dona la seva kakera"), ("$like <personatge>", "L'afegeix a la llista de m'agrada"),
                        ("$ll", "La llista de m'agrada: amb 5, desbloqueges el $daily"), ("$left", "Quants personatges queden per reclamar al servidor"),
                        ("$top", "Els 1.000 personatges més reclamats"), ("$ts", "Rànquing del servidor")]),
    ("Wishlist i tirades", [("$wishl <nom>", "Desig bloquejat: el $wishpurge no el treu"), ("$wishpurge", "Treu de la wishlist els reclamats i els repetits"),
                            ("$wishsort", "Ordena la wishlist per ordre alfabètic"), ("$wishdm y y y", "Wishlist privada i avís per privat"),
                            ("$ru", "Tirades que et queden"), ("$ku", "Quan pots tornar a reaccionar a kakera"),
                            ("$d <sèrie>", "Treu una sèrie de les teves tirades"), ("$dl", "Les sèries que has tret de les tirades"),
                            ("$ad <sèrie>", "Torna a les teves tirades una sèrie que havies tret"), ("$qs full", "Desa una còpia de la wishlist i de les llistes")]),
    ("Intercanvis", [("$give @persona", "Regala personatges (separa'ls amb $)"), ("$trade @persona", "Intercanvi amb confirmació dels dos")]),
    ("Esferes", [("$oh", "Esferes (4 al dia)"), ("$oc", "Troba la vermella: fes servir l'ajudant"), ("$oq", "Joc d'esferes"), ("$ot", "Joc d'esferes"),
                 ("$opp <personatge>", "Millora un personatge amb esferes"), ("$opz", "Els teus personatges que tenen esferes"),
                 ("$or <personatge>", "Et torna les esferes d'un personatge (perds el 10 %)"), ("$os", "Botiga d'esferes: millores per a tots"),
                 ("$shoprefund", "Et torna les esferes de la botiga (cada 15 dies)"), ("$ohu", "Quan tornen les esferes i les millores")]),
    ("Torre (kakeratower)", [("$kt", "Pisos, beneficis i preu del següent"), ("$towervalue", "Quant costa cada pis"),
                             ("$build 8", "Pis: reaccions a kakera −4 % de poder"), ("$build 11", "Pis: +1 tirada per hora"),
                             ("$build <número>", "Construeix el pis que triïs"), ("$destroy", "Desfà la torre i et torna la kakera")]),
    ("Kakeraloots", [("$kl", "Un kakeraloot (500 de kakera)"), ("$kl <número>", "Diversos alhora"), ("$lk", "Les teves estadístiques i inventari"),
                     ("$infokl", "Què donen els kakeraloots"), ("$quantity <número>", "Puja nivells de quantitat"), ("$quality <número>", "Puja nivells de qualitat")]),
    ("Kakera, insígnies i claus", [("$kakerareward", "Recompenses de la kakera"), ("$badges", "Les teves insígnies"), ("$bronze", "Insígnia de bronze (i $silver, $gold…)"),
                                   ("$bku", "Bonus de les claus daurades"), ("$infokeys", "Com funcionen les claus"), ("$givekakera @persona <número>", "Regala kakera")]),
    ("Pokéslot", [("$p", "Tirada de Pokémon: cada 2 h, per a tots els servidors"), ("$pd", "La teva Pokédex"),
                  ("$arl", "Allibera els repetits i et dona tirades"), ("$r <pokémon>", "N'allibera un i et dona tirades"),
                  ("$sh <pokémon>", "Cacera de shiny d'aquest Pokémon"), ("$pl <pokémon>", "El protegeix perquè l'$arl no l'alliberi")]),
    ("Jocs i ajuda", [("$quiz", "Preguntes per a tothom"), ("$jankenpon @persona", "Pedra, paper, tisora"),
                      ("$greentea", "Joc de paraules: troba'n una, ràpid"), ("$blacktea", "Joc de paraules per torns"),
                      ("$s <paraula>", "Busca qualsevol ordre de Mudae"), ("$overview", "La teva configuració de Mudae"), ("$bonus", "Els bonus que tens"),
                      ("$sl", "Ànimes bessones: personatges amb 10 claus o més"), ("$me @persona <personatge>", "Intercanvi: pots afegir-hi kakera ($ número)")]),
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
    out.innerHTML = `<details class="fold" data-fold="mine-series"${foldOpen('mine-series') ? ' open' : ''}><summary><h3 class="og-h">`
      + esc(T('Les teves sèries · clica per afegir-la')) + `</h3><span class="muted">${counts.length}</span></summary><div class="serieslist fold-body">`
      + counts.map(([s, n]) => `<button type="button" data-add="${esc(s)}">${esc(label(s))} <span>${n}</span></button>`).join('') + '</div></details>';
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
  if (h === null || h >= 7) age.innerHTML = '<span class="tag bad">' + esc(T('Toca posar al dia el harem: $mma+s')) + '</span> ' + esc(age.textContent);
})();
const serieName = () => serieSel.value === '__new__' ? newIn.value.trim() : serieSel.value;
// totes les sèries d'una vegada: a+ diu la sèrie de cada personatge, o els propietaris, s per privat
const ALL = [...serieSel.options].map(o => o.value).filter(v => v !== '__new__' && !SPECIALS.includes(v));
const allNew = document.getElementById('updAllNew');
const newSeries = () => allNew.value.split('$').map(x => x.trim()).filter(x => x && !ALL.some(a => a.toLowerCase() === x.toLowerCase()));
const allList = () => [...ALL, ...newSeries()];
// Què has enganxat, mirant el text: la wishlist, moltes sèries o el harem («Nom - Sèrie»), o una sèrie (títol «Sèrie 12/34»)
function detect(text) {
  const ls = text.split(/\r?\n/).map(l => l.replace(/^\d+\.\s+/, '').trim()).filter(Boolean);
  if (ls.slice(0, 4).some(l => /wishlist/i.test(l))) return { kind: 'wishlist' };
  if (ls.filter(l => / - /.test(l)).length > ls.length / 2) return { kind: ls.some(l => /=>.* - /.test(l)) ? 'totes' : 'harem' };
  const title = ls.slice(0, 6).map(l => l.match(/^(.*\S)\s+\d+\s*\/\s*\d+$/)).find(Boolean);
  return title ? { kind: 'serie', serie: title[1] } : {};
}
function syncUpd() {
  serieBox.hidden = kindSel.value !== 'serie';
  newIn.hidden = serieSel.value !== '__new__';
  const esp = kindSel.value === 'serie' && SPECIALS.includes(serieName());
  document.getElementById('updEsp').hidden = !esp;
  const cmd = kindSel.value === 'harem' ? '$mma+s' : kindSel.value === 'wishlist' ? '$wl' : kindSel.value === 'totes' ? '$imaa+os ' + allList().join(' $ ')
    : esp ? '$imaso ' : '$imaso ' + (serieName() || '<' + T('sèrie') + '>');
  updCmd.textContent = cmd; updCmd.dataset.copy = cmd;
  const all = kindSel.value === 'totes', allCopy = document.getElementById('updAllCopy');  // l'ordre sencera ben a la vista, amb les sèries
  document.getElementById('updAll').hidden = !all;
  document.getElementById('updAllCmd').textContent = allList().length ? cmd : T('Encara no tens cap sèrie: escriu-les a «Sèries noves».');
  allCopy.dataset.copy = cmd; allCopy.hidden = !allList().length;
  updBtn.disabled = location.protocol === 'file:' || !txt.value.trim() || (kindSel.value === 'serie' && !serieName());
}
[kindSel, serieSel, newIn, txt, allNew].forEach(el => el.addEventListener('input', syncUpd));
// Tria sol «Què enganxes» (i la sèrie, si la tens). newIfMissing: una sèrie que no tens passa a «Sèrie nova…» amb el seu nom
// (des de la casella d'Avui; aquí no, perquè per a una col·lecció especial enganxes el $imaso d'una altra sèrie)
function autoKind(newIfMissing) {
  const d = detect(txt.value); if (!d.kind) return false;
  kindSel.value = d.kind;
  const opt = d.serie && [...serieSel.options].find(o => o.value.toLowerCase() === d.serie.toLowerCase());
  if (opt) serieSel.value = opt.value;
  else if (d.serie && newIfMissing) { serieSel.value = '__new__'; newIn.value = d.serie; }
  [kindSel, serieSel].forEach(el => { try { localStorage.setItem('mudae-' + el.id, el.value); } catch (e) {} });
  syncUpd(); updMsg.innerHTML = '<span class="tag">' + esc(T('Detectat: {k}', { k: kindSel.selectedOptions[0].textContent })) + '</span>';
  return true;
}
// En enganxar, es desa sol si és segur (Victor, 01/11/2026): el harem, la wishlist, totes les sèries o una sèrie que ja tens.
// Una sèrie nova o una col·lecció especial (hi enganxes el $imaso d'una altra sèrie) les tries tu i cliques «Desar i actualitzar».
function autoSave() {
  const d = detect(txt.value);
  if (updBtn.disabled || !d.kind) return false;
  if (d.kind === 'serie' && (serieSel.value === '__new__' || !d.serie || serieName().toLowerCase() !== d.serie.toLowerCase())) return false;
  updBtn.click(); return true;
}
txt.addEventListener('paste', () => setTimeout(() => {
  if (autoKind(false) && !autoSave()) updMsg.innerHTML += ' ' + esc(T('Comprova la sèrie i clica «Desar i actualitzar».'));
}));
// el que vas triar l'última vegada (què enganxes i quina sèrie), per no haver-ho de canviar cada cop
[kindSel, serieSel].forEach(el => {
  try { const v = localStorage.getItem('mudae-' + el.id); if ([...el.options].some(o => o.value === v)) el.value = v; } catch (e) {}
  el.addEventListener('change', () => { try { localStorage.setItem('mudae-' + el.id, el.value); } catch (e) {} });
});
if (location.protocol === 'file:') updMsg.innerHTML = OFF_MSG();
try {  // el que no s'ha pogut posar al dia la vegada anterior (es desa abans de recarregar)
  const w = sessionStorage.getItem('mudae-avis');
  if (w) { sessionStorage.removeItem('mudae-avis'); updMsg.innerHTML = '<span class="tag good">' + esc(T('Desat')) + '</span> <span class="tag bad">' + esc(T('Atenció')) + '</span> ' + esc(w); }
} catch (e) {}
updBtn.onclick = async () => {
  updBtn.disabled = true; updMsg.textContent = T('Desant…');
  try {
    const j = await postJSON('/dades', { kind: kindSel.value, serie: serieName(), text: txt.value, noves: kindSel.value === 'totes' ? newSeries() : [] });
    if (j.ok) { try { if (j.warn) sessionStorage.setItem('mudae-avis', j.warn); } catch (e) {} location.reload(); }
    else { updMsg.innerHTML = ERR_MSG(j.error); syncUpd(); }
  } catch (err) { updMsg.innerHTML = NET_MSG(); syncUpd(); }
};
syncUpd();
"""


JS_MILL = r"""
// Pestanya «Millores»: llegeix el que enganxes de Mudae ($k, $kt, $opp) i recomana el següent pas. Es guarda al navegador.
const BADGES = __BADGES__, PLAN = __PLAN__, TOWER = __TOWER__, PORD = __PORD__, SORD = __SORD__, PDESC = __PDESC__, SDESC = __SDESC__;
// Esferes: les millores 1-5 d'un personatge tenen 6 nivells, les 6-10 un (wiki, «Perks / Level»); la botiga ($os), 10 nivells de +4.000
const PCOST = [200, 400, 600, 800, 1000, 2000];
const perkMax = n => n <= 5 ? 6 : 1, perkCost = (n, l) => n <= 5 ? PCOST[l] : 1000;
const shopMax = () => 10, shopCost = (n, l) => 4000 * (l + 1);
const KEYS_OFTEN = 10;  // shortcut: llindar triat per Claude (una clau cada cop que et surt un personatge teu); canviar-lo si Victor en vol un altre
const num = t => Number(String(t).replace(/[^\d]/g, ''));
const fmtK = n => n.toLocaleString(document.documentElement.lang);
// Què diu el text enganxat: kakera o esferes que tens, nivells d'insígnies, cost del pròxim pis, nivells de les esferes
// Nivells del $opp o del $os: Mudae posa [MAX] en lloc de [LVL n] quan ja és al màxim
const LVL = /\[(?:LVL\s*(\d+)|MAX)\]/;
const levelsOf = (text, max) => [...text.matchAll(new RegExp(LVL, 'g'))].slice(0, 10).map((m, i) => m[1] != null ? +m[1] : max(i + 1));
function readMudae(text, tv) {
  const out = {}, have = text.match(/(?:Tienes|You have|Tens|Vous avez|Hai|Du hast|Você tem)\s+(\d[\d.,]*)/i);
  if (/\$limroul/i.test(text) && /\$kl\b/i.test(text)) {  // $bonus: «· Rolls por hora: +8 (6 $k + 1 $kl + 1 $kt)», una línia per bonus
    out.bonus = text.split(/\r?\n/).map(l => l.match(/^\s*(?:\d+\.\s*)?·\s*([^:]+?)\s*:\s*(.+?)\s*$/)).filter(Boolean).map(m => [m[1], m[2]]);
  } else if (/\$shoprefund/i.test(text) && LVL.test(text)) {  // $os: la botiga d'esferes
    out.shop = levelsOf(text, shopMax);
    if (have) out.sp = num(have[1]);
  } else if (LVL.test(text)) {
    const rows = text.split(/\r?\n/).map(l => l.trim());
    const first = rows.findIndex(l => LVL.test(l));
    out.opp = { name: first > 0 ? rows[first - 1] : '', levels: levelsOf(text, perkMax) };
    if (have) out.sp = num(have[1]);
  } else if (/\$build/i.test(text)) {
    const cost = text.match(/(?:cuesta|costs?|coûte|costa|kostet)\s+(\d[\d.,]*)/i);
    if (cost) out.ktCost = num(cost[1]);
    if (cost && tv) out.nf = Math.max(0, Math.round(num(cost[1]) / tv) - 1);
    if (have) out.k = num(have[1]);
  } else if (/\$bronze/i.test(text) && !/rewards/i.test(text)) {
    out.b = {};
    // «Sapphire II · 15.000»: el nivell que tens i el preu del següent. Cada servidor pot tenir els seus preus ($badgevalue).
    const seen = {};
    for (const [name] of BADGES) {
      const m = text.match(new RegExp('\\b' + name + '\\s+(IV|III|II|I)\\b'));
      out.b[name] = m ? ['I', 'II', 'III', 'IV'].indexOf(m[1]) + 1 : 0;
      const p = text.match(new RegExp('\\b' + name + '(?:\\s+(?:IV|III|II|I))?\\s*·\\s*(\\d[\\d.,]*)'));
      if (p) seen[name] = num(p[1]) / (out.b[name] + 1);
    }
    out.bp = {};  // preu per nivell sense el −25 % del Ruby IV (el $k ja el porta descomptat)
    for (const n in seen) out.bp[n] = Math.round(seen[n] / (out.b.Ruby >= 4 ? 0.75 : 1));
    if (have) out.k = num(have[1]);
  } else if (/^.+?·.*\(\d+\)\s*$/m.test(text) && !/Capital:/i.test(text)) {  // $mmy=: «Lanie Parish · 🔑 (1)», els que tenen claus
    out.keys = text.split(/\r?\n/).map(l => l.match(/^\s*(.+?)\s*·.*\((\d+)\)\s*$/)).filter(Boolean)
      .map(m => [m[1], +m[2]]).sort((a, b) => b[1] - a[1]);
  } else if (/Capital:/i.test(text)) {  // $tu: «Capital: 4.532:kakera:» i, a la part de les esferes, «Capital: 1.725 :sp:»
    [...text.matchAll(/Capital:\s*(\d[\d.,]*)\s*(:sp:)?/gi)].forEach((m, i) => { if (m[2] || i > 0) out.sp = num(m[1]); else out.k = num(m[1]); });
  }
  return out;
}
// Del $bonus, el que fa servir el tauler: poder màxim i el que gasta cada botó de kakera, esferes del $oh i hores del $rt
// shortcut: etiquetes en castellà (les de Victor) i en anglès suposat; mirar-ho amb un $bonus en anglès
function bonusFacts(B) {
  const get = (label, value) => { const r = (B || []).find(([l]) => label.test(l)), m = r && r[1].match(value); return m ? +m[1] : null; };
  return { pmax: get(/m[aá]x.*kakera|kakera.*m[aá]x/i, /(\d+)\s*%/), pcost: get(/(?:cost|coût).*kakera/i, /(\d+)\s*%/),
    oh: get(/\$oh\b/i, /\+\s*(\d+)/), rt: get(/\$rt\b/i, /(\d+)\s*h/), wl: get(/(?:espacio|slot).*(?:lista de deseos|wishlist)|wishlist.*slot/i, /\+\s*(\d+)/) };
}
// Cada torre té 12 pisos, un de cada; la següent s'obre en acabar-la, amb els mateixos pisos però amb menys efecte (wiki de Mudae)
function towerNext(M) {
  const tw = M.t || 0, built = tw * 12 + M.f.length;
  return { tw, built, cost: M.tv ? M.tv * (built + 1) * (tw >= 75 ? 10 : 1) : null, rec: TOWER.find(n => !M.f.includes(n)) };  // de la torre 76 en amunt, x10
}
// El $kt diu quants pisos tens en total: d'aquí surt la torre on ets. Si canvies de torre, els pisos marcats eren de l'anterior.
function towerFromKt(M, nf) { const t = Math.floor(nf / 12); return { t, f: t === (M.t || 0) ? M.f : [] }; }
// Esferes posades a un personatge: les millores 1-5 costen 200, 400, 600, 800, 1.000 i 2.000 per nivell; les 6-10, 1.000
const spent = lv => lv.reduce((s, l, i) => s + (i < 5 ? PCOST.slice(0, l).reduce((a, b) => a + b, 0) : l ? 1000 : 0), 0);
// $or <personatge> et torna el 90 % de tot el que hi has posat ($spheres). Si té esferes a les 1-9 (que només funcionen quan et surt),
// val la pena si, tornant-hi la 10 (si la tenia), et queden almenys 1.000 per a la 10 d'un altre personatge. null si no.
function refundPlan(lv) {
  const total = spent(lv), back = Math.floor(total * 0.9), free = back - (lv[9] >= 1 ? 1000 : 0);
  return lv.slice(0, 9).some(v => v > 0) && free >= 1000 ? { total, back, free } : null;
}
// A la botiga cada nivell costa més que l'anterior: el nivell més barat d'entre les que et serveixen (si costen igual, per ordre), fins a `count`
function cheapestBuys(levels, order, max, cost, count = 3) {
  const lv = Array.from({ length: 10 }, (_, i) => levels[i] || 0), out = [];
  while (out.length < count) {
    const open = order.filter(n => lv[n - 1] < max(n));
    if (!open.length) break;
    const n = open.reduce((a, b) => cost(b, lv[b - 1]) < cost(a, lv[a - 1]) ? b : a);
    out.push({ n, l: lv[n - 1] + 1, cost: cost(n, lv[n - 1]) });
    lv[n - 1]++;
  }
  return out;
}
// Les compres següents, per ordre de recomanació: la primera millora que no és al màxim, i així fins a `count`
function nextBuys(levels, order, max, cost, count = 3) {
  const lv = Array.from({ length: 10 }, (_, i) => levels[i] || 0), out = [];
  while (out.length < count) {
    const n = order.find(n => lv[n - 1] < max(n));
    if (!n) break;
    out.push({ n, l: lv[n - 1] + 1, cost: cost(n, lv[n - 1]) });
    lv[n - 1]++;
  }
  return out;
}
// El preu base de la torre ($towervalue) és del servidor: l'escrius tu o surt del $kt si has marcat els pisos que tens (preu = base × (pisos + 1))
function towerGuess(M) {
  const { tw, built } = towerNext(M), v = M.ktCost && built ? Math.round(M.ktCost / (built + 1) / (tw >= 75 ? 10 : 1)) : 0;
  return v >= 5000 && v <= 50000 ? v : null;
}
// Llocs de la wishlist: 7 de base (wiki de Mudae) més els del $bonus; sense $bonus, no ho sabem
const wlSlots = () => { const v = bonusFacts(M.bonus).wl; return v == null ? null : 7 + v; };
function millLoad() { try { return Object.assign({ b: {}, f: [], tv: null }, JSON.parse(localStorage.getItem('mudae-millores') || '{}')); } catch (e) { return { b: {}, f: [], tv: null }; } }
let M = millLoad();
const millSave = () => { try { localStorage.setItem('mudae-millores', JSON.stringify(M)); } catch (e) {} };
const $id = id => document.getElementById(id);
function millDraw() {
  const k = M.k ?? null;
  $id('millSum').textContent = [k !== null ? T('Kakera: {n}', { n: fmtK(k) }) : '', M.sp != null ? T('Esferes: {n}', { n: fmtK(M.sp) }) : '']
    .filter(Boolean).join(' · ') || T('Encara no has enganxat res');
  // el que encara no ha enganxat: cada servidor té els seus preus i cada jugador els seus nivells
  const miss = [!M.bp && '$k', M.nf == null && '$kt', !M.opp && '$opp <personatge>', !M.shop && '$os', !M.keys && '$mmy=', !M.bonus && '$bonus'].filter(Boolean);
  $id('millMiss').textContent = miss.length ? T('Per començar, enganxa: {list}', { list: miss.join(' · ') }) : '';
  // $bonus: el que en treus i, plegada, la llista sencera (el valor en negreta i d'on surt, a part)
  const BF = bonusFacts(M.bonus), li = s => '<li>' + esc(s) + '</li>', bl = [];
  if (BF.pmax && BF.pcost) bl.push(li(T('Amb el poder ple ({p} %) pots reaccionar a {n} botons de kakera: cadascun en gasta {c} %.', { p: BF.pmax, n: Math.floor(BF.pmax / BF.pcost), c: BF.pcost })));
  if (BF.oh != null) {
    const n = Math.round(BF.oh / 20), seen = Object.values(M.opps || {}).filter(lv => lv[9] >= 1).length;
    bl.push(li(T('El primer $oh del dia et dona +{e} esferes: tens {n} personatges amb la millora 10 (+20 cadascun).', { e: BF.oh, n })
      + (seen < n ? ' ' + T("Al tauler n'hi ha {m}: enganxa el $opp dels que falten.", { m: seen }) : '')));
  }
  if (BF.rt != null) bl.push(li(T('El $rt et torna cada {n} h: el tauler ja ho fa servir a «Què toca ara».', { n: BF.rt })));
  if (BF.wl != null) bl.push(li(T('Tens {n} llocs a la wishlist (7 de base i {b} de bonus).', { n: 7 + BF.wl, b: BF.wl })));
  window.RT_H = BF.rt; if (window.tick) tick();  // el $rt de «Què toca ara» torna quan diu el $bonus
  $id('bonusNext').innerHTML = !M.bonus ? li(T('Enganxa el $bonus per veure-hi el resum.')) : bl.join('');
  $id('bonusFold').hidden = !M.bonus;
  $id('bonusAll').innerHTML = (M.bonus || []).map(([l, v]) => { const i = v.indexOf(' (');
    return `<tr><td>${esc(l)}</td><td><b>${esc(i > 0 ? v.slice(0, i) : v)}</b>${i > 0 ? ' <span class="muted" style="font-size:13px">' + esc(v.slice(i + 1)) + '</span>' : ''}</td></tr>`; }).join('');
  // insígnies: el nivell de cada una i els 3 passos següents del pla, amb el preu (−25 % amb el Ruby IV)
  document.querySelectorAll('select[data-badge]').forEach(s => { s.value = String(M.b[s.dataset.badge] || 0); });
  // els preus, només del $k: cada servidor té els seus ($badgevalue)
  const lvl = Object.fromEntries(BADGES.map(([n]) => [n, M.b[n] || 0])), price = n => (M.bp && M.bp[n]) || null;
  const next = [];
  let sum = 0;
  for (const [n, l] of PLAN) {
    if (lvl[n] >= l || l !== lvl[n] + 1 || next.length >= 3) continue;
    const cost = price(n) && Math.round(price(n) * l * (lvl.Ruby >= 4 ? 0.75 : 1));
    if (cost) sum += cost;
    lvl[n] = l;
    next.push(`<li><button class="cmd" data-copy="$${n.toLowerCase()}">$${n.toLowerCase()}</button> <b>${esc(n)} ${['I', 'II', 'III', 'IV'][l - 1]}</b>` + (cost ? ` · ${fmtK(cost)}` : '')
      + (cost && k !== null ? ` <span class="tag ${k >= sum ? 'good' : ''}">${esc(k >= sum ? T('La pots pagar') : T('Et falten {n}', { n: fmtK(sum - k) }))}</span>` : '') + '</li>');
  }
  if (next.length && !M.bp) next.push('<li class="muted">' + esc(T('Enganxa el $k per saber què costen al teu servidor.')) + '</li>');
  $id('badgeNext').innerHTML = next.join('') || '<li>' + esc(T('Ja les tens totes al màxim.')) + '</li>';
  // Totes al màxim: ja no hi ha res a fer. Insígnies baixa plegada (recorda si l'obres) i les esferes pugen al costat de la torre;
  // el consell de la kakera passa a la torre, que és on va ara. Si no, sempre oberta i sense recordar res (sense data-fold).
  const allMax = !next.length, bx = $id('badgeBox');
  $id('badgeMax').hidden = !allMax; $id('kakeraNext').hidden = !allMax;
  $id('kakeraNext').textContent = T("Ara, la kakera: primer acaba la torre 1 (cada pis és un bonus nou) i després els kakeraloots ($kl), que donen tirades, llocs de wishlist i menys espera del $rt i no s'acaben mai. Si n'obres molts, puja el $quantity: cada nivell, +1 % de doblar el que surt.");
  if (allMax) {
    $id('millG2').append($id('perkBox')); $id('v-mill').append(bx);
    if (bx.dataset.fold !== 'badges-max') { bx.dataset.fold = 'badges-max'; bx.open = foldOpen('badges-max'); }
  } else { $id('millG2').prepend(bx); $id('shopBox').before($id('perkBox')); delete bx.dataset.fold; bx.open = true; }
  // torre: els pisos marcats són els de la torre on ets (M.t torres acabades), el que en diu el $kt i el pis recomanat
  document.querySelectorAll('input[data-floor]').forEach(i => { i.checked = M.f.includes(+i.dataset.floor); });
  $id('towerValue').value = M.tv ?? '';
  const { tw, built, cost, rec } = towerNext(M);
  let html = cost == null ? T('Torre {t} · pisos {n}/12 · per saber què costa el següent, marca els pisos que tens i enganxa el $kt (o escriu el preu base)', { t: tw + 1, n: M.f.length })
    : T('Torre {t} · pisos {n}/12 · el següent costa {c}', { t: tw + 1, n: M.f.length, c: fmtK(cost) });
  if (cost != null && k !== null) html += ' · ' + (k >= cost ? T('el pots pagar') : T('et falten {n}', { n: fmtK(cost - k) }));
  html = esc(html);
  if (M.nf != null && M.nf !== built) html += '<br><span class="tag bad">' + esc(T('Segons el $kt tens {n} pisos: marca quins són a la llista', { n: M.nf })) + '</span>';
  if (rec) html += `<br>${esc(T('Recomanat:'))} <button class="cmd" data-copy="$build ${rec}">$build ${rec}</button>`;
  $id('towerNext').innerHTML = html;
  // esferes: els nivells del personatge que has enganxat
  $id('perkWho').textContent = M.opp ? T('Nivells de {name}', { name: M.opp.name || '?' }) : T('Enganxa un $opp <personatge> per veure-hi els nivells');
  document.querySelectorAll('[data-perk]').forEach(el => {
    const v = M.opp ? M.opp.levels[+el.dataset.perk - 1] : undefined;
    el.textContent = v === undefined ? '' : 'LVL ' + v + (+el.dataset.perk <= 5 ? '/6' : '/1');
    el.className = 'tag' + (v === undefined ? '' : (+el.dataset.perk <= 5 ? v >= 6 : v >= 1) ? ' good' : '');
  });
  // les compres d'esferes recomanades, amb el que costen sumades i si les pots pagar
  const pay = (sum, have = M.sp) => have != null ? ` <span class="tag ${have >= sum ? 'good' : ''}">${esc(have >= sum ? T('La pots pagar') : T('Et falten {n}', { n: fmtK(sum - have) }))}</span>` : '';
  const buyList = (buys, cmd, desc) => { let sum = 0; const said = new Set(); return buys.map(b => { sum += b.cost;
    const d = said.has(b.n) ? '' : desc[b.n]; said.add(b.n);  // la descripció, només el primer cop que surt cada millora
    return `<li><button class="cmd" data-copy="${esc(cmd)}">${esc(cmd)}</button> <b>${esc(T('Millora {n}', { n: b.n }))} · LVL ${b.l}</b> · ${fmtK(b.cost)}` + pay(sum)
      + (d ? `<br><span class="muted" style="font-size:13px">${esc(d)}</span>` : '') + '</li>'; }).join(''); };
  // Què val més la pena (wiki de Mudae, «Ourospheres»): la 10 dona +20 esferes cada dia a cada personatge que la té, encara que no et surti,
  // fins a 100 personatges (després, 8), i es paga en 50 dies. Les 1-9 només funcionen quan et surt el personatge: només als que et surten sovint.
  const keyOf = Object.fromEntries(M.keys || []), kOpp = M.opp ? keyOf[M.opp.name] || 0 : 0, often = kOpp >= KEYS_OFTEN;
  const has10 = Object.entries(M.opps || {}).filter(([, lv]) => lv[9] >= 1).map(([n]) => n);
  const n10 = Math.max(has10.length, BF.oh != null ? Math.round(BF.oh / 20) : 0);  // el $bonus els compta tots, també els que no has enganxat
  // on posar la 10: personatges de les teves col·leccions (no els regalaràs), primer els que tenen més claus
  const mineColl = typeof COLL === 'undefined' ? [] : Object.values(COLL).flat().filter(c => c.st === 'Meu').map(c => c.n);
  // fins on la 10: a tots els personatges de les teves col·leccions, i com a molt a 100 (després en dona 8). Llavors, para i passa a la botiga.
  const keepers = new Set(mineColl).size, goal10 = Math.min(100, keepers || 100), done10 = n10 >= goal10;
  const cands = [...new Set(mineColl)].filter(n => !has10.includes(n) && n !== (M.opp && M.opp.name)).sort((a, b) => (keyOf[b] || 0) - (keyOf[a] || 0)).slice(0, 3);
  // la millora 10 a uns quants personatges, amb el que costen sumades i si les pots pagar amb `budget` esferes
  const put10 = (names, budget) => names.map((n, i) => `<li><button class="cmd" data-copy="${esc('$opp ' + n)}">${esc('$opp ' + n)}</button> <b>${esc(T('Millora {n}', { n: 10 }))}</b> · ${fmtK(1000)}`
    + pay(1000 * (i + 1), budget) + (i ? '' : `<br><span class="muted" style="font-size:13px">${esc(PDESC[10])}</span>`) + '</li>').join('');
  let perkHtml;
  const R = M.opp && !often && !done10 ? refundPlan(M.opp.levels) : null;
  if (!M.opp) perkHtml = li(T('Enganxa un $opp <personatge> per veure-hi els nivells'));
  else if (R) {  // té esferes a les 1-9 i no et surt sovint: millor tornar-les ($or) i posar-les a la 10
    const name = M.opp.name, budget = R.back + (M.sp || 0);
    perkHtml = `<li><button class="cmd" data-copy="${esc('$or ' + name)}">${esc('$or ' + name)}</button> <b>${esc(T('Reemborsa les esferes'))}</b> · `
      + esc(T("et torna {b} de les {t} que hi has posat (perds el 10 %), també les de la millora 10", { b: fmtK(R.back), t: fmtK(R.total) })) + '</li>'
      + put10([name, ...cands].slice(0, Math.max(1, Math.min(4, Math.floor(budget / 1000)))), budget);
  } else {
    perkHtml = buyList(nextBuys(M.opp.levels, often ? PORD : [10], perkMax, perkCost), ('$opp ' + M.opp.name).trim(), PDESC);
    if (!perkHtml && !often)  // aquest ja té la 10: la 10, a uns altres, o, si ja la tenen tots, la botiga
      perkHtml = done10 ? li(T('Para de gastar en personatges que no et surten sovint: ara, la botiga d\'esferes ($os).'))
        : put10(cands) || li(T('Posa la millora 10 a un altre personatge teu que no regalaràs: $opp <personatge>.'));
    perkHtml ||= li(T('Aquest personatge ja el tens millorat del tot.'));
  }
  if (M.opp) {  // per què
    perkHtml += '<li class="muted">' + esc(often
      ? T('Et surt sovint ({n} claus): les millores 1-5 hi valen la pena.', { n: kOpp })
      : done10 ? T('Ja tens la millora 10 a {n} personatges: tots els de les teves col·leccions, o 100. A partir d\'aquí, la botiga ($os) i les millores 1-9 als que et surten sovint.', { n: n10 })
      : T('Primer, la millora 10 a tots els teus personatges que no regalaràs (en tens {n} de {t}): +20 esferes cada dia cadascun, encara que no et surtin. Les altres només funcionen quan et surt el personatge: posa-les als que tinguin 10 claus o més.', { n: n10, t: goal10 })) + '</li>';
  }
  const top5 = (M.keys || []).slice(0, 5);
  $id('perkKeys').innerHTML = !M.keys ? esc(T('Enganxa el $mmy= per saber quins et surten més.'))
    : top5.map(([name, n]) => `<button class="cmd" data-copy="${esc('$opp ' + name)}">${esc(name)}</button> <span class="muted" style="font-size:13px">${esc(n === 1 ? T('1 clau') : T('{n} claus', { n }))}</span>`).join('<br>')
      + (top5.length && top5[0][1] < KEYS_OFTEN ? '<br><span class="muted" style="font-size:13px">' + esc(T('Encara cap no et surt prou: de moment, la millora 10.')) + '</span>' : '');
  $id('perkNext').innerHTML = perkHtml;
  // els personatges que has enganxat amb el $opp (Mudae no té cap llista dels millorats)
  const ops = Object.entries(M.opps || {}).filter(([, lv]) => lv.some(v => v > 0));
  const p10 = ops.filter(([, lv]) => lv[9] >= 1).length;
  $id('perkList').innerHTML = !ops.length ? esc(T('Encara no has enganxat cap personatge millorat.'))
    : esc(T('Amb la millora 10: {n} (+{e} esferes al primer $oh del dia)', { n: p10, e: p10 * 20 })) + '<br>'
      + ops.map(([name, lv]) => `<button class="cmd" data-copy="${esc('$opp ' + name)}">${esc(name)}</button> <span class="muted" style="font-size:13px">${lv.map((v, i) => v ? (i + 1) + ':' + v : '').filter(Boolean).join(' ')}</span>`).join('<br>');
  document.querySelectorAll('[data-shop]').forEach(el => {
    const v = M.shop ? M.shop[+el.dataset.shop - 1] : undefined;
    el.textContent = v === undefined ? '' : 'LVL ' + v + '/10';
    el.className = 'tag' + (v >= 10 ? ' good' : '');
  });
  // La botiga, quan ja tens la 10 a tots (abans, la 10 es paga més de pressa). I només el que et serveix: la 2 (megaesferes) va sola;
  // la 1-9 reforça aquella millora dels personatges, així que només si algun personatge teu la té; la 10, si en tens algun de millorat del tot.
  const opsLv = Object.values(M.opps || {});
  const useful = n => n === 2 || (n === 10 ? opsLv.some(lv => lv.every((v, i) => v >= perkMax(i + 1))) : opsLv.some(lv => lv[n - 1] > 0));
  $id('shopNext').innerHTML = (done10 ? li(T('Ara sí: la botiga. Et recomano només les que reforcen millores que ja tens als teus personatges.'))
      : '<li class="muted">' + esc(T('Encara no: abans, la millora 10 als teus personatges (en tens {n} de {t}). Cada 1.000 esferes et donen +20 cada dia, i gairebé tota la botiga només reforça les millores 1-9, que funcionen quan et surt el personatge.', { n: n10, t: goal10 })) + '</li>')
    + (!M.shop ? li(T('Enganxa el $os per veure-hi els nivells'))
    : buyList(cheapestBuys(M.shop, SORD.filter(useful), shopMax, shopCost), '$os', SDESC) || li(T('Ja tens al màxim les que et serveixen. Les altres reforcen millores que cap personatge teu no té.')));
  if (window.pendDraw) pendDraw();
}
const millText = $id('millText'), millMsg = $id('millMsg');  // en enganxar-hi, ho llegeix pasteAll (la casella d'Avui)
// Desa el que diu el text ($k, $kt, $opp, $os, $mmy= o $tu); false si no és res d'això. També el fa servir la casella d'Avui.
function millRead(text) {
  const got = readMudae(text, M.tv);
  if (!Object.keys(got).length) return false;
  if (got.nf != null) Object.assign(got, towerFromKt(M, got.nf));
  if (got.opp && got.opp.name) got.opps = Object.assign({}, M.opps, { [got.opp.name]: got.opp.levels });
  Object.assign(M, got);
  if (got.ktCost && !M.tv && (M.tv = towerGuess(M))) M.nf = towerNext(M).built;
  millSave(); millDraw();
  return true;
}
document.addEventListener('change', ev => {
  const el = ev.target;
  if (el.dataset && el.dataset.badge) M.b[el.dataset.badge] = +el.value;
  else if (el.dataset && el.dataset.floor) {
    M.f = [...document.querySelectorAll('input[data-floor]:checked')].map(i => +i.dataset.floor);
    if (M.f.length === 12) { M.t = (M.t || 0) + 1; M.f = []; }  // torre acabada: la següent comença buida
    if (!M.tv && (M.tv = towerGuess(M))) M.nf = towerNext(M).built;
  }
  else if (el.id === 'towerValue') M.tv = num(el.value) ? Math.min(50000, Math.max(5000, num(el.value))) : null;
  else return;
  millSave(); millDraw();
});
millDraw();
// el $tu enganxat a «Què toca ara» també posa al dia la kakera i les esferes d'aquí
window.millTu = text => { const got = readMudae(text, M.tv); if (got.k == null && got.sp == null) return; Object.assign(M, got); millSave(); millDraw(); };
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


JS_ORD = r"""
// «Ordres d'ús habitual»: el cercador filtra per ordre, pel que fa o pel nom del grup, i obre els grups on n'hi ha
const ordQ = document.getElementById('ordQ'), ordGroups = [...document.querySelectorAll('#ordres details.og')];
ordQ.addEventListener('input', () => {
  const k = ordQ.value.trim().toLowerCase();
  let any = false;
  ordGroups.forEach(g => {
    const all = g.querySelector('h3').textContent.toLowerCase().includes(k);
    let n = 0;
    g.querySelectorAll('.oline').forEach(l => { l.hidden = !!k && !all && !l.textContent.toLowerCase().includes(k); n += !l.hidden; });
    g.hidden = !n; g.open = k ? !!n : foldOpen(g.dataset.fold, g.hasAttribute('data-open')); any = any || !!n;
  });
  document.getElementById('ordNone').hidden = any;
});
"""


JS_PD = r"""
// Pokédex: només el recompte del $pd sumrare. La llista sencera, Mudae només la dona de 10 en 10 (més de 100 pàgines): no val la pena.
const RARE = [T('Comú'), T('Poc comú'), T('Rar'), T('Molt rar'), T('Llegendari'), T('Ultraentitat o paradoxa')];
const pdNum = t => +String(t).replace(/\D/g, '');
// «5 (1/90)» per raresa i «100 / 1.025» al final: {rare: {raresa: [tens, total]}, have, total}, o null si no n'és un
function pdSumrare(text) {
  const ls = [...text.matchAll(/^\s*([1-6])[^\d(\n]{0,8}\(\s*(\d[\d.,]*)\s*\/\s*(\d[\d.,]*)\s*\)\s*$/gm)];  // «5🌟 (1/90)»: l'emoji pot ser-hi o no
  const tot = text.match(/^\s*(\d[\d.,]*)\s*\/\s*(\d[\d.,]*)\s*$/m);
  if (ls.length < 2 || !tot || pdNum(tot[2]) < 1000) return null;  // el total és el de la Pokédex, no el d'una sèrie
  const rare = {}; ls.forEach(m => { rare[m[1]] = [pdNum(m[2]), pdNum(m[3])]; });  // només les que diu Mudae (no posa les que no en tens cap)
  return { rare, have: pdNum(tot[1]), total: pdNum(tot[2]) };
}
let PDS = null; try { PDS = JSON.parse(localStorage.getItem('mudae-pd-sumrare')); } catch (e) {}
const bar = (a, b) => `<div class="track"><div class="seg mine" style="width:${b ? (100 * a / b).toFixed(2) : 0}%"></div></div>`;
function pdDraw() {
  document.getElementById('pdHead').textContent = PDS ? T('{a} de {b}', { a: fmtK(PDS.have), b: fmtK(PDS.total) }) : '$pd sumrare';
  document.getElementById('pdSum').innerHTML = !PDS ? '<p class="muted" style="margin:0;font-size:14px">' + esc(T('Encara no has enganxat el $pd sumrare.')) + '</p>'
    : Object.keys(PDS.rare).sort().map(r => { const [a, b] = PDS.rare[r];
        return `<div class="pdrow"><span>${esc(RARE[r - 1] || r)}</span>${bar(a, b)}<span class="val">${fmtK(a)}/${fmtK(b)} · <b>${b ? Math.round(100 * a / b) : 0}%</b></span></div>`; }).join('');
}
function pdRead(text) {  // true si era un $pd sumrare
  const r = pdSumrare(text); if (!r) return false;
  PDS = r; try { localStorage.setItem('mudae-pd-sumrare', JSON.stringify(PDS)); } catch (e) {}
  pdDraw(); return true;
}
const pdText = document.getElementById('pdText'), pdMsg = document.getElementById('pdMsg');
pdText.addEventListener('paste', () => setTimeout(() => {
  if (!pdRead(pdText.value)) { pdMsg.innerHTML = '<span class="tag bad">' + esc(T('No ho entenc')) + '</span> ' + esc(T('Enganxa el missatge sencer del $pd sumrare.')); return; }
  pdText.value = ''; pdMsg.innerHTML = '<span class="tag good">' + esc(T('Pokédex al dia')) + '</span>';
}));
pdDraw();
"""


JS_PASTE = r"""
// La casella d'Avui: hi enganxes qualsevol missatge de Mudae i el tauler l'envia on toca
const tuText = document.getElementById('tuText'), tuMsg = document.getElementById('tuMsg');
const goTo = (v, name) => `<button type="button" class="go" data-goto="${v}">${esc(name)} →</button>`;
function pasteAll(text, area = tuText, msg = tuMsg) {  // area i msg: la casella on has enganxat i on surt el missatge
  if (!text.trim()) return;
  const tu = readTu(text), prev = txt.value;
  if (('claim' in tu || 'rolls' in tu) && Object.keys(tu).length > 2) {  // el $tu: marca les caselles (i la kakera i les esferes de Millores)
    window.millTu(text);
    const P = readPower(text);
    if (P) try { localStorage.setItem('mudae-poder', JSON.stringify({ ...P, t: Date.now() })); } catch (e) {}
    RS = learnResets(RS, readResets(text), Date.now());  // el rellotge, amb els reinicis del teu servidor
    try { localStorage.setItem('mudae-reinici', JSON.stringify(RS)); } catch (e) {}
    powerDraw(); resetDraw(); applyTu(tu, Date.now()); pendDraw(); area.value = '';
    msg.innerHTML = '<span class="tag good">' + esc(T('Marcat segons el $tu')) + '</span>';
    return;
  }
  if ('p' in tu && Object.keys(tu).length === 2) {  // «Tiempo restante antes de su próximo $p: 42 min.»
    applyTu({ p: tu.p }, Date.now()); area.value = '';
    msg.innerHTML = '<span class="tag good">' + esc(T('Pokéslot marcat')) + '</span>';
    return;
  }
  if (pdRead(text)) {  // abans que «Posar al dia»: «100 / 1.025» sembla el títol d'una sèrie
    area.value = '';
    msg.innerHTML = '<span class="tag good">' + esc(T('Pokédex al dia')) + '</span> ' + goTo('coll', T('Col·leccions'));
    return;
  }
  if (/\/\s*1[.,]?\d{3}\s*-\s*\S+\s+\d+\s*\/\s*\d+/.test(text)) {  // una pàgina del $pd («100 / 1.025 - Page 1 / 10»): no és una sèrie
    msg.innerHTML = '<span class="tag bad">' + esc(T('No ho entenc')) + '</span> ' + esc(T('Per a la Pokédex, enganxa el $pd sumrare (un sol missatge), no les pàgines del $pd.'));
    return;
  }
  const set = /\$set|setting|configuraci|ajuste/i.test(text) ? readSettings(text) : null;
  if (set) {  // $settings: cada quant hi ha reclam i a quin minut es reinicia tot
    RS.iv = set.iv; if (set.m != null) RS.m = set.m;
    try { localStorage.setItem('mudae-reinici', JSON.stringify(RS)); } catch (e) {}
    resetDraw(); tick(); pendDraw(); area.value = '';
    msg.innerHTML = '<span class="tag good">' + esc(T('Reclam cada {n} min', { n: set.iv })) + '</span>';
    return;
  }
  if (millRead(text)) {  // $k, $kt, $opp, $os, $mmy= o $bonus
    area.value = '';
    msg.innerHTML = '<span class="tag good">' + esc(T('Llegit a Millores')) + '</span> ' + (area === tuText ? goTo('mill', T('Millores')) : '');
    return;
  }
  txt.value = text;
  if (autoKind(true)) {  // el harem, la wishlist o sèries: van a «Posar al dia», que ho desa sol si és segur (una sèrie nova, no)
    area.value = ''; msg.innerHTML = '';
    go('eines'); document.getElementById('upd').scrollIntoView({ block: 'start' });
    if (!autoSave()) updMsg.innerHTML += ' ' + esc(serieSel.value === '__new__' && kindSel.value === 'serie'
      ? T('És una sèrie nova: comprova el nom i clica «Desar i actualitzar».') : T('Comprova la sèrie i clica «Desar i actualitzar».'));
    return;
  }
  txt.value = prev; syncUpd();
  msg.innerHTML = '<span class="tag bad">' + esc(T('No ho entenc')) + '</span> ' + esc(T('Enganxa el missatge sencer de Mudae.'));
}
tuText.addEventListener('paste', () => setTimeout(() => pasteAll(tuText.value)));
// la de Millores, igual: no cal cap botó
millText.addEventListener('paste', () => setTimeout(() => pasteAll(millText.value, millText, millMsg)));
// «Pendent!»: el que depèn del teu servidor o de tu i encara no has enganxat, amb l'ordre que ho diu. No suposem res.
const NEED = __NEED__;
function pendDraw() {
  const out = [], add = (cmd, why) => out.push(`<li><button class="cmd" data-copy="${esc(cmd)}">${esc(cmd.trim())}</button> <span class="muted">${esc(why)}</span></li>`);
  if (RS.m == null || !RS.next) add('$tu', T('quan tornen les tirades i el reclam, i el poder de kakera'));
  if (!RS.iv) add('$settings', T('cada quant hi ha reclam al servidor'));
  if (!M.bonus) add('$bonus', T('cada quant pots fer el $rt i quants llocs tens a la wishlist'));
  if (!M.bp) add('$k', T('les teves insígnies i quant costen al servidor'));
  if (!M.opps) add('$opz', T('quins personatges tenen esferes (enganxa cada pàgina)'));
  if (M.nf == null && !M.ktCost) add('$kt', T('la torre: quants pisos tens i quant costa el següent'));
  else if (M.nf == null) out.push(`<li>${goTo('mill', T('Millores'))} <span class="muted">${esc(T('la torre: marca els pisos que tens (o escriu el preu base de la torre)'))}</span></li>`);
  NEED.forEach(([cmd, why]) => add(cmd, why));
  document.getElementById('pendList').innerHTML = out.join('');
  document.getElementById('pend').hidden = !out.length;
}
pendDraw();
"""


JS_WL = r"""
// Muntador de wishlist: què treure (ja és teu o d'algú altre) i què posar-hi (lliures de prioritat alta, per ordre)
// els llocs de la wishlist depenen de tu (insígnies, torre, kakeraloots): surten del $bonus, no els suposem
const WL = __WL__, MAX = wlSlots();
const wlBox = document.getElementById('wlBuild');
const bad = WL.wl.filter(w => !w.ok), keep = WL.wl.filter(w => w.ok);
const free = MAX == null ? 0 : MAX - keep.length;
if (MAX != null) document.getElementById('wlMax').textContent = '/' + MAX;
if (MAX != null && MAX > WL.wl.length) {  // avís a Avui: llocs lliures
  const n = MAX - WL.wl.length, box = document.getElementById('alertsBox'), ul = document.getElementById('alertsList');
  ul.insertAdjacentHTML('beforeend', `<li><span class="tag accent">${esc(T('Wishlist'))}</span><span>${esc(T('Tens {n} llocs lliures. Mira «Per omplir» a la wishlist.', { n }))}</span><button type="button" class="go" data-goto="wl">${esc(T('Wishlist'))} →</button></li>`);
  box.hidden = false; document.getElementById('alertsN').textContent = ul.children.length;
}
// la col·lecció que vols completar (o totes): te'n recomana els primers, ja ordenats pel servidor
const SER = [...new Set(WL.cand.map(c => c.s))].sort((a, b) => a.localeCompare(b));
let only = ''; try { only = localStorage.getItem('mudae-wlser') || ''; } catch (e) {}
if (!SER.includes(only)) only = '';
const cands = () => WL.cand.filter(c => !only || c.s === only);
let pick = new Set(cands().slice(0, Math.max(free, 0)).map(c => c.n));
const note = s => '<p class="muted" style="margin:0;font-size:14px">' + s + '</p>';
function drawWl() {
  let html = '';
  if (bad.length) {
    // $wishpurge treu d'una vegada els reclamats (teus o d'altres) i els repetits; els posats amb $wishl (bloquejats) es queden
    html += `<div class="give"><div class="give-row"><b>${esc(T('1. Treu-ne {n}', { n: bad.length }))}</b><button class="copybtn" data-copy="$wishpurge">${esc(T('Copiar ordre'))}</button></div><code>$wishpurge</code>`
      + note(esc(T("Treu els que ja estan reclamats: {names}. Si en vols guardar algun, torna'l a posar amb $wishl (bloquejat) i el $wishpurge no el tocarà.", { names: bad.map(w => w.n).join(', ') }))) + '</div>';
  }
  if (free > 0) {
    const chosen = WL.cand.filter(c => pick.has(c.n)).map(c => c.n);
    const cmd = '$wish ' + chosen.join(' $ ');
    html += `<div class="give"><div class="give-row"><b>${bad.length ? '2. ' : ''}${esc(T(free === 1 ? 'Omple {n} lloc' : 'Omple {n} llocs', { n: free }))}</b>`
      + `<span class="muted">${esc(T("tria'n {n}: {c} de {n}", { n: free, c: chosen.length }))}</span>`
      + (chosen.length ? `<button class="copybtn" data-copy="${esc(cmd)}">${esc(T('Copiar ordre'))}</button>` : '') + '</div>'
      + `<label class="muted" style="font-size:14px">${esc(T('Col·lecció que vols completar:'))} <select id="wlSer">`
      + `<option value="">${esc(T("Totes: la més a prop d'acabar"))}</option>`
      + SER.map(s => `<option value="${esc(s)}"${s === only ? ' selected' : ''}>${esc(s)}</option>`).join('') + '</select></label>'
      + (WL.cand.length ? note(esc(T("Ja te n'he triat els recomanats; clica un nom per canviar-lo."))) : note(esc(T('No queda cap personatge lliure a les teves col·leccions.'))))
      + `<div class="chips">${cands().slice(0, 40).map(c => `<label class="ch ${pick.has(c.n) ? 'wish' : 'free'}" data-tip="${esc(c.s)}">`
        + `<input type="checkbox" data-wlc="${esc(c.n)}" ${pick.has(c.n) ? 'checked' : ''} hidden>${esc(c.n)}</label>`).join('')}</div>`
      + (chosen.length ? `<code>${esc(cmd)}</code>` : '')
      + (chosen.length > free ? `<span class="tag bad">${esc(T("N'has triat {n} de més", { n: chosen.length - free }))}</span>` : '')
      + '</div>';
  }
  if (MAX == null) html += note(esc(T('Enganxa el $bonus a la casella d\'Avui per saber quants llocs tens a la wishlist i quins pots omplir.')));
  if (!WL.wl.some(w => w.star && w.ok))
    html += note(esc(T('No tens desig estrella: posa-hi el que vulguis més amb')) + ` <button class="cmd" data-copy="$sw ">$sw &lt;${esc(T('nom'))}&gt;</button>`);
  if (bad.length || MAX != null) html += note(esc(bad.length || free > 0 ? T('Quan ho hagis fet, enganxa el teu $wl a «Posar al dia».') : T('La wishlist està plena i tot està lliure: res a fer.')));
  wlBox.innerHTML = html;
}
wlBox.addEventListener('change', ev => {
  if (ev.target.id === 'wlSer') {  // una altra col·lecció: torna a triar els recomanats
    only = ev.target.value; try { localStorage.setItem('mudae-wlser', only); } catch (e) {}
    pick = new Set(cands().slice(0, free).map(c => c.n)); return drawWl();
  }
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
    return t(GRUPS[g])


def give_cmd(ts):
    me = f"<@{DISCORD_ID}>" if DISCORD_ID else f"@{ME}"
    return f"$give {me} " + " $ ".join(tr["name"] for tr in ts)


def wish_candidates(d):
    """Lliures de les teves col·leccions per posar a la wishlist: preferits, la col·lecció més a prop d'acabar i prioritat alta primer."""
    pct = {s["name"]: s["Meu"] / s["total"] for s in d["series"]}
    rank = {n: i for i, n in enumerate(PREFERITS)}
    wl, seen, out = {w["name"] for w in d["wishlist"]}, set(), []
    for r in sorted((r for r in d["rows"] if r["state"] == "Lliure" and r["name"] not in wl),
                    key=lambda r: (rank.get(r["name"], len(rank)), -pct[r["serie"]], r["prio"] != "Alta", r["serie"], r["name"])):
        if r["name"] not in seen:
            seen.add(r["name"])
            out.append({"n": r["name"], "s": r["serie"]})
    return out


def write_dashboard(d, hist):
    rows, series = d["rows"], d["series"]
    owned = sum(s["Meu"] for s in series)
    total = sum(s["total"] for s in series)
    pct = round(100 * owned / total) if total else 0
    wl_bad = [w for w in d["wishlist"] if not w["ok"]]
    pend = sum(tr["status"] in ("Pendent", "Enviat") for tr in d["trades"])

    # pestanyes, per ordre d'ús: el que fas cada hora primer, el que fas de tant en tant al final
    tabs = {"avui": t("Avui"), "coll": t("Col·leccions"), "wl": t("Wishlist"), "trades": t("Intercanvis"),
            "mill": t("Millores"), "eines": t("Posar al dia")}

    def alert(tag, cls, text, go):  # cada avís porta a la pestanya on s'arregla
        return (f'<li><span class="tag {cls}">{e(tag)}</span><span>{text}</span>'
                f'<button type="button" class="go" data-goto="{go}">{e(tabs[go])} →</button></li>')
    alerts = [alert(t("Wishlist"), "bad", f'<b>{e(w["name"])}</b>: {e(w["todo"])}', "wl") for w in wl_bad]
    for kind, r, prev in d["changes"]:
        if kind == "nou":
            continue
        if r["owner"] == ME:
            alerts.append(alert(t("Aconseguit"), "good", f'<b>{e(r["name"])}</b> · {e(r["serie"])}', "coll"))
        elif not r["owner"]:
            alerts.append(alert(t("Lliure"), "accent", e(t("{name} ha quedat lliure (era de {prev}): posa'l a la wishlist", name=r["name"], prev=prev)), "wl"))
        else:
            alerts.append(alert(t("Canvi"), "", e(t("{name} ara és de {owner} (abans: {prev})", name=r["name"], owner=r["owner"], prev=prev or t("lliure"))), "coll"))
    news = sum(k == "nou" for k, _, _ in d["changes"])
    if news:
        alerts.append(alert(t("Nous"), "", e(t("{n} personatges nous a les teves sèries", n=news)), "coll"))
    if d["missing"]:
        alerts.append(alert(t("Harem"), "", e(t("{n} personatges teus no surten al harem: torna a enganxar el $mma+s a «Posar al dia»", n=len(d["missing"]))), "eines"))

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
                          "cand": wish_candidates(d)}, ensure_ascii=False).replace("</", r"<\/")

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

    # «Què toca ara», una cua: a «Ara», el que pots fer, per grups i per ordre d'importància; a «Fets» (plegat), el que has fet, per l'hora
    # que torna. Cada tasca marcada hi queda fins que es pot tornar a fer (data-torna); el JavaScript la mou (data-home: el seu grup).
    groups, foot = {}, {"reclam": f'<p class="tip">{e(t("Abans, fes totes les tirades. Si en vols més:"))} {cmd_html("$us · $rolls")}</p>'}
    for g, que, ordre, com, torna in RUTINA:
        if g not in groups:
            title = group_title(g)  # a «Cada hora» i «Pots reclamar», el JavaScript hi posa a sota el minut de les tirades i fins quan pots reclamar
            groups[g] = [f'<div class="grp{" claim" if g == "reclam" else ""}" data-period="{g}"><h3>{e(title[:1] + title[1:].lower())}'
                         + ('<small></small>' if g in ("hora", "reclam") else "") + '</h3><div class="tl">']
        key = TU_KEY[ordre] if torna is not None else None  # el seu nom i la línia del $tu que la marca
        box = f'<input type="checkbox" data-id="{key}" data-torna="{torna}" data-tu="{key}">' if key else '<span class="bullet">·</span>'
        home = f' data-home="{g}"' if key else ""
        groups[g].append(f'<label class="task"{home}>{box}<span><b>{e(t(que))}</b> {cmd_html(ordre)}'
                         f'<span class="desc">{e(t(com))}</span><span class="when"></span></span></label>')
    card = lambda g: "".join(groups[g]) + "</div>" + foot.get(g, "") + "</div>"
    now_g = [g for g, *_, torna in RUTINA if torna is not None]
    checklist = ('<div class="plan"><div class="now-col">' + "".join(card(g) for g in dict.fromkeys(now_g))
                 + f'<p class="muted none" id="nowNone" hidden>{e(t("Res a fer ara: a «Fets» tens quan torna cada cosa."))}</p></div>'
                 + f'<div class="grp later-col"><h3>{e(t("Estat"))}<small id="laterInfo"></small></h3><div id="state"></div>'
                 + f'<details class="fold done-fold" data-fold="fets" hidden><summary><h3>{e(t("Fets"))}</h3><span class="muted" id="laterN"></span></summary><div class="fold-body" id="later"></div></details>'
                 + f'<p class="muted none" id="laterNone">{e(t("Quan facis una cosa, marca-la: s’amagarà fins que la puguis tornar a fer. El $tu enganxat les marca soles."))}</p></div></div>'
                 + f'<details class="fold" data-fold="altres"><summary><h3>{e(t("Altres moments"))}</h3></summary><div class="groups fold-body">'
                 + "".join(card(g) for g in groups if g not in now_g) + "</div></details>")

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
    ordres = "".join(  # un grup plegable per tema: tancats d'entrada, el cercador els obre
        f'<details class="og fold" data-fold="ord-{e(g)}"{" data-open" if i == 0 else ""}><summary><h3>{e(t(g))}</h3><span class="muted">{len(items)}</span></summary><div class="fold-body">'
        + "".join(f'<div class="oline">{cmd_html(c)}<span>{e(t(desc))}</span></div>' for c, desc in items) + "</div></details>"
        for i, (g, items) in enumerate(ORDRES))
    oc = f"oc.html?lang={LANG}"
    need = ([["$mma+s", t("el teu harem, amb la sèrie de cada personatge")]] if not (DADES / "harem.txt").exists() else [])         + ([["$wl", t("la teva wishlist")]] if not (DADES / "wishlist.txt").exists() else [])         + ([["$imaso ", t("una sèrie que col·leccionis (o, a «Posar al dia», «Totes les teves sèries»)")]] if not series else [])
    # botons per copiar les ordres que entén la casella d'Avui; $imaa+os, amb totes les teves sèries (com a «Posar al dia»)
    all_cmd = "$imaa+os " + " $ ".join(x["name"] for x in series if x["name"] not in spec)
    mill_cmds = " ".join(cmd_html(c) for c in ("$k", "$kt", "$opz", "$opp <personatge>", "$os", "$mmy=", "$bonus", "$tu"))  # el que llegeix Millores
    paste_cmds = (" ".join(cmd_html(c) for c in ("$tu", "$settings", "$p", "$k", "$kt", "$opz", "$opp <personatge>", "$os", "$mmy=", "$bonus", "$mma+s", "$wl", "$imaso <sèrie>"))
                  + (f' <button class="cmd" data-copy="{e(all_cmd)}" data-tip="{e(t("Totes les teves sèries"))}">$imaa+os</button>' if len(all_cmd) > 9 else "")
                  + " " + cmd_html("$pd sumrare"))
    roman = ["0", "I", "II", "III", "IV"]
    badge_rows = "".join(
        f'<tr><td><b>{n}</b><br><span class="muted" style="font-size:12.5px">{e(t(desc))}</span></td>'
        f'<td style="text-align:right"><select data-badge="{n}" aria-label="{n}">'
        + "".join(f'<option value="{i}">{r}</option>' for i, r in enumerate(roman)) + "</select></td></tr>"
        for n, _, desc in BADGES)
    tower_rows = "".join(
        f'<li><label><input type="checkbox" data-floor="{n}"> <span>[{n}] {e(t(desc))}</span></label></li>' for n, desc in TOWER)
    perk_rows = "".join(
        f'<div class="og"><h3>{e(t(g))}</h3>' + "".join(
            f'<div class="oline"><code>{n}</code><span>{e(t(PERKS[n]))}</span><span data-perk="{n}"></span></div>' for n in ns) + "</div>"
        for g, ns in PERK_GROUPS)
    shop_rows = "".join(
        f'<div class="oline"><code>{n}</code><span>{e(t(SHOP[n]))}</span><span data-shop="{n}"></span></div>' for n in SHOP_ORDER)
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
<meta name="color-scheme" content="dark"><meta name="theme-color" content="#0e0f15">
<title>{e(TITLE_TEXT)}</title>{favicon}
<style>{STYLE}</style></head><body><main>
<header class="hdr fu">
 <div class="hdr-l">
  {logo}
  <div><h1 class="logo-name">{e(TITLE_TEXT)}</h1><div class="logo-sub">{T_("Dades del {date}", date=f"{dt.datetime.now():%d/%m/%Y %H:%M}")}</div></div>
 </div>
 <div class="hdr-r">
  <span class="pill" id="toolMsg" role="status"></span>
  <button type="button" class="pill" id="notiBtn" aria-pressed="false" hidden><span class="dot"></span><span>{T_("Notificacions: no")}</span></button>
  <button type="button" class="pill" id="refaBtn">{T_("Actualitzar el tauler")}</button>
  <details class="more" id="more"><summary class="pill" aria-label="{T_("Més opcions")}" title="{T_("Més opcions")}">···</summary>
   <div class="menu">
    <button type="button" id="lnkBtn">{T_("Crear accés directe a l'escriptori")}</button>
    <button type="button" id="stopBtn"><span class="dot ok"></span>{T_("Tancar el servidor")}</button>
   </div></details>
 </div>
</header>
<div class="bar">
 <nav class="nav" role="tablist" aria-label="{T_("Seccions")}">{"".join(
     f'<button type="button" role="tab" data-go="{k}" aria-controls="v-{k}">{e(v)}'
     + (f'<span class="n">{len(wl_bad)}</span>' if k == "wl" and wl_bad else "") + '</button>' for k, v in tabs.items())}</nav>
 <p id="clock" class="clock"></p>
</div>
<div class="view" id="v-avui" role="tabpanel">
<section class="resum glass" aria-labelledby="hola">
 <h2 class="hola" id="hola">{T_("Hola, {name}", name=ME) if ME else T_("Hola")}</h2>
 <div class="fets">
  <button type="button" class="fet" data-goto="coll"><span class="fet-l">{T_("Col·lecció")}</span><span class="fet-v">{pct}%</span><span class="fet-s">{T_("{a} de {b}", a=owned, b=total)}</span></button>
  <button type="button" class="fet" data-goto="trades"><span class="fet-l">{T_("Harem")}</span><span class="fet-v">{len(d["harem"])}</span><span class="fet-s">{T_("{n} sobrants", n=len(d["sobrants"]))}</span></button>
  <button type="button" class="fet" data-goto="wl"><span class="fet-l">{T_("Wishlist")}</span><span class="fet-v">{len(d["wishlist"])}<span id="wlMax"></span></span><span class="fet-s">{T_("tot bé") if not wl_bad else T_("{n} per arreglar", n=len(wl_bad))}</span></button>
  <button type="button" class="fet" data-goto="trades"><span class="fet-l">{T_("Intercanvis")}</span><span class="fet-v">{len(d["trades"])}</span><span class="fet-s">{T_("{n} oberts", n=pend)}</span></button>
 </div>
 <div class="paste"><textarea id="tuText" rows="1" spellcheck="false" aria-label="{T_("Enganxa un missatge de Mudae")}" placeholder="{T_("Enganxa aquí la resposta de Mudae a qualsevol d'aquestes ordres…")}"></textarea>
  <span id="tuMsg" role="status"></span>
  <div class="paste-cmds"><span class="muted">{T_("Copia l'ordre:")}</span> {paste_cmds}</div></div>
 <div class="pend" id="pend" hidden><b>{T_("Pendent!")}</b> <span class="muted">{T_("El tauler encara no sap això del teu servidor ni de tu. Copia cada ordre, envia-la a Mudae i enganxa'n la resposta a dalt:")}</span><ul id="pendList"></ul></div>
</section>
<section class="panel today">
 <div class="today-head"><h2>{T_("Què toca ara")}</h2></div>
 {checklist}
</section>
<div class="g2">
 <section class="panel" id="alertsBox"{"" if alerts else " hidden"}><div class="panel-head"><h2>{T_("Avisos")}</h2><span class="muted" id="alertsN" style="font-size:13px">{len(alerts)}</span></div><ul class="alerts" id="alertsList">{"".join(alerts)}</ul></section>
 <details class="panel fold ocbox" id="ocFold" data-fold="oc"><summary><h2>{T_("Ajudant del $oc")}</h2><span class="muted" style="font-size:13px">{T_("Troba la vermella: fes servir l'ajudant")}</span></summary>
  <div class="fold-body"><a href="{oc}" style="font-size:13px">{T_("Obrir a part")}</a><iframe src="{oc}" title="{T_("Ajudant del $oc")}"></iframe></div></details>
</div>
</div>
<div class="view" id="v-coll" role="tabpanel" hidden>
<section class="panel" id="coll">
 <div class="panel-head"><h2>{T_("Col·leccions")}</h2>
  <div class="legend"><span><i style="background:var(--accent)"></i>{T_("Teus")}</span><span><i style="background:var(--warm)"></i>{T_("D'algú altre")}</span><span><i style="background:var(--sunk)"></i>{T_("Lliures")}</span></div></div>
 <div class="bars">{"".join(bars)}</div>
 <div class="album">
  <div class="album-head"><h3 id="albumTitle"></h3>
   <span class="give-row" style="margin:0"><label class="toggle"><input type="checkbox" id="hideMine"> {T_("Amaga els que ja tinc")}</label>
    <button type="button" class="tab" id="collDel">{T_("Deixar de fer-la")}</button></span></div>
  <div class="legend">
   <span class="ch wish"><span class="star">★</span> {T_("A la wishlist")}</span><span class="ch alta">{T_("Lliure · prioritat alta")}</span>
   <span class="ch other">{T_("D'algú altre")}<small>{T_("qui el té")}</small></span><span class="ch free">{T_("Lliure")}</span><span class="ch mine">{T_("Teu")}</span></div>
  <div id="chips" class="chips"></div>
  <div class="give-row" id="gotBar" hidden><b id="gotCount"></b><button class="primary" id="gotBtn">{T_("Ja són meus")}</button>
   <button type="button" class="tab" id="gotClear">{T_("Desmarca")}</button><span id="gotMsg" role="status"></span></div>
 </div>
</section>
<details class="panel fold" id="esp" data-fold="esp">
 <summary><h2>{T_("Col·lecció especial")}</h2><span class="muted" style="font-size:13px">{T_("Un grup triat per tu: per exemple, els Eevee de la Pokédex")}</span></summary>
 <div class="fold-body"><div class="give-row"><input type="text" id="espName" list="dlEsp" spellcheck="false" autocomplete="off" placeholder="{T_("Nom del grup")}" aria-label="{T_("Nom del grup")}">
  <datalist id="dlEsp">{"".join(f'<option value="{e(x)}">' for x in spec)}</datalist></div>
 <textarea id="espText" rows="5" spellcheck="false" aria-label="{T_("Personatges")}" placeholder="{T_("Un personatge per línia, com surt a Mudae")}"></textarea>
 <div class="give-row"><button class="primary" id="espBtn" disabled>{T_("Desar la col·lecció")}</button><span id="espMsg" role="status"></span></div>
 <p class="muted" style="margin:0;font-size:13px">{T_("Els teus es marquen sols. Per saber de qui són els altres, ves a «Posar al dia», tria la col·lecció i enganxa el $imaso de la sèrie d'on surten. Per canviar-la, escriu el nom d'una que ja tens: s'hi carreguen els seus personatges.")}</p>
</div></details>
<details class="panel fold" id="pd" data-fold="pokedex">
 <summary><h2>{T_("Pokédex")}</h2><span class="muted" id="pdHead" style="font-size:13px">$pd sumrare</span></summary>
 <div class="fold-body">
  <textarea id="pdText" rows="2" spellcheck="false" aria-label="$pd sumrare" placeholder="{T_("Enganxa el $pd sumrare de Mudae: quants en tens de cada raresa, en un sol missatge…")}"></textarea>
  <span id="pdMsg" role="status"></span>
  <div id="pdSum"></div>
 </div></details>
</div>
<div class="view g2" id="v-wl" role="tabpanel" hidden>
 <section class="panel" id="wlp"><div class="panel-head"><h2>{T_("Wishlist")}</h2><span class="muted" style="font-size:13px">{T_("{n} de 15 llocs (l'estrella també compta)", n=len(d["wishlist"]))}</span></div>
  <table>{wl}</table>
 </section>
 <section class="panel"><div class="panel-head"><h2>{T_("Què has de fer")}</h2></div><div id="wlBuild"></div></section>
</div>
<div class="view" id="v-trades" role="tabpanel" hidden>
<section class="panel">
 <div class="panel-head"><h2>{T_("A qui demanar")}</h2>
  <p class="recvbar"><button id="recvBtn" class="primary" disabled>{T_("Ja són meus")}</button><button id="freeBtn" class="tab" disabled title="{T_("Algú se n'ha divorciat: ara són lliures")}">{T_("Alliberar")}</button><span id="recvMsg" role="status">{T_("Marca els que t'han donat o els que algú ha alliberat.")}</span></p></div>
 <div class="owners">{trades}</div>
</section>
<section class="panel" id="mine">
 <div class="panel-head"><h2>{T_("Buscar entre els meus")}</h2><span class="muted" style="font-size:13px">{T_("{n} personatges al teu harem", n=len(d["harem"]))}</span></div>
 <label for="mineQ" class="muted" style="font-size:14px">{T_("Enganxa sèries o noms, separats per / , comes o línies (com la llista que t'envia un amic). No cal escriure-les bé del tot.")}</label>
 <textarea id="mineQ" rows="3" spellcheck="false" placeholder="Naruto / South Park / full metal alchimist / Rezero"></textarea>
 <div id="mineOut"></div>
</section>
<div class="g2">
<section class="panel" id="cls">
 <div class="panel-head"><h2>{T_("Sèrie dels teus personatges")}</h2><span class="muted" style="font-size:13px">{T_("dels que no col·lecciones")}</span></div>
 {cls_html}
 <div class="give-row"><input type="text" id="clsName" list="dlNames" spellcheck="false" autocomplete="off" placeholder="{T_("Personatge")}" aria-label="{T_("Personatge")}">
  <span>{T_("és de")}</span><input type="text" id="clsSerie" list="dlSeries" spellcheck="false" autocomplete="off" placeholder="{T_("Sèrie")}" aria-label="{T_("Sèrie")}">
  <button class="primary" id="clsBtn" disabled>{T_("Desar")}</button><span id="clsMsg" role="status"></span></div>
 {datalists}
</section>
<section class="panel"><h2>{T_("Sobrants per sèrie")}</h2><table>{sob}</table></section>
</div>
</div>
<div class="view" id="v-mill" role="tabpanel" hidden>
<section class="panel">
 <div class="panel-head"><h2>{T_("Què tens")}</h2><span class="muted" id="millSum" style="font-size:13px"></span></div>
 <div class="paste"><textarea id="millText" rows="1" spellcheck="false" aria-label="{T_("Enganxa un missatge de Mudae")}" placeholder="{T_("Enganxa aquí la resposta de Mudae a qualsevol d'aquestes ordres…")}"></textarea>
  <span id="millMsg" role="status"></span>
  <div class="paste-cmds"><span class="muted">{T_("Copia l'ordre:")}</span> {mill_cmds}</div></div>
 <p class="muted" id="millMiss" style="margin:0;font-size:13px"></p>
</section>
<section class="panel"><div class="panel-head"><h2>{T_("Els teus bonus")}</h2><span class="muted" style="font-size:13px">$bonus</span></div>
 <ul class="steps" id="bonusNext"></ul>
 <details class="fold" data-fold="bonus" id="bonusFold" hidden><summary><h3 class="og-h">{T_("Tots els bonus")}</h3></summary><div class="fold-body"><table id="bonusAll"></table></div></details>
</section>
<div class="g2" id="millG2">
 <details class="panel fold" id="badgeBox" open><summary><h2>{T_("Insígnies")}</h2><span class="muted" style="font-size:13px">$kakerareward</span> <span class="tag good" id="badgeMax" hidden>{T_("Totes al màxim")}</span></summary><div class="fold-body">
  <h3 class="og-h">{T_("Les següents que et recomano")}</h3><ul class="steps" id="badgeNext"></ul>
  <table>{badge_rows}</table>
  <p class="muted" style="margin:0;font-size:13px">{T_("Primer, gastar menys poder per reacció (Gold IV) i el Ruby IV, que abaixa un 25 % el preu de totes les que queden; després, més tirades (Sapphire).")}</p>
 </div></details>
 <section class="panel"><div class="panel-head"><h2>{T_("Torre")}</h2><span class="muted" style="font-size:13px">$kt</span></div>
  <p class="clock" id="towerNext" style="margin:0"></p>
  <p class="muted" id="kakeraNext" style="margin:0;font-size:13px" hidden></p>
  <ul class="steps">{tower_rows}</ul>
  <p class="muted" style="margin:0;font-size:13px">{T_("Marca els pisos que ja tens de la torre on ets. Cada torre té 12 pisos, un de cada: quan l'acabes, comença la següent, amb els mateixos pisos però amb menys efecte (els de la llista són els de la primera). Cada pis costa més que l'anterior, triïs el que triïs: construeix primer els de dalt. Amb $destroy et torna la kakera.")}
   <label style="white-space:nowrap">{T_("Preu base ($towervalue):")} <input type="text" id="towerValue" inputmode="numeric" style="width:90px" aria-label="$towervalue"></label></p>
 </section>
</div>
<section class="panel" id="perkBox"><div class="panel-head"><h2>{T_("Esferes")}</h2><span class="muted" id="perkWho" style="font-size:13px"></span></div>
 <h3 class="og-h">{T_("Les següents que et recomano")}</h3><ul class="steps" id="perkNext"></ul>
 <div class="g2">
  <div><h3 class="og-h">{T_("Els que et surten més (claus, $mmy=)")}</h3><p id="perkKeys" style="margin:0;font-size:14px;line-height:1.9"></p></div>
  <div><h3 class="og-h">{T_("Personatges que has millorat")}</h3><p id="perkList" style="margin:0;font-size:14px;line-height:1.9"></p>
   <p class="muted" style="margin:6px 0 0;font-size:13px">{T_("El $opz et mostra els personatges que tenen esferes, un per pàgina: enganxa cada pàgina (amb la fletxa de Mudae passes a la següent) i el tauler se'n recorda.")}</p></div>
 </div>
 <details class="fold" data-fold="perks"><summary><h3 class="og-h">{T_("Totes les millores")}</h3></summary><div class="fold-body">
 <div class="ogrid">{perk_rows}</div>
 <p class="muted" style="margin:0;font-size:13px">{T_("Les 5 primeres tenen 6 nivells i les altres 5, un. Si millores un personatge del tot, les 5 primeres pugen un nivell més. Fes-les servir en personatges que no regalaràs: si te'n divorcies, perds el 10 %. Amb $or <personatge> et tornen les esferes d'un personatge, també amb un 10 % menys.")}</p></div></details>
</section>
<section class="panel" id="shopBox"><div class="panel-head"><h2>{T_("Botiga d'esferes")}</h2><span class="muted" style="font-size:13px">$os</span></div>
 <h3 class="og-h">{T_("Les següents que et recomano")}</h3><ul class="steps" id="shopNext"></ul>
 <details class="fold" data-fold="shop"><summary><h3 class="og-h">{T_("Totes les millores")}</h3></summary><div class="fold-body">
 <div class="og">{shop_rows}</div>
 <p class="muted" style="margin:0;font-size:13px">{T_("Valen per a tots els personatges, però gairebé totes fan més forta una millora dels personatges: primer gasta les esferes amb el $opp. Cada nivell costa 4.000 més que l'anterior, i amb $shoprefund et tornen les esferes (un cop cada 15 dies).")}</p></div></details>
</section>
</div>
<div class="view" id="v-eines" role="tabpanel" hidden>
<section class="panel" id="upd">
 <div class="panel-head"><h2>{T_("Posar al dia")}</h2><span id="updAge" class="muted" style="font-size:13px"></span></div>
 <div class="give-row">{T_("Què enganxes:")}
  <select id="updKind" aria-label="{T_("Què enganxes:")}"><option value="harem">{T_("El teu harem ($mma+s)")}</option><option value="wishlist">{T_("La wishlist ($wl)")}</option><option value="totes">{T_("Totes les teves sèries")}</option><option value="serie">{T_("Una sèrie ($imaso)")}</option></select>
  <span id="updSerieBox" hidden>{T_("Sèrie:")} <select id="updSerie" aria-label="{T_("Sèrie:")}">{serie_opts}<option value="__new__">{T_("Sèrie nova…")}</option></select>
  <input type="text" id="updNew" spellcheck="false" autocomplete="off" aria-label="{T_("Nom de la sèrie, com surt a Mudae")}" placeholder="{T_("Nom de la sèrie, com surt a Mudae")}" hidden></span>
  <span>{T_("Ordre:")} <button class="cmd" id="updCmd" data-copy="$mma+s">$mma+s</button></span></div>
 <p class="muted" id="updEsp" style="margin:0;font-size:13px" hidden>{T_("Col·lecció especial: enganxa el $imaso de la sèrie d'on surten els personatges (totes les pàgines). Només se n'agafaran els d'aquesta col·lecció.")}</p>
 <div class="give" id="updAll" hidden><div class="give-row"><b>{T_("Ordre amb totes les teves sèries")}</b><button class="copybtn" id="updAllCopy" data-copy="">{T_("Copiar ordre")}</button></div><code id="updAllCmd"></code>
  <span class="muted" style="font-size:13px">{T_("1. Copia l'ordre: són les teves sèries separades amb $. 2. Envia-la a Mudae. 3. Enganxa a sota tot el que et contesti per privat (tots els missatges).")}</span>
  <label class="give-row">{T_("Sèries noves:")} <input type="text" id="updAllNew" spellcheck="false" autocomplete="off" style="flex:1;min-width:220px;width:auto" placeholder="{T_("Harry Potter $ Pingu $ ABBA")}"></label>
  <span class="muted" style="font-size:13px">{T_("Escriu les que vulguis afegir, separades amb $ i com surten a Mudae: entren a l'ordre i es creen soles quan enganxes la resposta.")}</span></div>
 <textarea id="updText" rows="4" spellcheck="false" aria-label="{T_("Posar al dia")}" placeholder="{T_("Enganxa aquí el missatge sencer que t'ha enviat Mudae per privat")}"></textarea>
 <div class="give-row"><button class="primary" id="updBtn" disabled>{T_("Desar i actualitzar")}</button><span id="updMsg" role="status"></span></div>
</section>
<section class="panel" id="ordres"><div class="panel-head"><h2>{T_("Ordres d'ús habitual")}</h2>
 <input type="search" id="ordQ" spellcheck="false" autocomplete="off" aria-label="{T_("Busca una ordre")}" placeholder="{T_("Busca una ordre: wish, kakera, $tu…")}"></div>
 <div class="ogrid">{ordres}</div><p class="muted" id="ordNone" style="margin:0;font-size:14px" hidden>{T_("Cap ordre amb aquest text.")}</p></section>
</div>
<footer class="foot glass">Developed by <a href="https://github.com/Vlarac03">Vlarac</a> · {T_("Tauler de Mudae")} · {T_("Projecte de fans, no afiliat amb Mudae")}</footer>
</main><div id="tip"></div>
<script>{JS_HEAD.replace("__I18N__", js(STRINGS.get(LANG, {})))}
// Botons «Ja són meus» i «Alliberar» (algú se n'ha divorciat) dels intercanvis: només funcionen amb el servidor local engegat
const recvBtn = document.getElementById('recvBtn'), freeBtn = document.getElementById('freeBtn'), recvMsg = document.getElementById('recvMsg');
const picked = () => [...document.querySelectorAll('.recv:checked')].map(i => i.value);
if (location.protocol === 'file:') {{
  // obert com a fitxer: si el servidor està engegat, hi saltem; si no, avisem
  document.querySelectorAll('.recv').forEach(i => i.disabled = true);
  fetch('http://127.0.0.1:{PORT}/ping').then(r => r.json()).then(() => location.replace('http://127.0.0.1:{PORT}/dashboard.html'))
    .catch(() => {{ recvMsg.innerHTML = OFF_MSG(); }});
}} else {{
  document.addEventListener('change', ev => {{ if (ev.target.classList.contains('recv')) recvBtn.disabled = freeBtn.disabled = !picked().length; }});
  const send = path => async () => {{
    recvBtn.disabled = freeBtn.disabled = true; recvMsg.textContent = T('Desant…');
    try {{
      const j = await postJSON(path, {{ names: picked() }});
      if (j.ok) location.reload(); else {{ recvMsg.innerHTML = ERR_MSG(j.error); recvBtn.disabled = freeBtn.disabled = false; }}
    }} catch (err) {{ recvMsg.innerHTML = NET_MSG(); recvBtn.disabled = freeBtn.disabled = false; }}
  }};
  recvBtn.onclick = send('/rebut');
  freeBtn.onclick = send('/lliure');
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
const refaBtn = document.getElementById('refaBtn');
if (location.protocol === 'file:') {{ lnkBtn.disabled = stopBtn.disabled = refaBtn.disabled = true; }}
// Torna a fer el tauler i l'Excel, com l'Actualitza.bat, i mira si hi ha una versió nova a GitHub: si n'hi ha, pregunta i la posa
const wait = ms => new Promise(r => setTimeout(r, ms));
refaBtn.onclick = async () => {{
  refaBtn.disabled = true; toolMsg.textContent = T('Actualitzant…');
  try {{
    const j = await postJSON('/novetats'); if (!j.ok) throw j.error;
    if (j.warn && confirm(T('Hi ha una versió nova del tauler a GitHub. La poso ara? Les teves dades no es toquen.'))) {{
      toolMsg.textContent = T('Posant la versió nova…');
      const u = await postJSON('/actualitza'); if (!u.ok) throw u.error;
      await wait(3000);  // el servidor nou substitueix el vell: esperem que respongui
      for (let k = 0; k < 15; k++) try {{ await fetch('/ping', {{ cache: 'no-store' }}); break; }} catch (e) {{ await wait(1000); }}
    }}
    return location.reload();
  }} catch (err) {{ toolMsg.innerHTML = typeof err === 'string' ? ERR_MSG(err) : NET_MSG(); }}
  refaBtn.disabled = false;
}};
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
// els avisos i les xifres del resum porten a la pestanya on es fa la feina
document.addEventListener('click', ev => {{
  const b = ev.target.closest('[data-goto]'); if (!b) return;
  go(b.dataset.goto); scrollTo({{ top: 0 }});
}});
// menú «···» de la capçalera: es tanca clicant fora, amb Esc o en triar-ne una opció
const more = document.getElementById('more');
document.addEventListener('click', ev => {{ if (more.open && (!more.contains(ev.target) || ev.target.closest('.menu button'))) more.open = false; }});
document.addEventListener('keydown', ev => {{ if (ev.key === 'Escape' && more.open) {{ more.open = false; more.querySelector('summary').focus(); }} }});

</script><script>{JS_TODAY}</script>
<script>{JS_COLL.replace("__DATA__", js(coll))}</script>
<script>{JS_UPD.replace("__META__", js(read_meta())).replace("__SPECIALS__", js(spec))}</script>
<script>{JS_ESP}</script>
<script>{JS_MILL.replace("__BADGES__", js([[n, p] for n, p, _ in BADGES])).replace("__PLAN__", js(BADGE_PLAN)).replace("__TOWER__", js([n for n, _ in TOWER])).replace("__PORD__", js([n for _, ns in PERK_GROUPS for n in ns])).replace("__SORD__", js(SHOP_ORDER)).replace("__PDESC__", js({n: t(x) for n, x in PERKS.items()})).replace("__SDESC__", js({n: t(x) for n, x in SHOP.items()}))}</script>
<script>{JS_CLS}</script>
<script>{JS_ORD}</script>
<script>{JS_PD}</script>
<script>{JS_PASTE.replace("__NEED__", js(need))}</script>
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


# Windows no deixa posar \ / : * ? " < > | als noms de fitxer: es desen amb els signes d'amplada completa (：) i es llegeixen tal qual
TO_FILE = str.maketrans('\\/:*?"<>|', '＼／：＊？＂＜＞｜')
FROM_FILE = {v: k for k, v in TO_FILE.items()}


def series_path(serie):
    if not serie or serie.startswith("."):
        raise ValueError(t("Posa el nom de la sèrie."))
    return DADES / "series" / f"{serie.translate(TO_FILE)}.txt"


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
    found = {n.casefold(): o for n, o in parse_owners(drop_header(text, check=False)).items()}
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


def update_all(text, new=()):
    """$imaa+os S1 $ S2 $ …: totes les teves sèries d'una vegada. Cada línia és com al $imaso amb la sèrie al final:
    «Kate Beckett  💞 => pep - Castle» o «Kevin Ryan - Castle». El títol («Castle + Pingu + ABBA   12/19») diu quants n'hi ha.

    Una sèrie de la qual surten menys personatges dels que ja tenies no es toca (falta un tros del missatge).
    Les col·leccions especials s'agafen pel nom dels seus personatges, siguin de la sèrie que siguin.
    new: sèries noves escrites al tauler; es creen amb el nom que surt a Mudae (si hi surten).
    Torna un avís amb el que no s'ha posat al dia, o "".
    """
    spec = specials()
    files = {f.stem.translate(FROM_FILE): f for f in (DADES / "series").glob("*.txt")}
    known = {x.casefold() for x in files}
    new = list(dict.fromkeys(n.strip() for n in new if n.strip() and n.strip().casefold() not in known))
    normal = sorted([s for s in files if s not in spec] + new, key=len, reverse=True)  # la més llarga primer: «Castle» no es menja «Richard Castle»
    spec_names = sorted({n.casefold() for s in spec if s in files for n in parse_owners(files[s].read_text(encoding="utf-8-sig"))},
                        key=len, reverse=True)
    # «Pers. sin serie principal en este paquete»: Mudae hi posa una altra sèrie («Meemaw - Young Sheldon»), però ja és a la teva
    name_of = lambda body: re.split(r"\s*(?:💞|=>)", body)[0].strip().casefold()
    member = {n.casefold(): s for s in files if s not in spec for n in parse_owners(files[s].read_text(encoding="utf-8-sig"))}
    rows = text.splitlines()
    title = next((r for r in rows[:6] if IMASO_TITLE.match(r.strip())), None)
    got, extra, real, lost = {}, [], {}, []
    for raw in rows:
        line = re.sub(r"^\d+\.\s+", "", raw.replace("⭐", "").strip())
        low = line.casefold()
        s = next((s for s in normal if low.endswith(" - " + s.casefold())), None)
        if s:
            if s in new:  # el nom com el diu Mudae, no com l'has escrit
                real.setdefault(s, line[-len(s):])
            got.setdefault(s, []).append(line[:-len(s) - 3].strip())  # el que queda és una línia del $imaso de sempre
        elif " - " in line and (body := line.rsplit(" - ", 1)[0].strip()) and name_of(body) in member:
            got.setdefault(member[name_of(body)], []).append(body)
        elif " - " in line and any(low.startswith(n + " ") for n in spec_names):
            extra.append(line.rsplit(" - ", 1)[0].strip())
        elif " - " in line:  # de la llista «sin serie principal» i encara no és a cap sèrie teva: no sabem de quina és
            lost.append(re.split(r"\s*(?:💞|=>)", line.rsplit(" - ", 1)[0])[0].strip())
    read = sum(map(len, got.values())) + len(extra)
    done, short = 0, []
    for s, found in got.items():
        old = len(lines(files[s])) if s in files else 0
        if len(found) < old:
            short.append(t("{serie} ({n} de {old})", serie=s, n=len(found), old=old))
            continue
        save_paste("serie", real.get(s, s), "\n".join(found))
        done += 1
    for s in spec:
        if s in files and extra:
            try:
                update_special(s, "\n".join(extra))
                done += 1
            except ValueError:  # cap personatge d'aquesta especial al text: no la toquem
                pass
    if not done and not short:
        raise ValueError(t("No he trobat cap personatge de les teves sèries. Copia tots els missatges que t'ha enviat Mudae."))
    missing = [s for s in normal if s not in got]
    warn = []
    total = int(re.search(r"(\d+)\s*$", title).group(1)) if title else read
    if read < total:
        warn.append(t("Mudae diu {total} personatges i n'he llegit {n}.", total=total, n=read))
    if lost:
        warn.append(t("No sé de quina de les teves sèries són {names}: posa al dia aquella sèrie amb «Una sèrie ($imaso)».", names=", ".join(lost[:10])))
    if short:
        warn.append(t("No he tocat {series}: en falten. Copia tots els missatges de Mudae.", series=", ".join(short)))
    if missing:
        warn.append(t("No hi surten: {series}.", series=", ".join(sorted(missing))))
    return " ".join(warn)


def save_paste(kind, serie, text, new=()):
    """Desa el que has enganxat al tauler al fitxer de dades/ que toca. Si no s'hi troba cap nom, no toca res."""
    if kind == "harem":
        classify_harem(text)  # els que no s'han pogut llegir surten a «Sèrie dels teus personatges»
        return
    elif kind == "totes":
        return update_all(text, new)
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


def classify_harem(text):
    """$mma+s: el harem sencer, «Personatge - Sèrie» per línia. Desa els noms a harem.txt i la sèrie de cada un a series-sobrants.json.

    Algunes sèries porten un guionet al nom («Vigilante - Boku no Hero Academia: Illegals»): es talla pel primer
    « - » que deixa a l'esquerra un nom del teu harem (o pel primer, si el harem no el té).
    """
    harem = {n for n, _ in lines(DADES / "harem.txt")}
    got, bad = {}, []
    for raw in text.splitlines():
        line = re.sub(r"^\d+\.\s+", "", raw.replace("⭐", "").strip())
        if not line or SKIP.search(line):
            continue
        parts = line.split(" - ")
        if len(parts) < 2:
            bad.append(line)
            continue
        cut = next((k for k in range(1, len(parts)) if " - ".join(parts[:k]).strip() in harem), 1)
        name, serie = " - ".join(parts[:cut]).strip(), " - ".join(parts[cut:]).strip()
        if name and serie:
            got[name] = serie
    if not got:
        raise ValueError(t("No he trobat cap «Personatge - Sèrie». Copia el missatge del $mma+s sencer."))
    (DADES / "harem.txt").write_text("\n".join(got) + "\n", encoding="utf-8")
    for f in (DADES / "series").glob("*.txt"):  # els que eren teus i ja no hi són: els has donat o te n'has divorciat
        rows = f.read_text(encoding="utf-8-sig").splitlines()
        new = []
        for r in rows:
            m = CLAIM.match(re.sub(r"^\d+\.\s+", "", r.strip()))
            gone = m and m.group(2).casefold() == ME.casefold() and m.group(1).strip() not in got
            new.append(f"{m.group(1).strip()} 💞 => algú" if gone else r)  # de qui és ara, ho dirà el pròxim $imaso
        if new != rows:
            f.write_text("\n".join(new) + "\n", encoding="utf-8")
    path = DADES / "series-sobrants.json"
    old = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
    by_name = {n: sr for sr, ns in old.items() for n in ns if sr != ALTRES}  # el que ja sabíem i ara no surt, es queda
    by_name.update(got)
    m = {}
    for n, sr in sorted(by_name.items(), key=lambda kv: (kv[1], kv[0])):
        m.setdefault(sr, []).append(n)
    path.write_text(json.dumps(m, ensure_ascii=False, indent=1), encoding="utf-8")
    meta = read_meta()
    meta["harem"] = dt.date.today().isoformat()
    (DADES / "actualitzat.json").write_text(json.dumps(meta, ensure_ascii=False, indent=1), encoding="utf-8")
    return len(got), bad


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


def delete_series(name):
    """Ja no fas aquesta col·lecció: fora el seu fitxer (i de les especials). Els personatges del harem no es toquen."""
    path = series_path(name)
    if not path.exists():
        raise ValueError(t("No trobo la col·lecció «{name}».", name=name))
    path.unlink()
    spec = specials()
    if name in spec:
        (DADES / "especials.json").write_text(json.dumps([s for s in spec if s != name], ensure_ascii=False, indent=1), encoding="utf-8")
    meta = read_meta()
    if meta.get("series", {}).pop(name, None):
        (DADES / "actualitzat.json").write_text(json.dumps(meta, ensure_ascii=False, indent=1), encoding="utf-8")


def mark_free(name):
    """Algú s'ha divorciat de `name`: queda lliure al fitxer de la seva sèrie, com si ho hagués dit el $imaso."""
    found = False
    for f in sorted((DADES / "series").glob("*.txt")):
        text = f.read_text(encoding="utf-8-sig").splitlines()
        for k, raw in enumerate(text):
            line = re.sub(r"^\d+\.\s+", "", raw.strip())
            m = CLAIM.match(line) or HEART_ONLY.match(line)
            if m and m.group(1).strip() == name:
                text[k] = name
                f.write_text("\n".join(text) + "\n", encoding="utf-8")
                found = True
    if not found:
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


REPO_ZIP = "https://github.com/Vlarac03/mudae-dashboard/archive/refs/heads/main.zip"


def github_files():
    """La darrera versió de GitHub: {camí: bytes}. Mai res de dades/ ni camins que surtin de la carpeta."""
    import io
    import urllib.request
    import zipfile
    with urllib.request.urlopen(REPO_ZIP, timeout=30) as r:
        return zip_files(zipfile.ZipFile(io.BytesIO(r.read())))


def zip_files(z):
    out = {}
    for n in z.namelist():
        p = n.split("/", 1)[1] if "/" in n else ""  # treu la carpeta «mudae-dashboard-main/»
        if not p or p.endswith("/") or p.startswith(("dades/", "/")) or ".." in p.split("/") or ":" in p or "\\" in p:
            continue
        data = z.read(n)
        if p.endswith(".bat"):  # cmd llegeix malament els .bat sense CRLF
            data = data.replace(b"\r\n", b"\n").replace(b"\n", b"\r\n")
        out[p] = data
    return out


def changed_files(files):
    """Els fitxers de GitHub que aquí falten o són diferents (sense comptar els salts de línia: el ZIP i Git els poden canviar)."""
    norm = lambda b: b.replace(b"\r\n", b"\n")
    return [p for p, data in files.items() if not (DIR / p).is_file() or norm((DIR / p).read_bytes()) != norm(data)]


def do_action(action, body):
    """El que fan els botons del tauler. S'executa en un procés nou, sempre amb la versió actual del codi."""
    if action == "rebut":
        for n in body["names"]:
            mark_received(str(n))
    elif action == "esborra":
        delete_series(str(body["name"]))
    elif action == "lliure":
        for n in body["names"]:
            mark_free(str(n))
    elif action == "donat":
        mark_given([str(n) for n in body["names"]], str(body.get("to") or ""))
    elif action == "classifica":
        classify(str(body["name"]), str(body["serie"]))
    elif action == "dressera":
        make_shortcut()
    elif action == "dades":
        return save_paste(str(body["kind"]), str(body.get("serie") or ""), str(body["text"]), [str(n) for n in body.get("noves") or []])
    elif action == "especial":
        save_special(str(body["name"]), str(body["text"]))
    elif action == "refa":
        pass  # només tornar a fer el tauler i l'Excel (main() ho fa després de cada acció)
    elif action == "novetats":  # «Actualitzar el tauler»: quants fitxers han canviat a GitHub ("" si cap); el tauler es refà igualment
        if (DIR / ".git").exists():
            return ""  # una còpia de git (la de Victor) es posa al dia amb git pull, no aquí
        try:
            n = len(changed_files(github_files()))
        except OSError:
            return ""  # shortcut: sense internet, com si no n'hi hagués; si cal, avisar-ne al tauler
        return str(n) if n else ""
    elif action == "actualitza":  # posa la versió de GitHub; dades/ no es toca mai, i no s'esborra res
        if (DIR / ".git").exists():
            raise ValueError(t("Aquesta carpeta és una còpia de git: posa-la al dia amb git pull."))
        files = github_files()
        for p in changed_files(files):
            (DIR / p).parent.mkdir(parents=True, exist_ok=True)
            (DIR / p).write_bytes(files[p])
        import subprocess
        # aquest procés encara té el codi vell carregat: el tauler i l'Excel els fa el nou
        raise SystemExit(subprocess.run([sys.executable, str(DIR / "mudae.py")]).returncode)
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
                warn = next((l[6:] for l in r.stdout.splitlines() if l.startswith("AVIS: ")), "")
                res = {"ok": True, "warn": warn} if r.returncode == 0 else {"ok": False, "error": err}
            except Exception as ex:
                res = {"ok": False, "error": str(ex)}
            self.reply(res)
            if self.path == "/actualitza" and res["ok"] and str((DIR / "mudae.py").stat().st_mtime_ns) != VERSION:
                # aquest servidor és codi vell: n'engega un de nou, que li demanarà que plegui (com fa Actualitza.bat)
                subprocess.Popen([sys.executable, str(DIR / "mudae.py"), "tauler"], env={**os.environ, "MUDAE_SENSE_NAVEGADOR": "1"},
                                 creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))

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
    if not os.environ.pop("MUDAE_SENSE_NAVEGADOR", None):  # el servidor nou d'una actualització: el tauler ja és obert
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
        warn = None
        if args[:1] == ["accio"]:  # el servidor: python mudae.py accio <nom> < JSON
            if sys.stdin and hasattr(sys.stdin, "reconfigure"):
                sys.stdin.reconfigure(encoding="utf-8")
            warn = do_action(args[1], json.loads(sys.stdin.read() or "{}"))
        if args[:1] == ["rebut"]:  # python mudae.py rebut "Nom 1" "Nom 2"
            for n in args[1:]:
                mark_received(n)
        if args[:1] == ["donat"]:  # python mudae.py donat "@persona" "Nom 1" "Nom 2"
            mark_given(args[2:], args[1])
        if args[:1] == ["classifica"]:  # python mudae.py classifica "Personatge" "Sèrie"
            classify(args[1], args[2])
        run()
        if warn:
            print("AVIS: " + warn)  # el servidor el passa al tauler
    except (RuntimeError, ValueError, KeyError) as ex:
        sys.exit(str(ex))
    if args[:1] == ["serve"]:
        serve()


if __name__ == "__main__":
    main()
