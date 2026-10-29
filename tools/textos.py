"""Llista tots els textos de la interfície (en català, que fan de clau) i diu quins falten a i18n.py.

Executa: python tools/textos.py      (la prova tests/test_mudae.py fa servir keys())
"""
import ast
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))


def keys():
    """Textos que el programa pot mostrar: crides t()/T_() a Python, T() a JavaScript, i les taules de textos."""
    import mudae
    src = (ROOT / "mudae.py").read_text(encoding="utf-8")
    found = set()
    for node in ast.walk(ast.parse(src)):
        if isinstance(node, ast.Call) and getattr(node.func, "id", None) in ("t", "T_") and node.args \
                and isinstance(node.args[0], ast.Constant) and isinstance(node.args[0].value, str):
            found.add(node.args[0].value)
    for m in re.finditer(r"""\bT\((['"])((?:\\.|(?!\1).)*)\1""", src):  # T('…') i T("…") dins del JavaScript
        found.add(m.group(2).replace("\\'", "'").replace('\\"', '"'))
    found |= {"Omple {n} lloc", "Omple {n} llocs"}  # T(free === 1 ? … : …)
    for g, que, ordre, com, _ in mudae.RUTINA:
        found |= {que, com, *(p for p in ordre.split(" · ") if not p.startswith("$"))}
    found |= set(mudae.GRUPS.values())
    for g, items in mudae.ORDRES:
        found |= {g, *(desc for _, desc in items)}
    for _, cmds in mudae.ORDRES:
        for c, _ in cmds:
            found |= set(re.findall(r"<([^>]+)>", c))
    for _, _, ordre, _, _ in mudae.RUTINA:
        found |= set(re.findall(r"<([^>]+)>", ordre))
    found |= {mudae.ALTRES, "Meu", "Lliure", "D'algú altre", "Alta", "Mitjana", "Baixa", "Pendent", "Enviat", "Acceptat",
              "Rebutjat", "Guardar", "Per intercanviar", "Divorciar", "persona", "Resum", "Col·lecció", "Sobrants", "Wishlist",
              "Intercanvis", "Rutina"}
    return {k for k in found if k and not k.startswith("$")}


if __name__ == "__main__":
    from i18n import STRINGS
    ks = sorted(keys())
    for lang in STRINGS:
        miss = [k for k in ks if k not in STRINGS[lang]]
        print(f"{lang}: {len(ks) - len(miss)}/{len(ks)} traduïts")
        for k in miss:
            print("   ", repr(k))
