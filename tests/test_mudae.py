"""Proves del que llegeix i desa mudae.py. Executa: python -m unittest discover tests"""
import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import mudae  # noqa: E402


class Base(unittest.TestCase):
    """Cada prova té la seva carpeta dades/ buida, i l'usuari és «jo»."""

    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.old = (mudae.DADES, mudae.ME)
        mudae.DADES, mudae.ME = self.tmp, "jo"
        (self.tmp / "series").mkdir()
        (self.tmp / "series-sobrants.json").write_text("{}", encoding="utf-8")

    def tearDown(self):
        mudae.DADES, mudae.ME = self.old
        shutil.rmtree(self.tmp)

    def write(self, rel, text):
        (self.tmp / rel).write_text(text, encoding="utf-8")


class Llegir(Base):
    def test_ignora_capcaleres_i_numeracio(self):
        self.write("harem.txt", "1. MudaeAPP — 18:02\nHarem de jo\n2. Remus Lupin\nPage 1 / 21\n$mms\nHedwig ⭐\nGinger Torte | RANDOM\n")
        self.assertEqual(mudae.lines(self.tmp / "harem.txt"), [("Remus Lupin", False), ("Hedwig", True), ("Ginger Torte", False)])

    def test_wishlist_amb_marques(self):
        self.write("wishlist.txt", "ggezw's Wishlist - 15/15 $wl, 1/1 $sw\nAlbedo ✅ ⭐ +100%\nNero ✅ +200%\nRukia Kuchiki ✅\n"
                                   "Yoruichi Shihōin ✅ +30%\nRemus Lupin\n")
        self.assertEqual(mudae.lines(self.tmp / "wishlist.txt"), [("Albedo", True), ("Nero", False), ("Rukia Kuchiki", False),
                                                                  ("Yoruichi Shihōin", False), ("Remus Lupin", False)])

    def test_fitxer_que_no_existeix(self):
        self.assertEqual(mudae.lines(self.tmp / "no.txt"), [])

    def test_propietaris(self):
        self.write("series/HP.txt", "Aragog 💞 => jo\nDobby\nSan  💞\nAragog 💞 => jo\n")
        got = [(r["name"], r["owner"]) for r in mudae.read_series()]
        self.assertEqual(got, [("Aragog", "jo"), ("Dobby", ""), ("San", "algú")])  # sense duplicats


class Majuscules(Base):
    def test_nom_amb_majuscula(self):
        mudae.ME = "Jo"  # a Configura.bat amb majúscula; Mudae escriu «jo»
        self.write("harem.txt", "Aragog\n")
        self.write("series/HP.txt", "Aragog => jo\nDobby => pep\n")
        self.assertEqual([r["owner"] for r in mudae.read_series()], ["Jo", "pep"])
        self.assertEqual(mudae.state(mudae.read_series()[0]["owner"]), "Meu")
        mudae.mark_given(["Aragog"], "@anna")
        self.assertEqual((self.tmp / "series/HP.txt").read_text(encoding="utf-8"), "Aragog 💞 => anna\nDobby => pep\n")


class Desar(Base):
    def test_privat_conserva_propietaris(self):
        self.write("harem.txt", "Kiki\n")
        old = "Aragog 💞 => jo\nDudley 💞 => pep\nDobby\nNorris 💞 => anna\n"
        new = "1. Aragog  💞\nDudley 💞\nDobby 💞\nNorris\nKiki 💞\n"
        self.assertEqual(mudae.merge_owners(new, old).splitlines(),
                         ["Aragog 💞 => jo", "Dudley 💞 => pep", "Dobby 💞 => algú", "Norris", "Kiki 💞 => jo"])

    def test_imaso_sense_capcalera(self):
        paste = ("1. Harry Potter 3/94\n\nHP\nHarry Potter and the Philosopher's Stone\n"
                 "4 más - Escriba $ima? para mostrar todos los alias.\n \n\n"
                 "Aberforth Dumbledore\nAragog => jo\n2. Dudley Dursley => pep\nWinky (HP)\n")
        mudae.save_paste("serie", "Harry Potter", paste)
        got = [(r["name"], r["owner"]) for r in mudae.read_series()]
        self.assertEqual(got, [("Aberforth Dumbledore", ""), ("Aragog", "jo"), ("Dudley Dursley", "pep"), ("Winky (HP)", "")])

    def test_imaso_sense_alies_no_perd_personatges(self):
        mudae.save_paste("serie", "ABBA", "1. ABBA 1/4\n\nAgnetha\nAnni-Frid => jo\nBenny\nBjörn\n")
        self.assertEqual([r["name"] for r in mudae.read_series()], ["Agnetha", "Anni-Frid", "Benny", "Björn"])

    def test_text_sense_noms_no_toca_res(self):
        self.write("harem.txt", "Kiki\n")
        with self.assertRaises(ValueError):
            mudae.save_paste("harem", "", "MudaeAPP — 18:02\nPage 1 / 2\n")
        self.assertEqual((self.tmp / "harem.txt").read_text(encoding="utf-8"), "Kiki\n")

    def test_desa_i_apunta_la_data(self):
        mudae.save_paste("wishlist", "", "jo's Wishlist - 2/15\nSnape ⭐\nDobby\n")
        self.assertEqual([n for n, _ in mudae.lines(self.tmp / "wishlist.txt")], ["Snape", "Dobby"])
        self.assertIn("wishlist", json.loads((self.tmp / "actualitzat.json").read_text(encoding="utf-8")))

    def test_nom_de_serie_perillos(self):
        with self.assertRaises(ValueError):
            mudae.save_paste("serie", "../fora", "Dobby\n")

    def test_nom_de_serie_amb_dos_punts(self):
        mudae.save_paste("serie", "Re:Zero / Hyouketsu?", "Emilia\n")
        self.assertEqual([(r["serie"], r["name"]) for r in mudae.read_series()], [("Re:Zero / Hyouketsu?", "Emilia")])
        self.assertEqual([f.parent for f in (self.tmp / "series").glob("*.txt")], [self.tmp / "series"])  # el / no fa carpetes

    def test_series_del_harem(self):
        self.write("harem.txt", "Agnetha\nOboro Shirakumo\nAloe\n")
        self.write("series-sobrants.json", '{"Vella": ["Aloe"]}')
        paste = ("1. Agnetha - ABBA\nOboro Shirakumo - Vigilante - Boku no Hero Academia: Illegals\n"
                 "2. Endeavor - Boku no Hero Academia\nlinia sense guionet\n")
        mudae.save_paste("harem", "", paste)
        got = json.loads((self.tmp / "series-sobrants.json").read_text(encoding="utf-8"))
        self.assertEqual(got, {"ABBA": ["Agnetha"], "Boku no Hero Academia": ["Endeavor"], "Vella": ["Aloe"],
                               "Vigilante - Boku no Hero Academia: Illegals": ["Oboro Shirakumo"]})
        self.assertEqual([n for n, _ in mudae.lines(self.tmp / "harem.txt")], ["Agnetha", "Oboro Shirakumo", "Endeavor"])
        self.assertIn("harem", json.loads((self.tmp / "actualitzat.json").read_text(encoding="utf-8")))
        with self.assertRaises(ValueError):
            mudae.save_paste("harem", "", "res de Mudae aquí\n")

    def test_classifica(self):
        self.write("harem.txt", "Aloe\n")
        mudae.classify("Aloe", "Ishuzoku Reviewers")
        mudae.classify("Aloe", "Una altra")
        self.assertEqual(json.loads((self.tmp / "series-sobrants.json").read_text(encoding="utf-8")), {"Una altra": ["Aloe"]})
        with self.assertRaises(ValueError):
            mudae.classify("Goku", "Dragon Ball")

    def test_rebut_i_donat(self):
        self.write("harem.txt", "Hedwig\n")
        self.write("series/HP.txt", "Hedwig 💞 => jo\nDobby 💞 => pep\n")
        mudae.mark_received("Dobby")
        mudae.mark_given(["Hedwig"], "@anna")
        self.assertEqual((self.tmp / "series/HP.txt").read_text(encoding="utf-8"), "Hedwig 💞 => anna\nDobby 💞 => jo\n")
        self.assertEqual([n for n, _ in mudae.lines(self.tmp / "harem.txt")], ["Dobby"])


class Especials(Base):
    def test_crear_amb_el_que_ja_sabem(self):
        self.write("harem.txt", "Eevee\n")
        self.write("series/Pokedex.txt", "Vaporeon 💞 => pep\nJolteon\n")
        mudae.save_special("Eevees", "Eevee\nVaporeon $ Jolteon\n")
        self.assertEqual(mudae.specials(), ["Eevees"])
        got = [(r["name"], r["owner"]) for r in mudae.read_series() if r["serie"] == "Eevees"]
        self.assertEqual(got, [("Eevee", "jo"), ("Vaporeon", "pep"), ("Jolteon", "")])

    def test_posar_al_dia_nomes_agafa_els_del_grup(self):
        mudae.save_special("Eevees", "Eevee\nVaporeon\nJolteon\n")
        paste = "1. Pokédex 3/1025\n\nPokémon\n2 más - Escriba $ima? para mostrar todos los alias.\n \n\nBulbasaur => anna\nEevee => pep\nVaporeon => jo\n"
        mudae.save_paste("serie", "Eevees", paste)
        got = [(r["name"], r["owner"]) for r in mudae.read_series()]
        self.assertEqual(got, [("Eevee", "pep"), ("Vaporeon", "jo"), ("Jolteon", "")])  # sense Bulbasaur
        with self.assertRaises(ValueError):
            mudae.save_paste("serie", "Eevees", "Pikachu => anna\n")

    def test_no_trepitja_una_serie(self):
        self.write("series/Pingu.txt", "Pingu\n")
        with self.assertRaises(ValueError):
            mudae.save_special("Pingu", "Pinga\n")
        self.assertEqual(mudae.specials(), [])


class Idiomes(unittest.TestCase):
    def test_tots_els_textos_traduits(self):
        sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tools"))
        import re
        import textos
        from i18n import STRINGS
        for lang, table in STRINGS.items():
            self.assertEqual(sorted(k for k in textos.keys() if k not in table), [], f"falten textos en {lang}")
            for k, v in table.items():  # les variables {n}, {name}… han de ser les mateixes
                self.assertEqual(set(re.findall(r"\{\w+\}", k)), set(re.findall(r"\{\w+\}", v)), f"{lang}: {k}")

    def test_excel_en_qualsevol_idioma(self):
        self.assertEqual(mudae.canon("High"), "Alta")
        self.assertEqual(mudae.canon("Pendiente"), "Pendent")
        self.assertEqual(mudae.canon("Alta"), "Alta")


class Seguretat(unittest.TestCase):
    def test_excel_no_executa_noms_com_a_formules(self):
        from openpyxl import Workbook
        ws = Workbook().active
        self.assertEqual(mudae.put(ws, 1, 1, "=HYPERLINK(\"http://x\")").data_type, "s")
        self.assertEqual(mudae.put(ws, 2, 1, "Harry Potter").data_type, "s")

    def test_servidor_nomes_serveix_el_tauler_i_a_127(self):
        import threading
        import urllib.error
        import urllib.request
        import webbrowser
        webbrowser.open = lambda *a, **k: None
        port = 8789
        threading.Thread(target=mudae.serve, kwargs={"port": port}, daemon=True).start()
        base = f"http://127.0.0.1:{port}"
        for _ in range(50):
            try:
                urllib.request.urlopen(base + "/ping", timeout=1)
                break
            except OSError:
                import time
                time.sleep(0.1)

        def status(path, host=None):
            req = urllib.request.Request(base + path, headers={"Host": host} if host else {})
            try:
                return urllib.request.urlopen(req, timeout=3).status
            except urllib.error.HTTPError as err:
                return err.code
        self.assertEqual(status("/ping"), 200)
        self.assertEqual(status("/dades/config.json"), 404)       # les dades no es serveixen
        self.assertEqual(status("/mudae.py"), 404)
        self.assertEqual(status("/"), 404)                        # sense llistat de la carpeta
        self.assertEqual(status("/ping", host="evil.example:8789"), 403)  # DNS rebinding
        quit_req = urllib.request.Request(base + "/quit", data=b"{}", headers={"Content-Type": "application/json"})
        urllib.request.urlopen(quit_req, timeout=3)


class Ordres(unittest.TestCase):
    def test_botons_de_copiar(self):
        self.assertIn('data-copy="$oh"', mudae.cmd_html("$oh (×4)"))
        self.assertIn('data-copy="$wishremove "', mudae.cmd_html("$wishremove <nom>"))
        self.assertIn('data-copy="$build 8"', mudae.cmd_html("$build 8"))
        self.assertNotIn("data-copy", mudae.cmd_html("clic al cor"))


if __name__ == "__main__":
    unittest.main()
