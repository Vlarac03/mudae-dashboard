<p align="center"><img src="assets/icona-256.png" alt="" width="72"></p>

<h1 align="center">Tauler de Mudae</h1>

<p align="center">Porta les teves col·leccions de <a href="https://mudae.net">Mudae</a> al dia, sense perdre cap tirada.<br>
<b>Descarrega'l:</b> botó verd <b>Code → Download ZIP</b>, a dalt d'aquesta pàgina.<br>
<a href="README.md">English</a> · <b>Català</b> · <a href="README.es.md">Español</a></p>

![El tauler: xifres de la col·lecció, què toca ara i l'ajudant del $oc](assets/tauler.png)

Un tauler que funciona al teu ordinador per al bot de Discord Mudae. Hi enganxes el que t'envia Mudae i et porta les col·leccions, la wishlist i els intercanvis, et diu què toca fer a cada hora, t'ajuda amb el joc d'esferes `$oc` i et manté al dia un Excel. En català, castellà i anglès.

## Què fa

- **Què toca ara**: rellotge de tirades i reclams, i una llista on cada tasca queda marcada fins que la pots tornar a fer (`$daily` i `$dk` cada 20 h, `$vote` cada 12 h, `$oh` a mitjanit UTC, tirades cada hora…).
- **Notificacions**: mentre el tauler és obert, t'avisa quan tornes a tenir tirades, i et diu si pots reclamar i reaccionar a kakera. Fan un so i es queden a la pantalla fins que les tanques. S'activen i es desactiven amb el botó **Notificacions** de dalt.
- **Ajudant del `$oc`**: marques els colors que surten i et diu quina casella clicar per trobar la vermella.
- **Col·leccions**: un àlbum per sèrie amb els teus personatges, els lliures, els de la wishlist i els d'altres jugadors (i de qui són). Cliques els que aconsegueixes per marcar-los com a teus.
- **Col·leccions especials**: un grup que tries tu, encara que no sigui una sèrie de Mudae (per exemple, els Eevee de la Pokédex). Escrius el nom i els personatges, i el `$imaso` de la sèrie d'on surten t'hi posa de qui és cada un.
- **Millores**: enganxes el `$k`, el `$kt`, un `$opp <personatge>`, el `$os` o el `$mmy=` i et diu quina insígnia, quin pis de la torre, quina millora d'esferes i quina de la botiga d'esferes et toca, i quanta kakera et falta.
- **Wishlist**: t'avisa dels que ja són teus o que ha reclamat algú, i et prepara les ordres `$wishremove` i `$wish`.
- **Intercanvis i regals**: l'ordre `$give` per a cada jugador, i un cercador que entén llistes de sèries mal escrites («full metal alchimist», «Rezero»…).
- **Posar al dia**: enganxes el `$mma+s` (el teu harem, amb la sèrie de cada personatge), el `$wl` o un `$imaso` i cliques un botó. T'avisa si fa més d'una setmana que no poses al dia el harem.
- **Ordres d'ús habitual**: totes, amb el que fan i un botó de copiar.
- **Excel**: el mateix en un full de càlcul, que respecta el que hi escrius (prioritats, notes, estat dels intercanvis).

| Col·leccions | Cercador per a regals |
|---|---|
| ![Àlbum d'una col·lecció](assets/col-leccions.png) | ![Cercador de sèries](assets/cercador.png) |

## Instal·lació (5 minuts)

1. Instal·la [Python](https://www.python.org/downloads/) 3.12 o més nou. A la primera pantalla de l'instal·lador, marca «Add python.exe to PATH» (recomanat, tot i que sense també funciona).
2. Clica el botó verd **Code → Download ZIP** a dalt d'aquesta pàgina. Abans de descomprimir-lo: clic dret al ZIP → **Propietats** → marca **«Desbloqueja»** → D'acord (si no, Windows pot bloquejar els `.bat`). Després descomprimeix la carpeta on la vulguis tenir (per exemple, a Documents). No l'obris des de dins del zip.
3. Doble clic a **`Configura.bat`**. Tria l'idioma (el 3 és el català), escriu el teu nom a Mudae (el que surt a «Pertenece a …» als teus personatges) i, si vols, el del servidor. Crea l'accés directe **Mudae** a l'escriptori i obre el tauler.
4. Al tauler, ves a **«Posar al dia»** i enganxa-hi, per ordre:
   - **El teu harem**: escriu `$mma+s` a Discord, copia el missatge que t'arriba per privat i enganxa'l triant «El teu harem».
   - **La teva wishlist**: `$wl`, triant «La wishlist».
   - **Cada sèrie que col·lecciones**: `$imaso <sèrie>` (per exemple, `$imaso Harry Potter`), triant «Una sèrie» → «Sèrie nova…» i escrivint-ne el nom tal com surt a Mudae.
   - **Després, totes de cop**: tria «Totes les teves sèries», copia l'ordre que surt (`$imaa+os Sèrie 1 $ Sèrie 2 $ …`) i enganxa tots els missatges que et torna Mudae. No cal triar què enganxes: el tauler ho endevina.

A partir d'aquí, obre el tauler amb l'accés directe **Mudae** de l'escriptori.

## Cada dia

- **Què toca ara** et diu què fer i quan. Ves marcant el que fas.
- **Quan aconsegueixis un personatge**: a «Col·leccions», tria la sèrie, clica'l i clica **Ja són meus**.
- **Quan regalis**: a «Buscar entre els meus», busca'ls, copia el `$give` i, quan Mudae ho confirmi, clica **Ja els he donat**.
- **Un cop per setmana**: torna a enganxar el `$mma+s` a «Posar al dia» (el tauler t'avisa).

## Bo de saber

- **Tanca l'Excel** abans de fer servir els botons del tauler: si és obert, no es pot desar.
- Els botons fan servir un servidor petit que només escolta al teu ordinador (`127.0.0.1`) i només accepta peticions del tauler. Es tanca sol un parell de minuts després de tancar la pestanya, o amb el botó **Tancar el servidor**.
- Pots enganxar els missatges tal qual: les capçaleres de Discord, «Page 1 / 21» i la numeració s'ignoren.
- **Versions noves**: el botó **Actualitzar el tauler** mira si n'hi ha una a GitHub i, si li dius que sí, la posa sola. La carpeta `dades/` no es toca. (Si tens una versió d'abans d'aquest botó, baixa el ZIP un cop i copia'n els fitxers damunt dels teus.)
- **No esborris la carpeta `lib`**: hi ha el que necessita el programa per escriure l'Excel.

## Les teves dades

El tauler no es connecta a Discord ni a Mudae i no demana cap compte. Només llegeix el que hi enganxes i ho desa a la carpeta `dades/`, al teu ordinador.

## Configuració (`dades/config.json`)

| Clau | Què és |
|---|---|
| `usuari` | El teu nom a Mudae. El posa `Configura.bat`. |
| `idioma` | `ca`, `es` o `en`. Per canviar-lo, torna a obrir `Configura.bat`. |
| `servidor` | Opcional: surt al títol del tauler i de l'Excel. |
| `discord_id` | Opcional: el número que et dona `$myid`, perquè les ordres `$give` que copies et mencionin bé. |
| `hores_reclam` | Opcional: les hores en què es pot reclamar al vostre servidor. Per defecte, `[1, 4, 7, 10, 13, 16, 19, 22]` (cada 3 hores, al minut 02). |
| `rt_hores` | Opcional: cada quant pots fer servir `$rt`. Per defecte, `50`; treu-ne 10 hores per cada insígnia Esmeralda II, III i IV (`20` amb totes tres). |
| `preferits` | Opcional: personatges que el tauler et proposa primer quan s'allibera un lloc a la wishlist, per exemple `["Fawkes", "Howl"]`. |

## Si alguna cosa no va

- **«Could not update»**: tanca l'Excel i torna a obrir Mudae.
- **Un botó diu «Sense connexió» o «Botó apagat»**: obre el tauler amb l'accés directe Mudae, no obrint `dashboard.html` directament.
- **Windows diu «Control intel·ligent d'aplicacions ha bloquejat un fitxer»** (o SmartScreen avisa) en obrir un `.bat`: esborra la carpeta, clic dret al ZIP → Propietats → **Desbloqueja** i torna'l a descomprimir (pas 2).
- **«Python 3.12 or newer is needed»** en obrir `Configura.bat`: instal·la Python (pas 1) i torna-ho a provar.

## Com està fet

- `mudae.py`: tot el programa, amb la llibreria estàndard de Python i openpyxl. Llegeix `dades/`, escriu l'Excel i `dashboard.html`, i fa de servidor local.
- `i18n.py`: els textos en castellà i anglès (el català fa de clau).
- `oc.html`: l'ajudant del `$oc`, sense dependències.
- `lib/`: openpyxl i et_xmlfile, inclosos perquè funcioni sense `pip install`.
- `Configura.bat` / `Actualitza.bat`: la primera configuració i l'accés directe.

Proves:

```bash
python -m unittest discover tests
node tests/test_oc.mjs
node tests/test_rutina.mjs
node tests/test_millores.mjs
```

Textos sense traduir: `python tools/textos.py`. Un `.zip` net per passar-lo a algú (sense cap dada personal): `python tools/fer_zip.py`.

## Llicència

[MIT](LICENSE) © 2026 Victor Lara (Vlarac).

Projecte de fans: no està afiliat amb Mudae ni amb Discord, ni té el seu suport.

<p align="center"><sub>Developed by <a href="https://github.com/Vlarac03">Vlarac</a></sub></p>
