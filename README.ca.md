<p align="center"><img src="assets/icona-256.png" alt="" width="72"></p>

<h1 align="center">Tauler de Mudae</h1>

<p align="center">Porta les teves col·leccions de <a href="https://mudae.net">Mudae</a> al dia, sense perdre cap tirada.<br>
<b>Descarrega'l:</b> botó verd <b>Code → Download ZIP</b>, a dalt d'aquesta pàgina.<br>
<a href="README.md">English</a> · <b>Català</b> · <a href="README.es.md">Español</a></p>

![La pestanya Avui: xifres de la col·lecció, el rellotge de tirades i reclam, la casella per enganxar amb el requadre Pendent!, i què toca ara](assets/tauler.png)

Un tauler que funciona al teu ordinador per al bot de Discord Mudae. Hi enganxes el que t'envia Mudae i et porta les col·leccions, la wishlist i els intercanvis, et diu què toca fer a cada hora, t'ajuda amb el joc d'esferes `$oc` i et manté al dia un Excel. En català, castellà i anglès.

## Què fa

- **Pestanyes per ordre d'ús**: Avui, Col·leccions, Wishlist, Intercanvis, Millores i Posar al dia. El rellotge de tirades i reclam es veu des de totes, els avisos et porten a la pestanya on s'arreglen, i els blocs llargs es pleguen i es queden oberts o tancats com els vas deixar.
- **Una sola casella per a tot**: a dalt de la pestanya Avui hi enganxes qualsevol missatge de Mudae (`$tu`, `$settings`, `$bonus`, `$k`, `$kt`, `$opp`, `$opz`, `$os`, `$mmy=`, `$mma+s`, `$wl`, `$imaso`…) i el tauler el llegeix així que l'enganxes, sense cap botó, i l'envia on toca. A sota hi ha un botó per copiar cada ordre que entén i un requadre **Pendent!** amb les ordres que encara li falten: no suposa res del que depèn del teu servidor o de tu (hores de tirades i reclams, hores del `$rt`, llocs de la wishlist, preus de les insígnies i de la torre). El harem, la wishlist i les sèries es desen de seguida; una sèrie nova o una col·lecció especial t'espera a «Posar al dia» perquè la comprovis.
- **Què toca ara**: enganxes el `$tu` i la llista es marca sola, Pokéslot (`$p`) inclòs. Funciona com una cua: a «Ara» només hi ha el que pots fer, primer el més important (el reclam a dalt quan pots reclamar, després les tirades i la kakera de cada hora, el que és un cop al dia i els comodins que guardes per quan calguin: `$rt`, `$dk`), el que ja has fet s'amaga fins que el pots tornar a fer, i un requadre «Estat» et diu sí o no del reclam, les tirades, la kakera (amb quantes reaccions et dona el poder), el Pokéslot, el `$rt` i el `$dk`, i quan torna cada cosa (`$daily` i `$dk` cada 20 h, `$vote` cada 12 h, `$oh` a mitjanit UTC, tirades al minut del teu servidor…). El minut de les tirades i el pròxim reclam surten del `$tu`, cada quant hi ha reclam del `$settings` i les hores del `$rt` del `$bonus`.
- **Notificacions**: mentre el tauler és obert, t'avisa quan tornes a tenir tirades (i et diu si pots reclamar i reaccionar a kakera), i quan tornes a tenir el reclam, el `$daily`, el `$vote`, el `$dk`, el `$rt` o el `$p`. Fan un so i es queden a la pantalla fins que les tanques. S'activen i es desactiven amb el botó **Notificacions** de dalt.
- **Ajudant del `$oc`**: marques els colors que surten i et diu quina casella clicar per trobar la vermella. És plegat a la pestanya Avui: s'obre amb un clic.
- **Col·leccions**: un àlbum per sèrie amb els teus personatges, els lliures, els de la wishlist i els d'altres jugadors (i de qui són). Cliques els que aconsegueixes per marcar-los com a teus.
- **Col·leccions especials**: un grup que tries tu, encara que no sigui una sèrie de Mudae (per exemple, els Eevee de la Pokédex). Escrius el nom i els personatges, i el `$imaso` de la sèrie d'on surten t'hi posa de qui és cada un.
- **Pokédex**: enganxes el `$pd sumrare` (un sol missatge) i veus quants Pokémon tens de cada raresa. Si enganxes el missatge del temps que falta per al pròxim `$p`, es marca la tasca del Pokéslot.
- **Millores**: enganxes el `$k`, el `$kt`, un `$opp <personatge>`, el `$os`, el `$mmy=` o el `$bonus` i et diu quina insígnia, quin pis de la torre, quina millora d'esferes i quina de la botiga d'esferes et toca, i quanta kakera et falta. El `$bonus` hi afegeix un resum: quants botons de kakera fas amb el poder ple, les esferes del primer `$oh`, les hores del `$rt` i els llocs de la wishlist. Els preus només surten del que enganxes (`$k`, `$kt`), perquè cada servidor té els seus. Per a les esferes: primer la millora 10 a tots els personatges que no regalaràs (fins a 100), les 1-9 només als que et surten sovint, un pla amb el `$or` quan surt a compte, i et diu quan has de parar i passar a la botiga d'esferes. La seva casella funciona com la d'Avui.
- **Wishlist**: t'avisa dels que ja són teus o que ha reclamat algú, i et prepara les ordres `$wishremove` i `$wish`.
- **Intercanvis i regals**: l'ordre `$give` per a cada jugador, i un cercador que entén llistes de sèries mal escrites («full metal alchimist», «Rezero»…).
- **Posar al dia**: enganxes el `$mma+s` (el teu harem, amb la sèrie de cada personatge), el `$wl` o un `$imaso` i es desa així que l'enganxes (una sèrie nova o una col·lecció especial t'espera perquè la comprovis i cliquis el botó). T'avisa si fa més d'una setmana que no poses al dia el harem.
- **Ordres d'ús habitual**: per temes (tirades, reclamar, la teva col·lecció, harem i perfil, intercanvis, esferes, torre, kakera, Pokéslot…), cada grup plegat, amb un cercador («wish», «kakera»…), el que fa cada ordre i un botó de copiar.
- **Excel**: el mateix en un full de càlcul, que respecta el que hi escrius (prioritats, notes, estat dels intercanvis).

| Col·leccions | Intercanvis i regals |
|---|---|
| ![Àlbum d'una col·lecció](assets/col-leccions.png) | ![A qui demanar i el cercador per a regals](assets/cercador.png) |

## Instal·lació (5 minuts)

1. Instal·la [Python](https://www.python.org/downloads/) 3.12 o més nou. A la primera pantalla de l'instal·lador, marca «Add python.exe to PATH» (recomanat, tot i que sense també funciona).
2. Clica el botó verd **Code → Download ZIP** a dalt d'aquesta pàgina. Abans de descomprimir-lo: clic dret al ZIP → **Propietats** → marca **«Desbloqueja»** → D'acord (si no, Windows pot bloquejar els `.bat`). Després descomprimeix la carpeta on la vulguis tenir (per exemple, a Documents). No l'obris des de dins del zip.
3. Doble clic a **`Configura.bat`**. Tria l'idioma (el 3 és el català), escriu el teu nom a Mudae (el que surt a «Pertenece a …» als teus personatges) i, si vols, el del servidor. Crea l'accés directe **Mudae** a l'escriptori i obre el tauler.
4. Al tauler, enganxa el que t'envia Mudae a la casella de dalt de la pestanya **Avui**. Endevina què és cada missatge i el desa sol. Per ordre:
   - **El teu harem**: escriu `$mma+s` a Discord, copia el missatge que t'arriba per privat i enganxa'l.
   - **La teva wishlist**: `$wl`.
   - **Cada sèrie que col·lecciones**: `$imaso <sèrie>` (per exemple, `$imaso Harry Potter`). Una sèrie nova t'espera: comprova'n el nom a «Posar al dia» i clica **Desar i actualitzar**.
   - **Després, totes de cop**: a «Posar al dia», tria «Totes les teves sèries», copia l'ordre que surt (`$imaa+os Sèrie 1 $ Sèrie 2 $ …`) i enganxa tots els missatges que et torna Mudae.
   - **El teu servidor i tu**: el requadre **Pendent!**, sota la casella, diu quines ordres falten (`$tu`, `$settings`, `$bonus`, `$k`…). Enganxa-les igual.

A partir d'aquí, obre el tauler amb l'accés directe **Mudae** de l'escriptori.

## Cada dia

- **Què toca ara** et diu què fer i quan. Ves marcant el que fas.
- **Quan aconsegueixis un personatge**: a «Col·leccions», tria la sèrie, clica'l i clica **Ja són meus**.
- **Quan regalis**: a «Intercanvis» → «Buscar entre els meus», busca'ls, copia el `$give` i, quan Mudae ho confirmi, clica **Ja els he donat**.
- **Un cop per setmana**: torna a enganxar el `$mma+s` a la casella d'Avui (el tauler t'avisa).

## Bo de saber

- **Tanca l'Excel** abans de fer servir els botons del tauler: si és obert, no es pot desar.
- Els botons fan servir un servidor petit que només escolta al teu ordinador (`127.0.0.1`) i només accepta peticions del tauler. Es tanca sol un parell de minuts després de tancar la pestanya, o amb **Tancar el servidor**, al menú **···** de dalt.
- Pots enganxar els missatges tal qual: les capçaleres de Discord, «Page 1 / 21» i la numeració s'ignoren.
- **Versions noves**: el botó **Actualitzar el tauler** mira si n'hi ha una a GitHub i, si li dius que sí, la posa sola. La carpeta `dades/` no es toca. (Si tens una versió d'abans d'aquest botó, baixa el ZIP un cop i copia'n els fitxers damunt dels teus.)
- **Has perdut l'accés directe?** Al menú **···** de dalt hi ha **Crear accés directe a l'escriptori**.
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
