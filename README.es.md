<p align="center"><img src="assets/icona-256.png" alt="" width="72"></p>

<h1 align="center">Panel de Mudae</h1>

<p align="center">Lleva tus colecciones de <a href="https://mudae.net">Mudae</a> al día, sin perder ninguna tirada.<br>
<b>Descárgalo:</b> botón verde <b>Code → Download ZIP</b>, arriba de esta página.<br>
<a href="README.md">English</a> · <a href="README.ca.md">Català</a> · <b>Español</b></p>

![El panel: cifras de la colección, qué toca ahora y el ayudante del $oc](assets/tauler.png)

Un panel que funciona en tu ordenador para el bot de Discord Mudae. Pegas lo que te envía Mudae y te lleva las colecciones, la wishlist y los intercambios, te dice qué toca hacer cada hora, te ayuda con el juego de esferas `$oc` y te mantiene al día un Excel. En español, catalán e inglés.

## Qué hace

- **Qué toca ahora**: reloj de tiradas y reclamos, y una lista donde cada tarea queda marcada hasta que la puedes volver a hacer (`$daily` y `$dk` cada 20 h, `$vote` cada 12 h, `$oh` a medianoche UTC, tiradas cada hora…).
- **Notificaciones**: mientras el panel está abierto, te avisa cuando vuelves a tener tiradas, y te dice si puedes reclamar y reaccionar a kakera. Se activan y se desactivan con el botón **Notificaciones** de arriba.
- **Ayudante del `$oc`**: marcas los colores que salen y te dice en qué casilla hacer clic para encontrar la roja.
- **Colecciones**: un álbum por serie con tus personajes, los libres, los de la wishlist y los de otros jugadores (y de quién son). Haces clic en los que consigues para marcarlos como tuyos.
- **Colecciones especiales**: un grupo que eliges tú, aunque no sea una serie de Mudae (por ejemplo, los Eevee de la Pokédex). Escribes el nombre y los personajes, y el `$imaso` de la serie de la que salen te pone de quién es cada uno.
- **Mejoras**: pegas el `$k`, el `$kt`, un `$opp <personaje>`, el `$os` o el `$mmy=` y te dice qué insignia, qué piso de la torre, qué mejora de esferas y qué mejora de la tienda de esferas te toca, y cuánta kakera te falta.
- **Wishlist**: te avisa de los que ya son tuyos o que ha reclamado alguien, y te prepara los comandos `$wishremove` y `$wish`.
- **Intercambios y regalos**: el comando `$give` para cada jugador, y un buscador que entiende listas de series mal escritas («full metal alchimist», «Rezero»…).
- **Poner al día**: pegas el `$mma+s` (tu harem, con la serie de cada personaje), el `$wl` o un `$imaso` y haces clic en un botón. Te avisa si hace más de una semana que no pones al día el harem.
- **Comandos de uso habitual**: todos, con lo que hacen y un botón de copiar.
- **Excel**: lo mismo en una hoja de cálculo, que respeta lo que escribes en ella (prioridades, notas, estado de los intercambios).

| Colecciones | Buscador para regalos |
|---|---|
| ![Álbum de una colección](assets/col-leccions.png) | ![Buscador de series](assets/cercador.png) |

## Instalación (5 minutos)

1. Instala [Python](https://www.python.org/downloads/) 3.12 o más nuevo. En la primera pantalla del instalador, marca «Add python.exe to PATH» (recomendado, aunque sin él también funciona).
2. Haz clic en el botón verde **Code → Download ZIP** arriba de esta página. Antes de descomprimirlo: clic derecho al ZIP → **Propiedades** → marca **«Desbloquear»** → Aceptar (si no, Windows puede bloquear los `.bat`). Después descomprime la carpeta donde la quieras tener (por ejemplo, en Documentos). No la abras desde dentro del zip.
3. Doble clic en **`Configura.bat`**. Elige el idioma (el 2 es el español), escribe tu nombre en Mudae (el que sale en «Pertenece a …» en tus personajes) y, si quieres, el del servidor. Crea el acceso directo **Mudae** en el escritorio y abre el panel.
4. En el panel, ve a **«Poner al día»** y pega, por orden:
   - **Tu harem**: escribe `$mma+s` en Discord, copia el mensaje que te llega por privado y pégalo eligiendo «Tu harem».
   - **Tu wishlist**: `$wl`, eligiendo «La wishlist».
   - **Cada serie que coleccionas**: `$imaso <serie>` (por ejemplo, `$imaso Harry Potter`), eligiendo «Una serie» → «Serie nueva…» y escribiendo su nombre tal como sale en Mudae.
   - **Después, todas de golpe**: elige «Todas tus series», copia el comando que sale (`$imaa+os Serie 1 $ Serie 2 $ …`) y pega todos los mensajes que te devuelve Mudae. No hace falta elegir qué pegas: el panel lo adivina.

A partir de aquí, abre el panel con el acceso directo **Mudae** del escritorio.

## Cada día

- **Qué toca ahora** te dice qué hacer y cuándo. Ve marcando lo que haces.
- **Cuando consigas un personaje**: en «Colecciones», elige la serie, haz clic en él y en **Ya son míos**.
- **Cuando regales**: en «Buscar entre los míos», búscalos, copia el `$give` y, cuando Mudae lo confirme, haz clic en **Ya los he dado**.
- **Una vez por semana**: vuelve a pegar el `$mma+s` en «Poner al día» (el panel te avisa).

## Bueno saber

- **Cierra el Excel** antes de usar los botones del panel: si está abierto, no se puede guardar.
- Los botones usan un servidor pequeño que solo escucha en tu ordenador (`127.0.0.1`) y solo acepta peticiones del panel. Se cierra solo un par de minutos después de cerrar la pestaña, o con el botón **Cerrar el servidor**.
- Puedes pegar los mensajes tal cual: las cabeceras de Discord, «Page 1 / 21» y la numeración se ignoran.
- **No borres la carpeta `lib`**: tiene lo que necesita el programa para escribir el Excel.

## Tus datos

El panel no se conecta a Discord ni a Mudae y no pide ninguna cuenta. Solo lee lo que pegas y lo guarda en la carpeta `dades/`, en tu ordenador.

## Configuración (`dades/config.json`)

| Clave | Qué es |
|---|---|
| `usuari` | Tu nombre en Mudae. Lo pone `Configura.bat`. |
| `idioma` | `es`, `ca` o `en`. Para cambiarlo, vuelve a abrir `Configura.bat`. |
| `servidor` | Opcional: sale en el título del panel y del Excel. |
| `discord_id` | Opcional: el número que te da `$myid`, para que los comandos `$give` que copias te mencionen bien. |
| `hores_reclam` | Opcional: las horas en que se puede reclamar en vuestro servidor. Por defecto, `[1, 4, 7, 10, 13, 16, 19, 22]` (cada 3 horas, en el minuto 02). |
| `rt_hores` | Opcional: cada cuánto puedes usar `$rt`. Por defecto, `50`; quita 10 horas por cada insignia Esmeralda II, III y IV (`20` con las tres). |
| `preferits` | Opcional: personajes que el panel te propone primero cuando se libera un hueco en la wishlist, por ejemplo `["Fawkes", "Howl"]`. |

## Si algo no va

- **«Could not update»**: cierra el Excel y vuelve a abrir Mudae.
- **Un botón dice «Sin conexión» o «Botón apagado»**: abre el panel con el acceso directo Mudae, no abriendo `dashboard.html` directamente.
- **Windows dice «Control Inteligente de Aplicaciones ha bloqueado un archivo»** (o SmartScreen avisa) al abrir un `.bat`: borra la carpeta, haz clic derecho al ZIP → Propiedades → **Desbloquear** y vuelve a descomprimirlo (paso 2).
- **«Python 3.12 or newer is needed»** al abrir `Configura.bat`: instala Python (paso 1) y vuelve a probar.

## Cómo está hecho

- `mudae.py`: todo el programa, con la librería estándar de Python y openpyxl. Lee `dades/`, escribe el Excel y `dashboard.html`, y hace de servidor local.
- `i18n.py`: los textos en español e inglés (el catalán hace de clave).
- `oc.html`: el ayudante del `$oc`, sin dependencias.
- `lib/`: openpyxl y et_xmlfile, incluidos para que funcione sin `pip install`.
- `Configura.bat` / `Actualitza.bat`: la primera configuración y el acceso directo.

Pruebas:

```bash
python -m unittest discover tests
node tests/test_oc.mjs
node tests/test_rutina.mjs
node tests/test_millores.mjs
```

Textos sin traducir: `python tools/textos.py`. Un `.zip` limpio para pasárselo a alguien (sin datos personales): `python tools/fer_zip.py`.

## Licencia

[MIT](LICENSE) © 2026 Victor Lara (Vlarac).

Proyecto de fans: no está afiliado con Mudae ni con Discord, ni cuenta con su apoyo.

<p align="center"><sub>Developed by <a href="https://github.com/Vlarac03">Vlarac</a></sub></p>
