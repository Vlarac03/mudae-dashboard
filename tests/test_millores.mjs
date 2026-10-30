// Proves de la pestanya «Millores»: llegir el $k, el $kt i el $opp. Executa: node tests/test_millores.mjs
import { readFileSync } from 'node:fs';
import assert from 'node:assert/strict';

const py = readFileSync(new URL('../mudae.py', import.meta.url), 'utf8');
const js = py.match(/JS_MILL = r"""([\s\S]*?)"""/)[1];
const BADGES = ['Bronze', 'Silver', 'Gold', 'Sapphire', 'Ruby', 'Emerald', 'Diamond'].map(n => [n, 0]);
const num = js.match(/const num = [^\n]*/)[0];
const readMudae = new Function('BADGES', num + '\n' + js.match(/function readMudae[\s\S]*?\n}\n/)[0] + 'return readMudae;')(BADGES);

const k = `¿Cómo coleccionar kakera en el servidor? (Cambie las opciones con $togglekakera)
Tienes 28.223!
Convierte tu kakera en insignias de poder:
 Bronze IV · ¡Máximo alcanzado!
 Silver IV · ¡Máximo alcanzado!
 Gold IV · ¡Máximo alcanzado!
 Sapphire II · 15.000
 Ruby I · 14.000
Kakeratower: $kt
Desbloquea un nivel con $bronze, $silver...`;
assert.deepEqual(readMudae(k, 5000), { b: { Bronze: 4, Silver: 4, Gold: 4, Sapphire: 2, Ruby: 1, Emerald: 0, Diamond: 0 }, k: 28223 });

const kt = `Gasta tupara armar los 12 pisos de la torre:
Tu nivel actual es
El siguiente nivel cuesta 40.000
Tienes 615
Arma un piso usando $build <numero del beneficio>
 [1]  +2 espacios en la lista de deseos`;
assert.deepEqual(readMudae(kt, 5000), { nf: 7, k: 615 });
assert.deepEqual(readMudae(kt.replace('cuesta 40.000', 'costs 80,000'), 10000), { nf: 7, k: 615 });

const opp = `Mejora los beneficios del personaje seleccionado.
Tienes 1.780
Harry Potter
[LVL 2] Probabilidad de aparición incrementada: 0 > 15%
[LVL 0] Valor de kakera base incrementado: 0 > 20
[LVL 0] Probabilidad de obtener +1 botón de kakera: 0 > 7%
[LVL 6] Probabilidad de obtener +1 llave: 0 > 4%
[LVL 0] Cantidad de spheres: 0 > 3
[LVL 0] a
[LVL 0] b
[LVL 0] c
[LVL 0] d
[LVL 1] El primer $oh del día genera +20 spheres`;
assert.deepEqual(readMudae(opp, 5000), { opp: { name: 'Harry Potter', levels: [2, 0, 0, 6, 0, 0, 0, 0, 0, 1] }, sp: 1780 });

assert.deepEqual(readMudae('Hola, això no és de Mudae', 5000), {});
assert.deepEqual(readMudae('Kakera rewards\nBonus Bronze I + Silver I + Gold I\n$bronze', 5000), {});  // el $kakerareward no són els teus nivells
// torre: 12 pisos per torre; el $kt diu el total i d'aquí surt la torre on ets
const fn = name => js.match(new RegExp(`function ${name}[\\s\\S]*?\\n}\\n`))[0];
const TOWER = [11, 8, 7, 6, 5, 1, 4, 2, 9, 10, 3, 12];
const [towerNext, towerFromKt] = new Function('TOWER', fn('towerNext') + fn('towerFromKt').replace(/\n$/, '') + '\nreturn [towerNext, towerFromKt];')(TOWER);
assert.deepEqual(towerNext({ f: [11, 8], tv: 5000 }), { tw: 0, built: 2, cost: 15000, rec: 7 });
assert.deepEqual(towerNext({ t: 1, f: [], tv: 5000 }), { tw: 1, built: 12, cost: 65000, rec: 11 });  // torre 2: torna a començar per l'11
assert.deepEqual(towerFromKt({ f: [11, 8] }, 14), { t: 1, f: [] });  // 14 pisos: torre 2, els marcats eren de la 1
assert.deepEqual(towerFromKt({ t: 1, f: [11, 8] }, 14), { t: 1, f: [11, 8] });
console.log('millores: totes les proves passen');
