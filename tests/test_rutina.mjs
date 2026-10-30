// Proves de «quan es pot tornar a fer» de la llista del tauler. Executa: node tests/test_rutina.mjs
import { readFileSync } from 'node:fs';
import assert from 'node:assert/strict';

const py = readFileSync(new URL('../mudae.py', import.meta.url), 'utf8');
const js = py.match(/JS_TODAY = """([\s\S]*?)"""/)[1];
const until = new Function('CLAIM', js.match(/function until[\s\S]*?\n}\n/)[0] + 'return until;')([1, 4, 7, 10, 13, 16, 19, 22]);
const at = (d, h, m, s = 0) => +new Date(2026, 9, d, h, m, s);  // hora local, 29/10/2026

assert.equal(until('hora', at(29, 14, 30)), at(29, 15, 2));            // marcada a mitja hora: fins a les xx:02
assert.equal(until('hora', at(29, 15, 1)), at(29, 15, 2));
assert.equal(until('hora', at(29, 15, 2, 30)), at(29, 16, 2));         // just després del reinici: la següent
assert.equal(until('hora', at(29, 23, 30)), at(30, 0, 2));             // canvi de dia
assert.equal(until('reclam', at(29, 16, 5)), at(29, 19, 2));
assert.equal(until('reclam', at(29, 23, 0)), at(30, 1, 2));            // l'últim del dia passa a l'endemà
assert.equal(until('20', at(29, 9, 0)), at(30, 5, 0));                 // $daily / $dk: 20 h després
const utc = new Date(until('utc', at(29, 14, 30)));
assert.equal(utc.getUTCHours(), 0); assert.equal(utc.getUTCDate(), 30); // $oh: mitjanit UTC
// El $tu enganxat: minuts fins que cada tasca torna (0 = ja la pots fer)
// JS_TODAY no és r"""…""": al fitxer de Python les barres van doblades (\\d) i al tauler surten soles (\d)
const readTu = new Function(js.match(/function readTu[\s\S]*?\n}\n/)[0].replaceAll('\\\\', '\\') + 'return readTu;')();
const tuEs = `jo, no puedes reclamar hasta dentro de 2h 18 min.
Tienes 0 rolls restantes.
El siguiente reinicio será en 18 min.
Siguiente reinicio de $daily en 19h 52 min.

¡Puedes reaccionar a kakera en este momento!
Poder: 110%
Cada botón de kakera consume 36% de su poder de reacción.
Tus personajes con 10+ llaves, consumen la mitad del poder (18%)
Capital: 4.532:kakera:

¡$rt está disponible!
Siguiente $dk en 19h 52 min.
¡Puedes votar en este momento!

0 $oh restantes para hoy, 0 $oc, 0 $oq y 0 $ot.
14h 16 min hasta la recarga.
Capital: 1.725 :sp:
=> $tuarrange`;
// els minuts que diu Mudae: tirades en 18 min i esferes en 14 h 16 min (cada servidor reinicia a una hora diferent)
assert.deepEqual(readTu(tuEs), { claim: 138, rolls: 18, daily: 1192, kakera: 0, rt: 0, dk: 1192, vote: 0, oh: 856, oc: 856, tu: true });
const tuEn = `jo, you can claim right now! The next claim reset is in 1h 14 min.
You have 8 rolls left. Next rolls reset in 18 min.
$daily is available!
You can't react to kakera for 41 min.
You may vote again in 3h 05 min.
2 $oh left for today, 1 $oc`;
assert.deepEqual(readTu(tuEn), { claim: 0, rolls: 0, daily: 0, kakera: 41, vote: 185, oh: 0, oc: 0, tu: true });
assert.deepEqual(readTu("You have 0 rolls left.\n$rt is not ready"), { rolls: true, rt: true, tu: true });  // sense minuts: com si ho marquessis a mà
assert.deepEqual(readTu('Hola'), {});
console.log('rutina: totes les proves passen');
