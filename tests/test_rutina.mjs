// Proves de «quan es pot tornar a fer» de la llista del tauler. Executa: node tests/test_rutina.mjs
import { readFileSync } from 'node:fs';
import assert from 'node:assert/strict';

const py = readFileSync(new URL('../mudae.py', import.meta.url), 'utf8');
const js = py.match(/JS_TODAY = """([\s\S]*?)"""/)[1];
// Les hores del servidor no se suposen: surten del $tu (minut de les tirades, pròxim reclam) i del $settings (cada quant hi ha reclam)
const fn = name => js.match(new RegExp(`function ${name}[\\s\\S]*?\\n}\\n`))[0].replaceAll('\\\\', '\\');
const consts = js.match(/const MIN = [^\n]*/)[0] + '\n' + js.match(/const lastClaim = [^\n]*/)[0] + '\n' + js.match(/const tuMins = [^\n]*/)[0].replaceAll('\\\\', '\\');
const T = new Function('RS', 'window', [consts, ...['nextRoll', 'nextClaim', 'until', 'readResets', 'learnResets', 'readSettings'].map(fn)].join('\n')
  + 'return { nextRoll, nextClaim, lastClaim, until, readResets, learnResets, readSettings };');
const at = (d, h, m, s = 0) => +new Date(2026, 9, d, h, m, s);  // hora local, 29/10/2026

// sense res enganxat, no sabem res: les tasques marcades queden marcades fins que enganxes el $tu
let t = T({}, {});
assert.equal(t.until('hora', at(29, 14, 30)), Infinity);
assert.equal(t.until('reclam', at(29, 14, 30)), Infinity);
assert.equal(t.until('rt', at(29, 14, 30)), Infinity);                 // el $rt, sense $bonus
assert.equal(t.until('20', at(29, 9, 0)), at(30, 5, 0));               // $daily / $dk: 20 h després (de Mudae, igual a tots els servidors)
const utc = new Date(t.until('utc', at(29, 14, 30)));
assert.equal(utc.getUTCHours(), 0); assert.equal(utc.getUTCDate(), 30); // $oh: mitjanit UTC
// un $tu enganxat a les 14:43:40 que diu «tirades en 18 min» i «reclam en 2 h 18 min»: reinici al minut 02, reclam a les 17:02
const R = t.learnResets({}, { roll: 18, claim: 138 }, at(29, 14, 43, 40));
assert.deepEqual(R, { m: 2, next: at(29, 17, 2) });
assert.deepEqual(t.learnResets({}, { roll: 18, claim: 0 }, at(29, 14, 44, 10)), { m: 2 });  // Mudae arrodoneix: el mig minut ho aguanta
t = T(R, { RT_H: 20 });
assert.equal(t.until('hora', at(29, 14, 50)), at(29, 15, 2));
assert.equal(t.until('hora', at(29, 15, 2, 30)), at(29, 16, 2));         // just després del reinici: la següent
assert.equal(t.until('hora', at(29, 23, 30)), at(30, 0, 2));             // canvi de dia
assert.equal(t.until('reclam', at(29, 16, 5)), at(29, 17, 2));
assert.equal(t.until('reclam', at(29, 17, 30)), Infinity);              // passat el reclam i sense $settings: no sabem el següent
assert.equal(t.lastClaim(at(29, 17, 30)), at(29, 17, 2));                // però sí que fa poc que es pot reclamar
assert.equal(t.until('rt', at(29, 9, 0)), at(30, 5, 0));                // el $rt, amb les 20 h del $bonus
// amb el $settings (reclam cada 180 min), els reclams van de 3 en 3 hores des del que deia el $tu, endavant i enrere
t = T({ ...R, iv: 180 }, {});
assert.equal(t.until('reclam', at(29, 17, 30)), at(29, 20, 2));
assert.equal(t.until('reclam', at(29, 23, 0)), at(29, 23, 2));
assert.equal(t.until('reclam', at(29, 23, 30)), at(30, 2, 2));           // passa a l'endemà
assert.equal(t.until('reclam', at(29, 9, 0)), at(29, 11, 2));            // abans del $tu també
assert.equal(t.lastClaim(at(29, 9, 0)), at(29, 8, 2));
// del $tu: d'aquí quant es reinicien el reclam i les tirades
assert.deepEqual(t.readResets('jo, no puedes reclamar hasta dentro de 2h 18 min.\nTienes 0 rolls restantes.\nEl siguiente reinicio será en 18 min.\nSiguiente reinicio de $daily en 19h 52 min.'), { claim: 138, roll: 18 });
assert.deepEqual(t.readResets('jo, you can claim right now! The next claim reset is in 1h 14 min.\nYou have 8 rolls left. Next rolls reset in 18 min.'), { claim: 74, roll: 18 });
// $settings de Victor (01/11/2026): cada quant hi ha reclam i el minut dels reinicis
const settings = `🛠️ Ajustes de Servidor 🛠️
(Server sin premium)

· Prefijo: $ ($prefix)
· Idioma: es ($lang)
· Reinicio de reclamos: Cada 180 min. ($setclaim)
· Minuto exacto de cada reinicio: xx:02 ($setinterval)
· Reinicio cambiado: por +0 min. ($shifthour)
· Rolls por hora: 8 ($setrolls)
· Tiempo antes de que se expire la reacción para reclamar: 45 sec. ($settimer)
· Rolls sniping: 2 ($togglesnipe) => 8 sec.
· Calculación de kakera: Reclamos y ranking de favoritos (y número de personajes reclamados) ($togglekakeraclaim/$togglekakeralike)`;
assert.deepEqual(t.readSettings(settings), { iv: 180, m: 2 });
assert.deepEqual(t.readSettings('Claim reset: every 2h ($setclaim)'), { iv: 120, m: null });  // en anglès: suposat
assert.equal(t.readSettings('Rolls: 10 ($setrolls)'), null);
// El $tu enganxat: minuts fins que cada tasca torna (0 = ja la pots fer)
// JS_TODAY no és r"""…""": al fitxer de Python les barres van doblades (\\d) i al tauler surten soles (\d)
const readTu = new Function((js.match(/const tuMins = [^\n]*/)[0] + '\n' + js.match(/function readTu[\s\S]*?\n}\n/)[0]).replaceAll('\\\\', '\\') + 'return readTu;')();
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
// el Pokéslot ($p): format suposat, sense cap $tu real que en parli
assert.deepEqual(readTu('You can use $p in 1h 10 min.\nYou have 0 rolls left.'), { p: 70, rolls: true, tu: true });
assert.deepEqual(readTu('Tiempo restante antes de su próximo $p: 42 min.'), { p: 42, tu: true });  // el $p en espera (de Victor, 31/10/2026)
// el poder de kakera, per a la notificació de les tirades
const readPower = new Function(js.match(/function readPower[\s\S]*?\n}\n/)[0].replaceAll('\\\\', '\\') + 'return readPower;')();
assert.deepEqual(readPower(tuEs), { p: 110, c: 36, h: 18 });
assert.deepEqual(readPower(tuEs.replace('(18%)', '(15%)')), { p: 110, c: 36, h: 15 });  // el 10+ es llegeix, no es calcula
assert.deepEqual(readPower('Power: 85%\nEach kakera button consumes 36% of your reaction power.'), { p: 85, c: 36, h: 18 });  // sense la línia 10+: la meitat
assert.equal(readPower(tuEn), null);
// el $settings no s'ha de confondre amb el $tu (la casella d'Avui mira primer si és un $tu: cal reclam i tirades)
const tuS = readTu(settings);
assert.ok(!(('claim' in tuS || 'rolls' in tuS) && Object.keys(tuS).length > 2), JSON.stringify(tuS));
console.log('rutina: totes les proves passen');
