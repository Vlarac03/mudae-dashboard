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
console.log('rutina: totes les proves passen');
