// Proves de l'ajudant del $oc amb dues partides de debò. Executa: node tests/test_oc.mjs
import { readFileSync } from 'node:fs';
import assert from 'node:assert/strict';

const html = readFileSync(new URL('../oc.html', import.meta.url), 'utf8');
const src = html.match(/<script id="logic">([\s\S]*?)<\/script>/)[1];
const mod = { exports: {} };
new Function('module', src)(mod);
const { candidates, best } = mod.exports;
const at = (r, c) => (r - 1) * 5 + (c - 1);
const sorted = a => [...a].sort((x, y) => x - y);

// Partida 1: blava f2c2, blava f3c4, taronja f5c1 -> la vermella és a f4c1
let b = Array(25).fill(null);
b[at(2, 2)] = 'blue';
assert.equal(candidates(b).length, 10);
assert.equal(best(b).pick, at(3, 4));
b[at(3, 4)] = 'blue';
assert.deepEqual(sorted(candidates(b)), sorted([at(1, 5), at(4, 1), at(5, 1), at(5, 3)]));
b[at(5, 1)] = 'orange';
assert.deepEqual(candidates(b), [at(4, 1)]);
b[at(4, 1)] = 'red';
assert.deepEqual(sorted(best(b).oranges), sorted([at(3, 1), at(4, 2)]));

// Partida 2: blava al centre, turquesa f1c2, blava f1c1 -> f5c2 o f4c5
b = Array(25).fill(null);
b[at(3, 3)] = 'blue';
assert.equal(candidates(b).length, 8);
b[at(1, 2)] = 'teal';
b[at(1, 1)] = 'blue';
assert.deepEqual(sorted(candidates(b)), sorted([at(5, 2), at(4, 5)]));

console.log('oc.html: totes les proves passen');
