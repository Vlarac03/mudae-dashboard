// Proves de l'ajudant del $oq amb la partida de Victor (02/11/2026). Executa: node tests/test_oq.mjs
import { readFileSync } from 'node:fs';
import assert from 'node:assert/strict';

const html = readFileSync(new URL('../oq.html', import.meta.url), 'utf8');
const src = html.match(/<script id="logic">([\s\S]*?)<\/script>/)[1];
const mod = { exports: {} };
new Function('module', src)(mod);
const { solutions, best } = mod.exports;
const at = (r, c) => (r - 1) * 5 + (c - 1);

// Sense res marcat, totes les combinacions (25 sobre 4)
let b = Array(25).fill(null);
assert.equal(solutions(b).length, 12650);

// La partida de la captura: porpres a f1c1, f3c1, f5c3 i f5c4
const PUR = [at(1, 1), at(3, 1), at(5, 3), at(5, 4)];
const COL = ['blue', 'teal', 'green', 'yellow', 'orange'];
const near = (i, j) => i !== j && Math.abs(Math.floor(i / 5) - Math.floor(j / 5)) <= 1 && Math.abs(i % 5 - j % 5) <= 1;
const color = (i, pur) => pur.includes(i) ? 'purple' : COL[pur.filter(j => near(i, j)).length];
for (const [r, c, col] of [[2, 2, 'green'], [2, 4, 'blue'], [3, 3, 'blue'], [4, 2, 'green'], [4, 3, 'green'], [4, 4, 'green'], [5, 3, 'purple'], [5, 2, 'teal']]) {
  assert.equal(color(at(r, c), PUR), col);  // la regla del joc encaixa amb el que va sortir
  b[at(r, c)] = col;
}
const res = best(b);
assert.ok(solutions(b).some(s => PUR.every(p => s.has(p))));
assert.equal(res.prob[at(5, 4)], 1);  // amb això, la f5c4 era segur porpra
assert.equal(res.prob[at(5, 4)], Math.max(...res.prob.filter(x => x != null)));

// Partida de Victor (04/11/2026): 3 porpres i la vermella (la 4a, que compta com a porpra per als colors)
const P2 = [at(1, 1), at(2, 4), at(3, 4), at(5, 1)];
b = Array(25).fill(null);
for (const [r, c, col] of [[2, 2, 'teal'], [3, 1, 'blue'], [3, 3, 'green'], [4, 2, 'teal'], [1, 1, 'purple'], [2, 4, 'purple'], [3, 4, 'purple'], [5, 1, 'red']]) {
  assert.equal(color(at(r, c), P2).replace('purple', at(r, c) === at(5, 1) ? 'red' : 'purple'), col);
  b[at(r, c)] = col;
}
const fin = best(b);
assert.equal(fin.found, 4);
assert.equal(fin.near[fin.pick], 2);  // la millor que queda és verda (f2c3, f2c5 o f3c5)
assert.ok([at(2, 3), at(2, 5), at(3, 5)].includes(fin.pick));
assert.equal(fin.near[at(3, 2)], 0);
b[at(5, 1)] = 'purple';  // marcada com a porpra en lloc de vermella, el mateix
assert.equal(best(b).pick, fin.pick);

// Jugant sempre la recomanada, quantes partides es guanyen amb 7 clics (les porpres no en gasten)
let wins = 0, seed = 7;
const rnd = () => (seed = (seed * 1103515245 + 12345) % 2147483648) / 2147483648;
for (let g = 0; g < 40; g++) {
  const pur = []; while (pur.length < 4) { const x = Math.floor(rnd() * 25); if (!pur.includes(x)) pur.push(x); }
  const bd = Array(25).fill(null); let clicks = 0;
  while (clicks < 7 && bd.filter(c => c === 'purple').length < 3) {
    const i = best(bd).pick; bd[i] = color(i, pur);
    if (bd[i] !== 'purple') clicks++;
  }
  if (bd.filter(c => c === 'purple').length >= 3) wins++;
}
assert.ok(wins >= 30, `només ${wins} de 40`);
console.log(`oq.html: totes les proves passen (${wins} de 40 partides guanyades)`);
