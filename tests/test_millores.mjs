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

const fn = name => js.match(new RegExp(`function ${name}[\\s\\S]*?\\n}\\n`))[0];
// $os: la botiga d'esferes (10 bonus; alguns ocupen dues línies)
const os = `Mejora los bonos de los ourobeneficios. Estos bonos se aplican a TODOS los personajes. Escribe $shoprefund para reembolsar las spheres. Escribe $oo para los fragmentos.
Cada bono tiene 10 niveles (su costo se incrementa por +4,000 cada nivel).
Tienes 1.725
[LVL 0] Una porción del bono...
[LVL 3] Una megasphere tiene un 1/50...
[LVL 0] El botón de kakera adicional...
[LVL 0] Cuando obtienes una llave...
[LVL 0] Cuando un botón de kakera entrega spheres...
[LVL 0] Los deseos que aparezcan gracias al beneficio 6...
Tus deseos reclamados que aparecen gracias al beneficio 6 pueden dar +1 llave omega adicional: 0% > 50%
[LVL 0] Todo el kakera del Caos...
[LVL 0] Si aparecen en personajes completamente mejorados...
Al hacer click en botones de kakera...
[LVL 0] Cada día, se pueden cliquear más botones de sphere...
Spheres cliqueados gracias al beneficio 9 dan más spheres: 0% > 10%
[LVL 10] El primer $oh del día...`;
assert.deepEqual(readMudae(os, 5000), { shop: [0, 3, 0, 0, 0, 0, 0, 0, 0, 10], sp: 1725 });
// les compres d'esferes recomanades: la primera de l'ordre que no és al màxim, nivell a nivell
const costs = js.match(/const PCOST[^\n]*\n[^\n]*\n[^\n]*/)[0];
const [nextBuys, perkMax, perkCost, shopMax, shopCost] = new Function(costs + '\n' + fn('nextBuys') + 'return [nextBuys, perkMax, perkCost, shopMax, shopCost];')();
const PORD = [10, 4, 5, 2, 3, 8, 9, 1, 6, 7];
assert.deepEqual(nextBuys([2, 0, 0, 6, 0, 0, 0, 0, 0, 1], PORD, perkMax, perkCost),  // l'Harry Potter de dalt: 10 i 4 fets, toca la 5
  [{ n: 5, l: 1, cost: 200 }, { n: 5, l: 2, cost: 400 }, { n: 5, l: 3, cost: 600 }]);
assert.deepEqual(nextBuys([0, 0, 0, 5, 0, 0, 0, 0, 0, 0], PORD, perkMax, perkCost, 2), [{ n: 10, l: 1, cost: 1000 }, { n: 4, l: 6, cost: 2000 }]);
assert.deepEqual(nextBuys([6, 6, 6, 6, 6, 1, 1, 1, 1, 1], PORD, perkMax, perkCost), []);  // millorat del tot
assert.deepEqual(nextBuys([0, 3], [2, 4], shopMax, shopCost, 2), [{ n: 2, l: 4, cost: 16000 }, { n: 2, l: 5, cost: 20000 }]);

assert.deepEqual(readMudae('Hola, això no és de Mudae', 5000), {});
assert.deepEqual(readMudae('Kakera rewards\nBonus Bronze I + Silver I + Gold I\n$bronze', 5000), {});  // el $kakerareward no són els teus nivells
// torre: 12 pisos per torre; el $kt diu el total i d'aquí surt la torre on ets
const TOWER = [11, 8, 7, 6, 5, 1, 4, 2, 9, 10, 3, 12];
const [towerNext, towerFromKt] = new Function('TOWER', fn('towerNext') + fn('towerFromKt').replace(/\n$/, '') + '\nreturn [towerNext, towerFromKt];')(TOWER);
assert.deepEqual(towerNext({ f: [11, 8], tv: 5000 }), { tw: 0, built: 2, cost: 15000, rec: 7 });
assert.deepEqual(towerNext({ t: 1, f: [], tv: 5000 }), { tw: 1, built: 12, cost: 65000, rec: 11 });  // torre 2: torna a començar per l'11
assert.deepEqual(towerFromKt({ f: [11, 8] }, 14), { t: 1, f: [] });  // 14 pisos: torre 2, els marcats eren de la 1
assert.deepEqual(towerFromKt({ t: 1, f: [11, 8] }, 14), { t: 1, f: [11, 8] });
// «Posar al dia» endevina què has enganxat
const upd = py.match(/JS_UPD = r"""([\s\S]*?)"""/)[1];
const detect = new Function(upd.match(/function detect[\s\S]*?\n}\n/)[0] + 'return detect;')();
assert.deepEqual(detect("jo's Wishlist - 2/15\nSnape\n"), { kind: 'wishlist' });
assert.deepEqual(detect('Castle + Pingu   3/7\n(bundle)\nKate Beckett  💞 => jo - Castle\nKevin Ryan - Castle\nPingu - Pingu\n'), { kind: 'totes' });
assert.deepEqual(detect('Kate Beckett - Castle\nPingu - Pingu\n'), { kind: 'harem' });
assert.deepEqual(detect('1. Harry Potter 2/4\n\nHP\n\nDobby  💞 => pep\nWinky\n'), { kind: 'serie', serie: 'Harry Potter' });
assert.deepEqual(detect('Hola'), {});
console.log('millores: totes les proves passen');
