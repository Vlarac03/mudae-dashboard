// Proves de la pestanya «Millores»: llegir el $k, el $kt i el $opp. Executa: node tests/test_millores.mjs
import { readFileSync } from 'node:fs';
import assert from 'node:assert/strict';

const py = readFileSync(new URL('../mudae.py', import.meta.url), 'utf8');
const js = py.match(/JS_MILL = r"""([\s\S]*?)"""/)[1];
const BADGES = ['Bronze', 'Silver', 'Gold', 'Sapphire', 'Ruby', 'Emerald', 'Diamond'].map(n => [n, 0]);
const num = js.match(/const num = [^\n]*/)[0];
const costs = js.match(/const PCOST[^\n]*\n[^\n]*\n[^\n]*/)[0], lvl = js.match(/const LVL[^\n]*\n[^\n]*/)[0];
const readMudae = new Function('BADGES', [num, costs, lvl, js.match(/function readMudae[\s\S]*?\n}\n/)[0]].join('\n') + 'return readMudae;')(BADGES);

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
// el preu que surt és el del nivell següent: Sapphire II → el III costa 15.000 (5.000 per nivell, els de sèrie)
assert.deepEqual(readMudae(k, 5000), { b: { Bronze: 4, Silver: 4, Gold: 4, Sapphire: 2, Ruby: 1, Emerald: 0, Diamond: 0 }, bp: { Sapphire: 5000, Ruby: 7000 }, k: 28223 });
// un servidor amb altres preus ($badgevalue) i el Ruby IV, que ja ve descomptat al $k: el preu per nivell es desa sense descompte
assert.deepEqual(readMudae('Tienes 10!\n Ruby IV · ¡Máximo alcanzado!\n Emerald · 1.500\n Diamond II · 4.500\nDesbloquea un nivel con $bronze', 5000).bp,
  { Emerald: 2000, Diamond: 2000 });
// el $k de Victor (30/10/2026): totes al IV
const kMax = `Tienes 6.221!\nConvierte tu kakera en insignias de poder:\n` + BADGES.map(([n]) => ` ${n} IV · ¡Máximo alcanzado!`).join('\n') + `\nDesbloquea un nivel con $bronze, $silver...`;
assert.deepEqual(readMudae(kMax, 5000), { b: Object.fromEntries(BADGES.map(([n]) => [n, 4])), bp: {}, k: 6221 });
// $mmy=: els personatges amb claus, de més a menys
assert.deepEqual(readMudae('Harem de jo\n\nLanie Parish ·  🔑 (1)\nHello Kitty ·   (12)\nRui (HOHD) ·   (3)', 5000),
  { keys: [['Hello Kitty', 12], ['Rui (HOHD)', 3], ['Lanie Parish', 1]] });
// $tu: les dues «Capital», la de la kakera i la de les esferes
assert.deepEqual(readMudae('Poder: 110%\nCapital: 4.532:kakera:\n\n0 $oh restantes para hoy\nCapital: 1.725 :sp:\n=> $tuarrange', 5000), { k: 4532, sp: 1725 });

const kt = `Gasta tupara armar los 12 pisos de la torre:
Tu nivel actual es
El siguiente nivel cuesta 40.000
Tienes 615
Arma un piso usando $build <numero del beneficio>
 [1]  +2 espacios en la lista de deseos`;
assert.deepEqual(readMudae(kt, 5000), { ktCost: 40000, nf: 7, k: 615 });
assert.deepEqual(readMudae(kt.replace('cuesta 40.000', 'costs 80,000'), 10000), { ktCost: 80000, nf: 7, k: 615 });
assert.deepEqual(readMudae(kt, null), { ktCost: 40000, k: 615 });  // sense el preu base de la torre, no sabem quants pisos tens

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
// el $opp de debò (Victor, 30/10/2026): al màxim posa [MAX], no [LVL 1]
const oppMax = `Mejora los beneficios del personaje seleccionado.
Un bono se aplica cuando es mejorado completamente. Filtrar con variables funciona.
Escribe $infosphere para ver más información.
Tienes 1.725 :sp:
character image
Harry Potter
[LVL 0]  Probabilidad de aparición incrementada...
[LVL 0]  Valor de kakera base incrementado: 0 > 20
[MAX]  Probabilidad de obtener +1 botón de kakera
[LVL 0]  Probabilidad de obtener +1 llave: 0 > 4%
[LVL 0]  Cantidad de spheres...
[LVL 0]  Un personaje aleatorio...
[LVL 0]  Los botones de kakera...
[LVL 0]  Aparece con 4 botones...
[LVL 0]  Un botón de sphere...
[MAX]  El primer $oh del día genera +20 spheres y tiene un +1% de probabilidad de otorgar 1 $oq`;
assert.deepEqual(readMudae(oppMax, 5000), { opp: { name: 'Harry Potter', levels: [0, 0, 6, 0, 0, 0, 0, 0, 0, 1] }, sp: 1725 });
// el $opz de Victor (01/11/2026): els personatges amb esferes, un per pàgina, amb el mateix format que el $opp
const opz = `1. $opz
2. MudaeAPP — 12:29
Mejora los beneficios del personaje seleccionado.
Un bono se aplica cuando es mejorado completamente. Filtrar con variables funciona.
Escribe $infosphere para ver más información.
Tienes 145
Harry Potter
[LVL 0] Probabilidad de aparición incrementada para los personajes adyacentes a este en tu $wishlist: 0 > 15%
[LVL 0] Valor de kakera base incrementado: 0 > 20
[LVL 0] Probabilidad de obtener +1 botón de kakera: 0 > 7%
[LVL 0] Probabilidad de obtener +1 llave: 0 > 4%
[LVL 0] Cantidad de spheres ganadas por botón de kakera (excepto púrpura) cliqueado por ti en este personaje: 0 > 3
[LVL 0] Un personaje aleatorio de tu lista de deseos podría aparecer automáticamente después de que tu rolleas este personaje (2% de probabilidades). Si no está reclamado, el personaje tendrá protección de deseos. Si ya ha sido reclamado por ti, obtendrás 3 llaves Omega (revisa $ok)
[LVL 0] Los botones de kakera se pueden convertir en kakera del Caos cuando tu rolleas este personaje (1% de probabilidad por kakera excluyendo rojo, luz, oscuridad y arcoíris)
[LVL 0] Aparece con 4 botones de kakera (sin púrpura) y costando la mitad de poder cuando rolleado por primera vez ese día, solo tu puedes hacer click. El descuento aplica solo a los primeros 40 cliqueos del día. Las spheres adquiridas por el beneficio 5 se duplicarán en estos botones tras realizar 40 clicks.
[LVL 0] Un botón de sphere aparece cuando rolleas este personaje por primera vez ese día. 1/7 de obtener 1 $oq por click. Hasta 10 clicks en spheres por día, solo tú puedes hacer click.
[MAX] El primer $oh del día genera +20 spheres y tiene un +1% de probabilidad de otorgar 1 $oq`;
assert.deepEqual(readMudae(opz, 5000), { opp: { name: 'Harry Potter', levels: [0, 0, 0, 0, 0, 0, 0, 0, 0, 1] }, sp: 145 });
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
assert.equal(readMudae(os.replace('[LVL 3]', '[MAX]'), 5000).shop[1], 10);
// $or et torna el 90 % del que hi has posat: l'Harry Potter de Victor (01/11/2026), millora 4 al LVL 4 i la 10 → 3.000, et torna 2.700
const refundPlan = new Function([costs, js.match(/const spent = [^\n]*/)[0], fn('refundPlan')].join('\n') + 'return refundPlan;')();
const lv = o => Array.from({ length: 10 }, (_, i) => o[i + 1] || 0);
assert.deepEqual(refundPlan(lv({ 4: 4, 10: 1 })), { total: 3000, back: 2700, free: 1700 });  // hi tornes la 10 i en queden 1.700
assert.equal(refundPlan(lv({ 10: 1 })), null);             // només la 10: no hi ha res a tornar
assert.equal(refundPlan(lv({ 4: 1, 10: 1 })), null);       // 1.200 → 1.080, i amb la 10 de nou en queden 80: no val la pena
assert.deepEqual(refundPlan(lv({ 8: 1, 9: 1 })), { total: 2000, back: 1800, free: 1800 });
// les compres d'esferes recomanades: la primera de l'ordre que no és al màxim, nivell a nivell
const [nextBuys, perkMax, perkCost, shopMax, shopCost] = new Function(costs + '\n' + fn('nextBuys') + 'return [nextBuys, perkMax, perkCost, shopMax, shopCost];')();
const PORD = [10, 4, 5, 2, 3, 8, 9, 1, 6, 7];
assert.deepEqual(nextBuys([2, 0, 0, 6, 0, 0, 0, 0, 0, 1], PORD, perkMax, perkCost),  // l'Harry Potter de dalt: 10 i 4 fets, toca la 5
  [{ n: 5, l: 1, cost: 200 }, { n: 5, l: 2, cost: 400 }, { n: 5, l: 3, cost: 600 }]);
assert.deepEqual(nextBuys([0, 0, 0, 5, 0, 0, 0, 0, 0, 0], PORD, perkMax, perkCost, 2), [{ n: 10, l: 1, cost: 1000 }, { n: 4, l: 6, cost: 2000 }]);
assert.deepEqual(nextBuys([6, 6, 6, 6, 6, 1, 1, 1, 1, 1], PORD, perkMax, perkCost), []);  // millorat del tot
assert.deepEqual(nextBuys([0, 3], [2, 4], shopMax, shopCost, 2), [{ n: 2, l: 4, cost: 16000 }, { n: 2, l: 5, cost: 20000 }]);
// la botiga: el nivell més barat d'entre les que et serveixen; si costen igual, la primera de l'ordre
const cheapestBuys = new Function(costs + '\n' + fn('cheapestBuys') + 'return cheapestBuys;')();
assert.deepEqual(cheapestBuys([0, 0, 0, 0], [2, 4], shopMax, shopCost), [{ n: 2, l: 1, cost: 4000 }, { n: 4, l: 1, cost: 4000 }, { n: 2, l: 2, cost: 8000 }]);
assert.deepEqual(cheapestBuys([0, 3, 0, 0], [2, 4], shopMax, shopCost, 2), [{ n: 4, l: 1, cost: 4000 }, { n: 4, l: 2, cost: 8000 }]);

// $bonus de Victor (01/11/2026), dues pàgines («1.» i «2.»): una línia per bonus, i el que en fa servir el tauler
const bonus = `1. Bonos del Jugador

 · Rolls por hora: +8 (6 $k + 1 $kl + 1 $kt)
 · Tiempo de espera para $rt: 20h (-30h $k -0h $kl)
 · Límite $limroul de animanga: -6 $wa/$ha (6 $kl + 0 $kt)
 · Poder máximo de Kakera: 110% ($kt)
 · Costo de poder por botón de kakera: 36% (-60% $k -4% $kt)
 · Bono adicional para botones de kakera: +0% ($bk)
2. · Probabilidad de obtener una llave adicional en deseos: +10% ($kt)
 · $oh bono diario: +40 spheres, 2% de obtener 1 $oq ($op) y 0% $ot ($shop)`;
const B = readMudae(bonus + '\n · Espacios de lista de deseos: +8 (6 $k + 0 $kl + 2 $kt)\n · Espacios de Starwish: +0 (0 $kl + 0 $sw)', 5000).bonus;
assert.equal(B.length, 10);
assert.deepEqual(B[0], ['Rolls por hora', '+8 (6 $k + 1 $kl + 1 $kt)']);
assert.deepEqual(B[6], ['Probabilidad de obtener una llave adicional en deseos', '+10% ($kt)']);
assert.deepEqual(readMudae(bonus.replaceAll('·', '•'), 5000).bonus.slice(0, 1), [['Rolls por hora', '+8 (6 $k + 1 $kl + 1 $kt)']]);  // amb • en lloc de ·
assert.equal(readMudae('Bonos del Jugador\n$limroul $kl', 5000).bonus, undefined);  // cap línia entesa: no desa un $bonus buit
// les dues pàgines del $bonus, cadascuna sola (Victor, 01/11/2026: la 2 sola no s'entenia), i ajuntades
const [pg1, pg2] = bonus.split('\n2. ');
assert.equal(readMudae(pg1, 5000).bonus.length, 6);
assert.deepEqual(readMudae('2. ' + pg2, 5000).bonus.map(([l]) => l), ['Probabilidad de obtener una llave adicional en deseos', '$oh bono diario']);
const mergeBonus = new Function(js.match(/const mergeBonus = [^\n]*/)[0] + '\nreturn mergeBonus;')();
const both = mergeBonus(readMudae(pg1, 5000).bonus, readMudae('2. ' + pg2, 5000).bonus);
assert.equal(both.length, 8);
assert.equal(mergeBonus(both, [['$oh bono diario', '+80 spheres']]).find(([l]) => l === '$oh bono diario')[1], '+80 spheres');  // la nova mana
assert.equal(readMudae('[LVL 3] Una megasphere tiene un 1/50\n · Megaspheres: 0', 5000).bonus, undefined);  // el $os no és el $bonus
// el $bonus tal com surt en copiar-lo de Discord (Victor, 01/11/2026): cada línia comença amb el codi de l'emoji
const clip = `:addroll: · Rolls por hora: +8 (6 $k + 1 $kl + 1 $kt)
:wlslot: · Espacios de lista de deseos: +8 (6 $k + 0 $kl + 2 $kt)
:wlslot: · Bono de aparición para deseos: +150% ($k + $bw + slash)
:sw: · % bono adicional de aparición para $starwish: +150% ($kt + $bw + $tuto) (= 300%)
:sw: · Espacios de Starwish:  +0 (0 $kl + 0 $sw)
:wishprotect: · Probabilidad de aparición de Wishprotect: 1/10.000 ($kl)
:rtcd: · Tiempo de espera para $rt: 20h (-30h $k -0h $kl)
:disablemore: · Límite $limroul de animanga: -6 $wa/$ha (6 $kl + 0 $kt)
:disablemore: · Límite $limroul de juegos: -4 $wg/$hg (4 $kl + 0 $kt)
:morekakera: · Bono por kakera obtenido: +0% (premium + slash)
:morekakera: · Bono de kakera de llaves doradas: 4.500 (0 $kt, 0 $bk)
:morekakera: · Poder máximo de Kakera: 110% ($kt)
:morekakera: · Costo de poder por botón de kakera: 36% (-60% $k -4% $kt)
:morekakera: · Bono adicional para botones de kakera: +0% ($bk)
:morekakera: · Bono adicional para botones de kakera en starwishes: +16% ($sw)
:kakeraL: · Kakera aleatorio por kakera de Luz: 4-5 (1 $kt)
:bku: · Probabilidad de completar + reiniciar $bku en $sw: +0% ($kl) (este intervalo: 10%)
:chaoskey: · Probabilidad de obtener una llave adicional en deseos: +10% ($kt)
:sp: · Spheres adicionales: +0 (spheres cliqueadas + premium)
:sp: · $oh bono diario: +80 spheres, 4% de obtener 1 $oq ($op) y 0% $ot ($shop)
:spM: · Megaspheres: 0 recompensas y 0% de probabilidad de ser gratis ($shop)`;
const noEmoji = new Function(js.match(/const noEmoji = [^\n]*/)[0] + '\nreturn noEmoji;')();
assert.equal(readMudae(clip, 5000).bonus, undefined);  // sense treure els emojis, no se n'entén cap línia
assert.equal(readMudae(noEmoji(clip), 5000).bonus.length, 21);
assert.equal(noEmoji(' <:sp:123> · Capital: 1.725 :sp:'), ' · Capital: 1.725 :sp:');  // només el del principi de la línia
const bonusFacts = new Function(fn('bonusFacts') + 'return bonusFacts;')();
assert.deepEqual(bonusFacts(B), { pmax: 110, pcost: 36, oh: 40, rt: 20, wl: 8 });  // 7 de base + 8 = els 15 de Victor
assert.deepEqual(bonusFacts(readMudae(noEmoji(clip), 5000).bonus), { pmax: 110, pcost: 36, oh: 80, rt: 20, wl: 8 });
assert.deepEqual(bonusFacts(undefined), { pmax: null, pcost: null, oh: null, rt: null, wl: null });
assert.deepEqual(readMudae('Hola, això no és de Mudae', 5000), {});
assert.deepEqual(readMudae('Kakera rewards\nBonus Bronze I + Silver I + Gold I\n$bronze', 5000), {});  // el $kakerareward no són els teus nivells
// torre: 12 pisos per torre; el $kt diu el total i d'aquí surt la torre on ets
const TOWER = [11, 8, 7, 6, 5, 1, 4, 2, 9, 10, 3, 12];
const [towerNext, towerFromKt] = new Function('TOWER', fn('towerNext') + fn('towerFromKt').replace(/\n$/, '') + '\nreturn [towerNext, towerFromKt];')(TOWER);
assert.deepEqual(towerNext({ f: [11, 8], tv: 5000 }), { tw: 0, built: 2, cost: 15000, rec: 7 });
assert.deepEqual(towerNext({ t: 1, f: [], tv: 5000 }), { tw: 1, built: 12, cost: 65000, rec: 11 });  // torre 2: torna a començar per l'11
assert.deepEqual(towerFromKt({ f: [11, 8] }, 14), { t: 1, f: [] });  // 14 pisos: torre 2, els marcats eren de la 1
assert.deepEqual(towerFromKt({ t: 1, f: [11, 8] }, 14), { t: 1, f: [11, 8] });
assert.equal(towerNext({ f: [8], tv: null }).cost, null);  // sense preu base, no s'inventa el preu
// el preu base surt del $kt si has marcat els pisos: 7 pisos i el següent costa 40.000 → 5.000
const towerGuess = new Function('TOWER', fn('towerNext') + fn('towerGuess') + 'return towerGuess;')(TOWER);
assert.equal(towerGuess({ f: [1, 2, 4, 7, 8, 10, 11], ktCost: 40000 }), 5000);
assert.equal(towerGuess({ f: [], ktCost: 40000 }), null);  // sense pisos marcats no ho sabem
assert.equal(towerGuess({ f: [8], ktCost: 300000 }), null);  // fora de 5.000-50.000: els pisos marcats no quadren
// «Posar al dia» endevina què has enganxat
const upd = py.match(/JS_UPD = r"""([\s\S]*?)"""/)[1];
const detect = new Function(upd.match(/function detect[\s\S]*?\n}\n/)[0] + 'return detect;')();
assert.deepEqual(detect("jo's Wishlist - 2/15\nSnape\n"), { kind: 'wishlist' });
assert.deepEqual(detect('Castle + Pingu   3/7\n(bundle)\nKate Beckett  💞 => jo - Castle\nKevin Ryan - Castle\nPingu - Pingu\n'), { kind: 'totes' });
assert.deepEqual(detect('Kate Beckett - Castle\nPingu - Pingu\n'), { kind: 'harem' });
assert.deepEqual(detect('1. Harry Potter 2/4\n\nHP\n\nDobby  💞 => pep\nWinky\n'), { kind: 'serie', serie: 'Harry Potter' });
assert.deepEqual(detect('Hola'), {});

// Pokédex: el $pd sumrare de Victor (31/10/2026), quants en té de cada raresa en un sol missatge
const pdJs = py.match(/JS_PD = r"""([\s\S]*?)"""/)[1];
const pdSumrare = new Function([pdJs.match(/const pdNum[^\n]*/)[0], pdJs.match(/function pdSumrare[\s\S]*?\n}\n/)[0]].join('\n') + 'return pdSumrare;')();
assert.deepEqual(pdSumrare('vlarac\n\n5 (1/90)\n\n4 (1/165)\n\n3 (8/337)\n\n2 (15/225)\n\n1 (75/177)\n100 / 1.025'),
  { rare: { 1: [75, 177], 2: [15, 225], 3: [8, 337], 4: [1, 165], 5: [1, 90] }, have: 100, total: 1025 });
// tal com surt copiat de Discord, amb l'emoji 🌟 i la línia «Imagen» (de Victor, 31/10/2026)
assert.deepEqual(pdSumrare('Mudae\nAPP\n — 12:33\n\nvlarac\n\n5🌟 (1/90)\n\n4🌟 (1/165)\n\n3🌟 (8/337)\n\n2🌟 (15/225)\n\n1🌟 (75/177)\nImagen\n100 / 1.025'),
  { rare: { 1: [75, 177], 2: [15, 225], 3: [8, 337], 4: [1, 165], 5: [1, 90] }, have: 100, total: 1025 });
assert.equal(pdSumrare('vlarac\n Krabby x2\n100 / 1.025 - Page 1 / 10'), null);  // una pàgina del $pd no és el sumrare
assert.equal(pdSumrare('Castle 5/8'), null);
console.log('millores: totes les proves passen');
