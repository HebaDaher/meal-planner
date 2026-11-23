// Builds products.json (the app's built-in list of real products) and test-products.json
// from the free Open Food Facts database. Runs on GitHub Actions — see .github/workflows/products.yml
import fs from 'node:fs';
const TARGET = 500;
const FIELDS = 'code,product_name,product_name_en,brands,quantity,serving_size,serving_quantity,ingredients_text,ingredients_text_en,ingredients_tags,allergens_tags,traces_tags,nutriments';
const UA = 'MealPlannerIngredientChecker/1.0 (HCI university project; github.com/HebaDaher/meal-planner)';
const sleep = ms => new Promise(r => setTimeout(r, ms));

// --- use the app's own checking code, taken straight from index.html ---
const src = fs.readFileSync('index.html', 'utf8');
const g = (a, b) => { const i = src.indexOf(a); const j = src.indexOf(b, i); if (i < 0 || j < 0) throw new Error('missing ' + a); return src.slice(i, j); };
const code = [g('const ALLERGENS={', '\nconst sevWord'), g('function matchAllergen', 'function recipeStatus'), g('const OFF_MAP=', 'const SRC='), g('const MAY_RE=', 'function makeProduct'), g('function fromOFF', '/* built-in list')].join('\n');
const app = new Function(code + ';return {fromOFF,evaluate};')();

const ean = c => { if (!/^\d{8}$|^\d{12,13}$/.test(c)) return false; const d = c.padStart(13, '0').split('').map(Number); let s = 0; for (let i = 0; i < 12; i++) s += d[i] * (i % 2 ? 3 : 1); return (10 - s % 10) % 10 === d[12]; };
const usable = x => ean(x.code) && (x.product_name_en || x.product_name) &&
  (((x.ingredients_text_en || '').length > 15 && /[a-z]/i.test(x.ingredients_text_en)) || (x.ingredients_tags || []).filter(t => t.startsWith('en:')).length >= 3);

async function page(country, p) {
  const u = `https://world.openfoodfacts.org/api/v2/search?countries_tags_en=${country}&states_tags=en:ingredients-completed&sort_by=unique_scans_n&page_size=100&page=${p}&fields=${FIELDS}`;
  for (let t = 0; t < 5; t++) {
    try { const r = await fetch(u, { headers: { 'User-Agent': UA } }); if (r.ok) return await r.json(); console.log(country, p, 'HTTP', r.status); }
    catch (e) { console.log(country, p, String(e)); }
    await sleep(15000 * (t + 1));
  }
  return null;
}

const seen = new Map();
for (const [country, label] of [['united-arab-emirates', 'UAE'], ['saudi-arabia', 'Saudi Arabia'], ['united-kingdom', 'UK']]) {
  for (let p = 1; p <= 12; p++) {
    const j = await page(country, p);
    const prods = (j && j.products) || [];
    let added = 0;
    for (const x of prods) if (!seen.has(x.code) && usable(x)) { x.__src = label; seen.set(x.code, x); added++; }
    console.log(label, 'page', p, '+', added, 'usable, total', seen.size);
    if (!prods.length || seen.size >= TARGET) break;
    await sleep(7000);
  }
  if (seen.size >= TARGET) break;
}

const profiles = { T: { name: 'X', allergens: { tree: 5 }, sugar: null }, P: { name: 'X', allergens: { peanut: 5 }, sugar: null }, M: { name: 'X', allergens: { milk: 4 }, sugar: null }, G: { name: 'X', allergens: { gluten: 4 }, sugar: null }, D: { name: 'X', allergens: {}, sugar: 10 } };
const W = { bad: 'N', unc: 'U', safe: 'S' };
const db = {}, list = [];
for (const x of [...seen.values()].slice(0, TARGET)) {
  const n = x.nutriments || {};
  const rec = { product_name: x.product_name || '', product_name_en: x.product_name_en || '', brands: (x.brands || '').split(',').slice(0, 2).join(','), quantity: x.quantity || '', serving_size: x.serving_size || '', serving_quantity: x.serving_quantity || '',
    ingredients_text: x.ingredients_text_en ? '' : (x.ingredients_text || '').slice(0, 900), ingredients_text_en: (x.ingredients_text_en || '').slice(0, 900),
    ingredients_tags: (x.ingredients_tags || []).filter(t => t.startsWith('en:')).slice(0, 60), allergens_tags: x.allergens_tags || [], traces_tags: x.traces_tags || [],
    nutriments: { sugars_serving: n.sugars_serving, sugars_100g: n.sugars_100g } };
  db[x.code] = rec;
  const p = app.fromOFF(rec, x.code); const r = {};
  for (const [k, P] of Object.entries(profiles)) { const ev = app.evaluate(p, P, null); r[k] = (ev.status === 'unc' && ev.why.length === 1 && ev.why[0] === 'over') ? 'O' : W[ev.status]; }
  list.push({ code: x.code, name: (p.name || '').slice(0, 70), from: x.__src, contains: (x.allergens_tags || []).map(t => t.replace(/^\w+:/, '')), mayContain: (x.traces_tags || []).map(t => t.replace(/^\w+:/, '')), sugarPerServing: p.sugar, expected: { treeNuts: r.T, peanut: r.P, milk: r.M, gluten: r.G, diabetes10g: r.D } });
}
fs.writeFileSync('products.json', JSON.stringify(db));
fs.writeFileSync('test-products.json', JSON.stringify({ built: new Date().toISOString(), source: 'Open Food Facts (openfoodfacts.org), ODbL licence', legend: { N: 'Not safe', U: 'Uncertain', O: 'Over sugar limit', S: 'Safe' }, products: list }, null, 1));
console.log('wrote', list.length, 'products');
if (list.length < 100) process.exit(1);
