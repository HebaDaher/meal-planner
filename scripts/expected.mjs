// Re-computes the expected results in test-products.json from products.json with the app's current code (no network needed).
import fs from 'node:fs';
const src=fs.readFileSync('index.html','utf8');const g=(a,b)=>{const i=src.indexOf(a);const j=src.indexOf(b,i);if(i<0||j<0)throw new Error('missing '+a);return src.slice(i,j);};
const app=new Function([g('const ALLERGENS={','\nconst sevWord'),g('function matchAllergen','function recipeStatus'),g('const OFF_MAP=','const SRC='),g('const MAY_RE=','function makeProduct'),g('function fromOFF','/* built-in list')].join('\n')+';return {fromOFF,evaluate};')();
const db=JSON.parse(fs.readFileSync('products.json','utf8'));const t=JSON.parse(fs.readFileSync('test-products.json','utf8'));
const profiles={treeNuts:{allergens:{tree:5}},peanut:{allergens:{peanut:5}},milk:{allergens:{milk:4}},gluten:{allergens:{gluten:4}},diabetes10g:{allergens:{},sugar:10}};
const W={bad:'N',unc:'U',safe:'S'};let changed=0;
for(const x of t.products){const p=app.fromOFF(db[x.code],x.code);for(const [k,P] of Object.entries(profiles)){P.name='X';const ev=app.evaluate(p,P,null);const v=(ev.status==='unc'&&ev.why.length===1&&ev.why[0]==='over')?'O':W[ev.status];if(x.expected[k]!==v){changed++;console.log('changed',x.code,x.name,k,x.expected[k],'->',v);}x.expected[k]=v;}}
t.checked=new Date().toISOString();fs.writeFileSync('test-products.json',JSON.stringify(t,null,1));console.log('results changed:',changed);
