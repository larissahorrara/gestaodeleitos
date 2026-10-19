// Extrai do index.html as bases fictícias embutidas e os indicadores calculados pelo próprio painel.
// Uso: node scripts/extrair_dados.js   →   build/bases.json e build/calculos.json
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const raiz = path.join(__dirname, '..');
const html = fs.readFileSync(path.join(raiz, 'index.html'), 'utf8');
const scripts = [...html.matchAll(/<script>\n([\s\S]*?)\n<\/script>/g)].map((m) => m[1]);

const blocoDados = scripts.find((s) => s.startsWith('window.__DADOS__'));
const blocoCalc = scripts.find((s) => s.includes('dash.calc(') && !s.includes('window.dash ='));
if (!blocoDados || !blocoCalc) throw new Error('Não encontrei os dados ou os cálculos no index.html');

const ctx = { window: {}, console };
vm.runInNewContext(blocoDados, ctx);
const bases = ctx.window.__DADOS__;
const D = JSON.parse(JSON.stringify(bases));

ctx.dash = { calc: (id, def) => { D[id] = def.fn(...def.inputs.map((i) => D[i])); } };
vm.runInNewContext(blocoCalc, ctx);

fs.mkdirSync(path.join(raiz, 'build'), { recursive: true });
fs.writeFileSync(path.join(raiz, 'build', 'bases.json'), JSON.stringify(bases));
fs.writeFileSync(path.join(raiz, 'build', 'calculos.json'), JSON.stringify(D));
console.log('Bases:', Object.keys(bases).map((k) => `${k} (${bases[k].length})`).join(', '));
console.log('Indicadores calculados:', Object.keys(D).length - Object.keys(bases).length);
