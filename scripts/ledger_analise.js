/*
 * Agrega o livro-razão por placa (ledger_placas.json) por representante e compara com o painel publicado.
 * Uso: node scripts/ledger_analise.js [ateISO]
 */
const fs = require('fs');
const path = require('path');
const L = JSON.parse(fs.readFileSync(path.join(__dirname, 'ledger_placas.json'), 'utf8'));
const ATE = process.argv[2] || '2026-08-13';
const MES = ATE.slice(0, 7);

const nrm = s => String(s || '').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
  .replace(/[^A-Z ]/g, ' ').replace(/\s+/g, ' ').trim();
const titulo = s => String(s || '').trim().toLowerCase().replace(/\s+/g, ' ').split(' ')
  .map(w => ['de', 'da', 'do', 'e', 'dos', 'das'].includes(w) ? w : w.charAt(0).toUpperCase() + w.slice(1)).join(' ');

const placas = Object.values(L);

// --- canonicalização: o PPM trunca nomes em 30 chars. Nomes completos vêm da BASE.
const completos = new Map();
placas.forEach(x => { const c = x.base && x.base.consultor; if (c) completos.set(nrm(c), titulo(c)); });
function canon(nome) {
  if (!nome) return '';
  const k = nrm(nome);
  if (completos.has(k)) return completos.get(k);
  const c = [...completos.keys()].filter(v => v !== k && v.length >= 12 && k.length >= 12 && (v.startsWith(k) || k.startsWith(v)));
  return c.length === 1 ? completos.get(c[0]) : titulo(nome);
}

// --- agrega
const EM_VIGOR = new Set(['Ativo', 'Inadimplente']);
const acc = {};
for (const x of placas) {
  if (!x.base) continue;                      // só o que existe na carteira do Siprov
  const nome = canon(x.rep);
  const a = acc[nome] || (acc[nome] = { nome, carteira: 0, vendasMes: 0, fontes: {} });
  a.fontes[x.repFonte] = (a.fontes[x.repFonte] || 0) + 1;
  if (EM_VIGOR.has(x.base.situacao)) a.carteira++;
  const d = x.base.adesao;
  if (d && d.slice(0, 7) === MES && d <= ATE) a.vendasMes++;
}
const lista = Object.values(acc).sort((a, b) => b.carteira - a.carteira);

// --- painel publicado, para comparar
const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const D = JSON.parse(html.match(/<script id="dashboard-data" type="application\/json">([\s\S]*?)<\/script>/)[1]);
const painel = {}; D.representante.forEach(r => painel[nrm(r.nome)] = r);

console.log('=== POR REPRESENTANTE — livro-razão (placas) x painel publicado ===');
console.log('carteira            vendas ' + MES);
console.log('ledger painel  dif |ledger painel  dif | representante');
let difC = 0, difV = 0;
for (const r of lista) {
  const p = painel[nrm(r.nome)];
  const pc = p ? p.carteira !== undefined ? p.carteira : p.total : 0;
  const pv = p ? p['vendas_' + ({ '01': 'janeiro', '08': 'agosto', '07': 'julho', '06': 'junho', '05': 'maio' }[MES.slice(5)])] || 0 : 0;
  difC += Math.abs(r.carteira - pc); difV += Math.abs(r.vendasMes - pv);
  if (r.carteira >= 10 || Math.abs(r.vendasMes - pv) >= 3)
    console.log(String(r.carteira).padStart(6), String(pc).padStart(6), String(r.carteira - pc).padStart(4),
      '|', String(r.vendasMes).padStart(6), String(pv).padStart(6), String(r.vendasMes - pv).padStart(4), '|', r.nome);
}
console.log('\ntotais: ledger carteira =', lista.reduce((s, r) => s + r.carteira, 0),
  '| painel =', D.kpis.carteira_qtde);
console.log('        ledger vendas', MES, '=', lista.reduce((s, r) => s + r.vendasMes, 0),
  '| painel =', D.kpis['vendas_agosto'].qtde);
console.log('        representantes: ledger =', lista.length, '| painel =', D.representante.length);

console.log('\n=== KAUAN (detalhe da origem de cada placa) ===');
const k = lista.find(r => /KAUAN/i.test(r.nome));
console.log(k ? JSON.stringify(k, null, 1) : 'não encontrado');
