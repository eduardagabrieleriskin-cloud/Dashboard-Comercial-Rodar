/*
 * Livro-razão por PLACA: junta as três fontes numa linha só por placa.
 *
 *   COTAÇÃO (Controle de Cotações V2) -> SUBSCRIÇÃO (Controle de Subscrição V2) -> BASE (Siprov)
 *
 * A placa é a única chave que existe nas três (o Nº da cotação NÃO serve: quando uma cotação vira
 * subscrição ela sai do Controle de Cotações — os dois arquivos são conjuntos disjuntos).
 *
 * Serve para responder, sem ambiguidade: de quem é cada placa, quando foi cotada, quando foi vendida
 * e qual a situação dela hoje. Gera scripts/ledger_placas.json.
 *
 * Uso: node scripts/ledger_placas.js [ateISO]
 */
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const DOWNLOADS = path.join(process.env.USERPROFILE, 'Downloads');
const ATE = process.argv[2] || '2026-08-13';
const OUT = path.join(__dirname, 'ledger_placas.json');

const np = s => { const p = String(s == null ? '' : s).toUpperCase().replace(/[^A-Z0-9]/g, ''); return p.length >= 6 ? p : ''; };
const iso = s => { const m = String(s || '').match(/(\d{2})\/(\d{2})\/(\d{4})/); return m ? m[3] + '-' + m[2] + '-' + m[1] : null; };
const limpaFranquia = s => String(s || '').replace(/^\d+\s*-?\s*/, '').trim();
const placasDe = cel => String(cel || '').split(/[\s,\/]+/).map(np).filter(Boolean);

function maisRecente(regex) {
  const a = fs.readdirSync(DOWNLOADS).filter(f => regex.test(f) && !f.startsWith('~$'))
    .map(f => ({ f, full: path.join(DOWNLOADS, f), m: fs.statSync(path.join(DOWNLOADS, f)).mtimeMs }))
    .sort((x, y) => y.m - x.m);
  return a[0] || null;
}
const linhas = (file, skip) => {
  const wb = XLSX.readFile(file);
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false }).slice(skip);
};

const fCot = maisRecente(/^Controle_de_Cota.*\.xlsx$/i);
const fSub = maisRecente(/^Controle_de_Subscri.*\.xlsx$/i);
const fBase = maisRecente(/^BASE_\d{8}.*\.xlsx$/i);
console.log('fontes:\n  cotações :', fCot && fCot.f, '\n  subscrição:', fSub && fSub.f, '\n  base      :', fBase && fBase.f, '\n  corte     :', ATE, '\n');

const L = {}; // placa -> registro
const reg = p => L[p] || (L[p] = { placa: p });

// ---- 1. COTAÇÃO: 0=Código 1=Associado 3=Placas 4=Data 6=Franquia 7=Representante 8=Status
let nCot = 0;
for (const r of linhas(fCot.full, 2)) {
  const d = iso(r[4]); if (!d) continue;
  for (const p of placasDe(r[3])) {
    const x = reg(p);
    if (!x.cot || d < x.cot.data) { // guarda a PRIMEIRA cotação da placa
      x.cot = { data: d, rep: String(r[7] || '').trim(), franquia: limpaFranquia(r[6]), status: String(r[8] || '').trim(), codigo: String(r[0] || '').split('/')[0] };
    }
    nCot++;
  }
}

// ---- 2. SUBSCRIÇÃO: 0=NºCotação 2=CPF 7=Placa(s) 16=Data Transmissão 24=Franquia 26=Representante 28=Status
let nSub = 0;
for (const r of linhas(fSub.full, 2)) {
  const d = iso(r[16]); if (!d) continue;
  for (const p of placasDe(r[7])) {
    const x = reg(p);
    if (!x.sub || d > x.sub.data) { // guarda a subscrição MAIS RECENTE (a que vale hoje)
      x.sub = { data: d, rep: String(r[26] || '').trim(), franquia: limpaFranquia(r[24]), status: String(r[28] || '').trim(), cotacao: String(r[0] || '').trim() };
    }
    nSub++;
  }
}

// ---- 3. BASE
const todas = linhas(fBase.full, 0);
const hdr = todas[1];
const ix = n => hdr.findIndex(h => String(h || '').trim().toUpperCase() === n);
const C = {
  placa: ix('VEÍCULO - PLACA DO VEÍCULO'), situacao: ix('BENEFÍCIO - SITUAÇÃO ATUAL'),
  adesao: ix('BENEFÍCIO - DATA DE ADESÃO'), consultor: ix('BENEFÍCIO - NOME DO CONSULTOR'),
  representante: ix('BENEFÍCIO - REPRESENTANTE'), valor: ix('BENEFÍCIO - VALOR DA MENSALIDADE AJUSTADA'),
  loja: ix('BENEFÍCIO - LOJA - NOME FANTASIA'), assoc: ix('ASSOCIADO - CPF/CNPJ'),
};
const NORM_SIT = { 'Endosso Ativo': 'Ativo', 'Endosso Inadimplente': 'Inadimplente' };
let nBase = 0, semPlaca = 0;
for (const r of todas.slice(2)) {
  const p = np(r[C.placa]); if (!p) { semPlaca++; continue; }
  let sit = String(r[C.situacao] || '').trim(); sit = NORM_SIT[sit] || sit;
  const x = reg(p);
  x.base = { situacao: sit, adesao: iso(r[C.adesao]), consultor: String(r[C.consultor] || '').trim(),
    representantePJ: String(r[C.representante] || '').trim(), loja: String(r[C.loja] || '').trim(),
    valor: parseFloat(String(r[C.valor] || '0').replace(/[^\d,.-]/g, '').replace(/\.(?=\d{3}(\D|$))/g, '').replace(',', '.')) || 0,
    assoc: String(r[C.assoc] || '').trim() };
  nBase++;
}

// ---- Cobertura do cruzamento
const todasPlacas = Object.values(L);
const naBase = todasPlacas.filter(x => x.base);
const cnt = (arr, f) => arr.filter(f).length;
console.log('placas distintas no total :', todasPlacas.length);
console.log('  em Cotação  :', cnt(todasPlacas, x => x.cot));
console.log('  em Subscrição:', cnt(todasPlacas, x => x.sub));
console.log('  na BASE     :', naBase.length, '(linhas sem placa na BASE:', semPlaca + ')');
console.log('\nDas placas da BASE, quantas achamos nas outras fontes:');
console.log('  BASE + Subscrição + Cotação:', cnt(naBase, x => x.sub && x.cot));
console.log('  BASE + Subscrição          :', cnt(naBase, x => x.sub && !x.cot));
console.log('  BASE + Cotação             :', cnt(naBase, x => !x.sub && x.cot));
console.log('  só BASE (órfã)             :', cnt(naBase, x => !x.sub && !x.cot));

// ---- Quem é o dono da placa (precedência explícita) e de onde veio
for (const x of todasPlacas) {
  const b = x.base || {};
  if (b.consultor) { x.rep = b.consultor; x.repFonte = 'BASE.consultor'; }
  else if (x.sub && x.sub.rep) { x.rep = x.sub.rep; x.repFonte = 'SUBSCRIÇÃO'; }
  else if (x.cot && x.cot.rep) { x.rep = x.cot.rep; x.repFonte = 'COTAÇÃO'; }
  else if (b.representantePJ) { x.rep = b.representantePJ; x.repFonte = 'BASE.PJ'; }
  else { x.rep = '(sem dono)'; x.repFonte = 'nenhuma'; }
  x.franquia = (x.sub && x.sub.franquia) || (x.cot && x.cot.franquia) || '';
}
console.log('\nOrigem do representante (placas na BASE):');
const porFonte = {}; naBase.forEach(x => porFonte[x.repFonte] = (porFonte[x.repFonte] || 0) + 1);
Object.entries(porFonte).sort((a, b) => b[1] - a[1]).forEach(([k, v]) => console.log('  ', String(v).padStart(5), k));

fs.writeFileSync(OUT, JSON.stringify(L));
console.log('\nledger salvo em', OUT, '(' + todasPlacas.length + ' placas)');
