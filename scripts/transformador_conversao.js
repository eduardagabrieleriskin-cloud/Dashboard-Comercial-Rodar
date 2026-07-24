/*
 * Calcula cotações mensais por consultor e a taxa de conversão por consultor.
 *
 * Fontes:
 *  - Controle de Subscrição V2 (PPM/AMSS): cada linha = 1 placa cotada, com Representante, Franquia, Data e Status
 *  - BASE do Siprov: define o FECHAMENTO (placa com situação Ativo/Inadimplente)
 *
 * Cruzamento por PLACA (não por nome) — garante que a conversão nunca passe de 100%,
 * pois os fechamentos são sempre um subconjunto das placas cotadas.
 *
 * Uso: build(subscricaoXlsx, baseXlsx, ateISO) -> { consultores, totais, meses }
 */
const XLSX = require('xlsx');

const S = { cotacao: 0, placa: 7, data: 16, franquia: 24, representante: 26, status: 28 };
const B = { situacao: 16, placa: 26 };
const MESES = ['2026-05', '2026-06', '2026-07'];

const np = s => (s || '').toString().toUpperCase().replace(/[^A-Z0-9]/g, '');
function iso(s) { if (!s) return null; const m = String(s).match(/(\d{2})\/(\d{2})\/(\d{4})/); return m ? m[3] + '-' + m[2] + '-' + m[1] : null; }
function titleCase(s) {
  return (s || '').toString().trim().toLowerCase().replace(/\s+/g, ' ')
    .split(' ').map(w => ['de','da','do','e','dos','das'].includes(w) ? w : (w.charAt(0).toUpperCase() + w.slice(1))).join(' ');
}
function ehTeste(nome) {
  const n = (nome || '').trim().toLowerCase();
  return n === 'eduarda' || n === 'yara' || n === 'teste' || /^teste?\b/.test(n) || n.includes('(teste)');
}

module.exports = function build(subscricaoXlsx, baseXlsx, ateISO) {
  const ate = ateISO || '2026-12-31';

  // fechamentos: placas ativas/inadimplentes na BASE
  const wbB = XLSX.readFile(baseXlsx);
  const base = XLSX.utils.sheet_to_json(wbB.Sheets[wbB.SheetNames[0]], { header: 1, raw: false }).slice(2);
  const fechadas = new Set();
  for (const r of base) {
    const p = np(r[B.placa]);
    if (p && ['Ativo', 'Inadimplente'].includes((r[B.situacao] || '').trim())) fechadas.add(p);
  }

  // cotações
  const wbS = XLSX.readFile(subscricaoXlsx);
  const sub = XLSX.utils.sheet_to_json(wbS.Sheets[wbS.SheetNames[0]], { header: 1, raw: false }).slice(2);

  const acc = {};
  for (const r of sub) {
    const d = iso(r[S.data]);
    if (!d || d > ate) continue;
    const nome = titleCase(r[S.representante]);
    if (!nome || ehTeste(nome)) continue;
    const mes = d.slice(0, 7);
    if (!MESES.includes(mes)) continue;                       // foco: maio a julho
    const a = acc[nome] || (acc[nome] = {
      nome, unidade: (r[S.franquia] || '(Sem Unidade)').toString().trim(),
      cot: {}, fech: {}, cotacoes: new Set(), total_cot: 0, total_fech: 0
    });
    a.cot[mes] = (a.cot[mes] || 0) + 1;
    a.total_cot++;
    if (r[S.cotacao]) a.cotacoes.add(String(r[S.cotacao]));
    if (fechadas.has(np(r[S.placa]))) { a.fech[mes] = (a.fech[mes] || 0) + 1; a.total_fech++; }
  }

  const consultores = Object.values(acc).map(a => {
    const o = {
      nome: a.nome, unidade: a.unidade,
      cotacoes_distintas: a.cotacoes.size,
      total_cotado: a.total_cot, total_fechado: a.total_fech,
      conversao: a.total_cot ? +(a.total_fech / a.total_cot).toFixed(4) : 0
    };
    MESES.forEach(m => { o['cot_' + m] = a.cot[m] || 0; o['fech_' + m] = a.fech[m] || 0; });
    return o;
  }).sort((x, y) => y.total_cotado - x.total_cotado);

  // sanidade: conversão nunca > 100%
  const invalidos = consultores.filter(c => c.total_fechado > c.total_cotado);
  if (invalidos.length) throw new Error('conversao >100% em: ' + invalidos.map(c => c.nome).join(', '));

  const tc = consultores.reduce((s, c) => s + c.total_cotado, 0);
  const tf = consultores.reduce((s, c) => s + c.total_fechado, 0);
  const totais = {
    total_cotado: tc, total_fechado: tf,
    conversao: tc ? +(tf / tc).toFixed(4) : 0,
    por_mes: MESES.reduce((o, m) => {
      const c = consultores.reduce((s, x) => s + x['cot_' + m], 0);
      const f = consultores.reduce((s, x) => s + x['fech_' + m], 0);
      o[m] = { cotado: c, fechado: f, conversao: c ? +(f / c).toFixed(4) : 0 };
      return o;
    }, {}),
    consultores: consultores.length
  };

  return { consultores, totais, meses: MESES };
};
