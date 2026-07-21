/*
 * Transforma uma exportação BASE do Siprov (associados/benefícios, .xlsx) no
 * conjunto de dados do painel (mesma estrutura do modelo _14) e devolve o objeto.
 *
 * Uso: const build = require('./transformador_base'); const data = build(caminhoXlsx, ateISO);
 *   caminhoXlsx: caminho do BASE_*.xlsx
 *   ateISO: data limite de adesão (ex '2026-07-20'); adesões depois disso são ignoradas nas vendas/série diária
 *
 * Regras de limpeza embutidas: remove testes (loja "Teste Rodar", nomes de teste,
 * situação Recusado/vazia) e garante que todo representante tenha unidade
 * (unidade vem do mapa consultor->franquia; sem mapa => descartado e logado).
 */
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const MAPA = JSON.parse(fs.readFileSync(path.join(__dirname, 'mapa_unidades.json'), 'utf8'));

// colunas (0-indexed) do BASE
const C = { situacao: 16, valorAjust: 19, placa: 26, cpfConsultor: 31, adesao: 32, loja: 36, consultor: 38 };
const MESES = ['2026-05', '2026-06', '2026-07'];

function titleCase(s) {
  return (s || '').toString().trim().toLowerCase().replace(/\s+/g, ' ')
    .split(' ').map(w => ['de','da','do','e','dos','das'].includes(w) ? w : (w.charAt(0).toUpperCase() + w.slice(1))).join(' ');
}
function ehTeste(consultor, loja) {
  const n = (consultor || '').trim().toLowerCase();
  if ((loja || '').trim().toLowerCase() === 'teste rodar') return true;
  return n === 'eduarda' || n === 'yara' || n === 'teste' || /^teste?\b/.test(n) || n.includes('(teste)');
}
function parseISO(d) {
  if (!d) return null;
  const p = d.toString().split('/');
  if (p.length !== 3) return null;
  return p[2] + '-' + p[1].padStart(2, '0') + '-' + p[0].padStart(2, '0');
}
function toNum(v) {
  if (v == null || v === '') return 0;
  const n = parseFloat(v.toString().replace(/[^\d,.-]/g, '').replace(/\.(?=\d{3}(\D|$))/g, '').replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

module.exports = function build(xlsxPath, ateISO) {
  const wb = XLSX.readFile(xlsxPath);
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false }).slice(2);
  const SITU_VALIDAS = ['Ativo', 'Inadimplente', 'Cancelado', 'Inativo', 'Pendente'];
  const dropConsultores = new Set();

  const regs = [];
  for (const r of rows) {
    const situacao = (r[C.situacao] || '').toString().trim();
    if (!SITU_VALIDAS.includes(situacao)) continue;              // fora Recusado/vazio
    const consultor = titleCase(r[C.consultor]);
    const loja = r[C.loja];
    if (ehTeste(r[C.consultor], loja)) continue;                 // fora testes
    const unidade = MAPA[(r[C.consultor] || '').toString().trim().toUpperCase()];
    if (!unidade) { if (consultor) dropConsultores.add(consultor); continue; } // sem unidade => fora (regra)
    regs.push({
      situacao, consultor, unidade,
      adesao: parseISO(r[C.adesao]),
      valor: toNum(r[C.valorAjust]),
    });
  }

  const ate = ateISO || '2026-12-31';
  const ehVendaMes = (reg, mes) => reg.adesao && reg.adesao.slice(0, 7) === mes && reg.adesao <= ate;
  const ehAtivo = reg => reg.situacao === 'Ativo' || reg.situacao === 'Inadimplente';

  // ---- KPIs ----
  function vendasMes(mes) {
    const v = regs.filter(x => ehVendaMes(x, mes));
    const valor = v.reduce((s, x) => s + x.valor, 0);
    const comValor = v.filter(x => x.valor > 0).length;
    return { qtde: v.length, valor: +valor.toFixed(2), ticket_medio: v.length ? +(valor / v.length).toFixed(2) : 0,
      cobertura_valor_n: comValor, cobertura_valor_pct: v.length ? +(comValor / v.length).toFixed(3) : 0 };
  }
  const vm = vendasMes('2026-05'), vj = vendasMes('2026-06'), vjl = vendasMes('2026-07');
  const cont = s => regs.filter(x => x.situacao === s);
  const ativos = cont('Ativo'), inad = cont('Inadimplente'), canc = cont('Cancelado'), inat = cont('Inativo'), pend = cont('Pendente');
  const carteira = regs.filter(ehAtivo);
  const somaVal = arr => +arr.reduce((s, x) => s + x.valor, 0).toFixed(2);
  const universo = ativos.length + inad.length + canc.length + inat.length + pend.length;
  // dias úteis do período de julho até "ate" para o ritmo
  const kpis = {
    data_ultima_venda: regs.filter(x => x.adesao && x.adesao <= ate).map(x => x.adesao).sort().pop() || ate,
    vendas_maio: vm, vendas_junho: vj, vendas_julho: vjl,
    var_maio_junho_pct: vm.qtde ? +((vj.qtde - vm.qtde) / vm.qtde).toFixed(4) : 0,
    var_junho_julho_ritmo_pct: 0,
    carteira_qtde: carteira.length, carteira_valor: somaVal(carteira),
    carteira_ticket_medio: carteira.length ? +(somaVal(carteira) / carteira.length).toFixed(2) : 0,
    carteira_cobertura_valor_n: carteira.filter(x => x.valor > 0).length,
    carteira_cobertura_valor_pct: carteira.length ? +(carteira.filter(x => x.valor > 0).length / carteira.length).toFixed(3) : 0,
    ativos_qtde: ativos.length, ativos_valor: somaVal(ativos),
    inadimplentes_qtde: inad.length, inadimplentes_valor: somaVal(inad),
    cancelados_qtde: canc.length, inativos_qtde: inat.length, pendentes_qtde: pend.length,
    total_universo_qtde: universo,
    pct_inadimplencia: universo ? +(inad.length / universo).toFixed(4) : 0,
    pct_perda: universo ? +((canc.length + inat.length) / universo).toFixed(4) : 0,
  };
  // variação de ritmo (vendas por dia útil) junho -> julho
  const diasUteis = (ini, fim) => { let c = 0, dd = new Date(ini + 'T12:00:00'), ee = new Date(fim + 'T12:00:00'); while (dd <= ee) { const w = dd.getUTCDay(); if (w !== 0 && w !== 6) c++; dd.setUTCDate(dd.getUTCDate() + 1); } return c; };
  const duJun = diasUteis('2026-06-01', '2026-06-30');
  const duJul = diasUteis('2026-07-01', ate < '2026-07-31' ? ate : '2026-07-31');
  const ritmoJun = duJun ? vj.qtde / duJun : 0;
  const ritmoJul = duJul ? vjl.qtde / duJul : 0;
  kpis.var_junho_julho_ritmo_pct = ritmoJun ? +((ritmoJul - ritmoJun) / ritmoJun).toFixed(4) : 0;

  // ---- agregação por unidade / representante ----
  function agrupar(keyFn, extraUnidade) {
    const o = {};
    for (const x of regs) {
      const k = keyFn(x);
      const g = o[k] || (o[k] = { nome: k, vendas_maio: 0, vendas_junho: 0, vendas_julho: 0,
        ativos: 0, valor_ativos: 0, inadimplentes: 0, valor_inadimplentes: 0, cancelados: 0, inativos: 0, pendentes: 0,
        total: 0, valor_total: 0, _unidade: x.unidade });
      if (ehVendaMes(x, '2026-05')) g.vendas_maio++;
      if (ehVendaMes(x, '2026-06')) g.vendas_junho++;
      if (ehVendaMes(x, '2026-07')) g.vendas_julho++;
      if (x.situacao === 'Ativo') { g.ativos++; g.valor_ativos += x.valor; }
      if (x.situacao === 'Inadimplente') { g.inadimplentes++; g.valor_inadimplentes += x.valor; }
      if (x.situacao === 'Cancelado') g.cancelados++;
      if (x.situacao === 'Inativo') g.inativos++;
      if (x.situacao === 'Pendente') g.pendentes++;
    }
    return Object.values(o).map(g => {
      g.total = g.ativos + g.inadimplentes;
      g.valor_total = +(g.valor_ativos + g.valor_inadimplentes).toFixed(2);
      g.valor_ativos = +g.valor_ativos.toFixed(2);
      g.valor_inadimplentes = +g.valor_inadimplentes.toFixed(2);
      const uni = g.total + g.cancelados + g.inativos + g.pendentes;
      g.pct_inadimplencia = uni ? +(g.inadimplentes / uni).toFixed(4) : 0;
      g.pct_perda = uni ? +((g.cancelados + g.inativos) / uni).toFixed(4) : 0;
      if (extraUnidade) g.unidade = g._unidade;
      delete g._unidade;
      return g;
    });
  }
  const unidade = agrupar(x => x.unidade, false).sort((a, b) => b.ativos - a.ativos);
  const representante = agrupar(x => x.consultor, true).sort((a, b) => b.ativos - a.ativos);

  // ---- série diária + médias móveis (dias úteis) ----
  const vendasPeriodo = regs.filter(x => x.adesao && x.adesao >= '2026-02-01' && x.adesao <= ate);
  const porDia = {};
  vendasPeriodo.forEach(x => { porDia[x.adesao] = (porDia[x.adesao] || 0) + 1; });
  const dmin = vendasPeriodo.map(x => x.adesao).sort()[0] || '2026-02-13';
  const dates = [], qtde = [], is_weekday = [];
  let d = new Date(dmin + 'T12:00:00'); const end = new Date(ate + 'T12:00:00');
  while (d <= end) {
    const iso = d.toISOString().slice(0, 10);
    dates.push(iso); qtde.push(porDia[iso] || 0);
    const dow = d.getUTCDay(); is_weekday.push(dow !== 0 && dow !== 6);
    d.setUTCDate(d.getUTCDate() + 1);
  }
  function mm(win) {
    const out = []; let ultimaMedia = 0;
    for (let i = 0; i < dates.length; i++) {
      if (is_weekday[i]) {
        const jan = []; let j = i;
        while (j >= 0 && jan.length < win) { if (is_weekday[j]) jan.push(qtde[j]); j--; }
        ultimaMedia = +(jan.reduce((s, v) => s + v, 0) / jan.length).toFixed(2);
      }
      out.push(ultimaMedia);
    }
    return out;
  }
  const daily = { dates, qtde, is_weekday, ma7: mm(7), ma15: mm(15), ma30: mm(30) };

  return {
    data: { kpis, unidade, representante, daily, meta: { gerado_em: '__HOJE__', periodo: 'Maio, Junho e Julho de 2026 (até ' + ate.split('-').reverse().join('/') + ')' } },
    diagnostico: { registros: regs.length, consultoresSemUnidade: [...dropConsultores] }
  };
};
