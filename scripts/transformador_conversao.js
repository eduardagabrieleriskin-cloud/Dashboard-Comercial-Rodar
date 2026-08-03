/*
 * Calcula cotações mensais por consultor e a taxa de conversão por consultor.
 *
 * Fontes:
 *  - Controle de Cotações V2 (PPM/AMSS): o relatório real do funil de cotação —
 *    cada linha = 1 placa cotada, com Representante, Franquia, Data e Status.
 *    (NÃO usar o Controle de Subscrição aqui: é uma etapa mais avançada do funil
 *    e conta duplicado — confirmado com a Daiane em junho: Subscrição dava 78,
 *    Cotações V2 dá 40, que é o número certo, batendo com o PPM.)
 *  - BASE do Siprov: define o FECHAMENTO (placa com situação Ativo/Inadimplente)
 *  - representantesBase: array já canonicalizado (nome + unidade) vindo do transformador_base,
 *    para que TODO consultor da carteira apareça nesta tabela — mesmo sem cotação casada.
 *
 * O nome do consultor no arquivo do PPM vem truncado em 30 caracteres, então o
 * casamento de nomes usa: (1) igual, (2) prefixo (cobre o truncamento),
 * (3) maior prefixo de tokens em comum, (4) primeiro+último token iguais.
 * Em caso de empate entre dois consultores candidatos, a cotação fica sem
 * atribuição (não conta para nenhum dos dois) — evita atribuir errado.
 *
 * Cruzamento cotação->fechamento por PLACA (não por nome) — garante que a
 * conversão nunca passe de 100%.
 *
 * Uso: build(cotacoesXlsx, baseXlsx, ateISO, representantesBase) -> { consultores, totais, meses }
 */
const XLSX = require('xlsx');

const S = { codigo: 0, associado: 1, placa: 3, data: 4, franquia: 6, representante: 7, status: 8 };
const B = { situacao: 16, placa: 26 };
const MESES = ['2026-05', '2026-06', '2026-07', '2026-08'];

const np = s => (s || '').toString().toUpperCase().replace(/[^A-Z0-9]/g, '');
function norm(s) { return (s || '').toString().toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Z ]/g, ' ').replace(/\s+/g, ' ').trim(); }
function tok(s) { return norm(s).split(' ').filter(Boolean); }
function commonPrefix(ta, tb) { let n = 0; while (n < ta.length && n < tb.length && ta[n] === tb[n]) n++; return n; }
function scoreNomes(cotName, baseName) {
  const nc = norm(cotName), nb = norm(baseName);
  if (nc === nb) return 1000;
  if (nc.length >= 28 && nb.startsWith(nc)) return 900;   // truncado a 30 chars no PPM
  if (nb.length >= 28 && nc.startsWith(nb)) return 900;
  const tc = tok(cotName), tb = tok(baseName);
  let s = commonPrefix(tc, tb);
  if (s < 2 && tc.length >= 2 && tb.length >= 2 && tc[0] === tb[0] && tc[tc.length - 1] === tb[tb.length - 1]) s = 2;
  return s;
}
function iso(s) { if (!s) return null; const m = String(s).match(/(\d{2})\/(\d{2})\/(\d{4})/); return m ? m[3] + '-' + m[2] + '-' + m[1] : null; }
// "Conjunto" pode ter mais de uma placa numa célula só (ex: "NED5C33 JBK1F44") — considera fechada se QUALQUER placa da cotação estiver ativa/inadimplente
function algumaPlacaFechada(celula, fechadas) {
  const placas = (celula || '').toString().split(/[\s,/]+/).map(np).filter(p => p.length >= 6);
  return placas.some(p => fechadas.has(p));
}

module.exports = function build(cotacoesXlsx, baseXlsx, ateISO, representantesBase) {
  const ate = ateISO || '2026-12-31';
  const diaCorte = new Date(ate + 'T12:00:00').getDate();
  function limiteMes(mes) {
    const [ano, m] = mes.split('-').map(Number);
    const diasNoMes = new Date(ano, m, 0).getDate();
    return mes + '-' + String(Math.min(diaCorte, diasNoMes)).padStart(2, '0');
  }

  // fechamentos: placas ativas/inadimplentes na BASE
  const wbB = XLSX.readFile(baseXlsx);
  const base = XLSX.utils.sheet_to_json(wbB.Sheets[wbB.SheetNames[0]], { header: 1, raw: false }).slice(2);
  const fechadas = new Set();
  for (const r of base) {
    const p = np(r[B.placa]);
    if (p && ['Ativo', 'Inadimplente'].includes((r[B.situacao] || '').trim())) fechadas.add(p);
  }

  // cotações (Controle de Cotações V2 — funil real, não a Subscrição)
  const wbC = XLSX.readFile(cotacoesXlsx);
  const cot = XLSX.utils.sheet_to_json(wbC.Sheets[wbC.SheetNames[0]], { header: 1, raw: false }).slice(2);

  const nomesBase = representantesBase.map(r => r.nome);
  const rawNomesCot = [...new Set(cot.map(r => (r[S.representante] || '').trim()).filter(Boolean))];

  // para cada nome bruto do PPM, acha o melhor consultor da BASE (evita atribuição ambígua)
  const destino = {};
  for (const cn of rawNomesCot) {
    const scored = nomesBase.map(bn => ({ bn, s: scoreNomes(cn, bn) })).filter(x => x.s > 0).sort((a, b) => b.s - a.s);
    if (!scored.length) continue;
    const top = scored[0].s;
    if (scored.filter(x => x.s === top).length > 1) continue; // empate -> descarta (ambíguo)
    destino[cn] = scored[0].bn;
  }

  // agrega cotações por consultor da BASE (usando o destino calculado)
  // "cliente" = mesmo associado cotado mais de uma vez (mesmo dentro do mês) conta 1x na Visão Cliente
  const acc = {};
  representantesBase.forEach(r => { acc[r.nome] = {
    nome: r.nome, unidade: r.unidade, cot: {}, fech: {}, total_cot: 0, total_fech: 0,
    cliCot: {}, cliFech: {}, total_cliCot: new Set(), total_cliFech: new Set()
  }; });

  for (const r of cot) {
    const d = iso(r[S.data]);
    if (!d) continue;
    const mes = d.slice(0, 7);
    if (!MESES.includes(mes) || d > limiteMes(mes)) continue;
    const nomeBaseAlvo = destino[(r[S.representante] || '').trim()];
    if (!nomeBaseAlvo) continue;
    const a = acc[nomeBaseAlvo];
    const cliente = norm(r[S.associado]);
    const fechou = algumaPlacaFechada(r[S.placa], fechadas);

    a.cot[mes] = (a.cot[mes] || 0) + 1;
    a.total_cot++;
    if (fechou) { a.fech[mes] = (a.fech[mes] || 0) + 1; a.total_fech++; }

    if (cliente) {
      (a.cliCot[mes] || (a.cliCot[mes] = new Set())).add(cliente);
      a.total_cliCot.add(mes + '|' + cliente);
      if (fechou) {
        (a.cliFech[mes] || (a.cliFech[mes] = new Set())).add(cliente);
        a.total_cliFech.add(mes + '|' + cliente);
      }
    }
  }

  const consultores = Object.values(acc).map(a => {
    const o = {
      nome: a.nome, unidade: a.unidade,
      total_cotado: a.total_cot, total_fechado: a.total_fech,
      conversao: a.total_cot ? +(a.total_fech / a.total_cot).toFixed(4) : null,
      total_cotado_cliente: a.total_cliCot.size, total_fechado_cliente: a.total_cliFech.size,
      conversao_cliente: a.total_cliCot.size ? +(a.total_cliFech.size / a.total_cliCot.size).toFixed(4) : null,
    };
    MESES.forEach(m => {
      const cm = a.cot[m] || 0, fm = a.fech[m] || 0;
      o['cot_' + m] = cm; o['fech_' + m] = fm;
      o['conv_' + m] = cm ? +(fm / cm).toFixed(4) : null;
      const ccm = a.cliCot[m] ? a.cliCot[m].size : 0, cfm = a.cliFech[m] ? a.cliFech[m].size : 0;
      o['cot_cliente_' + m] = ccm; o['fech_cliente_' + m] = cfm;
      o['conv_cliente_' + m] = ccm ? +(cfm / ccm).toFixed(4) : null;
    });
    return o;
  }).sort((x, y) => y.total_cotado - x.total_cotado);

  // sanidade: conversão nunca > 100%
  const invalidos = consultores.filter(c => c.total_fechado > c.total_cotado || c.total_fechado_cliente > c.total_cotado_cliente);
  if (invalidos.length) throw new Error('conversao >100% em: ' + invalidos.map(c => c.nome).join(', '));

  const tc = consultores.reduce((s, c) => s + c.total_cotado, 0);
  const tf = consultores.reduce((s, c) => s + c.total_fechado, 0);
  const tcc = consultores.reduce((s, c) => s + c.total_cotado_cliente, 0);
  const tfc = consultores.reduce((s, c) => s + c.total_fechado_cliente, 0);
  const totais = {
    total_cotado: tc, total_fechado: tf, conversao: tc ? +(tf / tc).toFixed(4) : 0,
    total_cotado_cliente: tcc, total_fechado_cliente: tfc, conversao_cliente: tcc ? +(tfc / tcc).toFixed(4) : 0,
    por_mes: MESES.reduce((o, m) => {
      const c = consultores.reduce((s, x) => s + x['cot_' + m], 0);
      const f = consultores.reduce((s, x) => s + x['fech_' + m], 0);
      const cc = consultores.reduce((s, x) => s + x['cot_cliente_' + m], 0);
      const fc = consultores.reduce((s, x) => s + x['fech_cliente_' + m], 0);
      o[m] = { cotado: c, fechado: f, conversao: c ? +(f / c).toFixed(4) : 0,
        cotado_cliente: cc, fechado_cliente: fc, conversao_cliente: cc ? +(fc / cc).toFixed(4) : 0 };
      return o;
    }, {}),
    consultores: consultores.length,
    sem_cotacao_registrada: consultores.filter(c => c.total_cotado === 0).length
  };

  return { consultores, totais, meses: MESES };
};
