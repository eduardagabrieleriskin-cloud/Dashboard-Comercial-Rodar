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
 * Uso: build(cotacoesXlsx, baseXlsx, ateISO, representantesBase, janela) -> { consultores, totais, meses }
 *   janela: [{iso,nome,campo}] mais antigo -> atual, vindo de res.data.kpis.janela (transformador_base.js)
 *   — os campos "vendas_<nome>" lidos de representantesBase têm que ser exatamente os mesmos nomes que
 *   o transformador_base gerou pra esta mesma janela de 4 meses (janela móvel, não mais fixa mai-ago).
 */
const XLSX = require('xlsx');

const S = { codigo: 0, associado: 1, placa: 3, data: 4, franquia: 6, representante: 7, status: 8 };

// índices das colunas do BASE resolvidos por nome (não por posição fixa) — o Siprov já
// mudou a ordem/qtde de colunas entre exportações (ver transformador_base.js).
function resolverColunasBase(headerRow) {
  const norm = s => (s || '').toString().trim().toUpperCase();
  const acha = nome => {
    const i = headerRow.findIndex(h => norm(h) === norm(nome));
    if (i === -1) throw new Error('coluna não encontrada no BASE: "' + nome + '" (leiaute do Siprov mudou?)');
    return i;
  };
  return { situacao: acha('BENEFÍCIO - SITUAÇÃO ATUAL'), placa: acha('VEÍCULO - PLACA DO VEÍCULO') };
}

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

module.exports = function build(cotacoesXlsx, baseXlsx, ateISO, representantesBase, janela) {
  const MESES = janela.map(j => j.iso);
  const MES_NOME = Object.fromEntries(janela.map(j => [j.iso, j.nome]));
  const ate = ateISO || '2026-12-31';
  const diaCorte = new Date(ate + 'T12:00:00').getDate();
  function limiteMes(mes) {
    const [ano, m] = mes.split('-').map(Number);
    const diasNoMes = new Date(ano, m, 0).getDate();
    return mes + '-' + String(Math.min(diaCorte, diasNoMes)).padStart(2, '0');
  }

  // fechamentos: placas ativas/inadimplentes na BASE
  const wbB = XLSX.readFile(baseXlsx);
  const todasLinhasBase = XLSX.utils.sheet_to_json(wbB.Sheets[wbB.SheetNames[0]], { header: 1, raw: false });

  // Detecta se há linha de título: alguns arquivos têm "BASE" na linha 0, outros começam direto com cabeçalho
  const ehCabecalho = (linha) => linha.some(h => h && /ASSOCIADO|BENEFÍCIO|VEÍCULO|ENDEREÇO/.test(h));
  const temTitulo = !ehCabecalho(todasLinhasBase[0]);
  const linhaHeader = temTitulo ? 1 : 0;
  const dataInicio = temTitulo ? 2 : 1;

  const B = resolverColunasBase(todasLinhasBase[linhaHeader]);
  const base = todasLinhasBase.slice(dataInicio);
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

  // FECHADO = placas efetivamente fechadas no mês (mesma fonte da tabela de Detalhamento), para as duas
  // tabelas mostrarem o mesmo número (Kauan em agosto = 26).
  // COTADO = cotações do arquivo + as fechadas. O Controle de Cotações só guarda o que NÃO converteu —
  // quando a cotação vira subscrição ela SAI do arquivo (verificado: zero interseção de nº de cotação entre
  // os dois). Somando as fechadas de volta, o denominador vira o funil real do mês e, por construção,
  // FECHADO nunca passa de COTADO.
  const vendasPorNome = {};
  representantesBase.forEach(r => { vendasPorNome[r.nome] = r; });

  const consultores = Object.values(acc).map(a => {
    const V = vendasPorNome[a.nome] || {};
    const o = { nome: a.nome, unidade: a.unidade };
    let tc = 0, tf = 0, tcc = 0, tfc = 0;
    MESES.forEach(m => {
      const fm = V['vendas_' + MES_NOME[m]] || 0;
      const cfm = V['vendas_' + MES_NOME[m] + '_cliente'] || 0;
      const cm = (a.cot[m] || 0) + fm;                                  // não convertidas + fechadas
      const ccm = (a.cliCot[m] ? a.cliCot[m].size : 0) + cfm;
      tc += cm; tf += fm; tcc += ccm; tfc += cfm;
      o['cot_' + m] = cm; o['fech_' + m] = fm;
      o['conv_' + m] = cm ? +(fm / cm).toFixed(4) : null;
      o['cot_cliente_' + m] = ccm; o['fech_cliente_' + m] = cfm;
      o['conv_cliente_' + m] = ccm ? +(cfm / ccm).toFixed(4) : null;
    });
    o.total_cotado = tc; o.total_fechado = tf;
    o.total_cotado_cliente = tcc; o.total_fechado_cliente = tfc;
    o.conversao = tc ? +(tf / tc).toFixed(4) : null;
    o.conversao_cliente = tcc ? +(tfc / tcc).toFixed(4) : null;
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
