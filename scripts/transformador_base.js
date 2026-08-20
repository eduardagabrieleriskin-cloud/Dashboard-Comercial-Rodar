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

// De-para de cadastros duplicados no Siprov (mesma pessoa gravada duas vezes). Confirmado contra o
// cadastro oficial REPRESENTANTES.xlsx, que traz UMA entrada só para a pessoa. Chave e valor em MAIÚSCULAS.
const ALIAS_CONSULTOR = {
  'CARLOS EDUARDO CORDEIRO N': 'CARLOS EDUARDO CORDEIRO', // cadastro oficial: 1 entrada, CPF 040.609.489-69
};

// nomes das colunas no cabeçalho do BASE (linha 2 da planilha) — resolvidos para índice
// em tempo de execução, porque o Siprov já mudou a ordem/qtde de colunas entre exportações
// (ex.: layout de 04/08 tinha índices diferentes do layout de 10/08). Buscar por nome evita
// que uma mudança de leiaute quebre o pipeline silenciosamente (tudo cairia a zero).
const NOMES_COLUNA = {
  cpfAssociado: 'ASSOCIADO - CPF/CNPJ',
  situacao: 'BENEFÍCIO - SITUAÇÃO ATUAL',
  valorAjust: 'BENEFÍCIO - VALOR DA MENSALIDADE AJUSTADA',
  placa: 'VEÍCULO - PLACA DO VEÍCULO',
  cpfConsultor: 'BENEFÍCIO - CPF/CNPJ DO CONSULTOR',
  adesao: 'BENEFÍCIO - DATA DE ADESÃO',
  loja: 'BENEFÍCIO - LOJA - NOME FANTASIA',
  consultor: 'BENEFÍCIO - NOME DO CONSULTOR',
  representante: 'BENEFÍCIO - REPRESENTANTE',
};
function resolverColunas(headerRow) {
  const norm = s => (s || '').toString().trim().toUpperCase();
  const idx = {};
  for (const [chave, nomeCol] of Object.entries(NOMES_COLUNA)) {
    const i = headerRow.findIndex(h => norm(h) === norm(nomeCol));
    if (i === -1) throw new Error('coluna não encontrada no BASE: "' + nomeCol + '" (leiaute do Siprov mudou?)');
    idx[chave] = i;
  }
  return idx;
}
const MESES = ['2026-05', '2026-06', '2026-07', '2026-08'];

function titleCase(s) {
  return (s || '').toString().trim().toLowerCase().replace(/\s+/g, ' ')
    .split(' ').map(w => ['de','da','do','e','dos','das'].includes(w) ? w : (w.charAt(0).toUpperCase() + w.slice(1))).join(' ');
}
function canonicalizeUnidade(s) {
  // remove o código numérico na frente ("689 - Kcor Londrina" -> "Kcor Londrina") e padroniza capitalização
  const semCodigo = (s || '').toString().trim().replace(/^\d+\s*-\s*/, '');
  return titleCase(semCodigo);
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
function normPlaca(v) {
  const p = (v == null ? '' : v).toString().toUpperCase().replace(/[^A-Z0-9]/g, '');
  return p.length >= 6 ? p : '';
}
function toNum(v) {
  if (v == null || v === '') return 0;
  const n = parseFloat(v.toString().replace(/[^\d,.-]/g, '').replace(/\.(?=\d{3}(\D|$))/g, '').replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

// sinais (opcional): mapas auxiliares para preencher o que o Siprov deixou em branco, todos chaveados por
// PLACA — a única chave que existe nas três fontes (o Nº da cotação não serve: quando uma cotação vira
// subscrição ela SAI do Controle de Cotações, os dois arquivos são conjuntos disjuntos).
//   placa2rep / placa2unidade         - Controle de Subscrição V2 (quem transmitiu a proposta)
//   placa2repCot / placa2unidadeCot   - Controle de Cotações V2 (quem cotou) — pega o que a subscrição não tem
//   cpfAssoc2unidade                  - CPF do associado -> franqueado ("Relatório de Subscrição"), mais fraco
// Precedência: o que a BASE informa SEMPRE vence; estes sinais só preenchem lacunas. Ver scripts/ledger_placas.js,
// que monta o livro-razão por placa usado para validar esta cascata.
module.exports = function build(xlsxPath, ateISO, sinais) {
  const { placa2rep = {}, placa2unidade = {}, placa2repCot = {}, placa2unidadeCot = {},
    cpfAssoc2unidade = null } = sinais || {};
  const wb = XLSX.readFile(xlsxPath);
  const todasLinhas = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false });

  // Detecta se há linha de título antes do cabeçalho: alguns arquivos têm "BASE" na linha 0,
  // outros começam direto com cabeçalho. Verifica se linha 0 contém palavras-chave de coluna.
  const ehCabecalho = (linha) => linha.some(h => h && /ASSOCIADO|BENEFÍCIO|VEÍCULO|ENDEREÇO/.test(h));
  const temTitulo = !ehCabecalho(todasLinhas[0]);
  const linhaHeader = temTitulo ? 1 : 0;
  const dataInicio = temTitulo ? 2 : 1;

  const C = resolverColunas(todasLinhas[linhaHeader]);
  const rows = todasLinhas.slice(dataInicio);
  const SITU_VALIDAS = ['Ativo', 'Inadimplente', 'Cancelado', 'Inativo', 'Pendente'];
  const dropConsultores = new Set();

  // "Endosso Ativo" / "Endosso Inadimplente" contam como Ativo/Inadimplente na carteira (sem rótulo separado)
  const NORMALIZA_SITUACAO = { 'Endosso Ativo': 'Ativo', 'Endosso Inadimplente': 'Inadimplente' };

  // A partir de 13/08/2026 a exportação do Siprov passou a vir com "NOME DO CONSULTOR" vazio na maioria
  // das linhas (62%), deixando só a razão social da agência em "REPRESENTANTE" (96% preenchido). Por isso
  // a chave passou a ser consultor || representante, e a unidade dessas PJs é derivada por votação: nas
  // linhas em que a MESMA PJ aparece junto de um consultor que ESTÁ no mapa, herda-se a unidade dele.
  // (Sem isso, 2.5k placas — 65% da carteira — caíam em "(Sem Representante)"/"(Sem Unidade)".)
  const votosUnidadePorRepresentante = {};
  for (const r of rows) {
    const rep = (r[C.representante] || '').toString().trim();
    const uni = MAPA[(r[C.consultor] || '').toString().trim().toUpperCase()];
    if (!rep || !uni) continue;
    const v = votosUnidadePorRepresentante[rep] || (votosUnidadePorRepresentante[rep] = {});
    v[uni] = (v[uni] || 0) + 1;
  }
  function unidadeDerivada(rep) {
    const v = votosUnidadePorRepresentante[rep];
    if (!v) return null;
    return Object.entries(v).sort((a, b) => b[1] - a[1])[0][0];
  }

  // O PPM trunca o nome do representante em 30 caracteres ("Queila de Camargo Reis dos San") e às vezes
  // perde acento ("Mendonca" x "Mendonça"). Sem canonicalizar contra os nomes COMPLETOS da BASE, cada
  // pessoa vira dois representantes na tabela. Só aceita o casamento quando há UM candidato — na dúvida
  // mantém o nome como veio, para nunca fundir duas pessoas diferentes.
  const nrmNome = s => (s || '').toString().toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^A-Z ]/g, ' ').replace(/\s+/g, ' ').trim();
  const canonPorNorm = new Map();
  for (const r of rows) {
    const n = titleCase(r[C.consultor]);
    if (n) canonPorNorm.set(nrmNome(n), n);
  }
  function canonicalizarNome(raw) {
    if (!raw) return '';
    const k = nrmNome(raw);
    if (canonPorNorm.has(k)) return canonPorNorm.get(k);
    const cands = [...canonPorNorm.keys()].filter(c => c !== k && c.length >= 12 && k.length >= 12
      && (c.startsWith(k) || k.startsWith(c)));
    return cands.length === 1 ? canonPorNorm.get(cands[0]) : titleCase(raw);
  }

  const regs = [];
  for (const r of rows) {
    let situacao = (r[C.situacao] || '').toString().trim();
    situacao = NORMALIZA_SITUACAO[situacao] || situacao;
    if (!SITU_VALIDAS.includes(situacao)) continue;              // fora Recusado/vazio
    const loja = r[C.loja];
    if (ehTeste(r[C.consultor], loja)) continue;                 // fora testes
    let consultorRaw = (r[C.consultor] || '').toString().trim();
    consultorRaw = ALIAS_CONSULTOR[consultorRaw.toUpperCase()] || consultorRaw;
    const repRaw = (r[C.representante] || '').toString().trim();
    const placa = normPlaca(r[C.placa]);
    // Quem vendeu, em camadas: o consultor da BASE manda; sem ele, busca a PESSOA pela placa na Subscrição;
    // só então cai na agência. Sem a camada da placa, 2.2k placas ficavam na PJ (com zero vendas) enquanto
    // as vendas iam para a pessoa — estoque e venda do mesmo negócio em linhas diferentes.
    const quemVendeu = titleCase(consultorRaw)
      || (placa && placa2rep[placa] ? canonicalizarNome(placa2rep[placa]) : '')
      || (placa && placa2repCot[placa] ? canonicalizarNome(placa2repCot[placa]) : '')
      || titleCase(repRaw);
    // unidade em camadas, da mais confiável para a mais ampla; sem nenhuma => "(Sem Unidade)", mas o
    // registro NÃO é descartado (descartar fazia o total da carteira ficar abaixo do real — bug de 10/08).
    const unidadeRaw = MAPA[consultorRaw.toUpperCase()]                                   // 1. consultor no mapa
      || MAPA[repRaw.toUpperCase()]                                                       // 2. agência no mapa
      || (placa && placa2unidade[placa])                                                  // 3. franquia da placa (subscrição)
      || (placa && placa2unidadeCot[placa])                                               // 4. franquia da placa (cotação)
      || (cpfAssoc2unidade && cpfAssoc2unidade[(r[C.cpfAssociado] || '').toString().replace(/\D/g, '')]) // 5. franqueado do associado
      || unidadeDerivada(repRaw);                                                         // 6. voto da agência
    if (!unidadeRaw && quemVendeu) dropConsultores.add(quemVendeu); // loga pra atualizar o mapa depois
    const unidade = unidadeRaw ? canonicalizeUnidade(unidadeRaw) : '(Sem Unidade)';
    regs.push({
      situacao, consultor: quemVendeu || '(Sem Representante)', unidade, placa,
      cpfAssociado: (r[C.cpfAssociado] || '').toString().trim(),
      adesao: parseISO(r[C.adesao]),
      valor: toNum(r[C.valorAjust]),
    });
  }

  // PLACA -> representante já canonicalizado. As vendas (transformador_vendas) usam este mapa para
  // atribuir cada subscrição à MESMA pessoa que detém a placa aqui — sem isso, estoque e venda do mesmo
  // negócio caem em linhas diferentes (e variações de nome tipo "Hugo Sigaki"/"Hugo Eity Felix Sigaki"
  // viram dois representantes distintos).
  const placaParaRepresentante = {};
  regs.forEach(x => { if (x.placa) placaParaRepresentante[x.placa] = x.consultor; });

  // PLACA -> BENEFÍCIO - DATA DE ADESÃO (BASE) — pedido em 17/08: contar "vendas" pela data de adesão
  // oficial do contrato, não pela "Data Transmissão/Cálculo" do Controle de Subscrição (que é quando a
  // proposta foi calculada no PPM, podendo ser antes ou depois da adesão real). Usada em transformador_vendas.js
  // como a data preferencial; cai para a data da Subscrição só quando a placa ainda não está na BASE.
  const placaParaAdesao = {};
  regs.forEach(x => { if (x.placa && x.adesao) placaParaAdesao[x.placa] = x.adesao; });

  const ate = ateISO || '2026-12-31';
  // meses fechados mostram total do mês inteiro; só o mês atual (o de "ate") é parcial por natureza.
  const mesAtual = ate.slice(0, 7);
  const diaCorte = new Date(ate + 'T12:00:00').getDate();
  const ehVendaMes = (reg, mes) => reg.adesao && reg.adesao.slice(0, 7) === mes && reg.adesao <= ate;
  const ehAtivo = reg => reg.situacao === 'Ativo' || reg.situacao === 'Inadimplente';

  // ---- KPIs ----
  function vendasMes(mes) {
    const v = regs.filter(x => ehVendaMes(x, mes));
    const valor = v.reduce((s, x) => s + x.valor, 0);
    const comValor = v.filter(x => x.valor > 0).length;
    const clientes = new Set(v.map(x => x.cpfAssociado)).size;
    return { qtde: v.length, qtde_cliente: clientes, valor: +valor.toFixed(2), ticket_medio: v.length ? +(valor / v.length).toFixed(2) : 0,
      cobertura_valor_n: comValor, cobertura_valor_pct: v.length ? +(comValor / v.length).toFixed(3) : 0 };
  }
  // comparação justa: TUDO na mesma data de corte (dia do mês atual) — pedido da Eduarda em 17/08. Sem
  // isso, Maio/Junho/Julho (mês fechado, 30 dias) não são comparáveis com Agosto (parcial, só até o dia
  // do corte). "diasNoMes"/"limiteMes" travam no último dia de meses mais curtos (ex: corte dia 31 em mês
  // de 30 dias).
  function diasNoMes(mesIso) { const [ano, m] = mesIso.split('-').map(Number); return new Date(ano, m, 0).getDate(); }
  function limiteMes(mesIso) { return mesIso + '-' + String(Math.min(diaCorte, diasNoMes(mesIso))).padStart(2, '0'); }
  function qtdeMesAteDia(mes, dia) {
    return regs.filter(x => x.adesao && x.adesao.slice(0, 7) === mes && x.adesao <= mes + '-' + String(dia).padStart(2, '0')).length;
  }
  function variacaoJusta(mesAnterior, qtdeAnteriorCheio, mesAtualCard, qtdeAtualCard) {
    if (mesAtualCard !== mesAtual) return qtdeAnteriorCheio ? +((qtdeAtualCard - qtdeAnteriorCheio) / qtdeAnteriorCheio).toFixed(4) : 0;
    const baseComparavel = qtdeMesAteDia(mesAnterior, diaCorte);
    return baseComparavel ? +((qtdeAtualCard - baseComparavel) / baseComparavel).toFixed(4) : 0;
  }
  // taxa de perda POR MÊS, sempre na mesma data de corte (para os 4 meses, inclusive o atual — no mês
  // atual o "limite" já é a própria data de corte, então dá o mesmo resultado de sempre): das placas
  // vendidas naquele mês até o dia do corte, quantas já viraram Cancelado/Inativo até hoje.
  function perdaAteCorte(mesIso) {
    const limite = limiteMes(mesIso);
    const v = regs.filter(x => x.adesao && x.adesao.slice(0, 7) === mesIso && x.adesao <= limite);
    const perdidos = v.filter(x => x.situacao === 'Cancelado' || x.situacao === 'Inativo').length;
    return { qtde: v.length, perda_qtde: perdidos, perda_pct: v.length ? +(perdidos / v.length).toFixed(4) : 0 };
  }
  const vm = vendasMes('2026-05'), vj = vendasMes('2026-06'), vjl = vendasMes('2026-07'), vag = vendasMes('2026-08');
  // "até dia X" de cada mês FECHADO, na mesma data de corte do mês atual — pedido da Eduarda em 17/08:
  // dá pra comparar Maio/Junho/Julho com Agosto na mesma régua (todos "até dia 16"), não só o total cheio.
  const mesesTodos = ['2026-05', '2026-06', '2026-07', '2026-08'];
  const ateCortePorMes = {};
  mesesTodos.filter(m => m !== mesAtual).forEach(m => { ateCortePorMes[m] = qtdeMesAteDia(m, diaCorte); });
  const perdaPorMes = {};
  mesesTodos.forEach(m => { perdaPorMes[m] = perdaAteCorte(m); });
  const cont = s => regs.filter(x => x.situacao === s);
  const ativos = cont('Ativo'), inad = cont('Inadimplente'), canc = cont('Cancelado'), inat = cont('Inativo'), pend = cont('Pendente');
  const carteira = regs.filter(ehAtivo);
  const somaVal = arr => +arr.reduce((s, x) => s + x.valor, 0).toFixed(2);
  const universo = ativos.length + inad.length + canc.length + inat.length + pend.length;
  // dias úteis do período de julho até "ate" para o ritmo
  const kpis = {
    data_ultima_venda: regs.filter(x => x.adesao && x.adesao <= ate).map(x => x.adesao).sort().pop() || ate,
    data_referencia: ate, mes_atual: mesAtual, dia_corte: diaCorte,
    vendas_maio: vm, vendas_junho: vj, vendas_julho: vjl, vendas_agosto: vag,
    var_maio_junho_pct: variacaoJusta('2026-05', vm.qtde, '2026-06', vj.qtde),
    var_junho_julho_pct: variacaoJusta('2026-06', vj.qtde, '2026-07', vjl.qtde),
    var_julho_agosto_pct: variacaoJusta('2026-07', vjl.qtde, '2026-08', vag.qtde),
    var_junho_julho_ritmo_pct: 0,
    ate_corte_por_mes: ateCortePorMes,
    perda_por_mes: perdaPorMes,
    // variação do MÊS ATUAL especificamente (qual das três acima corresponde a ele) — usada pra colorir o
    // card do mês corrente (verde/amarelo/vermelho) sem precisar saber, do lado do template, qual mês é.
    variacao_mes_atual: mesAtual === '2026-08' ? variacaoJusta('2026-07', vjl.qtde, '2026-08', vag.qtde)
      : mesAtual === '2026-07' ? variacaoJusta('2026-06', vj.qtde, '2026-07', vjl.qtde)
      : mesAtual === '2026-06' ? variacaoJusta('2026-05', vm.qtde, '2026-06', vj.qtde)
      : null,
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
      const g = o[k] || (o[k] = { nome: k, vendas_maio: 0, vendas_junho: 0, vendas_julho: 0, vendas_agosto: 0,
        _cliMaio: new Set(), _cliJunho: new Set(), _cliJulho: new Set(), _cliAgosto: new Set(),
        ativos: 0, valor_ativos: 0, inadimplentes: 0, valor_inadimplentes: 0, cancelados: 0, inativos: 0, pendentes: 0,
        total: 0, valor_total: 0, _unidade: x.unidade });
      if (ehVendaMes(x, '2026-05')) { g.vendas_maio++; g._cliMaio.add(x.cpfAssociado); }
      if (ehVendaMes(x, '2026-06')) { g.vendas_junho++; g._cliJunho.add(x.cpfAssociado); }
      if (ehVendaMes(x, '2026-07')) { g.vendas_julho++; g._cliJulho.add(x.cpfAssociado); }
      if (ehVendaMes(x, '2026-08')) { g.vendas_agosto++; g._cliAgosto.add(x.cpfAssociado); }
      if (x.situacao === 'Ativo') { g.ativos++; g.valor_ativos += x.valor; }
      if (x.situacao === 'Inadimplente') { g.inadimplentes++; g.valor_inadimplentes += x.valor; }
      if (x.situacao === 'Cancelado') g.cancelados++;
      if (x.situacao === 'Inativo') g.inativos++;
      if (x.situacao === 'Pendente') g.pendentes++;
    }
    return Object.values(o).map(g => {
      g.vendas_maio_cliente = g._cliMaio.size; g.vendas_junho_cliente = g._cliJunho.size; g.vendas_julho_cliente = g._cliJulho.size; g.vendas_agosto_cliente = g._cliAgosto.size;
      delete g._cliMaio; delete g._cliJunho; delete g._cliJulho; delete g._cliAgosto;
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
  const daily = { dates, qtde, is_weekday };

  return {
    data: { kpis, unidade, representante, daily, meta: { gerado_em: '__HOJE__', periodo: 'Maio a Agosto de 2026 (até ' + ate.split('-').reverse().join('/') + ')' } },
    placaParaRepresentante,
    placaParaAdesao,
    diagnostico: { registros: regs.length, consultoresSemUnidade: [...dropConsultores] }
  };
};
