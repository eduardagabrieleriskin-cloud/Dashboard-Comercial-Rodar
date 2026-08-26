/*
 * Atualiza o Painel Comercial Rodar Mutual a partir de:
 *  - BASE_*.xlsx mais recente do Siprov (carteira, vendas, unidades, representantes)
 *  - Controle_de_Cotações_*.xlsx mais recente do PPM (cotações e conversão por consultor —
 *    é o relatório real do funil de cotação; NÃO usar o Controle de Subscrição aqui,
 *    que é uma etapa mais avançada e conta duplicado)
 * ambos salvos na pasta Downloads.
 *
 * Fluxo: acha os arquivos mais novos -> transforma (até ONTEM) -> injeta no
 * template -> valida -> commit + push (GitHub e Vercel atualizam via git).
 *
 * NÃO acessa Siprov nem PPM. Quem exporta os arquivos é a usuária; este script publica.
 */
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const XLSX = require('xlsx');
const buildBase = require('./transformador_base');
const buildConversao = require('./transformador_conversao');
const buildVendas = require('./transformador_vendas');
const lerSubscricao = require('./leitor_subscricao');
const lerDescontoEspecial = require('./leitor_desconto_especial');

const MESES = ['2026-05', '2026-06', '2026-07', '2026-08'];
const MES_NOME = { '2026-05': 'maio', '2026-06': 'junho', '2026-07': 'julho', '2026-08': 'agosto' };

// casamento de nome truncado do PPM (30 chars) x nome completo da BASE — mesma heurística usada em
// transformador_vendas.js/transformador_conversao.js, reaproveitada aqui para os arquivos avulsos de
// campanha (ex.: "Desconto especial semana do dia X"), que trazem só o nome do representante, sem placa
// necessariamente cadastrada na BASE ainda.
function norm(s) { return (s || '').toString().toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Z ]/g, ' ').replace(/\s+/g, ' ').trim(); }
function tok(s) { return norm(s).split(' ').filter(Boolean); }
function commonPrefix(ta, tb) { let n = 0; while (n < ta.length && n < tb.length && ta[n] === tb[n]) n++; return n; }
function scoreNomes(raw, baseName) {
  const nc = norm(raw), nb = norm(baseName);
  if (nc === nb) return 1000;
  if (nc.length >= 28 && nb.startsWith(nc)) return 900;
  if (nb.length >= 28 && nc.startsWith(nb)) return 900;
  const tc = tok(raw), tb = tok(baseName);
  let s = commonPrefix(tc, tb);
  if (s < 2 && tc.length >= 2 && tb.length >= 2 && tc[0] === tb[0] && tc[tc.length - 1] === tb[tb.length - 1]) s = 2;
  return s;
}
function acharRepresentante(raw, placa, placaParaRepresentante, representantesBase) {
  if (placa && placaParaRepresentante[placa]) return placaParaRepresentante[placa];
  const scored = representantesBase.map(r => ({ nome: r.nome, s: scoreNomes(raw, r.nome) })).filter(x => x.s > 0).sort((a, b) => b.s - a.s);
  if (!scored.length) return null;
  if (scored.filter(x => x.s === scored[0].s).length > 1) return null; // empate -> descarta
  return scored[0].nome;
}

// Planilhas avulsas de campanha (ex.: "Desconto especial semana do dia 14.xlsx") — adesões que aconteceram
// FORA do Controle de Subscrição normal e por isso não entrariam na contagem principal. Pedido em 17/08:
// somar essas adesões ao mês corrente (mesma régua de corte). Retorna registros extras no mesmo formato
// usado por buildVendas, para entrarem nos mesmos recálculos de porRep/registros logo abaixo.
function lerExtrasDeCampanha(downloads, mesAtual, representantesBase, placaParaRepresentante) {
  const arqs = fs.readdirSync(downloads).filter(f => /^desconto especial.*\.xlsx$/i.test(f));
  const extras = [];
  for (const f of arqs) {
    // dia usado pra filtro de corte: extrai do nome do arquivo ("...semana do dia 14...") — a planilha em
    // si não traz data por linha. Sem casar, cai no dia 1 do mês (entra em qualquer corte, nunca é excluído
    // por engano; só arrisca contar mesmo se o corte for antes do dia real do arquivo).
    const diaMatch = f.match(/dia\s*(\d{1,2})/i);
    const dia = diaMatch ? diaMatch[1].padStart(2, '0') : '01';
    const linhas = lerDescontoEspecial(path.join(downloads, f));
    let casadas = 0;
    for (const l of linhas) {
      const alvo = acharRepresentante(l.representante, l.placa, placaParaRepresentante, representantesBase);
      if (!alvo) continue;
      extras.push({ alvo, mes: mesAtual, dataISO: mesAtual + '-' + dia, associado: norm(l.associado) });
      casadas++;
    }
    log('adesões extras de campanha lidas de ' + f + ': ' + linhas.length + ' (casadas: ' + casadas + ')');
  }
  return extras;
}

// Substitui as contagens de "vendas" (placas fechadas) da BASE pelas do Controle de
// Subscrição — validado com o Kauan/Daiane: a BASE só conta quem está "Ativo" hoje,
// subcontando quem fechou no mês mas depois teve status alterado. Subscrição conta a
// proposta pela data de transmissão, independente do status atual — é o número certo.
// Valor/ticket médio continuam vindo da BASE (não houve validação separada para R$).
function aplicarVendasDaSubscricao(res, subscricaoPath, ate, downloads) {
  const { porRep, semMatch, registros } = buildVendas(subscricaoPath, res.data.representante, MESES, res.placaParaRepresentante, ate, res.placaParaAdesao);

  // soma as adesões de campanha (fora da Subscrição) ao mês corrente antes de qualquer recálculo abaixo
  const mesAtual = ate.slice(0, 7);
  if (downloads) {
    const extras = lerExtrasDeCampanha(downloads, mesAtual, res.data.representante, res.placaParaRepresentante);
    extras.forEach(x => {
      const bucket = porRep[x.alvo] && porRep[x.alvo][x.mes];
      if (!bucket) return;
      bucket.placas++;
      bucket.clientes.add(x.associado);
      registros.push(x);
    });
  }

  res.data.representante.forEach(r => {
    MESES.forEach(mIso => {
      const nome = MES_NOME[mIso];
      const b = porRep[r.nome][mIso];
      r['vendas_' + nome] = b.placas;
      r['vendas_' + nome + '_cliente'] = b.clientes.size;
    });
  });

  // reagrega por unidade a partir dos representantes já atualizados
  const porUnidade = {};
  res.data.representante.forEach(r => {
    const u = porUnidade[r.unidade] || (porUnidade[r.unidade] = {});
    MESES.forEach(mIso => {
      const nome = MES_NOME[mIso];
      u['vendas_' + nome] = (u['vendas_' + nome] || 0) + r['vendas_' + nome];
      u['vendas_' + nome + '_cliente'] = (u['vendas_' + nome + '_cliente'] || 0) + r['vendas_' + nome + '_cliente'];
    });
  });
  res.data.unidade.forEach(u => Object.assign(u, porUnidade[u.nome] || {}));

  // recalcula os totais de KPI (qtde/qtde_cliente) a partir da Subscrição; valor/ticket seguem da BASE
  const K = res.data.kpis;
  MESES.forEach(mIso => {
    const nome = MES_NOME[mIso];
    const totalPlacas = res.data.representante.reduce((s, r) => s + r['vendas_' + nome], 0);
    const totalClientes = res.data.representante.reduce((s, r) => s + r['vendas_' + nome + '_cliente'], 0);
    K['vendas_' + nome].qtde = totalPlacas;
    K['vendas_' + nome].qtde_cliente = totalClientes;
  });

  // variação justa: mês fechado = total cheio; só o par que termina no mês atual usa o mesmo corte de dia
  function qtdeAteDia(mesIso, dia) {
    const limite = mesIso + '-' + String(dia).padStart(2, '0');
    return registros.filter(x => x.mes === mesIso && x.dataISO <= limite).length;
  }
  function variacaoJusta(mesA, qtdeACheio, mesB, qtdeB) {
    if (mesB !== K.mes_atual) return qtdeACheio ? +((qtdeB - qtdeACheio) / qtdeACheio).toFixed(4) : 0;
    const base = qtdeAteDia(mesA, K.dia_corte);
    return base ? +((qtdeB - base) / base).toFixed(4) : 0;
  }
  K.var_maio_junho_pct = variacaoJusta('2026-05', K.vendas_maio.qtde, '2026-06', K.vendas_junho.qtde);
  K.var_junho_julho_pct = variacaoJusta('2026-06', K.vendas_junho.qtde, '2026-07', K.vendas_julho.qtde);
  K.var_julho_agosto_pct = variacaoJusta('2026-07', K.vendas_julho.qtde, '2026-08', K.vendas_agosto.qtde);

  // idem para ate_corte_por_mes / variacao_mes_atual (cards do painel, semáforo): tinham ficado com os
  // valores calculados dentro de transformador_base.js a partir da BASE, sem a correção da Subscrição
  // acima — o card do mês atual mostrava um semáforo/variação que não batia com o "vendas_agosto.qtde"
  // corrigido aqui do lado. Refaz os dois com a mesma fonte (registros da Subscrição).
  MESES.forEach(mIso => { K.ate_corte_por_mes[mIso] = qtdeAteDia(mIso, K.dia_corte); });
  delete K.ate_corte_por_mes[K.mes_atual];
  const idx = MESES.indexOf(K.mes_atual);
  K.variacao_mes_atual = idx > 0 ? variacaoJusta(MESES[idx - 1], K['vendas_' + MES_NOME[MESES[idx - 1]]].qtde, K.mes_atual, K['vendas_' + MES_NOME[K.mes_atual]].qtde) : null;

  return semMatch;
}

// Recalcula SOMENTE o gráfico "Ritmo de Vendas — Diário" a partir do Relatório de Subscrição:
// 1 linha = 1 proposta transmitida, sem filtro de Atividade/Status. Em 25/08 são 68 linhas.
// Os cards e o detalhamento continuam sendo PLACAS por Data de Adesão da BASE; substituir tudo por este
// relatório reduziria indevidamente o mês, pois uma proposta pode conter mais de uma placa.
function aplicarRitmoDoRelatorioSubscricao(res, relSubscricaoPath, ate) {
  const lerRelatorioSubscricao = require('./leitor_relatorio_subscricao');
  const rows = lerRelatorioSubscricao(relSubscricaoPath);
  const porDia = {};
  for (const row of rows) {
    if (!row.data || row.data > ate) continue;
    porDia[row.data] = (porDia[row.data] || 0) + 1;
  }
  const datasComVenda = Object.keys(porDia).filter(d => d >= '2026-02-01');
  const dmin = datasComVenda.sort()[0] || res.data.daily.dates[0] || '2026-02-13';
  const dates = [], qtde = [], is_weekday = [];
  let d = new Date(dmin + 'T12:00:00'); const end = new Date(ate + 'T12:00:00');
  while (d <= end) {
    const iso = d.toISOString().slice(0, 10);
    dates.push(iso); qtde.push(porDia[iso] || 0);
    const dow = d.getUTCDay(); is_weekday.push(dow !== 0 && dow !== 6);
    d.setUTCDate(d.getUTCDate() + 1);
  }
  res.data.daily = { dates, qtde, is_weekday };
  return { linhas: rows.length, noCorte: Object.values(porDia).reduce((s, n) => s + n, 0) };
}

// data mais recente presente no arquivo de cotações (coluna "Data solicitação", índice 4) —
// evita que o corte "mesmo dia do mês" distorça a conversão quando o arquivo de cotações
// está mais antigo que a BASE (ex: BASE de hoje mas cotações de alguns dias atrás)
function maxDataCotacoes(xlsxPath) {
  const wb = XLSX.readFile(xlsxPath);
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false }).slice(2);
  let max = '';
  for (const r of rows) {
    const m = String(r[4] || '').match(/(\d{2})\/(\d{2})\/(\d{4})/);
    if (!m) continue;
    const iso = m[3] + '-' + m[2] + '-' + m[1];
    if (iso > max) max = iso;
  }
  return max;
}

const REPO = path.resolve(__dirname, '..');
const DOWNLOADS = process.env.USERPROFILE ? path.join(process.env.USERPROFILE, 'Downloads') : 'C:/Users/eduar/Downloads';
const OUT = path.join(REPO, 'index.html');
const TEMPLATE = path.join(__dirname, 'template_painel.html');
const MARKER = path.join(__dirname, '.ultimo_base.txt');
const LOG = path.join(__dirname, 'atualizar.log');

function log(m) { const l = '[' + new Date().toISOString() + '] ' + m; console.log(l); try { fs.appendFileSync(LOG, l + '\n'); } catch (e) {} }
function fmtBR(dt) { return String(dt.getDate()).padStart(2, '0') + '/' + String(dt.getMonth() + 1).padStart(2, '0') + '/' + dt.getFullYear(); }
// corte padrão = ONTEM, não hoje. Regra da Eduarda (19/08): o dia corrente ainda está incompleto no
// Siprov (adesão entra com atraso), então incluí-lo derruba artificialmente o ritmo do mês.
function isoOntem() { const d = new Date(); d.setDate(d.getDate() - 1); return d.toISOString().slice(0, 10); }

function acharMaisRecente(regex) {
  const arqs = fs.readdirSync(DOWNLOADS)
    .filter(f => regex.test(f))
    .map(f => ({ f, full: path.join(DOWNLOADS, f), m: fs.statSync(path.join(DOWNLOADS, f)).mtimeMs }))
    .sort((a, b) => b.m - a.m);
  return arqs[0] || null;
}

try {
  const base = acharMaisRecente(/^BASE_\d{8}.*\.xlsx$/i);
  if (!base) { log('nenhum BASE_*.xlsx em Downloads. Nada a fazer.'); process.exit(0); }
  const cotacoes = acharMaisRecente(/^Controle_de_Cota.*\.xlsx$/i);
  let subscricao = acharMaisRecente(/^Controle_de_Subscri.*\.xlsx$/i);
  if (subscricao) {
    // arquivo pode ter sido exportado vazio (só título+cabeçalho, 0 linhas de dado) — nesse caso
    // NÃO usar (senão zera as vendas de todo mundo); cai no fallback via BASE, com aviso.
    const nRows = lerSubscricao(subscricao.full).length;
    if (nRows === 0) { log('AVISO: ' + subscricao.f + ' está vazio (0 linhas de dado) — ignorando, vendas usam a BASE.'); subscricao = null; }
    else log(subscricao.f + ': ' + nRows + ' subscrições lidas');
  }

  const relSubscricao = acharMaisRecente(/^Relat[óo]rio de Subscri.*\.xlsx$/i);
  const assinatura = base.f + '|' + Math.round(base.m) + '|' + (cotacoes ? cotacoes.f + '|' + Math.round(cotacoes.m) : 'sem-cotacoes')
    + '|' + (subscricao ? subscricao.f + '|' + Math.round(subscricao.m) : 'sem-subscricao')
    + '|' + (relSubscricao ? relSubscricao.f + '|' + Math.round(relSubscricao.m) : 'sem-relatorio-subscricao');
  const marca = fs.existsSync(MARKER) ? fs.readFileSync(MARKER, 'utf8').trim() : '';
  if (marca === assinatura && !process.env.FORCAR) { log('fontes mais recentes já publicadas (' + base.f + (cotacoes ? ' + ' + cotacoes.f : '') + '). Nada novo.'); process.exit(0); }

  const ate = process.env.ATE_OVERRIDE || isoOntem(); // ATE_OVERRIDE=YYYY-MM-DD força o corte (ex: "até ontem" pedido manualmente)
  log('fonte BASE: ' + base.f + ' | fonte Cotações: ' + (cotacoes ? cotacoes.f : '(nenhuma)') + ' | até ' + ate);

  // "Relatório de Subscrição" (export diferente do Controle de Subscrição V2): traz Franqueado por
  // associado, usado só para descobrir a UNIDADE linha a linha quando o Siprov exporta o consultor vazio.
  // NÃO serve para contar vendas — ali cada linha é 1 proposta (pode cobrir várias placas), o que daria
  // ~metade das placas reais. Opcional: sem ele, a unidade cai no voto por agência.
  let cpfAssoc2unidade = null;
  if (relSubscricao) {
    const wbR = XLSX.readFile(relSubscricao.full);
    const rowsR = XLSX.utils.sheet_to_json(wbR.Sheets[wbR.SheetNames[0]], { header: 1, raw: false }).slice(2);
    const votos = {};
    for (const r of rowsR) {
      const cpf = String(r[3] || '').replace(/\D/g, '');
      const franq = String(r[9] || '').trim();
      if (!cpf || !franq) continue;
      (votos[cpf] || (votos[cpf] = {}))[franq] = (votos[cpf][franq] || 0) + 1;
    }
    cpfAssoc2unidade = {};
    for (const [cpf, v] of Object.entries(votos)) cpfAssoc2unidade[cpf] = Object.entries(v).sort((a, b) => b[1] - a[1])[0][0];
    log('unidade por associado carregada de ' + relSubscricao.f + ' (' + Object.keys(cpfAssoc2unidade).length + ' CPFs)');
  }

  // Controle de Subscrição V2: PLACA -> representante/franquia. Chave exata, usada para preencher o que o
  // Siprov exportou vazio (ver transformador_base). Colunas: 7=Placa(s), 24=Franquia, 26=Representante.
  const placa2rep = {}, placa2unidade = {};
  if (subscricao) {
    for (const r of lerSubscricao(subscricao.full)) {
      for (const pn of r.placas) {
        if (r.representante) placa2rep[pn] = r.representante;
        if (r.franquia) placa2unidade[pn] = r.franquia;
      }
    }
    log('placas mapeadas na Subscrição: ' + Object.keys(placa2rep).length + ' (representante) / ' + Object.keys(placa2unidade).length + ' (franquia)');
  }

  // Controle de Cotações V2: PLACA -> representante/franquia de quem COTOU. Última rede antes de cair na
  // agência — resgata placas antigas que não aparecem no Controle de Subscrição (janela mais curta).
  // Colunas: 3=Placas, 4=Data solicitação, 6=Franquia, 7=Representante.
  const placa2repCot = {}, placa2unidadeCot = {};
  if (cotacoes) {
    const wbC = XLSX.readFile(cotacoes.full);
    const rowsC = XLSX.utils.sheet_to_json(wbC.Sheets[wbC.SheetNames[0]], { header: 1, raw: false }).slice(2);
    for (const r of rowsC) {
      const rep = String(r[7] || '').trim(), franq = String(r[6] || '').replace(/^\d+\s*-\s*/, '').trim();
      for (const p of String(r[3] || '').split(/[\s,\/]+/)) {
        const pn = p.toUpperCase().replace(/[^A-Z0-9]/g, '');
        if (pn.length < 6) continue;
        if (rep && !placa2repCot[pn]) placa2repCot[pn] = rep;
        if (franq && !placa2unidadeCot[pn]) placa2unidadeCot[pn] = franq;
      }
    }
    log('placas mapeadas nas Cotações: ' + Object.keys(placa2repCot).length);
  }

  const res = buildBase(base.full, ate, { placa2rep, placa2unidade, placa2repCot, placa2unidadeCot, cpfAssoc2unidade });
  if (relSubscricao) {
    const ritmo = aplicarRitmoDoRelatorioSubscricao(res, relSubscricao.full, ate);
    log('ritmo diário recalculado por Data Transmissão de ' + relSubscricao.f + ' (' + ritmo.linhas + ' linhas; ' + ritmo.noCorte + ' até o corte)');
  } else {
    log('AVISO: sem Relatório de Subscrição; ritmo diário permanece pela Data de Adesão da BASE.');
  }
  if (res.diagnostico.consultoresSemUnidade.length)
    log('AVISO: consultores sem unidade no mapa (caem em "(Sem Unidade)", NÃO são descartados — atualizar mapa_unidades.json): '
      + res.diagnostico.consultoresSemUnidade.join(', '));
  res.data.meta.gerado_em = fmtBR(new Date());

  // Vendas contadas direto do BASE (coluna BENEFÍCIO - DATA DE ADESÃO), não cruzando com Subscrição.
  // Isso garante que os números de vendas batem exatamente com a contagem manual no BASE.
  // (Antes: aplicarVendasDaSubscricao substituía pelos números da Subscrição, que era um conjunto diferente.)
  if (subscricao) {
    log('Subscrição carregada (' + subscricao.f + ') mas vendas contadas direto do BASE (data de adesão).');
  }

  // representantes/unidades sem NADA no período (0 na carteira e 0 vendas nos 4 meses) só poluem a tabela
  const vazio = x => !x.total && MESES.every(m => !x['vendas_' + MES_NOME[m]]);
  const removidos = res.data.representante.filter(vazio).map(x => x.nome);
  res.data.representante = res.data.representante.filter(x => !vazio(x));
  res.data.unidade = res.data.unidade.filter(x => !vazio(x));
  if (removidos.length) log('representantes sem movimento no período (removidos da tabela): ' + removidos.join(', '));

  let avisoConversao = '';
  let conversao = { consultores: [], totais: { total_cotado: 0, total_fechado: 0, conversao: 0,
    total_cotado_cliente: 0, total_fechado_cliente: 0, conversao_cliente: 0,
    por_mes: {}, consultores: 0, sem_cotacao_registrada: 0 }, meses: ['2026-05', '2026-06', '2026-07'] };
  if (cotacoes) {
    const maxCot = maxDataCotacoes(cotacoes.full);
    const ateCotacoes = maxCot && maxCot < ate ? maxCot : ate;
    if (ateCotacoes !== ate) log('cotações mais antigas que a BASE (até ' + maxCot + ') — usando essa data como corte, não "hoje".');
    conversao = buildConversao(cotacoes.full, base.full, ateCotacoes, res.data.representante);
    conversao.fonte_rotulo = cotacoes.f;
    log('conversão calculada: ' + conversao.consultores.length + ' consultores (' + conversao.totais.sem_cotacao_registrada + ' sem cotação casada) | '
      + conversao.totais.total_fechado + '/' + conversao.totais.total_cotado + ' (' + (conversao.totais.conversao * 100).toFixed(1) + '%)');
  } else {
    // Sem Controle_de_Cotações no Downloads a conversão zeraria e a tabela sumiria da tela — pior que
    // mostrar o número anterior. Então reaproveitamos o bloco de conversão já publicado no index.html e
    // marcamos na própria seção de qual export ele veio, para ninguém ler como se fosse do dia.
    // (O "Relatório de Cotações" NÃO serve de substituto: leiaute diferente e ~80% das linhas sem placa,
    //  o que quebra o cruzamento cotação->fechamento.)
    log('AVISO: sem Controle_de_Cotações em Downloads.');
    try {
      const anterior = fs.readFileSync(OUT, 'utf8')
        .match(/<script id="conversao-data" type="application\/json">([\s\S]*?)<\/script>/);
      const prev = anterior && JSON.parse(anterior[1]);
      if (prev && prev.consultores && prev.consultores.length) {
        conversao = prev;
        avisoConversao = 'Cotações desatualizadas: não havia <b>Controle de Cotações V2</b> novo nesta atualização, '
          + 'então esta seção repete os números de ' + (prev.fonte_rotulo || 'uma exportação anterior')
          + '. Os demais indicadores do painel são da BASE de hoje.';
        log('conversão reaproveitada do publish anterior (' + prev.consultores.length + ' consultores) — seção marcada como desatualizada.');
      } else log('sem conversão anterior para reaproveitar — seção ficará vazia.');
    } catch (e) { log('não deu para reaproveitar a conversão anterior: ' + e.message); }
  }

  let html = fs.readFileSync(TEMPLATE, 'utf8')
    .replace('__DATA__', JSON.stringify(res.data))
    .replace('__CONVERSAO__', JSON.stringify(conversao));
  html = html.replace('__AVISO_CONVERSAO__', avisoConversao
    ? '<div class="warn-note" style="border-left:3px solid #d97706"><b>&#9888; </b>' + avisoConversao + '</div>' : '');

  // validação
  const j = html.match(/<script id="dashboard-data" type="application\/json">([\s\S]*?)<\/script>/);
  JSON.parse(j[1]);
  const jc = html.match(/<script id="conversao-data" type="application\/json">([\s\S]*?)<\/script>/);
  JSON.parse(jc[1]);
  if (html.length < 20000) throw new Error('HTML suspeitosamente pequeno');
  // "(Sem Unidade)" agora é um grupo válido (não descartamos mais o registro) — só avisa, não bloqueia o deploy.
  const semU = res.data.representante.filter(r => !r.unidade || r.unidade === '(Sem Unidade)');
  if (semU.length) log('AVISO: representantes em "(Sem Unidade)": ' + semU.map(r => r.nome).join(', '));
  // "fechado" agora é por PERÍODO (placas fechadas no mês), não por safra de cotação — passar de 100% é
  // esperado (o que fecha em agosto foi cotado em julho, e há venda sem cotação registrada no PPM).
  const acima100 = conversao.consultores.filter(c => c.total_fechado > c.total_cotado);
  if (acima100.length) log('conversão >100% (esperado na leitura por período): ' + acima100.map(c => c.nome).join(', '));

  fs.writeFileSync(OUT, html);
  log('index.html gerado (' + (html.length / 1024).toFixed(0) + 'KB) | carteira=' + res.data.kpis.carteira_qtde
    + ' | reps=' + res.data.representante.length + ' | unidades=' + res.data.unidade.length
    + ' | consultores-conversao=' + conversao.consultores.length);

  execSync('git add index.html scripts', { cwd: REPO });
  if (!execSync('git status --porcelain', { cwd: REPO }).toString().trim()) {
    log('sem mudanças para commitar.'); fs.writeFileSync(MARKER, assinatura); process.exit(0);
  }
  execSync('git commit -m "auto: painel a partir de ' + base.f + (subscricao ? ' + ' + subscricao.f : '') + (cotacoes ? ' + ' + cotacoes.f : '') + ' (dados ate ' + ate + ')"', { cwd: REPO });
  execSync('git push origin main', { cwd: REPO });
  fs.writeFileSync(MARKER, assinatura);
  log('PUBLICADO com sucesso (GitHub + Vercel via git).');
} catch (e) {
  log('ERRO (nada publicado): ' + e.message);
  process.exit(1);
}
