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

const MESES = ['2026-05', '2026-06', '2026-07', '2026-08'];
const MES_NOME = { '2026-05': 'maio', '2026-06': 'junho', '2026-07': 'julho', '2026-08': 'agosto' };

// Substitui as contagens de "vendas" (placas fechadas) da BASE pelas do Controle de
// Subscrição — validado com o Kauan/Daiane: a BASE só conta quem está "Ativo" hoje,
// subcontando quem fechou no mês mas depois teve status alterado. Subscrição conta a
// proposta pela data de transmissão, independente do status atual — é o número certo.
// Valor/ticket médio continuam vindo da BASE (não houve validação separada para R$).
function aplicarVendasDaSubscricao(res, subscricaoPath, ate) {
  const { porRep, semMatch, registros } = buildVendas(subscricaoPath, res.data.representante, MESES, res.placaParaRepresentante, ate);

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

  return semMatch;
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
function isoHoje() { return new Date().toISOString().slice(0, 10); }

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
    const wbSub = XLSX.readFile(subscricao.full);
    const nRows = XLSX.utils.sheet_to_json(wbSub.Sheets[wbSub.SheetNames[0]], { header: 1, raw: false }).slice(2).length;
    if (nRows === 0) { log('AVISO: ' + subscricao.f + ' está vazio (0 linhas de dado) — ignorando, vendas usam a BASE.'); subscricao = null; }
  }

  const assinatura = base.f + '|' + Math.round(base.m) + '|' + (cotacoes ? cotacoes.f + '|' + Math.round(cotacoes.m) : 'sem-cotacoes')
    + '|' + (subscricao ? subscricao.f + '|' + Math.round(subscricao.m) : 'sem-subscricao');
  const marca = fs.existsSync(MARKER) ? fs.readFileSync(MARKER, 'utf8').trim() : '';
  if (marca === assinatura) { log('fontes mais recentes já publicadas (' + base.f + (cotacoes ? ' + ' + cotacoes.f : '') + '). Nada novo.'); process.exit(0); }

  const ate = process.env.ATE_OVERRIDE || isoHoje(); // ATE_OVERRIDE=YYYY-MM-DD força o corte (ex: "até ontem" pedido manualmente)
  log('fonte BASE: ' + base.f + ' | fonte Cotações: ' + (cotacoes ? cotacoes.f : '(nenhuma)') + ' | até ' + ate);

  // "Relatório de Subscrição" (export diferente do Controle de Subscrição V2): traz Franqueado por
  // associado, usado só para descobrir a UNIDADE linha a linha quando o Siprov exporta o consultor vazio.
  // NÃO serve para contar vendas — ali cada linha é 1 proposta (pode cobrir várias placas), o que daria
  // ~metade das placas reais. Opcional: sem ele, a unidade cai no voto por agência.
  let cpfAssoc2unidade = null;
  const relSubscricao = acharMaisRecente(/^Relat[óo]rio de Subscri.*\.xlsx$/i);
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
    const wbSub = XLSX.readFile(subscricao.full);
    const rowsSub = XLSX.utils.sheet_to_json(wbSub.Sheets[wbSub.SheetNames[0]], { header: 1, raw: false }).slice(2);
    for (const r of rowsSub) {
      const rep = String(r[26] || '').trim(), franq = String(r[24] || '').trim();
      for (const p of String(r[7] || '').split(/[\s,\/]+/)) {
        const pn = p.toUpperCase().replace(/[^A-Z0-9]/g, '');
        if (pn.length < 6) continue;
        if (rep) placa2rep[pn] = rep;
        if (franq) placa2unidade[pn] = franq;
      }
    }
    log('placas mapeadas na Subscrição: ' + Object.keys(placa2rep).length + ' (representante) / ' + Object.keys(placa2unidade).length + ' (franquia)');
  }

  const res = buildBase(base.full, ate, { placa2rep, placa2unidade, cpfAssoc2unidade });
  if (res.diagnostico.consultoresSemUnidade.length)
    log('AVISO: consultores sem unidade no mapa (caem em "(Sem Unidade)", NÃO são descartados — atualizar mapa_unidades.json): '
      + res.diagnostico.consultoresSemUnidade.join(', '));
  res.data.meta.gerado_em = fmtBR(new Date());

  if (subscricao) {
    const semMatch = aplicarVendasDaSubscricao(res, subscricao.full, ate);
    log('vendas recalculadas a partir de ' + subscricao.f + ' (fonte validada — BASE subcontava quem saiu de "Ativo")'
      + (semMatch.length ? ' | consultores da Subscrição sem match: ' + semMatch.join(', ') : ''));
  } else {
    log('AVISO: sem Controle_de_Subscrição em Downloads — vendas usam a BASE (pode subcontar).');
  }

  // representantes/unidades sem NADA no período (0 na carteira e 0 vendas nos 4 meses) só poluem a tabela
  const vazio = x => !x.total && MESES.every(m => !x['vendas_' + MES_NOME[m]]);
  const removidos = res.data.representante.filter(vazio).map(x => x.nome);
  res.data.representante = res.data.representante.filter(x => !vazio(x));
  res.data.unidade = res.data.unidade.filter(x => !vazio(x));
  if (removidos.length) log('representantes sem movimento no período (removidos da tabela): ' + removidos.join(', '));

  let conversao = { consultores: [], totais: { total_cotado: 0, total_fechado: 0, conversao: 0,
    total_cotado_cliente: 0, total_fechado_cliente: 0, conversao_cliente: 0,
    por_mes: {}, consultores: 0, sem_cotacao_registrada: 0 }, meses: ['2026-05', '2026-06', '2026-07'] };
  if (cotacoes) {
    const maxCot = maxDataCotacoes(cotacoes.full);
    const ateCotacoes = maxCot && maxCot < ate ? maxCot : ate;
    if (ateCotacoes !== ate) log('cotações mais antigas que a BASE (até ' + maxCot + ') — usando essa data como corte, não "hoje".');
    conversao = buildConversao(cotacoes.full, base.full, ateCotacoes, res.data.representante);
    log('conversão calculada: ' + conversao.consultores.length + ' consultores (' + conversao.totais.sem_cotacao_registrada + ' sem cotação casada) | '
      + conversao.totais.total_fechado + '/' + conversao.totais.total_cotado + ' (' + (conversao.totais.conversao * 100).toFixed(1) + '%)');
  } else {
    log('AVISO: sem Controle_de_Cotações em Downloads — seção de cotações/conversão ficará vazia.');
  }

  let html = fs.readFileSync(TEMPLATE, 'utf8')
    .replace('__DATA__', JSON.stringify(res.data))
    .replace('__CONVERSAO__', JSON.stringify(conversao));

  // validação
  const j = html.match(/<script id="dashboard-data" type="application\/json">([\s\S]*?)<\/script>/);
  JSON.parse(j[1]);
  const jc = html.match(/<script id="conversao-data" type="application\/json">([\s\S]*?)<\/script>/);
  JSON.parse(jc[1]);
  if (html.length < 20000) throw new Error('HTML suspeitosamente pequeno');
  // "(Sem Unidade)" agora é um grupo válido (não descartamos mais o registro) — só avisa, não bloqueia o deploy.
  const semU = res.data.representante.filter(r => !r.unidade || r.unidade === '(Sem Unidade)');
  if (semU.length) log('AVISO: representantes em "(Sem Unidade)": ' + semU.map(r => r.nome).join(', '));
  const acima100 = conversao.consultores.filter(c => c.total_fechado > c.total_cotado);
  if (acima100.length) throw new Error('conversão >100% em: ' + acima100.map(c => c.nome).join(', '));

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
