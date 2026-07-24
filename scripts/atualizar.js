/*
 * Atualiza o Painel Comercial Rodar Mutual a partir de:
 *  - BASE_*.xlsx mais recente do Siprov (carteira, vendas, unidades, representantes)
 *  - Controle_de_Subscrição_*.xlsx mais recente do PPM (cotações e conversão por consultor)
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
const buildBase = require('./transformador_base');
const buildConversao = require('./transformador_conversao');

const REPO = path.resolve(__dirname, '..');
const DOWNLOADS = 'C:/Users/eduar/Downloads';
const OUT = path.join(REPO, 'index.html');
const TEMPLATE = path.join(__dirname, 'template_painel.html');
const MARKER = path.join(__dirname, '.ultimo_base.txt');
const LOG = path.join(__dirname, 'atualizar.log');

function log(m) { const l = '[' + new Date().toISOString() + '] ' + m; console.log(l); try { fs.appendFileSync(LOG, l + '\n'); } catch (e) {} }
function fmtBR(dt) { return String(dt.getDate()).padStart(2, '0') + '/' + String(dt.getMonth() + 1).padStart(2, '0') + '/' + dt.getFullYear(); }
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
  const subscricao = acharMaisRecente(/^Controle_de_Subscri.*\.xlsx$/i);

  const assinatura = base.f + '|' + Math.round(base.m) + '|' + (subscricao ? subscricao.f + '|' + Math.round(subscricao.m) : 'sem-subscricao');
  const marca = fs.existsSync(MARKER) ? fs.readFileSync(MARKER, 'utf8').trim() : '';
  if (marca === assinatura) { log('fontes mais recentes já publicadas (' + base.f + (subscricao ? ' + ' + subscricao.f : '') + '). Nada novo.'); process.exit(0); }

  const ate = isoOntem();
  log('fonte BASE: ' + base.f + ' | fonte Subscrição: ' + (subscricao ? subscricao.f : '(nenhuma)') + ' | até ' + ate);

  const res = buildBase(base.full, ate);
  if (res.diagnostico.consultoresSemUnidade.length)
    log('consultores sem unidade (descartados): ' + res.diagnostico.consultoresSemUnidade.join(', '));
  res.data.meta.gerado_em = fmtBR(new Date());

  let conversao = { consultores: [], totais: { total_cotado: 0, total_fechado: 0, conversao: 0, por_mes: {}, consultores: 0 }, meses: ['2026-05', '2026-06', '2026-07'] };
  if (subscricao) {
    conversao = buildConversao(subscricao.full, base.full, ate);
    log('conversão calculada: ' + conversao.consultores.length + ' consultores | ' + conversao.totais.total_fechado + '/' + conversao.totais.total_cotado + ' (' + (conversao.totais.conversao * 100).toFixed(1) + '%)');
  } else {
    log('AVISO: sem Controle_de_Subscrição em Downloads — seção de cotações/conversão ficará vazia.');
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
  const semU = res.data.representante.filter(r => !r.unidade || r.unidade === '(Sem Unidade)');
  if (semU.length) throw new Error('representantes sem unidade: ' + semU.map(r => r.nome).join(', '));
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
  execSync('git commit -m "auto: painel a partir de ' + base.f + (subscricao ? ' + ' + subscricao.f : '') + ' (dados ate ' + ate + ')"', { cwd: REPO });
  execSync('git push origin main', { cwd: REPO });
  fs.writeFileSync(MARKER, assinatura);
  log('PUBLICADO com sucesso (GitHub + Vercel via git).');
} catch (e) {
  log('ERRO (nada publicado): ' + e.message);
  process.exit(1);
}
