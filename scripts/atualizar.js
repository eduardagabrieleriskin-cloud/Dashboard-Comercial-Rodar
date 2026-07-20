/*
 * Atualiza o Painel Comercial Rodar Mutual a partir do HTML mais recente
 * baixado na pasta Downloads (Dashboard_Comercial_RodarMutual_*.html).
 *
 * Passos: pega o arquivo mais novo -> extrai os dados -> aplica limpeza
 * (remove testes, garante que todo representante tenha unidade, corrige login)
 * -> grava index.html -> valida -> commit + push no GitHub.
 *
 * NÃO acessa o Siprov. Quem baixa o arquivo é a usuária; este script só publica.
 */
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const REPO = path.resolve(__dirname, '..');
const DOWNLOADS = 'C:/Users/eduar/Downloads';
const OUT = path.join(REPO, 'index.html');
const MARKER = path.join(__dirname, '.ultimo_processado.txt');
const LOG = path.join(__dirname, 'atualizar.log');
const MAPA = JSON.parse(fs.readFileSync(path.join(__dirname, 'mapa_unidades.json'), 'utf8'));

function log(msg) {
  const line = '[' + new Date().toISOString() + '] ' + msg;
  console.log(line);
  try { fs.appendFileSync(LOG, line + '\n'); } catch (e) {}
}

/* ---------- 1. localizar o HTML mais recente ---------- */
function acharFonte() {
  const arqs = fs.readdirSync(DOWNLOADS)
    .filter(f => /^Dashboard_Comercial_RodarMutual.*\.html$/i.test(f))
    .map(f => ({ f, full: path.join(DOWNLOADS, f), m: fs.statSync(path.join(DOWNLOADS, f)).mtimeMs }))
    .sort((a, b) => b.m - a.m);
  return arqs[0] || null;
}

/* ---------- 2. limpeza de dados ---------- */
function ehTeste(nome) {
  const n = (nome || '').trim().toLowerCase();
  return n === 'eduarda' || n === 'yara' || n === 'teste' || /^teste?\b/.test(n) || n.includes('(teste)');
}
const SOMA = ['vendas_maio','vendas_junho','vendas_julho','ativos','valor_ativos','inadimplentes','valor_inadimplentes','cancelados','inativos','pendentes','total','valor_total'];

function limpar(d) {
  const removidos = [], reatribuidos = [];
  // 1) reatribui unidade pelo mapa canônico quando faltar
  for (const r of d.representante) {
    const semU = !r.unidade || r.unidade === '(Sem Unidade)';
    const canon = MAPA[(r.nome || '').trim().toUpperCase()];
    if (semU && canon) { r.unidade = canon; reatribuidos.push(r.nome + ' -> ' + canon); }
  }
  // 2) filtra: fora testes, "(Sem Representante)" e quem continua sem unidade
  const antes = d.representante.length;
  d.representante = d.representante.filter(r => {
    if (ehTeste(r.nome)) { removidos.push(r.nome + ' (teste)'); return false; }
    if (r.nome === '(Sem Representante)') { removidos.push('(Sem Representante)'); return false; }
    if (!r.unidade || r.unidade === '(Sem Unidade)') { removidos.push(r.nome + ' (sem unidade)'); return false; }
    return true;
  });

  // 3) reconstrói a tabela de UNIDADES a partir dos representantes limpos (consistência total)
  const uMap = {};
  for (const r of d.representante) {
    const u = uMap[r.unidade] || (uMap[r.unidade] = { nome: r.unidade });
    for (const f of SOMA) u[f] = (u[f] || 0) + (r[f] || 0);
  }
  for (const u of Object.values(uMap)) {
    const uni = u.total + u.cancelados + u.inativos + u.pendentes;
    u.pct_inadimplencia = uni ? +(u.inadimplentes / uni).toFixed(4) : 0;
    u.pct_perda = uni ? +((u.cancelados + u.inativos) / uni).toFixed(4) : 0;
  }
  d.unidade = Object.values(uMap).sort((a, b) => b.ativos - a.ativos);

  // 4) recalcula os KPIs de contagem/valor a partir dos representantes limpos
  const K = d.kpis, T = {};
  for (const f of SOMA) T[f] = d.representante.reduce((s, r) => s + (r[f] || 0), 0);
  K.ativos_qtde = T.ativos; K.ativos_valor = +T.valor_ativos.toFixed(2);
  K.inadimplentes_qtde = T.inadimplentes; K.inadimplentes_valor = +T.valor_inadimplentes.toFixed(2);
  K.cancelados_qtde = T.cancelados; K.inativos_qtde = T.inativos; K.pendentes_qtde = T.pendentes;
  K.carteira_qtde = T.ativos + T.inadimplentes;
  K.carteira_valor = +(T.valor_ativos + T.valor_inadimplentes).toFixed(2);
  K.carteira_ticket_medio = K.carteira_qtde ? +(K.carteira_valor / K.carteira_qtde).toFixed(2) : 0;
  K.total_universo_qtde = T.ativos + T.inadimplentes + T.cancelados + T.inativos + T.pendentes;
  K.pct_inadimplencia = K.total_universo_qtde ? +(K.inadimplentes_qtde / K.total_universo_qtde).toFixed(4) : 0;
  K.pct_perda = K.total_universo_qtde ? +((T.cancelados + T.inativos) / K.total_universo_qtde).toFixed(4) : 0;
  // vendas por mês: contagem recalculada dos reps; valores/ticket mantidos da origem (ajustando ticket)
  if (K.vendas_maio)  { K.vendas_maio.qtde  = T.vendas_maio;  K.vendas_maio.ticket_medio  = T.vendas_maio  ? +(K.vendas_maio.valor  / T.vendas_maio ).toFixed(2) : 0; }
  if (K.vendas_junho) { K.vendas_junho.qtde = T.vendas_junho; K.vendas_junho.ticket_medio = T.vendas_junho ? +(K.vendas_junho.valor / T.vendas_junho).toFixed(2) : 0; }
  if (K.vendas_julho) { K.vendas_julho.qtde = T.vendas_julho; K.vendas_julho.ticket_medio = T.vendas_julho ? +(K.vendas_julho.valor / T.vendas_julho).toFixed(2) : 0; }
  if (K.vendas_maio && K.vendas_junho) K.var_maio_junho_pct = K.vendas_maio.qtde ? +((K.vendas_junho.qtde - K.vendas_maio.qtde) / K.vendas_maio.qtde).toFixed(4) : 0;

  log('limpeza: ' + antes + ' -> ' + d.representante.length + ' reps | reatribuidos=' + reatribuidos.length
    + ' | removidos=' + removidos.length + (removidos.length ? ' [' + removidos.join('; ') + ']' : '')
    + ' | unidades=' + d.unidade.length + ' | carteira=' + K.carteira_qtde);
  return d;
}

/* ---------- 3. correção do login (sha256 puro) ---------- */
function corrigirLogin(html) {
  const oldSha = /async function sha256\(text\)\{[\s\S]*?crypto\.subtle[\s\S]*?\n\}/;
  if (oldSha.test(html)) {
    const puro = "function sha256(ascii){function rr(v,a){return (v>>>a)|(v<<(32-a));}var mathPow=Math.pow,maxWord=mathPow(2,32),result='',words=[],asciiBitLength=ascii.length*8;var hash=sha256.h=sha256.h||[],k=sha256.k=sha256.k||[],primeCounter=k.length,isComposite={};for(var candidate=2;primeCounter<64;candidate++){if(!isComposite[candidate]){for(var i=0;i<313;i+=candidate){isComposite[i]=candidate;}hash[primeCounter]=(mathPow(candidate,.5)*maxWord)|0;k[primeCounter++]=(mathPow(candidate,1/3)*maxWord)|0;}}ascii+='\\x80';while(ascii.length%64-56)ascii+='\\x00';for(i=0;i<ascii.length;i++){var j=ascii.charCodeAt(i);if(j>>8)return;words[i>>2]|=j<<((3-i)%4)*8;}words[words.length]=((asciiBitLength/maxWord)|0);words[words.length]=(asciiBitLength);for(j=0;j<words.length;){var w=words.slice(j,j+=16),oldHash=hash;hash=hash.slice(0,8);for(i=0;i<64;i++){var w15=w[i-15],w2=w[i-2];var a=hash[0],e=hash[4];var temp1=hash[7]+(rr(e,6)^rr(e,11)^rr(e,25))+((e&hash[5])^((~e)&hash[6]))+k[i]+(w[i]=(i<16)?w[i]:(w[i-16]+(rr(w15,7)^rr(w15,18)^(w15>>>3))+w[i-7]+(rr(w2,17)^rr(w2,19)^(w2>>>10)))|0);var temp2=(rr(a,2)^rr(a,13)^rr(a,22))+((a&hash[1])^(a&hash[2])^(hash[1]&hash[2]));hash=[(temp1+temp2)|0].concat(hash);hash[4]=(hash[4]+temp1)|0;}for(i=0;i<8;i++){hash[i]=(hash[i]+oldHash[i])|0;}}for(i=0;i<8;i++){for(j=3;j+1;j--){var b=(hash[i]>>(j*8))&255;result+=((b<16)?0:'')+b.toString(16);}}return result;}";
    html = html.replace(oldSha, puro);
    html = html.replace('const hash = await sha256(val);', 'const hash = sha256(val);');
    log('login: crypto.subtle substituido por sha256 puro');
  }
  return html;
}

/* ---------- 4. validação ---------- */
function validar(html) {
  const jsons = [...html.matchAll(/<script id="[^"]+" type="application\/json">([\s\S]*?)<\/script>/g)];
  for (const j of jsons) JSON.parse(j[1]); // lança se inválido
  if (!/id="dashboard-data"/.test(html)) throw new Error('faltou dashboard-data');
  if (html.length < 20000) throw new Error('HTML suspeitosamente pequeno');
  return true;
}

/* ---------- 5. main ---------- */
try {
  const fonte = acharFonte();
  if (!fonte) { log('nenhum arquivo Dashboard_Comercial_RodarMutual_*.html em Downloads. Nada a fazer.'); process.exit(0); }
  const marcador = fs.existsSync(MARKER) ? fs.readFileSync(MARKER, 'utf8').trim() : '';
  const assinatura = fonte.f + '|' + Math.round(fonte.m);
  if (marcador === assinatura) { log('arquivo mais recente (' + fonte.f + ') já foi publicado. Nada novo.'); process.exit(0); }

  log('fonte: ' + fonte.f);
  let html = fs.readFileSync(fonte.full, 'utf8');
  const m = html.match(/(<script id="dashboard-data" type="application\/json">)([\s\S]*?)(<\/script>)/);
  if (!m) throw new Error('dashboard-data não encontrado no arquivo de origem');
  const d = limpar(JSON.parse(m[2]));
  html = html.replace(m[0], m[1] + JSON.stringify(d) + m[3]);
  html = corrigirLogin(html);
  validar(html);
  fs.writeFileSync(OUT, html);
  log('index.html gravado (' + (html.length / 1024).toFixed(0) + 'KB)');

  execSync('git add index.html scripts', { cwd: REPO });
  const st = execSync('git status --porcelain', { cwd: REPO }).toString().trim();
  if (!st) { log('sem mudanças para commitar.'); fs.writeFileSync(MARKER, assinatura); process.exit(0); }
  execSync('git commit -m "auto: atualiza painel a partir de ' + fonte.f + '"', { cwd: REPO });
  execSync('git push origin main', { cwd: REPO });
  fs.writeFileSync(MARKER, assinatura);
  log('PUBLICADO com sucesso.');
} catch (e) {
  log('ERRO (nada publicado): ' + e.message);
  process.exit(1);
}
