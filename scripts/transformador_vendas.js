/*
 * Calcula "Vendas" (placas fechadas) por mês, por consultor/unidade, a partir do
 * Controle de Subscrição V2 (PPM) — validado com a Daiane/Kauan como a fonte
 * correta de "fechou" (conta a proposta pela data de transmissão, não depende
 * do status "Ativo" na foto de hoje da BASE, que subcontava).
 *
 * O nome do representante no arquivo do PPM vem truncado em 30 caracteres —
 * mesmo casamento de nomes usado no transformador_conversao.js.
 *
 * Uso: build(subscricaoXlsx, representantesBase, meses) -> Map nome -> {mes: {placa, clientes:Set}}
 */
const lerSubscricao = require('./leitor_subscricao');

function norm(s) {
  return (s || '').toString().replace(/^[A-Za-z]\/\s*/, '') // remove prefixo tipo "G/ " (grupo/franquia colado no nome)
    .toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Z ]/g, ' ').replace(/\s+/g, ' ').trim();
}
function tok(s) { return norm(s).split(' ').filter(Boolean); }
function commonPrefix(ta, tb) { let n = 0; while (n < ta.length && n < tb.length && ta[n] === tb[n]) n++; return n; }
function scoreNomes(rawName, baseName) {
  const nc = norm(rawName), nb = norm(baseName);
  if (nc === nb) return 1000;
  if (nc.length >= 28 && nb.startsWith(nc)) return 900;
  if (nb.length >= 28 && nc.startsWith(nb)) return 900;
  const tc = tok(rawName), tb = tok(baseName);
  let s = commonPrefix(tc, tb);
  if (s < 2 && tc.length >= 2 && tb.length >= 2 && tc[0] === tb[0] && tc[tc.length - 1] === tb[tb.length - 1]) s = 2;
  return s;
}
function iso(s) { if (!s) return null; const m = String(s).match(/(\d{2})\/(\d{2})\/(\d{4})/); return m ? m[3] + '-' + m[2] + '-' + m[1] : null; }

// placaParaRepresentante (opcional): PLACA -> representante já canonicalizado pela BASE. Quando a placa da
// subscrição existe na BASE, a venda é atribuída a QUEM DETÉM A PLACA lá — não ao nome escrito no PPM.
// Isso mantém estoque e venda do mesmo negócio na mesma linha e resolve, de uma vez: (a) variações de nome
// ("Hugo Sigaki" vs "Hugo Eity Felix Sigaki"), (b) divergência de atribuição entre os dois sistemas — a BASE
// vence por ser o sistema de registro da carteira. O casamento por nome fica só como fallback.
// ateISO: data de corte. OBRIGATÓRIA — sem ela o mês corrente entrava INTEIRO (inclusive dias posteriores
// ao corte), inflando as vendas e fazendo esta tabela divergir da de Conversão, que já respeitava o corte.
module.exports = function build(subscricaoXlsx, representantesBase, meses, placaParaRepresentante, ateISO) {
  const ate = ateISO || '2026-12-31';
  const P2R = placaParaRepresentante || {};
  const normPlaca = v => {
    const p = (v == null ? '' : v).toString().toUpperCase().replace(/[^A-Z0-9]/g, '');
    return p.length >= 6 ? p : '';
  };
  const rows = lerSubscricao(subscricaoXlsx);   // já normalizado: {placas[], data ISO, representante, status, ...}

  const nomesBase = representantesBase.map(r => r.nome);
  const rawNomes = [...new Set(rows.map(r => r.representante).filter(Boolean))];
  const destino = {};
  for (const rn of rawNomes) {
    const scored = nomesBase.map(bn => ({ bn, s: scoreNomes(rn, bn) })).filter(x => x.s > 0).sort((a, b) => b.s - a.s);
    if (!scored.length) continue;
    const top = scored[0].s;
    if (scored.filter(x => x.s === top).length > 1) continue; // empate -> descarta
    destino[rn] = scored[0].bn;
  }

  // por representante: { mes: { placas: n, clientes: Set } }
  const porRep = {};
  representantesBase.forEach(r => { porRep[r.nome] = {}; meses.forEach(m => porRep[r.nome][m] = { placas: 0, clientes: new Set() }); });

  const semMatch = new Set();
  const registros = []; // {alvo, mes, dataISO, associado} — para recálculos flexíveis (ex: comparação justa por dia)
  for (const r of rows) {
    const d = r.data;
    if (!d || d > ate) continue;
    // proposta RECUSADA não é venda (validado com a Eduarda em 14/08: Queila em agosto tinha 28 propostas,
    // 3 recusadas => 25 fechadas). Os demais status (Pendência Vistoria, Análise Rastreador, vazio…) contam.
    if (r.status.toUpperCase() === 'RECUSADO') continue;
    const mes = d.slice(0, 7);
    if (!meses.includes(mes)) continue;
    const raw = r.representante;
    // 1) pela PLACA na BASE (exato); 2) fallback: casamento por nome
    let alvo = null;
    for (const pn of r.placas) { if (P2R[pn]) { alvo = P2R[pn]; break; } }
    if (!alvo) alvo = destino[raw];
    if (!alvo || !porRep[alvo]) { if (raw) semMatch.add(raw); continue; }
    const bucket = porRep[alvo][mes];
    bucket.placas++;
    bucket.clientes.add(norm(r.associado));
    registros.push({ alvo, mes, dataISO: d, associado: norm(r.associado) });
  }

  return { porRep, semMatch: [...semMatch], registros };
};
