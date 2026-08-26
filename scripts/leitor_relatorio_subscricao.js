/*
 * Leitor do "Relatório de Subscrição" (PPM) — relatório DIFERENTE do "Controle de Subscrição V2":
 * 1 linha = 1 PROPOSTA (não 1 placa; uma proposta pode cobrir várias placas), sem coluna de Placa(s).
 * Colunas: N° Subscrição, N° Cotação, Associado, CPF/CNPJ, Telefone, E-mail, Tipo Veículo,
 * Data Transmissão, Valor, Franqueado, Representante (aqui vem o CPF do representante, não o nome),
 * Atividade, SLA.
 *
 * Confirmado com a Eduarda em 26/08/2026: "vendas por dia" = contagem de LINHAS deste relatório por
 * Data Transmissão, sem filtrar por Atividade/Status — reconciliado ao vivo (68 linhas em 25/08 = 68
 * vendas que ela via no Siprov, contra 37 que a BASE por Data de Adesão mostrava, subcontando).
 *
 * Devolve [{ nSubscricao, cotacao, associado, cpfAssociado, data(ISO), valor, franquia,
 *            cpfRepresentante, atividade }]
 */
const XLSX = require('xlsx');
const path = require('path');

const COLUNAS = {
  nSubscricao: ['N° Subscrição', 'Nº Subscrição'],
  cotacao: ['N° Cotação', 'Nº Cotação'],
  associado: ['Associado'],
  cpfAssociado: ['CPF/CNPJ'],
  data: ['Data Transmissão'],
  valor: ['Valor'],
  franquia: ['Franqueado'],
  cpfRepresentante: ['Representante'],
  atividade: ['Atividade'],
};

const nrm = s => String(s == null ? '' : s).trim().toUpperCase();

// aceita "25/08/2026 09:55", "25/08/2026 Ás 09:55" ou só "25/08/2026" — sempre D/M/AAAA (confirmado:
// este relatório nunca veio em formato americano nos exports vistos; se algum dia vier ambíguo, dá
// pra reaproveitar a mesma detecção coluna-inteira do leitor_subscricao.js).
function paraISO(v) {
  const m = String(v == null ? '' : v).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (!m) return null;
  const dia = +m[1], mes = +m[2], ano = +m[3];
  if (mes < 1 || mes > 12 || dia < 1 || dia > 31) return null;
  return ano + '-' + String(mes).padStart(2, '0') + '-' + String(dia).padStart(2, '0');
}

function toNum(v) {
  if (v == null || v === '') return 0;
  const n = parseFloat(v.toString().replace(/[^\d,.-]/g, '').replace(/\.(?=\d{3}(\D|$))/g, '').replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

module.exports = function lerRelatorioSubscricao(xlsxPath) {
  const wb = XLSX.readFile(xlsxPath);
  const linhas = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false });

  let iCab = -1;
  for (let i = 0; i < Math.min(3, linhas.length); i++) {
    const set = new Set((linhas[i] || []).map(nrm));
    if (set.has('DATA TRANSMISSÃO') && set.has('REPRESENTANTE')) { iCab = i; break; }
  }
  if (iCab === -1) throw new Error('cabeçalho não encontrado em ' + path.basename(xlsxPath) + ' (leiaute do "Relatório de Subscrição" mudou?)');

  const cab = (linhas[iCab] || []).map(nrm);
  const idx = {};
  for (const [chave, nomes] of Object.entries(COLUNAS)) {
    const i = cab.findIndex(h => nomes.some(n => nrm(n) === h));
    if (i === -1 && ['data', 'cpfRepresentante'].includes(chave))
      throw new Error('coluna obrigatória não encontrada no Relatório de Subscrição: ' + nomes[0]);
    idx[chave] = i;
  }
  const val = (r, k) => idx[k] >= 0 ? r[idx[k]] : undefined;

  const dados = linhas.slice(iCab + 1).filter(r => r && !r.every(c => c == null || c === ''));
  const out = [];
  for (const r of dados) {
    const data = paraISO(val(r, 'data'));
    if (!data) continue;
    out.push({
      nSubscricao: String(val(r, 'nSubscricao') || '').trim(),
      cotacao: String(val(r, 'cotacao') || '').trim(),
      associado: String(val(r, 'associado') || '').trim(),
      cpfAssociado: String(val(r, 'cpfAssociado') || '').replace(/\D/g, ''),
      data,
      valor: toNum(val(r, 'valor')),
      franquia: String(val(r, 'franquia') || '').replace(/^\d+\s*-\s*/, '').trim(),
      cpfRepresentante: String(val(r, 'cpfRepresentante') || '').replace(/\D/g, ''),
      atividade: String(val(r, 'atividade') || '').trim(),
    });
  }
  return out;
};
