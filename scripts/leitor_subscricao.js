/*
 * Leitor do "Controle de Subscrição V2" (PPM), tolerante às variações de export.
 *
 * O PPM já mudou o leiaute duas vezes em agosto/2026:
 *   - com linha de título antes do cabeçalho x sem linha de título;
 *   - data "14/08/2026 Ás 13:20" (BR) x "8/14/26" (americano).
 * Ler por índice fixo + slice(2) fez o painel publicar AGOSTO = 0 silenciosamente em 14/08.
 * Por isso aqui: acha a linha de cabeçalho, resolve as colunas por NOME e aceita os dois formatos.
 *
 * Devolve [{ cotacao, associado, cpf, placas[], data(ISO), franquia, representante, status }]
 */
const XLSX = require('xlsx');
const path = require('path');

const COLUNAS = {
  cotacao: ['Nº Cotação', 'N° Cotação'],
  associado: ['Associado'],
  cpf: ['CPF/CNPJ'],
  placas: ['Placa(s)', 'Placas'],
  data: ['Data Trans', 'Data Transmissão/Cálculo', 'Data Transmissão'],
  franquia: ['Franquia'],
  representante: ['Representante'],
  status: ['Status'],
};

const nrm = s => String(s == null ? '' : s).trim().toUpperCase();

// aceita "14/08/2026 ..." (dd/mm/yyyy) e "8/14/26" (M/D/YY). Desambigua pelo componente > 12.
function paraISO(v) {
  const s = String(v == null ? '' : v).trim();
  if (!s) return null;
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (!m) return null;
  let [, a, b, ano] = m;
  a = +a; b = +b; ano = +ano;
  if (ano < 100) ano += 2000;
  let dia, mes;
  if (a > 12) { dia = a; mes = b; }        // só pode ser dd/mm
  else if (b > 12) { mes = a; dia = b; }   // só pode ser M/D
  else { dia = a; mes = b; }               // ambíguo: assume dd/mm (padrão BR histórico do arquivo)
  if (mes < 1 || mes > 12 || dia < 1 || dia > 31) return null;
  return ano + '-' + String(mes).padStart(2, '0') + '-' + String(dia).padStart(2, '0');
}

function normPlaca(v) {
  const p = String(v == null ? '' : v).toUpperCase().replace(/[^A-Z0-9]/g, '');
  return p.length >= 6 ? p : '';
}

module.exports = function lerSubscricao(xlsxPath) {
  const wb = XLSX.readFile(xlsxPath);
  const linhas = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false });

  // cabeçalho = primeira linha (das 3 primeiras) que contenha "Placa(s)" e "Representante"
  let iCab = -1;
  for (let i = 0; i < Math.min(3, linhas.length); i++) {
    const set = new Set((linhas[i] || []).map(nrm));
    if (set.has('PLACA(S)') && set.has('REPRESENTANTE')) { iCab = i; break; }
  }
  if (iCab === -1) throw new Error('cabeçalho não encontrado em ' + path.basename(xlsxPath) + ' (leiaute do PPM mudou?)');

  const cab = (linhas[iCab] || []).map(nrm);
  const idx = {};
  for (const [chave, nomes] of Object.entries(COLUNAS)) {
    const i = cab.findIndex(h => nomes.some(n => nrm(n) === h));
    if (i === -1 && ['placas', 'data', 'representante'].includes(chave))
      throw new Error('coluna obrigatória não encontrada na Subscrição: ' + nomes[0]);
    idx[chave] = i;
  }
  const val = (r, k) => idx[k] >= 0 ? r[idx[k]] : undefined;

  const out = [];
  for (const r of linhas.slice(iCab + 1)) {
    if (!r || r.every(c => c == null || c === '')) continue;
    const data = paraISO(val(r, 'data'));
    if (!data) continue;                       // linha sem data válida não é subscrição
    out.push({
      cotacao: String(val(r, 'cotacao') || '').trim(),
      associado: String(val(r, 'associado') || '').trim(),
      cpf: String(val(r, 'cpf') || '').replace(/\D/g, ''),
      placas: String(val(r, 'placas') || '').split(/[\s,\/]+/).map(normPlaca).filter(Boolean),
      data,
      franquia: String(val(r, 'franquia') || '').replace(/^\d+\s*-\s*/, '').trim(),
      representante: String(val(r, 'representante') || '').trim(),
      status: String(val(r, 'status') || '').trim(),
    });
  }
  return out;
};
module.exports.paraISO = paraISO;
