/*
 * Lê planilhas avulsas de "campanha de desconto especial" (ex.: "Desconto especial semana do dia 14.xlsx") —
 * exportações manuais, sem cabeçalho, só 3 colunas fixas: Associado, Placa, Representante. Cada linha é
 * uma adesão que aconteceu FORA do Controle de Subscrição normal (campanha à parte) e por isso não seria
 * contada pelo pipeline principal — precisa ser somada à parte às vendas do mês.
 *
 * Uso: ler(caminhoXlsx) -> [{ associado, placa, representante }]
 */
const XLSX = require('xlsx');

function normPlaca(v) {
  const p = (v == null ? '' : v).toString().toUpperCase().replace(/[^A-Z0-9]/g, '');
  return p.length >= 6 ? p : '';
}

module.exports = function ler(xlsxPath) {
  const wb = XLSX.readFile(xlsxPath);
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false });
  const out = [];
  for (const r of rows) {
    const associado = (r[0] || '').toString().trim();
    const placa = normPlaca(r[1]);
    const representante = (r[2] || '').toString().trim();
    if (!associado || !placa || !representante) continue; // linhas em branco no fim da planilha
    out.push({ associado, placa, representante });
  }
  return out;
};
