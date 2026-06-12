// Extração de dados da planilha "Cópia de Financeiro - QG.xlsx" (área Franquias)
// Gera data_fr_brilhante.json, data_fr_marketing.json, data_fr_franchising.json
// e data_fr_areceber.json, um por aba do documento original.
//
// Uso: npm install && node extract-qg.js

const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const SRC = path.join(__dirname, 'Cópia de Financeiro - QG.xlsx');
const wb = XLSX.readFile(SRC, { cellDates: true });

// Valores monetários na planilha vêm como "432.76", "1,518.00", "R$1518" ou
// "R$ 2,527.71" — sempre ponto como separador decimal e vírgula como milhar.
function money(v) {
  const n = parseFloat(String(v || '').replace(/[R$\s,]/g, ''));
  return isNaN(n) ? 0 : n;
}

// Datas vêm como DD/MM/AAAA ou DD/MM/AA. Valores fora desse padrão (erros de
// digitação na planilha original) ficam em branco.
function dateIso(v) {
  const m = String(v || '').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (!m) return '';
  let [, d, mo, y] = m;
  if (y.length === 2) y = '20' + y;
  return y + '-' + mo.padStart(2, '0') + '-' + d.padStart(2, '0');
}

function status(v) {
  return String(v || '').trim().toLowerCase();
}

function txt(v) {
  return String(v || '').trim();
}

// Esquema das abas "A pagar brilhante" / "A pagar marketing" / "A pagar franchising"
// Colunas: Categoria, mês, Data de vencimento, Despesa-descrição, Valor a pagar,
// Valor pago, Status, Data do pagamento[, Ano (só "A pagar brilhante")]
function mapAP(row, includeAno) {
  const rec = {
    cat: txt(row[0]),
    mes: parseInt(row[1]) || 0,
    dv: dateIso(row[2]),
    desc: txt(row[3]),
    va: money(row[4]),
    vp: money(row[5]),
    st: status(row[6]),
    dp: dateIso(row[7]),
  };
  if (includeAno) rec.ano = parseInt(row[8]) || 0;
  return rec;
}

// Esquema da aba "A receber" (23 colunas na planilha, 2 sem título descartadas)
function mapAR(row) {
  return {
    tipo: txt(row[0]),
    loja: txt(row[1]),
    forn: txt(row[2]),
    de: dateIso(row[3]),
    dv1: dateIso(row[4]),
    dv2: dateIso(row[5]),
    nota: txt(row[6]),
    vNota: money(row[7]),
    totalNota: money(row[8]),
    boletoMkt: money(row[9]),
    boletoFran: money(row[10]),
    boletoBri1: money(row[11]),
    boletoBri2: money(row[12]),
    st1: status(row[13]),
    dp1: dateIso(row[14]),
    st2: status(row[15]),
    dp2: dateIso(row[16]),
    pagFran: money(row[17]),
    pagMkt: money(row[18]),
    // row[19] e row[20]: colunas sem título na planilha, descartadas
    saldoFran: money(row[21]),
    saldoBri: money(row[22]),
  };
}

// Extrai uma aba, pulando linhas totalmente vazias e gerando _id sequencial
// (a partir de 2, mesma convenção usada em data.json: linha 1 = cabeçalho).
function extract(sheetName, mapper) {
  const ws = wb.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row.some(c => String(c).trim() !== '')) continue;
    out.push({ _id: out.length + 2, ...mapper(row) });
  }
  return out;
}

const views = {
  data_fr_brilhante: extract('A pagar brilhante', r => mapAP(r, true)),
  data_fr_marketing: extract('A pagar marketing', r => mapAP(r, false)),
  data_fr_franchising: extract('A pagar franchising ', r => mapAP(r, false)),
  data_fr_areceber: extract('A receber', mapAR),
};

for (const [name, rows] of Object.entries(views)) {
  fs.writeFileSync(path.join(__dirname, name + '.json'), JSON.stringify(rows, null, 2));
  console.log(name + '.json:', rows.length, 'registros');
}
