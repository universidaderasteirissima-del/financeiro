// ============================================================
// RASTEIRISSIMA — Google Apps Script API (FRANQUIAS / QG)
//
// Este script é INDEPENDENTE do apps-script.js de Lojas Próprias.
// Não altera nada relacionado a Mercadorias a Pagar (lp_).
//
// COMO USAR:
// 1. Abra a planilha de Franquias (QG) no Google Sheets
// 2. Extensões > Apps Script
// 3. Apague o código existente e cole TODO este arquivo
// 4. Salvar (ícone de disquete)
// 5. Implantar > Nova implantação
// 6. Tipo: App da Web | Executar como: Eu | Acesso: Qualquer pessoa
// 7. Implantar e copiar a URL gerada
// 8. Colar essa URL em Configurações, estando logado como Franquias
// ============================================================

const SPREADSHEET_ID = '1023jag9C3galUDYPUA6fOCcxwbD3PFEbYM1RX_0mAF8';

// Nome das abas na planilha, por view do sistema
const VIEW_SHEETS = {
  brilhante:   'A pagar brilhante',
  marketing:   'A pagar marketing',
  franchising: 'A pagar franchising',
  areceber:    'A receber'
};

// Mapa: chave do sistema -> texto do cabeçalho esperado na planilha
// (a busca de cabeçalho ignora acentos, maiúsculas/minúsculas e
// espaços/hifens, então pequenas variações de escrita não quebram)
const HEADER_MAP = {
  brilhante: {
    cat:'Categoria', mes:'mês', dv:'Data de vencimento', desc:'Despesa-descrição',
    va:'Valor a pagar', vp:'Valor pago', st:'Status', dp:'Data do pagamento', ano:'Ano',
    criadoEm:'Criado em'
  },
  marketing: {
    cat:'Categoria', mes:'mês', dv:'Data de vencimento', desc:'Despesa-descrição',
    va:'Valor a pagar', vp:'Valor pago', st:'Status', dp:'Data do pagamento',
    criadoEm:'Criado em'
  },
  franchising: {
    cat:'Categoria', mes:'mês', dv:'Data de vencimento', desc:'Despesa-descrição',
    va:'Valor a pagar', vp:'Valor pago', st:'Status', dp:'Data do pagamento',
    criadoEm:'Criado em'
  },
  areceber: {
    // OBS: no sistema, "totalNota" é o valor direto da NF (menor) e "vNota" é
    // o valor calculado pelo % do fornecedor externo (geralmente maior). Na
    // planilha, a coluna "Valor da nota" deve receber o valor menor/direto e
    // "Total da Nota" o valor maior/calculado — por isso o mapeamento abaixo
    // é invertido em relação às chaves do sistema.
    // Cabeçalhos confirmados via diagnosticarCabecalhos() em 2026-06-29.
    tipo:'Tipo', loja:'Loja', forn:'Fornecedor', de:'Data de emissão', nota:'Nota',
    vNota:'Total da Nota', totalNota:'Valor da nota',
    boletoMkt:'Boleto marketing', boletoFran:'Boleto franchising',
    dv1:'Data de vencimento 1° Parcela', dp1:'Data de pagamento 1° Parcela', st1:'Status 1° Parcela',
    dv2:'Data de vencimento 2° Parcela', dp2:'Data de pagamento 2° Parcela', st2:'Status 2° Parcela',
    boletoBri1:'Boleto brilhante / fábrica 1° Parcela',
    boletoBri2:'Boleto brilhante / fábrica 2° Parcela',
    criadoEm:'Criado em'
  }
};

// Campos que são datas (DD/MM/AAAA na planilha, AAAA-MM-DD no sistema)
const DATE_FIELDS = {
  brilhante: ['dv','dp'],
  marketing: ['dv','dp'],
  franchising: ['dv','dp'],
  areceber: ['de','dv1','dp1','dv2','dp2']
};

// "Criado em": guarda data+hora (AAAA-MM-DDTHH:MM:SS), não só data — por
// isso não entra em DATE_FIELDS (que é só DD/MM/AAAA). O sistema grava
// como texto puro; se a coluna não estiver formatada como "Texto simples"
// e o Sheets converter sozinho pra um valor de data/hora nativo, getData()
// abaixo reconverte pro mesmo formato ISO com hora, pra não perder a hora
// nem quebrar a leitura no app.
const DATETIME_FIELDS = {
  brilhante: ['criadoEm'], marketing: ['criadoEm'], franchising: ['criadoEm'], areceber: ['criadoEm']
};

const COR_SISTEMA = '#FFF5E0';
const DELETED_SHEET_NAME = 'Excluídos';

// ============================================================
// ROTAS
// ============================================================
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'getData';
  const view = e && e.parameter && e.parameter.view;
  try {
    if (action === 'getData') return respond(getData(view));
    if (action === 'getDeleted') return respond(getDeleted());
    return respond({ error: 'Ação desconhecida: ' + action });
  } catch (err) {
    return respond({ error: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    if (action === 'add')     return respond(addRow(body.view, body.fields));
    if (action === 'update')  return respond(updateRow(body.view, body.rowIndex, body.fields));
    if (action === 'archive') return respond(archiveRow(body.view, body.rowIndex));
    if (action === 'restore') return respond(restoreRow(body.deletedRowIndex));
    return respond({ error: 'Ação desconhecida: ' + action });
  } catch (err) {
    return respond({ error: err.message });
  }
}

function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet(view) {
  const name = VIEW_SHEETS[view];
  if (!name) throw new Error('View desconhecida: ' + view);
  return getSheetByNameLoose(name);
}
// Tolerante a espaços extras no nome da aba (ex: "A pagar franchising "
// com espaço no final, como está na planilha original)
function getSheetByNameLoose(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.getSheets().find(function (s) { return s.getName().trim() === name.trim(); });
  }
  if (!sheet) throw new Error('Aba "' + name + '" não encontrada na planilha.');
  return sheet;
}

// ============================================================
// Normalização de cabeçalho — ignora acentos, caixa, espaços e
// hifens para ser tolerante a pequenas variações de escrita
// ============================================================
function normHdr(s) {
  return String(s || '').trim().toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^a-z0-9]/g, '');
}
function findHeaderIdx(headers, target) {
  const t = normHdr(target);
  for (let i = 0; i < headers.length; i++) {
    if (normHdr(headers[i]) === t) return i;
  }
  return -1;
}
function getHeaders(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h || '').trim(); });
}
function fmtDateOut(val) {
  if (val instanceof Date) return Utilities.formatDate(val, 'America/Sao_Paulo', 'dd/MM/yyyy');
  return val;
}
// AAAA-MM-DD (sistema) -> DD/MM/AAAA (planilha)
function isoToBr(v) {
  const parts = String(v || '').split('-');
  if (parts.length === 3) return parts[2] + '/' + parts[1] + '/' + parts[0];
  return v;
}
// DD/MM/AAAA (planilha) -> AAAA-MM-DD (sistema)
function brToIso(v) {
  const s = String(v || '').trim();
  const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (m) return m[3] + '-' + m[2] + '-' + m[1];
  return '';
}

// ============================================================
// getData: lê a aba e retorna já mapeado pelas chaves do sistema
// ============================================================
function getData(view) {
  const sheet = getSheet(view);
  const map = HEADER_MAP[view];
  const dateKeys = DATE_FIELDS[view] || [];
  const datetimeKeys = DATETIME_FIELDS[view] || [];
  const headers = getHeaders(sheet);
  const colIdx = {};
  Object.keys(map).forEach(function (key) {
    const idx = findHeaderIdx(headers, map[key]);
    if (idx >= 0) colIdx[key] = idx;
  });

  const data = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.every(function (c) { return c === '' || c === null; })) continue;
    const obj = { _rowIndex: i + 1 };
    Object.keys(colIdx).forEach(function (key) {
      let val = row[colIdx[key]];
      if (datetimeKeys.indexOf(key) >= 0) {
        // Preserva a hora (fmtDateOut sozinho perderia, só formata dd/MM/aaaa)
        if (val instanceof Date) val = Utilities.formatDate(val, 'America/Sao_Paulo', "yyyy-MM-dd'T'HH:mm:ss");
        obj[key] = (val !== null && val !== undefined) ? val : '';
        return;
      }
      val = fmtDateOut(val);
      if (dateKeys.indexOf(key) >= 0) val = brToIso(val);
      obj[key] = (val !== null && val !== undefined) ? val : '';
    });
    rows.push(obj);
  }
  return { success: true, data: rows };
}

// ============================================================
// addRow / updateRow — escreve só nas colunas mapeadas, preserva
// quaisquer outras colunas existentes na planilha (ex: Pagamento
// Franchising, Saldo Atual etc. que o sistema ainda não gerencia)
// ============================================================
function addRow(view, fields) {
  const sheet = getSheet(view);
  const map = HEADER_MAP[view];
  const dateKeys = DATE_FIELDS[view] || [];
  const headers = getHeaders(sheet);
  const rowIndex = sheet.getLastRow() + 1;
  const newRow = new Array(headers.length).fill('');
  Object.keys(fields || {}).forEach(function (key) {
    if (!map[key]) return;
    const idx = findHeaderIdx(headers, map[key]);
    if (idx < 0) return;
    let v = fields[key];
    if (dateKeys.indexOf(key) >= 0 && v) v = isoToBr(v);
    newRow[idx] = v;
  });
  const range = sheet.getRange(rowIndex, 1, 1, newRow.length);
  range.setValues([newRow]);
  range.setBackground(COR_SISTEMA);
  SpreadsheetApp.flush();
  return { success: true, rowIndex: rowIndex };
}

function updateRow(view, rowIndex, fields) {
  const sheet = getSheet(view);
  if (!rowIndex || rowIndex < 2) throw new Error('rowIndex inválido: ' + rowIndex);
  const map = HEADER_MAP[view];
  const dateKeys = DATE_FIELDS[view] || [];
  const headers = getHeaders(sheet);
  const existing = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  Object.keys(fields || {}).forEach(function (key) {
    if (!map[key]) return;
    const idx = findHeaderIdx(headers, map[key]);
    if (idx < 0) return;
    let v = fields[key];
    if (dateKeys.indexOf(key) >= 0 && v) v = isoToBr(v);
    existing[idx] = v;
  });
  const range = sheet.getRange(rowIndex, 1, 1, existing.length);
  range.setValues([existing]);
  range.setBackground(COR_SISTEMA);
  SpreadsheetApp.flush();
  return { success: true };
}

// ============================================================
// EXCLUÍDOS — aba única e compartilhada entre as 4 views.
// Guarda um snapshot JSON completo do registro para permitir
// restaurar exatamente como estava, em qualquer view de origem.
// ============================================================
function getOrCreateDeletedSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(DELETED_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DELETED_SHEET_NAME);
    const headers = ['Aba de Origem', 'Data da Exclusão', 'Resumo', 'Dados (JSON)'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#F8D7DA');
    sheet.setTabColor('#cc0000');
  }
  return sheet;
}

function archiveRow(view, rowIndex) {
  const sheet = getSheet(view);
  if (!rowIndex || rowIndex < 2) throw new Error('rowIndex inválido: ' + rowIndex);
  const map = HEADER_MAP[view];
  const dateKeys = DATE_FIELDS[view] || [];
  const headers = getHeaders(sheet);
  const rowData = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];

  const obj = {};
  Object.keys(map).forEach(function (key) {
    const idx = findHeaderIdx(headers, map[key]);
    if (idx < 0) return;
    let v = fmtDateOut(rowData[idx]);
    if (dateKeys.indexOf(key) >= 0) v = brToIso(v);
    obj[key] = v;
  });

  const resumo = Object.keys(obj).map(function (k) { return obj[k]; })
    .filter(function (v) { return v !== '' && v !== null && v !== undefined; })
    .slice(0, 4).join(' | ');
  const timestamp = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');

  const deletedSheet = getOrCreateDeletedSheet();
  deletedSheet.appendRow([view, timestamp, resumo, JSON.stringify(obj)]);

  sheet.deleteRow(rowIndex);
  SpreadsheetApp.flush();
  return { success: true };
}

function restoreRow(deletedRowIndex) {
  const deletedSheet = getOrCreateDeletedSheet();
  if (!deletedRowIndex || deletedRowIndex < 2) throw new Error('deletedRowIndex inválido: ' + deletedRowIndex);
  const rowVals = deletedSheet.getRange(deletedRowIndex, 1, 1, 4).getValues()[0];
  const view = rowVals[0];
  const fields = JSON.parse(rowVals[3] || '{}');

  const added = addRow(view, fields);

  deletedSheet.deleteRow(deletedRowIndex);
  SpreadsheetApp.flush();
  return { success: true, view: view, rowIndex: added.rowIndex };
}

function getDeleted() {
  const sheet = getOrCreateDeletedSheet();
  const data = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    rows.push({ _rowIndex: i + 1, view: row[0], data_exclusao: row[1], resumo: row[2] });
  }
  return { success: true, data: rows };
}

// ============================================================
// DIAGNÓSTICO — rode manualmente no editor do Apps Script
// (Executar > diagnosticarCabecalhos) para ver, no "Ver registros
// de execução", quais cabeçalhos foram encontrados/não encontrados
// em cada aba. Útil se algum campo não estiver sincronizando.
// ============================================================
function diagnosticarCabecalhos() {
  Object.keys(VIEW_SHEETS).forEach(function (view) {
    let sheet;
    try { sheet = getSheetByNameLoose(VIEW_SHEETS[view]); }
    catch (e) { Logger.log(view + ': ABA NÃO ENCONTRADA (' + VIEW_SHEETS[view] + ')'); return; }
    const headers = getHeaders(sheet);
    Logger.log('--- ' + view + ' (' + VIEW_SHEETS[view] + ') — cabeçalhos na planilha: ' + JSON.stringify(headers));
    Object.keys(HEADER_MAP[view]).forEach(function (key) {
      const idx = findHeaderIdx(headers, HEADER_MAP[view][key]);
      Logger.log('  ' + key + ' ("' + HEADER_MAP[view][key] + '") -> ' + (idx >= 0 ? 'coluna ' + (idx + 1) : 'NÃO ENCONTRADO'));
    });
  });
}
