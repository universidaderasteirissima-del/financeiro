// ============================================================
// RASTEIRISSIMA — Google Apps Script API
//
// COMO USAR:
// 1. Abra sua planilha Google Sheets
// 2. Clique em Extensões > Apps Script
// 3. Apague o código existente e cole TODO este arquivo
// 4. Clique em Salvar (ícone de disquete)
// 5. Clique em Implantar > Nova implantação
// 6. Tipo: App da Web
// 7. Executar como: Eu
// 8. Quem tem acesso: Qualquer pessoa
// 9. Clique em Implantar
// 10. Copie a URL gerada e cole nas Configurações do sistema
// ============================================================

const SHEET_NAME = 'Mercadorias a pagar lojas própr';

// ============================================================
// GET: retorna todos os dados
// ============================================================
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'getData';
  try {
    if (action === 'getData') return respond(getData());
    return respond({ error: 'Ação desconhecida: ' + action });
  } catch (err) {
    return respond({ error: err.message });
  }
}

// ============================================================
// POST: adiciona, atualiza ou exclui linha
// ============================================================
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    if (action === 'add')    return respond(addRow(body.row));
    if (action === 'update') return respond(updateRow(body.rowIndex, body.row));
    if (action === 'delete') return respond(deleteRow(body.rowIndex));
    return respond({ error: 'Ação desconhecida: ' + action });
  } catch (err) {
    return respond({ error: err.message });
  }
}

// ============================================================
// Helpers
// ============================================================
function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Aba "' + SHEET_NAME + '" não encontrada na planilha.');
  return sheet;
}

// ============================================================
// getData: lê todas as linhas e retorna como array de objetos
// ============================================================
function getData() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: true, data: [] };

  const headers = data[0];
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue; // Ignora linhas vazias

    const obj = { _rowIndex: i + 1 };
    headers.forEach((h, j) => {
      let val = row[j];
      // Formata datas para dd/MM/yyyy
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      }
      obj[h] = val !== null && val !== undefined ? String(val) : '';
    });
    rows.push(obj);
  }

  return { success: true, data: rows };
}

// ============================================================
// addRow: adiciona uma nova linha no final
// ============================================================
function addRow(rowData) {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, rowData.length).setValues([rowData]);
  return { success: true, rowIndex: lastRow + 1, message: 'Linha adicionada com sucesso.' };
}

// ============================================================
// updateRow: atualiza uma linha existente pelo índice
// ============================================================
function updateRow(rowIndex, rowData) {
  const sheet = getSheet();
  if (!rowIndex || rowIndex < 2) throw new Error('rowIndex inválido: ' + rowIndex);
  sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
  return { success: true, message: 'Linha ' + rowIndex + ' atualizada.' };
}

// ============================================================
// deleteRow: exclui uma linha pelo índice
// ============================================================
function deleteRow(rowIndex) {
  const sheet = getSheet();
  if (!rowIndex || rowIndex < 2) throw new Error('rowIndex inválido: ' + rowIndex);
  sheet.deleteRow(rowIndex);
  return { success: true, message: 'Linha ' + rowIndex + ' excluída.' };
}
