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

const SPREADSHEET_ID = '1a_SPv0YGuvAZd8aLYsIqeTIl22n7lZqQ1GE50Mgq1HY';
const SHEET_NAME = 'Mercadorias a pagar lojas próprias ';

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
    if (action === 'add')     return respond(addRow(body.row, body.pcts));
    if (action === 'update')  return respond(updateRow(body.rowIndex, body.row, body.pcts));
    if (action === 'delete')  return respond(deleteRow(body.rowIndex));
    if (action === 'archive') return respond(archiveRow(body.rowIndex));
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
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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
      try {
        if (val instanceof Date) {
          val = Utilities.formatDate(val, 'America/Sao_Paulo', 'dd/MM/yyyy');
        }
      } catch(e) { val = ''; }
      // Skip columns with date-formatted headers (extra columns from Excel like "28/05/2026")
      const hStr = String(h || '').trim();
      if (/^\d{2}\/\d{2}\/\d{4}$/.test(hStr)) return;
      obj[hStr] = (val !== null && val !== undefined) ? String(val) : '';
    });
    rows.push(obj);
  }

  return { success: true, data: rows };
}

// Cor bege claro para lançamentos feitos pelo sistema
const COR_SISTEMA = '#FFF5E0';

// ============================================================
// addRow: adiciona uma nova linha no final com cor bege
// ============================================================
function addRow(rowData, pcts) {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  const rowIndex = lastRow + 1;
  const range = sheet.getRange(rowIndex, 1, 1, rowData.length);
  range.setValues([rowData]);
  range.setBackground(COR_SISTEMA);
  if (pcts) applyPercentFormulas(sheet, rowIndex, pcts);
  return { success: true, rowIndex: rowIndex, message: 'Linha adicionada com sucesso.' };
}

// ============================================================
// updateRow: atualiza uma linha existente e mantém a cor bege
// ============================================================
function updateRow(rowIndex, rowData, pcts) {
  const sheet = getSheet();
  if (!rowIndex || rowIndex < 2) throw new Error('rowIndex inválido: ' + rowIndex);
  const range = sheet.getRange(rowIndex, 1, 1, rowData.length);
  range.setValues([rowData]);
  range.setBackground(COR_SISTEMA);
  if (pcts) applyPercentFormulas(sheet, rowIndex, pcts);
  return { success: true, message: 'Linha ' + rowIndex + ' atualizada.' };
}

// ============================================================
// applyPercentFormulas: escreve fórmulas (ex: =F123*3/100) nas
// células de MKT/Produto/Serviço, referenciando a célula do
// "Valor da nota fiscal" da própria linha. Assim a planilha
// mostra o valor em R$ E a fórmula/porcentagem usada no cálculo.
// ============================================================
function colToLetter(col) {
  let letter = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

function applyPercentFormulas(sheet, rowIndex, pcts) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const findCol = (patterns) => {
    for (const p of patterns) {
      for (let i = 0; i < headers.length; i++) {
        const h = String(headers[i] || '').trim().toLowerCase();
        if (h.indexOf(p) !== -1) return i + 1;
      }
    }
    return 0;
  };
  const valCol = findCol(['valor_da_nota', 'valor da nota', 'valor nf', 'valor']);
  if (!valCol) return;
  const valRef = colToLetter(valCol) + rowIndex;

  const targets = [
    { col: findCol(['mkt']), pct: pcts.mkt },
    { col: findCol(['produto']), pct: pcts.produto },
    { col: findCol(['serviço', 'servico']), pct: pcts.servico }
  ];
  targets.forEach(function(t) {
    if (t.col && t.pct) {
      sheet.getRange(rowIndex, t.col).setFormula('=' + valRef + '*' + t.pct + '/100');
    }
  });
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

// ============================================================
// EXCLUÍDOS — arquiva linha antes de deletar
// ============================================================
const DELETED_SHEET_NAME = 'Excluídos';

function getOrCreateDeletedSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(DELETED_SHEET_NAME);
  if (!sheet) {
    const mainSheet = getSheet();
    const headers = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
    sheet = ss.insertSheet(DELETED_SHEET_NAME);
    const fullHeaders = headers.concat(['Data de Exclusão']);
    sheet.getRange(1, 1, 1, fullHeaders.length).setValues([fullHeaders]);
    sheet.getRange(1, 1, 1, fullHeaders.length).setFontWeight('bold').setBackground('#F8D7DA');
    sheet.setTabColor('#cc0000');
  }
  return sheet;
}

function archiveRow(rowIndex) {
  const sheet = getSheet();
  if (!rowIndex || rowIndex < 2) throw new Error('rowIndex inválido: ' + rowIndex);

  const lastCol = sheet.getLastColumn();
  const rowData = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];

  const deletedSheet = getOrCreateDeletedSheet();
  const timestamp = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
  const archiveData = rowData.concat([timestamp]);

  deletedSheet.getRange(deletedSheet.getLastRow() + 1, 1, 1, archiveData.length).setValues([archiveData]);
  sheet.deleteRow(rowIndex);

  return { success: true, message: 'Arquivado em "Excluídos".' };
}

// ============================================================
// BACKUP DIÁRIO AUTOMÁTICO
// Cria uma cópia da aba principal todo dia e exclui o backup
// anterior. NUNCA toca na aba principal.
// ============================================================

const BACKUP_PREFIX = 'Backup_'; // Prefixo exclusivo para abas de backup

function criarBackupDiario() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // *** PROTEÇÃO 1: verificar que a aba principal existe antes de qualquer ação ***
  const mainSheet = ss.getSheetByName(SHEET_NAME);
  if (!mainSheet) {
    throw new Error('ERRO CRÍTICO: aba principal "' + SHEET_NAME + '" não encontrada! Backup cancelado por segurança.');
  }

  // Nome do backup de hoje
  const hoje = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd');
  const backupNome = BACKUP_PREFIX + hoje;

  // *** PROTEÇÃO 2: nunca permitir que o prefixo de backup seja igual ao nome da aba principal ***
  if (SHEET_NAME.startsWith(BACKUP_PREFIX)) {
    throw new Error('ERRO: nome da aba principal começa com "' + BACKUP_PREFIX + '". Altere a constante BACKUP_PREFIX.');
  }

  // Verificar se o backup de hoje já existe
  if (ss.getSheetByName(backupNome)) {
    Logger.log('Backup já existe para hoje: ' + backupNome);
    return 'Backup já existe: ' + backupNome;
  }

  // Criar backup copiando a aba principal
  const backupSheet = mainSheet.copyTo(ss);
  backupSheet.setName(backupNome);

  // Mover o backup para o final da planilha (longe da aba principal)
  ss.setActiveSheet(backupSheet);
  ss.moveActiveSheet(ss.getNumSheets());

  // *** PROTEÇÃO 3: excluir APENAS abas com prefixo Backup_, nunca a principal ***
  const todasAbas = ss.getSheets();
  let excluidas = [];
  for (const aba of todasAbas) {
    const nome = aba.getName();

    // Condições de segurança — pula se qualquer uma falhar:
    if (!nome.startsWith(BACKUP_PREFIX)) continue;  // não é backup
    if (nome === backupNome) continue;               // é o backup de hoje
    if (nome === SHEET_NAME) continue;               // *** NUNCA excluir a principal ***
    if (aba.getSheetId() === mainSheet.getSheetId()) continue; // dupla verificação por ID

    ss.deleteSheet(aba);
    excluidas.push(nome);
    Logger.log('Backup antigo excluído: ' + nome);
  }

  const msg = 'Backup criado: ' + backupNome + (excluidas.length > 0 ? ' | Excluídos: ' + excluidas.join(', ') : '');
  Logger.log(msg);
  return msg;
}

// ============================================================
// Configurar o trigger diário (execute UMA VEZ manualmente)
// ============================================================
function configurarBackupDiario() {
  // Remover triggers anteriores desta função para evitar duplicatas
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'criarBackupDiario') {
      ScriptApp.deleteTrigger(t);
      Logger.log('Trigger anterior removido.');
    }
  }

  // Criar trigger: todo dia às 2h da manhã (horário de Brasília)
  ScriptApp.newTrigger('criarBackupDiario')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .nearMinute(0)
    .inTimezone('America/Sao_Paulo')
    .create();

  const msg = 'Trigger configurado! Backup automático todo dia às 02:00 (Brasília).';
  Logger.log(msg);
  return msg;
}
