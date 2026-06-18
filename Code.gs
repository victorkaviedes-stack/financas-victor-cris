// ============================================================
//  FINANÇAS DE VICTOR E CRIS — Google Apps Script (Backend)
//  Cole este código em: script.google.com → Novo projeto
// ============================================================

// ⚠️ CONFIGURAÇÃO: Cole o ID da sua planilha aqui após criá-la
const SHEET_ID = 'COLE_O_ID_DA_SUA_PLANILHA_AQUI';

const SHEET_TRANS  = 'Transacoes';
const SHEET_CONFIG = 'Config';

// ------------------------------------------------------------
//  Ponto de entrada HTTP — roteador
// ------------------------------------------------------------
function doGet(e) {
  const action = e.parameter.action || '';
  let result;

  try {
    if      (action === 'getTransactions') result = getTransactions(e.parameter);
    else if (action === 'getConfig')       result = getConfig();
    else if (action === 'ping')            result = { ok: true };
    else                                   result = { error: 'Ação desconhecida: ' + action };
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let body, result;

  try {
    body = JSON.parse(e.postData.contents);
    const action = body.action || '';

    if      (action === 'addTransaction')    result = addTransaction(body.data);
    else if (action === 'deleteTransaction') result = deleteTransaction(body.id);
    else if (action === 'editTransaction')   result = editTransaction(body.id, body.data);
    else if (action === 'setConfig')         result = setConfig(body.data);
    else                                     result = { error: 'Ação desconhecida: ' + action };
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ------------------------------------------------------------
//  Helpers de planilha
// ------------------------------------------------------------
function getSheet(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === SHEET_TRANS) {
      sheet.appendRow(['id','descricao','valor','data','tipo','catId','catLabel','catIcon','mes','ano']);
      sheet.getRange(1,1,1,10).setFontWeight('bold').setBackground('#1D9E75').setFontColor('#ffffff');
    }
    if (name === SHEET_CONFIG) {
      sheet.appendRow(['chave','valor']);
      sheet.appendRow(['pin','1234']);
      sheet.getRange(1,1,1,2).setFontWeight('bold').setBackground('#1D9E75').setFontColor('#ffffff');
    }
  }
  return sheet;
}

function genId() {
  return Utilities.getUuid().replace(/-/g,'').slice(0,12);
}

// ------------------------------------------------------------
//  Transações
// ------------------------------------------------------------
function getTransactions(params) {
  const sheet = getSheet(SHEET_TRANS);
  const rows  = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { transactions: [] };

  const headers = rows[0];
  let trans = rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i] === '' ? null : row[i]);
    obj.valor = parseFloat(obj.valor) || 0;
    obj.mes   = parseInt(obj.mes)     || 0;
    obj.ano   = parseInt(obj.ano)     || 0;
    return obj;
  });

  if (params.mes && params.ano) {
    const m = parseInt(params.mes), a = parseInt(params.ano);
    trans = trans.filter(t => t.mes === m && t.ano === a);
  }

  return { transactions: trans };
}

function addTransaction(data) {
  const sheet = getSheet(SHEET_TRANS);
  const id    = genId();
  sheet.appendRow([
    id,
    data.descricao,
    parseFloat(data.valor),
    data.data,
    data.tipo,
    data.catId,
    data.catLabel,
    data.catIcon,
    parseInt(data.mes),
    parseInt(data.ano)
  ]);
  return { ok: true, id };
}

function deleteTransaction(id) {
  const sheet = getSheet(SHEET_TRANS);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { error: 'Transação não encontrada: ' + id };
}

function editTransaction(id, data) {
  const sheet   = getSheet(SHEET_TRANS);
  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0];

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      const rowNum = i + 1;
      const map    = { descricao:2, valor:3, data:4, tipo:5, catId:6, catLabel:7, catIcon:8, mes:9, ano:10 };
      Object.keys(data).forEach(k => {
        if (map[k]) sheet.getRange(rowNum, map[k]).setValue(data[k]);
      });
      return { ok: true };
    }
  }
  return { error: 'Transação não encontrada: ' + id };
}

// ------------------------------------------------------------
//  Configurações (PIN)
// ------------------------------------------------------------
function getConfig() {
  const sheet = getSheet(SHEET_CONFIG);
  const rows  = sheet.getDataRange().getValues();
  const config = {};
  rows.slice(1).forEach(row => config[row[0]] = row[1]);
  return { config };
}

function setConfig(data) {
  const sheet   = getSheet(SHEET_CONFIG);
  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0];

  Object.keys(data).forEach(key => {
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(data[key]);
        found = true;
        break;
      }
    }
    if (!found) sheet.appendRow([key, data[key]]);
  });

  return { ok: true };
}

// ------------------------------------------------------------
//  Utilitário: criar planilha automaticamente
//  Execute esta função UMA VEZ pelo menu do Apps Script
// ------------------------------------------------------------
function criarPlanilha() {
  const ss    = SpreadsheetApp.create('Finanças Victor e Cris');
  const url   = ss.getUrl();
  const id    = ss.getId();

  // Remove aba padrão
  const padrao = ss.getSheetByName('Sheet1') || ss.getSheetByName('Página1');
  if (padrao) ss.deleteSheet(padrao);

  // Cria abas
  const trans  = ss.insertSheet(SHEET_TRANS);
  trans.appendRow(['id','descricao','valor','data','tipo','catId','catLabel','catIcon','mes','ano']);
  trans.getRange(1,1,1,10).setFontWeight('bold').setBackground('#1D9E75').setFontColor('#ffffff');
  trans.setColumnWidth(1, 150);
  trans.setColumnWidth(2, 200);

  const config = ss.insertSheet(SHEET_CONFIG);
  config.appendRow(['chave','valor']);
  config.appendRow(['pin','1234']);
  config.getRange(1,1,1,2).setFontWeight('bold').setBackground('#1D9E75').setFontColor('#ffffff');

  // Mostra ID no log
  Logger.log('✅ Planilha criada!');
  Logger.log('ID da planilha: ' + id);
  Logger.log('URL da planilha: ' + url);
  Logger.log('');
  Logger.log('⚠️  Cole o ID abaixo na variável SHEET_ID no topo do Code.gs:');
  Logger.log(id);

  SpreadsheetApp.getUi().alert(
    '✅ Planilha criada!\n\n' +
    'ID: ' + id + '\n\n' +
    'Cole este ID na variável SHEET_ID no topo do Code.gs\n\n' +
    'URL: ' + url
  );
}
