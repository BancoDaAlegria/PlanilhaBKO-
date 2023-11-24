/** @OnlyCurrentDoc */

function myFunction() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E3:F7').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('>60 dias'), true);
};

function _60dias() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E9:F13').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('>60 dias'), true);
};

function primeiro() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E9:F13').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('>60 dias'), true);
};

function segundo() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E3:F7').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Zerados'), true);
};

function terceiro() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B46').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('PENDENTES DE CADASTRO'), true);
};

function quarto() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E21:F25').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('BASE CADASTRO'), true);
};

function quinto() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B27:D31').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('-11 vendas'), true);
};

function sexto() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E33:F37').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('ATIVAR!'), true);
};

function cadastroBase() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('K13').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('BASE CADASTRO'), true);
};
