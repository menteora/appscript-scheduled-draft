function SheetHelper(sheetName, headerRow) {
  this.headerRow = !headerRow ? 1 : headerRow;
  this.sheet = !sheetName ? SpreadsheetApp.getActiveSheet() : SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
}

SheetHelper.prototype.clearSheet = function () {
  if(this.sheet.getLastRow() !== 0) {
     this.sheet.deleteRows(this.headerRow, this.sheet.getLastRow());
  }
};

SheetHelper.prototype.alert = function (prompt) {
 SpreadsheetApp.getUi().alert(prompt);
};

SheetHelper.prototype.log = function (data) {
  Logger.log(data);
};

SheetHelper.prototype.logStringify = function (data) {
  Logger.log(JSON.stringify(data));
};