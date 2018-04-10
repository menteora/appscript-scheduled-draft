function SheetHelper(sheetName, headerRow) {  
  this.headerRow = !headerRow ? 1 : headerRow;
  this.sheet = !sheetName ? SpreadsheetApp.getActiveSheet() : SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
}

SheetHelper.prototype.clearSheet = function () {
  if(this.sheet.getLastRow() !== 0) {
     this.sheet.deleteRows(this.headerRow, this.sheet.getLastRow());
  } else { 
    this.toast("nothing to delete") 
  }
};

SheetHelper.prototype.deleteTriggers = function (triggerName) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggerName && triggers[i].getHandlerFunction() === triggerName) {
      ScriptApp.deleteTrigger(triggers[i]);
    } else {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
};

SheetHelper.prototype.getCurrentSheet = function () {
 return this.sheet;
};

SheetHelper.prototype.toast = function (message) {
 SpreadsheetApp.getActiveSpreadsheet().toast(message);
};

SheetHelper.prototype.alert = function (message) {
 SpreadsheetApp.getUi().alert(message);
};

SheetHelper.prototype.log = function (data) {
  Logger.log(data);
};

SheetHelper.prototype.logStringify = function (data) {
  Logger.log(JSON.stringify(data));
};