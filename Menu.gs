function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Scheduler')
      .addItem('Initialize', 'menuItem1')
      .addToUi();
}

function menuItem1() {
  initialize();
}