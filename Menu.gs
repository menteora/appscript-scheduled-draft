function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Scheduler')
      .addItem('Initialize', 'menuItem1')
      .addSeparator()
      .addItem('SetSchedule', 'menuItem2')
      .addToUi();
}

function menuItem1() {
  initialize();
}

function menuItem2() {
  setSchedule();
}