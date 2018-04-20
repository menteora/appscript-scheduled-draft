/* CHANGELOG

- added sheetfu google script library
- added momentjs pure javascript library
- getMessage().getId() not working properly, change when draft was modified

*/
/* Written originally by Amit Agarwal of labnol.org */
/* Original Post: https://ctrlq.org/code/19716-schedule-gmail-emails */

var sheetName = "draft";

function initialize() {
  
  /* Clear draft form Speadsheet */
  var mySheetHelper = new SheetHelper(sheetName, 2);  
  //mySheetHelper.clearSheet();
  
  var gridRange = mySheetHelper.getCurrentSheet().getDataRange();
  var grid = new Table(gridRange);
  
  /* Import Gmail Draft Messages into the Spreadsheet */
  var drafts = GmailApp.getDrafts();
  if (drafts.length > 0) {
    for (var i = 0; i < drafts.length; i++) {
      if (drafts[i].getMessage().getTo() !== "" && isNew(grid,drafts[i].getId()) ) {
        grid.add({
          "ID": drafts[i].getId(), 
          "TO": drafts[i].getMessage().getTo(), 
          "SUBJECT": drafts[i].getMessage().getSubject(), 
          "DATE": moment().format("DD/MM/YYYY HH:mm"), 
          "STATUS": ""});
        grid.commit();
      }
    }
  }
}

function isNew(table, id) {
  var records = table.select({"STATUS":"","ID":id});
  return (records.length > 0) ? false : true;
}

function sendEmail() {
  var mySheetHelper = new SheetHelper(sheetName);
  var gridRange = mySheetHelper.getCurrentSheet().getDataRange();
  var table = new Table(gridRange);
  var records = table.select({"STATUS": ""});
  
  for (var i = 0; i < records.length; i ++) {
    try {
      var schedule = moment(records[i].getFieldValue("DATE"), "DD/MM/YYYY HH:mm");
      if (schedule.isSameOrBefore(moment())){
        GmailApp.getDraft(records[i].getFieldValue("ID")).send();
        records[i].setFieldValue("STATUS","DELIVERED");
        records[i].setFieldValue("DETAILS", moment().format("DD/MM/YYYY HH:mm"));
        records[i].commit();
      }
    } catch (err){
      records[i].setFieldValue("STATUS","NOT DELIVERED");
      records[i].setFieldValue("DETAILS", JSON.stringify(err));
      records[i].commit();
    }
  }
}