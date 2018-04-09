/* CHANGELOG

- added sheetfu google script library
- added momentjs pure javascript library
- getMessage().getId() not working properly, change when draft was modified

*/
/* Written originally by Amit Agarwal of labnol.org */
/* Original Post: https://ctrlq.org/code/19716-schedule-gmail-emails */

var sheetName = "draft";

function testLoopThroughItems() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("draft");
    var gridRange = sheet.getDataRange();
    
    var table = new Table(gridRange);
    
    for (var i = 0; i < table.items.length; i ++) {
        var item = table.items[i];
        // This will print in gas console the first name of everyone in the Table.
        Logger.log(item.getFieldValue("ID"))    
    }
    
    // You can commit the whole table instead of committing per item too
    // table.commit()
}

function initialize() {

  /* Clear draft form Speadsheet */
  var mySheatHelper = new SheetHelper(sheetName, 2);
  mySheatHelper.clearSheet();
  
  var sheet = mySheatHelper.getCurrentSheet();
  var gridRange = sheet.getDataRange();
  var grid = new Table(gridRange);
  
  /* Delete all existing triggers */
  mySheatHelper.deleteTriggers("sendMails");

  /* Import Gmail Draft Messages into the Spreadsheet */
  var drafts = GmailApp.getDrafts();
  if (drafts.length > 0) {
    for (var i = 0; i < drafts.length; i++) {
      if (drafts[i].getMessage().getTo() !== "") {
        grid.add({
          "ID": drafts[i].getId(), 
          "TO": drafts[i].getMessage().getTo(), 
          "SUBJECT": drafts[i].getMessage().getSubject(), 
          "DATE": moment().format("DD/MM/YYYY HH.mm.ss"), 
          "STATUS": ""});
        grid.commit()
      }
    }
  }
}

/* Create time-driven triggers based on Gmail send schedule */
function setSchedule() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var data = sheet.getDataRange().getValues();
    var time = moment().format("DD/MM/YYYY HH.mm.ss");
    var code = [];
    for (var row in data) {
        if (row != 0) {
          var schedule = moment(data[row][3]).format("DD/MM/YYYY HH.mm.ss");
          SpreadsheetApp.getUi().alert(schedule);
          SpreadsheetApp.getUi().alert(time);
          SpreadsheetApp.getUi().alert(moment(schedule).format("DD/MM/YYYY HH.mm.ss"));
            if (schedule !== "") {
                if (schedule > time) {
                    ScriptApp.newTrigger("sendMails")
                        .timeBased()
                        .at(schedule)
                        .inTimezone(SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone())
                        .create();
                    code.push("Scheduled");
                } else {
                    code.push("Date is in the past");
                }
            } else {
                code.push("Not Scheduled");
            }
        }
    }
    for (var i = 0; i < code.length; i++) {
        sheet.getRange("E" + (i + 2)).setValue(code[i]);
    }
}


function sendMails() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var data = sheet.getDataRange().getValues();
    var time = new Date().getTime();
    for (var row = 1; row < data.length; row++) {
        if (data[row][4] == "Scheduled") {
            var schedule = data[row][3];
            if ((schedule != "") && (schedule.getTime() <= time)) {
                var message = GmailApp.getMessageById(data[row][0]);
                var body = message.getBody();
                var options = {
                    cc: message.getCc(),
                    bcc: message.getBcc(),
                    htmlBody: body,
                    replyTo: message.getReplyTo(),
                    attachments: message.getAttachments()
                }

                /* Send a copy of the draft message and move it to Gmail trash */
                GmailApp.sendEmail(message.getTo(), message.getSubject(), body, options);
                //message.moveToTrash();
                sheet.getRange("E" + (row + 1)).setValue("Delivered");
            }
        }
    }
}