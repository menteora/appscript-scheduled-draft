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
  mySheetHelper.clearSheet();
  
  var gridRange = mySheetHelper.getCurrentSheet().getDataRange();
  var grid = new Table(gridRange);
  
  /* Import Gmail Draft Messages into the Spreadsheet */
  var drafts = GmailApp.getDrafts();
  if (drafts.length > 0) {
    for (var i = 0; i < drafts.length; i++) {
      if (drafts[i].getMessage().getTo() !== "") {
        grid.add({
          "ID": drafts[i].getId(), 
          "TO": drafts[i].getMessage().getTo(), 
          "SUBJECT": drafts[i].getMessage().getSubject(), 
          "DATE": moment().format("DD/MM/YYYY HH:mm"), 
          "STATUS": ""});
        grid.commit()
      }
    }
  }
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
        /*
        Logger.log(records[i].getFieldValue("ID"));
        var draftId = records[i].getFieldValue("ID")
        var message = GmailApp.getDraft(draftId).getMessage();
        Logger.log(message);
        Logger.log(message.getBody());
        Logger.log(message.getTo());
        Logger.log((message.getBcc()) == "" ? Session.getActiveUser().getEmail() : message.getBcc() + ", " + Session.getActiveUser().getEmail());
        var body = message.getBody();
        var options = {
          cc: message.getCc(),
          bcc: (message.getBcc()) == "" ? Session.getActiveUser().getEmail() : message.getBcc() + ", " + Session.getActiveUser().getEmail(),
          htmlBody: body,
          replyTo: message.getReplyTo(),
          attachments: message.getAttachments()
        }
        
        GmailApp.sendEmail(message.getTo(), message.getSubject(), body, options);
        */
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