const LAST_HEADER = 10;
const FIRST_ROW = LAST_HEADER + 1;


function preview(){
  
    //variables from sheet
    var firstName = 0;
    var email = 1;
    var status = 3;//sent or not?
    var customMessage = 4
 
    //from where I am going to take the invitation template
    //var invitationTemplate = HtmlService.createTemplateFromFile("html/token_1");(Three guys)
    var invitationTemplate = HtmlService.createTemplateFromFile("html/token_2"); //()

    //from where I am going to extract the invitees data
    var invitees = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
    var inviteesRange = invitees.getRange(FIRST_ROW,1, invitees.getLastRow()-LAST_HEADER, invitees.getLastColumn()).getValues();//returns the rectangular grid of values for this range
    //invitation subject
    var subject = invitees.getRange(2,2,1,1).getValue();
  
    //to be able to select the right invitation to preview
    var activeRange = invitees.getActiveCell();
    const rowNumber = activeRange.getRow() - FIRST_ROW;
    const row = inviteesRange[rowNumber];
  
    //taking the information from the list of clients
    invitationTemplate.fn = row[firstName];
  
    invitationTemplate.customMessage = row[customMessage];
  
    Logger.log("custom message", customMessage)

    if(!row[email]){
      return;
    }
   //to DELETE
    //if(!row[customMessage]){
      //return;
    //}
  
    //to create the HTML output
    var htmlMessage = invitationTemplate.evaluate()
    
    htmlMessage.setWidth(1000);
    htmlMessage.setHeight(800);
    SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .showModalDialog(htmlMessage, subject);
  
}

//