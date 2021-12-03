function runInvitation(callback){
  
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert(
        'Please confirm that you want to send the invitations',
        'Are you sure you want to continue?',
        ui.ButtonSet.YES_NO);
        console.log("result", result)
    if(result == ui.Button.YES){
        Logger.log('The user clicked "Yes."');
        const result = callback();
        
    }
}

function sendInvitationEmailsGooglers(sheetName, html){
   
    //from where I am going to take the invitation template
    //var invitationTemplate = HtmlService.createTemplateFromFile("html/token_1");(Three guys)
    var invitationTemplate = HtmlService.createTemplateFromFile(html); //()

    //from where I am going to extract the invitees data
    var invitees = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    var columnMap = getColumnMap(invitees, 10); //to get the column names
    
    console.log("columnMap", columnMap)
     //variables from sheet
    var salutation = columnMap["Salutation"];
   
    var firstName = columnMap["First Name"];
    var lastName = columnMap["Last Name"];
    var email = columnMap["Email"];
    
    var cc = columnMap["Cc"];//in case I have one or more
    var status = columnMap["Status"];
    var SENT = "SENT";
    //custom message
    var customMessage = columnMap["Custom message?"];
    
     
    var inviteesRange = invitees.getRange(11,1, invitees.getLastRow()-10, invitees.getLastColumn()).getValues();//returns the rectangular grid of values for this range
 
    //invitation subject
    var subject = invitees.getRange(2,2,1,1).getValue();
    //invitation signature
    var signature = invitees.getRange(3,2,1,1).getValue();
    //replyTo
    var replyTo = invitees.getRange(4,2,1,1).getValue();
    //with copy to
    var bcc = invitees.getRange(5,2,1,1).getValue();
    
    //invites
    var count = 0;
    inviteesRange.forEach((row, index)=>{

        invitationTemplate.fn = row[firstName];
                          
        if(lastName){
          invitationTemplate.ln = row[lastName];
        }
  
         invitationTemplate.sl = salutation || salutation === 0 ? row[salutation]:"";
      
  
        if(customMessage){
          invitationTemplate.customMessage = row[customMessage];
        }
        
        if(row[status] === SENT){
            return;
        }
  
        if(!row[email]){
            return;
        }
   
     
        if(row[status]!== SENT){
          var htmlMessage = invitationTemplate.evaluate().getContent();
          
          GmailApp.sendEmail(
                row[email].toLowerCase().trim(),
                subject,
                "Your email does not support HTML",
                {
                    name: signature,
                    htmlBody: htmlMessage, 
                    replyTo: replyTo,
                    cc:`${row[cc]}`,
                    bcc:bcc,
                }
            );
            
            invitees.getRange(11 + index,  columnMap["Send?"]+1).setValue(SENT);//switch from row array to row spreadsheet
            // Make sure the cell is updated right away in case the script is interrupted
            SpreadsheetApp.flush();
        count++;
        } 
    })  
    Browser.msgBox("Successfully sent " + count + " mails to clients. You can find the sent mails in your Gmail 'Sent' folder.")
}


//this document holds only two sheets with different list of clients
function nameSheetSie(){
    runInvitation(()=>sendInvitationEmailsGooglers("Philipp_Justus_Sie", "html/token_Sie_PJ"))
}
function nameSheetDu(){
    runInvitation(()=>sendInvitationEmailsGooglers("Philipp_Justus_Du", "html/token_Du_PJ"))
   
}

function getColumnMap(sheet, row = 1){
  const columns = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  return columns.reduce((result, current, index) => {
    result[current] = index;
    return result;
  }, {})
}



