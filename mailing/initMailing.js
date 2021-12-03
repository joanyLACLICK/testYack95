function initAllSheets(){
    initSie();
    initDu();
   
}

function initSie(){
    initMailing("Philipp_Justus_Sie")
}
function initDu(){
    initMailing("Philipp_Justus_Du")
}


function initMailing(sheetName){
    //to hard code the headers and other properties
    
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var headersLocation = activeSpreadsheet.getSheetByName(sheetName);
    
    
  
    if(headersLocation != null){
        activeSpreadsheet.deleteSheet(headersLocation);
    }
    headersLocation = activeSpreadsheet.insertSheet();
    headersLocation.setName(sheetName);
   
    const headersSettings = [
        ["Settings"],
        ["Subject"],
        ["From"], 
        ["Reply To"],
        ["Bcc"],
        ["DefJam Id"]

    ]
    const startingRow = [
      "List of clients",
      "",
      "",
      "",
      ""
    ]
    
    const headerInvitation = [
        "First Name",
        "Email",
        "Status",
        "Send?",
        "Custom message?"
    ]
    
    const listClients = headersLocation.getRange(9,1,1,headerInvitation.length)
    listClients.setValues([startingRow]);
    listClients.setFontWeight("bold");
    listClients.setFontColor("#ffffff")
    listClients.setBackground("#185d72")
    listClients.setHorizontalAlignment("center")
    listClients.mergeAcross()
    
    headersLocation.autoResizeColumn(2)
    
    const subjectCell = headersLocation.getRange(2,2,1,1)
    subjectCell.setBackground("#fff2cc")
    const rangeSettings = headersLocation.getRange(1,1, headersSettings.length,1)
    rangeSettings.setValues(headersSettings);
    rangeSettings.setBackground("#efefef")
    const settingHeader = headersLocation.getRange(1,1,1,1)
    settingHeader.setFontColor("#ffffff")
    settingHeader.setBackground("#185d72")
    
    const a = rangeSettings.getValues()


    const rangeInvitation = headersLocation.getRange(10, 1, 1, headerInvitation.length);
    rangeInvitation.setValues([headerInvitation])
    rangeInvitation.setFontWeight("bold");
    rangeInvitation.setFontColor("#ffffff")
    rangeInvitation.setBackground("#185d72")
    const b = rangeInvitation.getValues()

}


