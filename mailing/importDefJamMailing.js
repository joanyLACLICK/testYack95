const SPREADSHEET_ID= "1vWyTqCwolzRV1nyNjN1lzNOoaohityfzmHHhQu1cCeo";

function importDefJamMailing(sheetName){
  
    var defJamData = loadDefJam(sheetName);
    var inviteesMap = loadInvitations(sheetName);
    var joinedInviteesMap = joinInvitees(defJamData, inviteesMap);
    saveInvitees(joinedInviteesMap, sheetName);
   
}

function loadDefJam(sheetName){//defjam sheet
    
    //connection with Defjam. Lets keep them as an array because it is the leading data => important in case we have to bring data from multiple DefJam
    var invitationSetting = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
    //var invitationSetting = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    
    var importDefJam = invitationSetting.getRange(6, 2, 1, 1).getDisplayValues()[0][0]; //not array of array
    
    
    var spreadSheet = SpreadsheetApp.openById(importDefJam)
    var location = spreadSheet.getSheetByName("ExportedDelegates");
 
    var defJamSheetData = location.getRange(2, 1, location.getLastRow()-1, location.getLastColumn()).getDisplayValues(); //array of arrays
    

    var defJamData = [];
    for(let i = 0; i < defJamSheetData.length; i++){

        var row = defJamSheetData[i].slice(0);

        var defJamStatus = {
            firstNameDJ:row[0],
            lastNameDJ:row[1],
            emailDefJam:row[2],
            invitationEmailDefJam:row[3],
            rsvpDefJam: row[4],
            defJamIndex: i
        }

        if(defJamStatus.emailDefJam){
            defJamData.push(defJamStatus);
        }

    }
    
    return defJamData
}

function loadInvitations(sheetName){
    //Connection with the Sheet. lets keep it as an object(map) because will need to be updated
    var invitees = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
    var inviteesRange = invitees.getRange(11, 1, invitees.getLastRow()-10, invitees.getLastColumn()).getDisplayValues();//arrays of arrays 
   
    var columnMap = getColumnMap(invitees, 10); //to get the column names
    

    //variables from sheet
    var firstName = columnMap["First Name"];
    var lastName = columnMap["Last Name"];
    var email = columnMap["Email"];
    var status = columnMap["Status"]; 
    var invitationSent = columnMap["Send?"];
   
    var inviteesData = {};
    
    for(let i=0; i<inviteesRange.length; i++){

        var row = inviteesRange[i].slice(0);
       
        var inviteesStatus = {
            firstName:row[firstName], 
            lastName:row[lastName],
            email:row[email],
            status:row[status],
            emailSent:row[invitationSent],
        }
        if(inviteesStatus.email){
            inviteesData[inviteesStatus.email.toLowerCase().trim()] = inviteesStatus;      
        }
    }
   
    return inviteesData

}

function joinInvitees(defJamData, inviteesMap){

    //loop through defjamData
    defJamData.forEach((element)=> {// For every One find th invitee 

        var invitee = inviteesMap[element.emailDefJam];

        //update the invitee if found, and then save back the value to the map
        if(invitee){
            invitee.status = element.rsvpDefJam
             //save it to the map
            inviteesMap[element.emailDefJam] = invitee;
        }
    })
    
    return inviteesMap

}


function saveInvitees(inviteesMap, sheetName){
    //Open sheet
    var destination = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
    var columnMap = getColumnMap(destination, 10);
    var inviteesData = destination.getRange(11, 1, destination.getLastRow()-10, destination.getLastColumn()).getDisplayValues();//arrays of arrays 
  
    //converting my array of object to array of arrays
    var email = columnMap["Email"];
    var status = columnMap["Status"]; 
  
    var result = inviteesData.map((row) => {
        const joinedInvitee = inviteesMap[row[email]];  
        if (joinedInvitee){
          row[status] = joinedInvitee.status; 
        }
                                  
        return row;
    })
    
    
    //clear values from sheet 
    if (destination.getLastRow()-10 > 0){
      destination.getRange(11, 1, destination.getLastRow()-10, destination.getLastColumn()).clear();
    }
    
    // save the new list of invitees to the sheet 
    if (result.length > 0){
      destination.getRange(11,1, result.length, result[0].length).setValues(result)
    }
    
   
}



function getColumnMap(sheet, row = 1){
  const columns = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  return columns.reduce((result, current, index) => {
    result[current] = index;
    return result;
  }, {})
}


