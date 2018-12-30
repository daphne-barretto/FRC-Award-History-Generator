var spreadsheet = SpreadsheetApp.openById("1XoVokErF4Ez2IpChoT2Qgcd6QkjK1adgFIAJIfQv7_0");
var optionsSheet = spreadsheet.getSheetByName("Options");

function onOpen(){
  
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('FRC Award History Generator')
      .addItem('Generate Sheet', 'createSheet')
      .addItem('Generat Report', 'createReport')
  .addToUi();
  
}

function deleteEmptyRowsInOptions(){

  for(var i = 2; i <= optionsSheet.getMaxRows(); i++){
    var teamNumber = optionsSheet.getRange(i, 1).getValue();
    if(!(teamNumber === parseInt(teamNumber, 10))){
      optionsSheet.deleteRow(i);
      i--;
    }
  }
  
}

function createReport(){

  var date = Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "MM/dd/yyyy HH:mm:ss")
  var doc = DocumentApp.create("Award History Report " + date);
  var body = doc.getBody();
  
  doc.addEditor("mars2614@gmail.com");
  
  deleteEmptyRowsInOptions();
  
  var numTeams = optionsSheet.getLastRow() - 1;
  
  for(var i = 0; i < numTeams; i++){
  
    var teamNumber = optionsSheet.getRange(i + 2, 1).getValue();
    var firstYear = optionsSheet.getRange(i + 2, 2).getValue();
    var lastYear = optionsSheet.getRange(i + 2, 3).getValue()
    
    var title = teamNumber + "'s Awards from " + firstYear + " to " + lastYear;
  
    var awardHistory = getAwardHistory(teamNumber, firstYear, lastYear);
  
    body.editAsText().appendText(title);
  
    for(var j = 0; j < awardHistory.length; j++){
      var nextAward = awardHistory[j][0] + " " + awardHistory[j][1] + " " + awardHistory[j][2];
      body.appendListItem(nextAward).setGlyphType(DocumentApp.GlyphType.BULLET);
    }
  
    body.editAsText().appendText("\n\n")
  
  }
    
}

function getAwardHistory(teamNumber, firstYear, lastYear){
  
  var awardHistory = [];
  // 2D array containing award history
  // FORMAT:
  // [
  //   [year, event, awardName]
  //   [year, event, awardName]
  // }
  
  var teamKey = "frc" + teamNumber;
  
  for(year = firstYear; year <= lastYear; year++){

    var url = "https://www.thebluealliance.com/api/v3/team/" + teamKey + "/awards/" + year;
      var options = {
        "method": "GET",
        "headers": {
          "X-TBA-Auth-Key": "ElyWdtB6HR7EiwdDXFmX2PDXQans0OMq83cdBcOhwri2TTXdMeYflYARvlbDxYe6"
        },
        "payload": {
        }
      };
      
      var response = JSON.parse(UrlFetchApp.fetch(url, options));
  
    for (var i =  0; i < response.length; i++){
      awardHistory.push(Array.prototype.concat(getEventName(response[i].event_key), [response[i].name]))
    }
  }
  
  return awardHistory;
  
}

function createSheet(){

  var teamNumber = optionsSheet.getRange("A2").getValue();
  var firstYear = optionsSheet.getRange("B2").getValue();
  var lastYear = optionsSheet.getRange("C2").getValue();
  
  var title = teamNumber + "'s Awards from " + firstYear + " to " + lastYear;
  var sheet = spreadsheet.getSheetByName(title)

  var awardHistory = [];
  // 2D array containing award history
  // FORMAT:
  // [
  //   [year, event, awardName]
  //   [year, event, awardName]
  // }

  // Delete pre-existing sheet of same parameters
  if (sheet != null) {
        spreadsheet.deleteSheet(sheet);
  }
 
  // Create new sheet
  spreadsheet.insertSheet(title);
  sheet = spreadsheet.getSheetByName(title);
  
  // Format sheet
  sheet.getRange("A1:C1").merge();
  sheet.getRange("A1:C1").setValue(title);
  sheet.getRange("A2:C2").setValues([["Year", "Event", "Award"]]);
  sheet.getRange("A1:C2").setHorizontalAlignment("center");
  sheet.deleteColumns(4, 23);
  sheet.deleteRows(3, 997);
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidths(2, 2, 300);

  awardHistory = getAwardHistory(teamNumber, firstYear, lastYear);
   
   Logger.log(awardHistory);
 
   // Populate Sheet
   sheet.getRange(3, 1, awardHistory.length, 3).setValues(awardHistory);
 
 }

 function getEventName(eventKey){
   
  var url = "https://www.thebluealliance.com/api/v3/event/" + eventKey;
    var options = {
      "method": "GET",
      "headers": {
        "X-TBA-Auth-Key": "ElyWdtB6HR7EiwdDXFmX2PDXQans0OMq83cdBcOhwri2TTXdMeYflYARvlbDxYe6"
      },
      "payload": {
      }
    };
    
    var response = JSON.parse(UrlFetchApp.fetch(url, options));

  return [response.year, response.name];
 
 }