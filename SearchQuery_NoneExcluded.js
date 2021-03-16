var SHEETNAME = 'GoogleAdsSearchQueries';
var ROWS = 1;

function CheckSpreadSheet() {
  var sheetIterator = DriveApp.getFilesByName(SHEETNAME);
  if (!sheetIterator.hasNext()) {
    return 0;
  } else {
    sheetId = sheetIterator.next().getId();
    return sheetId;
  }

}

function CreateSpreadSheet(name, rows, columns) {
  var name = name;
  var rows = rows;
  var columns = columns;
  var sheet = SpreadsheetApp.create(name, rows, columns);
  return sheet;
}

//процент от общей суммы
function Percentage(numberOne, numberTwo) {
  if (numberTwo != 0) {
    var sum = numberOne + numberTwo
    return numberOne*100/sum;
  } else {
    return 0;
  }
}

function main() {
  var columns = ["Date", "None", "Clicks", "Cost", "Excluded", "Clicks", "Cost", "Excluded%", "Clicks%", "Costs%"];
  
  var sheetId = CheckSpreadSheet();
  if(sheetId == 0) {
    CreateSpreadSheet(SHEETNAME, ROWS, columns.length);
    sheetId = CheckSpreadSheet();
  }
  
  
  var spreadSheet = SpreadsheetApp.openById(sheetId);
  sheet = spreadSheet.getSheets()[0];
  var range = sheet.getRange("A1:J1");
  
  for (var i = 0; i < columns.length; i++) {
    var cell = range.getCell(1, i+1);
    cell.setValue(columns[i]);
  }
  
  var report = AdsApp.report(
   "SELECT Query, QueryTargetingStatus, Clicks, Cost" +
   " FROM SEARCH_QUERY_PERFORMANCE_REPORT " +
   " DURING YESTERDAY");
  var rows = report.rows();
  var row = rows.next();
  
  var none_count = 0;
  var none_clicks = 0;
  var none_cost = 0;
  var excluded_count = 0;
  var excluded_clicks = 0;
  var excluded_cost = 0;
  while (rows.hasNext()) {
    var row = rows.next();
    if (row['QueryTargetingStatus'] == 'Excluded') {
      excluded_count++;
      excluded_clicks = excluded_clicks + parseInt(row['Clicks']);
      excluded_cost = excluded_cost + parseFloat(row['Cost']);
    } else {
      none_count++;
      none_clicks = none_clicks + parseInt(row['Clicks']);
      none_cost = none_cost + parseFloat(row['Cost']);
    }
  }
  none_cost = Math.round(none_cost);
  excluded_cost = Math.round(excluded_cost);
 

  var excludedPercent = Percentage(excluded_count, none_count);
  excludedPercent = Math.round(excludedPercent);
  var clicksPercent = Percentage(excluded_count, none_clicks);
  clicksPercent = Math.round(clicksPercent);
  var costPercent = Percentage(excluded_count, none_cost);
  costPercent = Math.round(costPercent);
  
  var today = new Date();
  var yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);
  
  var dayMonth = yesterday.getDate();
  //starts from 0;
  var month = yesterday.getMonth() + 1;
  var year = yesterday.getYear();
  var yesterdayNewFormat = dayMonth + '/' + month + '/' + year;
  var data = [yesterdayNewFormat, none_count, none_clicks, none_cost, excluded_count, excluded_clicks, excluded_cost, excludedPercent, clicksPercent, costPercent];
  
  var lastRow = sheet.getLastRow();
  var prevDateCell = sheet.getRange(lastRow, 1).getDisplayValue();
  if (prevDateCell == 'Date') {
    sheet.appendRow(data);
  } else {
      var dateArray = prevDateCell.split("/");
      var prevDateCell = dateArray[1] + '/' + dateArray[0] + '/' + dateArray[2];
      var prevDateCell = new Date(prevDateCell);
      Logger.log(yesterdayNewFormat);
      //var previousDay = newDate.getDate();
      //Logger.log(previousDay);
      Logger.log(yesterday);
      Logger.log(prevDateCell < yesterday);
      if(prevDateCell == yesterday) {
          sheet.appendRow(data);
      } else if(prevDateCell < yesterday) {
          Logger.log('Warning!');
          Logger.log('Previous day is missing!');
          Logger.log('Appending new data!');
          sheet.appendRow(data);
      } else {
          Logger.log('Data is already given!');
      }
    }
}
