/**
* This is a library for project.gs and schedule.gs
* If you have any problem, you can send email for me or 
* look up this document https://docs.google.com/document/d/1_4o12u4Ehaa1pFSKnKE8XW4IQvwu-phL5Qe5k_sc39Q/edit
* @author Jay
* @Time 2020.5.12
* @Email 22365194popo@gmail.com
*/

/**
* Being handler for project.gs
*/
function getSheetId(){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    
    return sheet.getSheetId();
}
  
/**
 * Being handler for project.gs
 */
function getSheetName(){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];

    return sheet.getSheetName(); 
}
  
/**
 * In a given user's calendar, look for occurrences of the given keyword
 * in events within the specified date range and return any such events
 * found.
 * @param {Sheet} sheet The sheet what you want to get the boundary.
 * @param {int} row1 To determine the end of row.
 * @param {int} col1 To determine the end of row.
 * @param {int} row2 To determine the end of col.
 * @param {int} col2 To determine the end of col.
 * @return {int} first, second Return two values, first for the boundary of row, second for the boundary of col.
 */
function getBoundary(sheet, row1, col1, row2, col2){
  var endOfRow;
  var endOfCol;
  var valueOfRows = sheet.getRange(row1, col1, sheet.getColumnWidth(1)-1, 1).getValues();
  var valueOfCols = sheet.getRange(row2, col2, 1, sheet.getRowHeight(1)-1).getValues();
 
  for (var row=0; row<sheet.getColumnWidth(1); ++row){ // 150
    if (!valueOfRows[row][0] || 0===valueOfRows[row][0].length){
      endOfRow = row;
      break;
    }
  }
  for (var col=0; col<sheet.getRowHeight(1); ++col){ // 21
    if (!valueOfCols[0][col] || 0===valueOfCols[0][col].length){ // Get boundary ''
      endOfCol = col;
      break;
    }
  } 
  return {
    first: endOfRow + 1, // Start at 2, so plus 1 to index end of row.
    second: endOfCol,
  };
}