/**
* This is a independent study schedule.
* If you have any problem, you can send email for me or 
* look up this document https://docs.google.com/document/d/1_4o12u4Ehaa1pFSKnKE8XW4IQvwu-phL5Qe5k_sc39Q/edit
* @author Jay
* @Time 2020.5.12
* @Email 22365194popo@gmail.com
*/

var ss2; 
var sheet2; 
var endOfRow; // boundary of row
var endOfCol; // boundary of col
// 108-1 -> academicYear-semester
var academicYear; 
var semester; 
var nsn; // new spread sheet name

/**
* Copies the sheet to a given spreadsheet, 
* which can be the same spreadsheet as the source. 
*/
function addSpreadSheet(){
  // Copy sheet.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sn = ss.getSheetName(); // 108-1
  
  academicYear = sn[0].concat(sn[1]).concat(sn[2]); // 108
  semester = sn[sn.length-1]; // Get the last one. eg: 108-1 -> 1
  changeValue(academicYear, semester);
  SpreadsheetApp.getActiveSpreadsheet().duplicateActiveSheet().setName(nsn); // Copy sheet.
}

/**
* Delete all cell value in new spread sheet.
*/
function deleteValue(){
  // Delete Sheet value.
  ss2 = SpreadsheetApp.getActiveSpreadsheet()
  sheet2 = ss2.getSheetByName(nsn);
  var maxRows = sheet2.getMaxRows();
  var maxCols = sheet2.getMaxColumns();
  var valueOfRows = ss2.getSheetValues(3, 2, maxRows, 1); // Check row we dont care title.
  var valueOfCols = ss2.getSheetValues(2, 2, 1, maxCols); // But check col we need title as based.
  
  for (var row=0; row<valueOfRows.length; ++row){ // 2020/09/11(三), 2020/09/13(五), ... , ''  len() = 1000
    if (isEmpty(valueOfRows[row][0])){ // Get boundary ''
      endOfRow = row;
      break;
    }
  }
  for (var col=0; col<valueOfCols[0].length; ++col){ // 日期, 演講者, ... , '' len() = 6
    if (isEmpty(valueOfCols[0][col])){ // Get boundary ''
      endOfCol = col;
      break;
    }
  }
  // if the range is not totally blank.
  if (!sheet2.getRange(3, 2, endOfRow, endOfCol).isBlank()){
     sheet2.getRange(3, 2, endOfRow, endOfCol).setValue(''); // clear sheet
  }
}

/**
* Set yellow and green background to white background.
*/
function setBackground(){
  // Set background color.
  ss2.getSheetByName(nsn).getRange(3, 2, endOfRow, 1).setBackground('white');
}

/**
* Changed the academic year and semester value in title.
*/
function setTitleName(){
  // Set name.
  var titleName = sheet2.getSheetValues(1, 1, 1, 1)[0]; // '108學年度第一學期專題演講時程表'
  var newTitleName = '';

  newTitleName  = newTitleName.concat(academicYear, '學年度第', numToChinese(semester), '學期專題演講時程表');
  sheet2.getRange(1, 1, 1, 1).setValue(newTitleName);
}

/**
* Get the filr from google drive.
*/
function getDriveFile(){
  var file = DriveApp.getFileById(academicYear + '.txt') // get file from drive
  var startDate = file.getBlob().getDataAsString(); // the date of new semester
}

/**
* Handler semester.
* @param {int} n a Acdemicyear
*/
function numToChinese(n){
  if (n == 1)
    return "一";
  else
    return "二";
}

/**
* To check string is empty or not.
* @param {str} str Input a string to check if empty or not.
*/
function isEmpty(str){
  return (!str || 0 === str.length);
}

/**
* Need to replace academic year and semester value,
* when semester or academic year changed. 
* @param {str} a Acdemicyear
* @param {str} s Semester
*/
function changeValue(a, s){
  var str = '';
  var pa = parseInt(a);
  var ps = parseInt(s);
  var ta = pa.toString();
  var ts = ps.toString();
  
  if (s == 1){ // 108-1 -> 108-2
    ps += 1;
    ts = ps.toString();
    nsn = str.concat(a, '-', ts);
    semester = ts; // convient for calling numToChinese function.
  }
  else{ // 108-2 -> 109-1
    pa = parseInt(a); // string to int.
    pa += 1;
    ps -= 1;
    ta = pa.toString(); // int to string.
    ts = ps.toString();
    nsn = str.concat(ta, '-', ts);
  }
}

function main(){
  addSpreadSheet();
  deleteValue();
  setBackground();
  setTitleName();
  //setDate();
}