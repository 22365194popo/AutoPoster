/**
* This is a fornt end project to handler the data from 
* speecher and pass the data to back end.
* If you have problem, you can send email for me or 
* look up this document https://docs.google.com/document/d/1_4o12u4Ehaa1pFSKnKE8XW4IQvwu-phL5Qe5k_sc39Q/edit
* @author Jay
* @Time 2020.5.12
* @Email 22365194popo@gmail.com
*/

var sheet = getSpreadSheet();
var endOfRow = library.getBoundary(sheet, 2, 2, 1, 2).first;
var endOfCol = library.getBoundary(sheet, 2, 2, 1, 2).second;
var values;
var title;

/**
* This method gets the active sheet in a spreadsheet.
*/
function getSpreadSheet(){
  // Gets the active sheet in a spreadsheet.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  return sheet;
}

/**
* This method gets the title of spreadsheet.
*/
function getSheetTitle(){
  // [電子郵件地址, 演講時間, 演講者, 演講題目, 學歷/公司, 交通工具]
  return sheet.getSheetValues(1, 2, endOfRow, endOfCol)[0];
}

/**
* This method gets the active sheet in a spreadsheet.
* input row and col to get the value.
*/
function getSheetValue(){ 
  // [電子郵件地址.value, 演講時間.value, 演講者.value, 演講題目.value, 學歷/公司.value, 交通工具.value]
  return sheet.getSheetValues(endOfRow, 2, 1, endOfCol)[0]; 
}

/**
* This method makes a post request with json payload.
* We need this json to do our poster.
*/
function postRequest(){
  values = getSheetValue();
  var uploadData = {
    '電子郵件': values[0],
    '演講時間': values[1],
    '演講者': values[2],
    '演講題目': values[3],
    '學歷/公司': values[4],
    '交通工具': values[5],
  };
  // Make a POST request with a JSON payload.
  var options = {
    'method' : 'post',
    'contentType': 'application/json',
    // Convert the JavaScript object to a JSON string.
    'payload' : JSON.stringify(uploadData)
  };
  var response = UrlFetchApp.fetch('http://d62d8f2e.ngrok.io', options);
}

/**
* This method just sending email to speecher.
* The content of email deponds on what speecher takes transportation. 
* @param {str} p To determine which case we need to handle.
* @param err Exception error.
*/
function sendEmail(p, err){
  var email;
  var message;
  var subject;
  
  if (err != err){
    if (p == s){ // handling for speecher.
      email = values[0][0];
      var transportation = values[5];
      // This constant is written in column C for rows for which an email
      // has been sent successfully.
      var EMAIL_SENT = 'EMAIL_SENT';
      
      switch (transportation){
        case '自行開車':
          message = '{}您好，很榮幸邀請您於{}前來擔任演講嘉賓。若有後續我們會再將相關資訊或資料表單寄給老師您，在麻煩您協助回覆。\
                      若有任何問題，歡迎回覆此郵件，謝謝。'.format(values[3], values[1]);
          break;
        case '高鐵':
          message = '{}您好，很榮幸邀請您於{}前來擔任演講嘉賓，在此向您說明關於交通部分之事項。本校對於高鐵乘車費用有補助，老師您可持高鐵票根予系辦負責人核銷。\
            另外，若有後續我們會再將相關資訊或資料表單寄給老師您，在麻煩您協助回覆。於此，再次感謝您前來擔任本次演講嘉賓，若有任何問題，歡迎回覆此郵件，謝謝。'.format(values[3], values[1])
          break;
        case '台鐵':
          message = '{}您好，很榮幸邀請您於{}前來擔任演講嘉賓，在此向您說明關於交通部分之事項。本校對於台鐵乘車費用有補助，煩老師抵達貴校後向系辦負責人告知台鐵乘坐之起終點站，方便進行核銷。\
          另外，若有後續我們會再將相關資訊或資料表單寄給老師您，在麻煩您協助回覆。於此，再次感謝您前來擔任本次演講嘉賓，若有任何問題，歡迎回覆此郵件，謝謝。'.format(values[3], values[1])
          break;
      }
        subject = '國立宜蘭大學-資工系辦公室';
        MailApp.sendEmail(email, subject, message);
        // Make sure the cell is updated right away in case the script is interrupted
        SpreadsheetApp.flush();
        Logger.log('Sending email successfully.');
      }
   }
  else{
    email = '22365194popo@gmail.com';
    subject = '海報生成系統';
    message = err;
    if (p == m){ // handing for sending email for me.
      MailApp.sendEmail(email, subject, message);
      Logger.log(err);
    }
    if (p == u){ // handing for updating spread sheet.
      MailApp.sendEmail(email, subject, message);
      Logger.log(err);
    }
  }
}

/**
* If speecher do form completely, we need to update 
* the independent study schedule.
*/
function updateSpreadSheet(){
  try{
    if (sheet.getName() == 'test') { // That things needs to be modified when sheet name changed.
      var ss2 = SpreadsheetApp.openById(library.getSheetId()) // every sheet has unique sheet id. '5.042573E7'
      var sheet2 = SpreadsheetApp.openById(library.getSheetName()); // Name of sheet. '108-1' 
      var date = sheet2.getSheetValues(3, 2, date.length, 1); // Get data col, type list.
      
      for (var row=0; row<data[0].length; ++row){ // Visit col.
        if ((values[0][1].concat('三')||values[0][1].concat('五')) ==  date[0][row]) { // if match the same data, eg: 2020/5/13(三) == 2020/5/13(三)
          if (date[row].getBackgrounds() == '#FFFFFF'){ // none-white is OK.
            for (var col=0; col<3; ++col){ // Update the data from sheet1 to sheet2。 update speecher、title、company
              sheet2.getRange(row, col).setValue(values[0][col]);
              break;
            }
          }
        }
      }
    }
  }catch(err){
    sendingEmail(u, err); // Prvent from filling in none-white cell. 
  }
}

/**
* This method is to run all function.
* If error occured, sys sent email to me.
*/
function main(){
  /*try{
    spreadSheet();
    postRequest();
    sendEmail(s);
    updateSpreadSheet();
  }catch(err){
    sendEmail(m, err);
  }finally {
    Logger.log("All successfully.");
  }*/
  postRequest();
}