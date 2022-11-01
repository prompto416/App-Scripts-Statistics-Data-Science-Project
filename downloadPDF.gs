function myColorFunction() {
  var sheetRead = 'IM_Future'
  var sheetWrite = 'IM_Future'
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // var range = ss.getSheetByName(sheetWrite).getRange(2,6,ss.getLastRow());
  // var cellRange = range.getValues();

  // for(i = 0; i<cellRange.length-1; i++){
  //   ss.getSheetByName("IM_Futures").getRange(i+2,6).setBackground("red");
  //   ss.getSheetByName("IM_Futures").getRange(i+2,6).setFontColor('white');
  // }
  ss.getSheetByName(sheetWrite).getRange(3,letterToColumn('E')).getValues();
  ss.getSheetByName(sheetWrite).getRange(3,letterToColumn('E')).setBackground('green');
  ss.getSheetByName(sheetWrite).getRange(3,letterToColumn('E')).setValue('hello')
  Logger.log(ss.getSheetByName(sheetWrite).getRange(3,letterToColumn('G')).getValue())
  Logger.log(ss.getSheetByName('Sheet2').getRange(2,letterToColumn('C')).getValue())
  
  for (i = 0; i < 14; i++){
    var before = ss.getSheetByName('Sheet2').getRange(2+i,letterToColumn('R')).getValue()
    if(before == 'i'){
      ss.getSheetByName('Sheet2').getRange(2+i,letterToColumn('C')).setBackground('#b6d7a8');
    }
    else if(before == 'd'){
      ss.getSheetByName('Sheet2').getRange(2+i,letterToColumn('C')).setBackground('#ea9999');
    }
    else if (before =='s'){
      Logger.log('stayed the same')
    }
    else {
      Logger.log('invalid input')
    }
  }
  //18 19 20 = Columns for color notation 
  Logger.log(ss.getSheetByName('Sheet2').getRange(1,letterToColumn('S')).getBackground())
  Logger.log(ss.getSheetByName('Sheet2').getRange(2,letterToColumn('S')).getBackground())
  
}



function generatePdf() {
  SpreadsheetApp.flush();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSpreadsheet = SpreadsheetApp.getActive(); // Get active spreadsheet.
  var sheets = sourceSpreadsheet.getSheets(); // Get active sheet.
  var sheetName = sourceSpreadsheet.getActiveSheet().getName();
  sheetName = 'Sheet2'
  var sourceSheet = sourceSpreadsheet.getSheetByName(sheetName);
  var pdfName = sheetName + ".pdf"; // Set the output filename as SheetName.
  var parents = DriveApp.getFileById(sourceSpreadsheet.getId()).getParents(); // Get folder containing spreadsheet to save pdf in.
  if (parents.hasNext()) {
    var folder = parents.next();
  } else {
    folder = DriveApp.getRootFolder();
  }
  var theBlob = createblobpdf(sheetName, pdfName);
  var newFile = folder.createFile(theBlob);
  var email = Session.getActiveUser().getEmail() || 'ppmo444@gmail.com';
  var custemail = sourceSheet.getRange('A1').getValue();
  email = email + "," + custemail;
  // Subject of email message
  const subject = `Your subject Attachement: ${sheetName}`;
  // Email Body can  be HTML too with your image
  const body = "body";
  if (MailApp.getRemainingDailyQuota() > 0)
    GmailApp.sendEmail('ppmo444@gmail.com', subject, body, {
      htmlBody: body,
      attachments: [theBlob]
    });
  // delete pdf if already exists
  var files = folder.getFilesByName(pdfName);
  while (files.hasNext()) {
    files.next().setTrashed(true);
  }
  sourceSpreadsheet.toast("Emailed to " + email, "Success");

}

function createblobpdf(sheetName, pdfName) {
  var sourceSpreadsheet = SpreadsheetApp.getActive();
  var sourceSheet = sourceSpreadsheet.getSheetByName(sheetName);
  var url = 'https://docs.google.com/spreadsheets/d/' + sourceSpreadsheet.getId() + '/export?exportFormat=pdf&format=pdf' // export as pdf / csv / xls / xlsx
    +    '&size=A4' // paper size legal / letter / A4
    +    '&portrait=true' // orientation, false for landscape
    +    '&fitw=true' // fit to page width, false for actual size
    +    '&sheetnames=true&printtitle=false' // hide optional headers and footers
    +    '&pagenum=RIGHT&gridlines=false' // hide page numbers and gridlines
    +    '&fzr=false' // do not repeat row headers (frozen rows) on each page
    +    '&horizontal_alignment=CENTER' //LEFT/CENTER/RIGHT
    +    '&vertical_alignment=TOP' //TOP/MIDDLE/BOTTOM
    +    '&gid=' + sourceSheet.getSheetId(); // the sheet's Id
  var token = ScriptApp.getOAuthToken();
  // request export url
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });
  var theBlob = response.getBlob().setName(pdfName);
  return theBlob;
};


function columnToLetter(column)
{
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function letterToColumn(letter)
{
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++)
  {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}
