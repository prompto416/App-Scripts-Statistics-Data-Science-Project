function investorFund_Seperator() {
  var sheetName = "Summary";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
 
 for (i = 0; i < 10 ; i++){
   var temp = ss.getSheetByName(sheetName).getRange(2+i,letterToColumn('B')).getValue()
  
   if ( ( temp.substring(0,3) == "MR." ) || ( temp.substring(0,3) == "นาย" ) || ( temp.substring(0,4) == "น.ส." )){
     Logger.log(temp+" = Investor")
   }
   else{
     Logger.log(temp +" = Fund")
   }
   
 }
  
}

function buySell_determinator() {
  var sheetName = "Summary";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
 
 for (i = 0; i < 10 ; i++){
  var temp = ss.getSheetByName(sheetName).getRange(2+i,letterToColumn('C')).getBackground();
  var tempName = ss.getSheetByName(sheetName).getRange(2+i,letterToColumn('B')).getBackground();
  if (temp == "#ff4500"){
    Logger.log(tempName+' = Sell')
  }
  else {
    Logger.log(tempName+' = Buy')
  }
  // #ff4500 Sell
  // #00ff00 Buy

  
   
   
 }
  
}


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
