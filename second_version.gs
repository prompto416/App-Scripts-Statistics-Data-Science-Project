function fundInvestor_Seperator(wantedOutput) {
  //CAUTION 1: THE SHEETS INDEX IS CURRENTLY FLEXIBLE 'FUND AND INVESTOR' ARE TO BE REPLACED ONCE THE WORK IS FINISHED
  var sheetWrite = 'FundAUTO';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // ss.getSheetByName(sheetWrite);
  var totalSheets = ss.getSheets().length;
  var codedSheets = 5;
  //NUMBER OF SHEETS THAT IS NOT USED AS A SAMPLE !!!!!!!FLEXIBLE ONCE COMPLETED

  // var example1 = SpreadsheetApp.getActiveSpreadsheet().getSheets()[totalSheets-5];
  // Logger.log(example1.getName());
var sampleSheets = totalSheets-4;

var fund = [];
var investor = [];
for (var i = 2; i < 212; i++){
  var fundOrInvestor_names = ss.getSheets()[totalSheets-5].getRange('B'+String(i)).getValue();
  var fundOrInvestor_names = fundOrInvestor_names.trim()
  if (fundOrInvestor_names.includes("นาย") || fundOrInvestor_names.includes("น.ส.") || fundOrInvestor_names.includes("ม.ล.")  || fundOrInvestor_names.includes("นาง") || fundOrInvestor_names.includes("MR")){
    if (!investor.includes(fundOrInvestor_names)){
      investor.push(fundOrInvestor_names);
    }
    
  }
  else {
    if(!fund.includes(fundOrInvestor_names)){
      fund.push(fundOrInvestor_names);
    }
    
  }
  
  //Logger.log(fundOrInvestor_names);
}



if (wantedOutput == 0){
  return fund;
}
else {
  return investor;
}
  
  
}

function stats(interestedMonth,interested,writeRow = 0){
  
  writeRow +=2;

  //ซื้้อถูก เขียวราคาขึ้น=true
  // ขายเเพง เเดงราคาลด = true
 var sell = "#ff4500";
 //sell เเดง buy เขียว
 var buy = '#00ff00';
 var sheetWrite = 'FundAUTO';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // ss.getSheetByName(sheetWrite);
  var totalSheets = ss.getSheets().length;
  var codedSheets = 5;


  var interestedCount = 0;
  var profitCount = 0;
  var lossCount = 0;

  var profitList = [];
  var lossList = [];

  for (var i = 2; i < 212 ;i++){
    var fundOrInvestor_names = ss.getSheets()[totalSheets-5].getRange('B'+String(i)).getValue();
    var buyOrSell = ss.getSheets()[totalSheets-5].getRange('C'+String(i)).getBackgrounds();
    var priceAction = ss.getSheets()[totalSheets-5].getRange('D'+String(i)).getValue();
    
    //options are 1, 3, 6 and 12 month/(s).
    if (interestedMonth == 1){
      var priceMonth = ss.getSheets()[totalSheets-5].getRange('E'+String(i)).getValue(); 
      var priceColumn = 3;
     
      
    }
    else if(interestedMonth == 3){
      var priceMonth = ss.getSheets()[totalSheets-5].getRange('F'+String(i)).getValue(); 
      var priceColumn = 8;
    
    }
    else if(interestedMonth == 6){
      var priceMonth = ss.getSheets()[totalSheets-5].getRange('G'+String(i)).getValue(); 
      var priceColumn = 13;
      
    }
    else if(interestedMonth = 12){
      var priceMonth = ss.getSheets()[totalSheets-5].getRange('H'+String(i)).getValue(); 
      var priceColumn = 18;
      
    }
    else {
      Logger.log('Invalid Price Change Period!');
      return 0;
    }
    if ((fundOrInvestor_names == interested) ){
      interestedCount += 1;

      if (buyOrSell == buy){
        
        if (priceAction < priceMonth){
          profitCount += 1;
          profitList.push((priceMonth-priceAction).toFixed(2))
          

        }
        else if (priceAction > priceMonth){
          lossCount += 1
          lossList.push((priceAction-priceMonth).toFixed(2))
        }
        
      }
      else if (buyOrSell == sell){
         if (priceAction > priceMonth){
          profitCount += 1;
          profitList.push((priceAction-priceMonth).toFixed(2))
        }
        else if(priceAction < priceMonth){
          lossCount += 1;
          lossList.push((priceMonth-priceAction).toFixed(2))
        }

      }
      
    }
  }
  
  fundOrInvestor_names = ss.getSheetByName(sheetWrite).getRange('A'+String(writeRow)).setValue(interested);
  fundOrInvestor_names = ss.getSheetByName(sheetWrite).getRange('B'+String(writeRow)).setValue(interestedCount);
  fundOrInvestor_names = ss.getSheetByName(sheetWrite).getRange(writeRow,priceColumn).setValue(profitCount/interestedCount);

  
 
  
  
  var profitSum = 0;
  for (var i= 0; i < profitList.length;i++){
    profitSum += parseFloat(profitList[i]);
  }
  var lossSum = 0;
  for (var i= 0; i < lossList.length;i++){
    lossSum += parseFloat(lossList[i]);
  }


  
  ss.getSheetByName(sheetWrite).getRange(writeRow,priceColumn+1).setValue(Math.max.apply(null,profitList));
  ss.getSheetByName(sheetWrite).getRange(writeRow,priceColumn+2).setValue(profitSum/profitList.length);
  ss.getSheetByName(sheetWrite).getRange(writeRow,priceColumn+3).setValue(Math.max.apply(null,lossList));
  ss.getSheetByName(sheetWrite).getRange(writeRow,priceColumn+4).setValue(lossSum/lossList.length);
  
  
  

  
  


}

//11:57:33
function noteBackgroundColor(){
  var sheetWrite = 'FundAUTO';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var totalSheets = ss.getSheets().length;
  var sell = "#ff4500";
  var buy = '#00ff00';
  //sell เเดง buy เขียว
  for (var i = 2 ; i < 212 ; i++){

    var buyOrSell = ss.getSheetByName("Summary").getRange('C'+String(i)).getBackground();
    
    if (buyOrSell == buy){
      ss.getSheetByName("Summary").getRange('Z'+String(i)).setValue("buy");
    }
    else if(buyOrSell == sell){
      ss.getSheetByName("Summary").getRange('Z'+String(i)).setValue("sell");
    }
    else {
      Logger.log('ERROR COLOR PRICE ACTION!')
    }
    Logger.log(buyOrSell);
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


