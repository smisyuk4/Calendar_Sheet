function transferDataToAnaliticList(sheet, list, bottomPoint, analiticList) {  
  //поиск даты в list
  var date;   
  var i=0;
  do{
    i++;
    date = list.getRange(bottomPoint-i, 1).getValue(); //если она пустая, то взять выше на одну строку
  }while(date == "" || date == "undefined")
    
  Logger.log(date);
  var arrayDate = date.split(" ");   
  var newArrayDate = arrayDate[0]+ " " +arrayDate[1]+ " " + arrayDate[2]+ " " + arrayDate[3]; //30 ноября 2019 г.
   
  //поиск ячейки в analiticList и запись формулы  
  var lastRowAnaliticList = analiticList.getLastRow();  
  analiticList.getRange(lastRowAnaliticList+1, 1).setValue(newArrayDate);  
  
  for (var i=0; i<6; i++){
  //поиск ячейки в list и её названия, сделать формулу
  var countClients = list.getRange(bottomPoint, 6+i).getA1Notation();  
  var formulaTransfer = "='Данные из календаря'!" + countClients; //='Данные из календаря'!F137   
  analiticList.getRange(lastRowAnaliticList+1, 2+i).setValue(formulaTransfer);
  }
  
}
