function sumData() {
  //подключение к таблице "Работа"
  var sheet = SpreadsheetApp.openById("");
  
  //подключение к нужной странице
  var list = sheet.getSheetByName("Данные из календаря");       
  var analiticList = sheet.getSheetByName("AnaliticList"); 
   
  //поиск нижней желтой строки (<----- Значения за месяц)
   var bottomPoint = list.getLastRow();     
    while(list.getRange(bottomPoint, 6).getBackground() !== "#ffff00"){
      bottomPoint--;    
    }
  
   //поиск верхней желтой строки (<----- Значения за месяц)
    var topPoint = bottomPoint;      
      do {
      topPoint--;    
      }while(list.getRange(topPoint, 6).getBackground() !== "#ffff00")
      
  //формирование имени ячеек диапазона, формулы и дальнейшая запись в ячейку
      for (var i=0; i<6; i++){
        var sumRangeTop = list.getRange(topPoint + 1, 6+i).getA1Notation();
        var sumRangeBottom = list.getRange(bottomPoint - 1, 6+i).getA1Notation();  
  
        var formulaSum = "=SUM(" + sumRangeTop + ":" + sumRangeBottom + ")"; //"=SUM(F3:F121)"   
        list.getRange(bottomPoint, 6+i).setFormula(formulaSum);  
      }  
    
  questionForTransfer(sheet, list, bottomPoint, analiticList);  
}
