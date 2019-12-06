function main() {
  //подключение к таблице "Работа"
  var sheet = SpreadsheetApp.openById("");
  
  //подключение к нужной странице
  var list = sheet.getSheetByName("Данные из календаря");   
  
   //подключение к "CF Trainer - google calendar"     
  var calendar = CalendarApp.getCalendarById("");
  
  
   //поиск загруженного диапазона в таблицу  
  var firstRowRange = list.getLastRow();  
  pullEvents(calendar, list);  
  var lastRowRange = list.getLastRow();
  var countRow = lastRowRange - firstRowRange;   
  upGradeData(list, countRow, firstRowRange);
}
  
 function upGradeData(list, countRow, firstRowRange){
   //изменение отображения даты, удаление "EET"
  for (var i=0; i<countRow; i++){    
    var getDateCell = list.getRange(firstRowRange+1+i, 1).getValue();   
    var arrayDate = getDateCell.split(" ");   
    var newArrayDate = arrayDate[0] + " " + arrayDate[1]+ " " + arrayDate[2]+ " " + arrayDate[3]+ " " + arrayDate[4];   
    list.getRange(firstRowRange+1+i, 1).setValue(newArrayDate);    
  }   
 
  //поиск ключевых слов в диапазоне ячеек ввод суммы совпадений в соответствующие ячейки
  var textGroup = "группа";
  var textBpt = "бпт";
  var textBptPositive = "бпт+";
  var textNotVisit0 = "пришёл";
  var textNotVisit1 = "пришел";
  var textNotVisit2 = "пришла";
  var textDocument = "справка";           
   
//отлавливание ошибок из-за отсутствия событий в календарном дне
try{   
  var search = list.getRange(firstRowRange+1, 2, countRow, 4).createTextFinder(textGroup).findAll();
  var searchResult = search.length;
  list.getRange(firstRowRange+1, 7).setValue(searchResult);  
  
  var search2 = list.getRange(firstRowRange+1, 2, countRow, 4).createTextFinder(textBpt).findAll();
  var searchResult2 = search2.length;
  list.getRange(firstRowRange+1, 8).setValue(searchResult2);  
  
  var search3 = list.getRange(firstRowRange+1, 2, countRow, 4).createTextFinder(textBptPositive).findAll();
  var searchResult3 = search3.length;
  list.getRange(firstRowRange+1, 9).setValue(searchResult3);  
  
  var search40 = list.getRange(firstRowRange+1, 2, countRow, 4).createTextFinder(textNotVisit0).findAll();
  var search41 = list.getRange(firstRowRange+1, 2, countRow, 4).createTextFinder(textNotVisit1).findAll();
  var search42 = list.getRange(firstRowRange+1, 2, countRow, 4).createTextFinder(textNotVisit2).findAll();  
  var searchResult4 = search40.length + search41.length + search42.length;
  list.getRange(firstRowRange+1, 10).setValue(searchResult4);  
  
  var search5 = list.getRange(firstRowRange+1, 2, countRow, 4).createTextFinder(textDocument).findAll();
  var searchResult5 = search5.length;
  list.getRange(firstRowRange+1, 11).setValue(searchResult5);   
    
  
  //подсчет тренировок персональных клиентов    
  var countColumn = 4;  
  var notEmptyCell=0;   
  var startRow = firstRowRange+1;
  var startCol = 2;
  
  for (var i=0; i<countRow; i++){
    for (var j=0; j<countColumn; j++){
      if (list.getRange(startRow+i, startCol+j).getValue() !== ""){
        notEmptyCell++;
      }      
    }
  }   
  
  var searchResult6 = notEmptyCell - searchResult - searchResult2 - searchResult5;
  list.getRange(firstRowRange+1, 6).setValue(searchResult6);    
    
  //изменение цвета строки с значениями за день
   list.getRange(firstRowRange+1, 6, 1, 6).setBackground("#ceebd0"); 
   list.getRange(firstRowRange+1, 12).setValue("<----- Значения за день");  
  
  //проверка с администрацией
  checkDataFromAdmSheet (list, searchResult6, firstRowRange);
  
  } //конец try
  catch(e){   
  }  
}
  
function pullEvents(calendar, list){  
  //проверка в каком режиме выполнять загрузку событий из календаря
  var modeLoadingEvents = list.getRange(1, 2).isChecked(); 
  
  if (modeLoadingEvents == true){
    var oneDay = new Date();       
  }
  else{
    //поиск даты для загрузки данных   
    var oneDay = list.getRange(1, 5).getValue();     
  }  
  
  //отлавливание ошибок из-за отсутствия событий в календарном дне
try{   
  var events = calendar.getEventsForDay(oneDay);       
  
  //поиск последней строки в таблице
  var lastRow = list.getLastRow();          
  
  //выяснение дня недели, месяца, года 
  var todayDate = oneDay.getDate();   
  var numMonth = oneDay.getMonth() + 1; 
  var fullYear = oneDay.getFullYear(); 
  
  /*
  0 - Воскресенье, 1 - Понедельник, ..., 6 - Суббота    
  0 - Январь, 1 - Февраль, ..., 11 - Декабрь
  */
   
  //учитывает высокосность года  
  var lastDayMonth = 28 + ((numMonth + Math.floor(numMonth / 8)) % 2) + 2 % numMonth + 
    Math.floor((1 + (1 - (fullYear % 4 + 2) % (fullYear % 4 + 1)) * 
      ((fullYear % 100 + 2) % (fullYear % 100 + 1)) + (1 - (fullYear % 400 + 2) % (fullYear % 400 + 1))) / numMonth) + 
        Math.floor(1/numMonth) - Math.floor(((1 - (fullYear % 4 + 2) % (fullYear % 4 + 1)) * 
          ((fullYear % 100 + 2) % (fullYear % 100 + 1)) + (1 - (fullYear % 400 + 2) % (fullYear % 400 + 1)))/numMonth);    
    
  if (todayDate == lastDayMonth){
    list.getRange(lastRow+2, 6, 1, 6).setBackground("#ffff00");  
    list.getRange(lastRow+2, 12).setValue("<----- Значения за месяц");
  }     
              
    //загрузка данных из календаря в таблицу
    list.getRange(lastRow+1, 1).setValue(events[0].getStartTime().toLocaleString("ru"));
    var oldTime = list.getRange(lastRow+1, 1).getValue();  
    
    //изменение цвета первой строки новой даты
    list.getRange(lastRow+1, 1, 1, 5).setBackground("#ffd2bd");      
    
    var startRow = lastRow+1;  
    var startColumn = 2;
    var j=lastRow+2;  
    
    for (var i=0; i<events.length; i++){      
      var newTime = events[i].getStartTime().toLocaleString("ru");
      var man = events[i].getTitle();       
      
      if (newTime == oldTime){
        //вправо        
        list.getRange(startRow, startColumn).setValue(man);
        startColumn++;
      }
      else{
        //вниз            
        list.getRange(j, 1).setValue(newTime) 
        list.getRange(j, 2).setValue(man)                
        j++;
        startRow++;
        startColumn = 3;
        oldTime = newTime;
      } 
    }        
    
     } //конец try
catch (e){
    Logger.log("В календаре нет событий для загрузки");
    return;
  }
}

function checkDataFromAdmSheet (list, searchResult6, firstRowRange){  
  //поиск ID для загрузки данных    
  var admID = list.getRange(1, 7).getValue(); 
  var admSheet;
  
  if (admID !== ""){
    //подключение к таблице администратора 
    var admSheet = SpreadsheetApp.openById(admID);
    
     //подключение к странице тренера
    var admList = admSheet.getSheetByName("СМ");     
    var rowDate = 1;
    var colFirstDate = 8; 
    /*
    col 8-9 = 1 число
    col 10-11 = 2 число
    col 12-13 = 3 число
    ....
    */    
    var oneDay = list.getRange(1, 5).getValue();     
    var todayDate = oneDay.getDate();
    
    var currentColDate = colFirstDate + (todayDate*2)-2;   
    var admDate = admList.getRange(rowDate, currentColDate).getValue();
    var rowCountTraining = admList.getLastRow();
    var admCountTraining = admList.getRange(rowCountTraining, currentColDate).getValue();   
      
    if(searchResult6 !== admCountTraining){      
       list.getRange(firstRowRange+1, 6).setBackground("#ff8c8c");      
    }        
  }
  else{
    //комментарий под ячейкой "персоналки"
    list.getRange(firstRowRange+2, 6).setValue("проверки не было");
  }  
}