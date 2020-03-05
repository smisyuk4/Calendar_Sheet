 /*
            цвета в таблице:
  #ffffff - белый
  #ff0000 - красный (долг)
  #00ffff - голубой (списана тренировка за неявку) 
  #ea9999 - розовый (оплата)
  #e6b8af - розовый (оплата #2)
  #e06666 - розовый (оплата #3)
  #fce5cd - оранжевый (оплата #4)  
  #00ff00 - зеленый (тренировка не отмечена на абонементе)  
  */ 

function checkBase() {  
  //подключение к таблице тренера "Работа"  
  var sheet = SpreadsheetApp.openById("");  
  //подключение к нужной странице
  var list = sheet.getSheetByName("Данные из календаря");  
  
  //поиск ID для загрузки данных    
  var admID = list.getRange(1, 7).getValues();  
  var admIDSheet = admID[0][0];
  //Logger.log(admIDSheet)
  
  if (admIDSheet !== ""){
    //подключение к таблице администратора 
    var admSheet = SpreadsheetApp.openById(admIDSheet);
  }
  
  //название страницы в админской таблице, ID чата с ботом в Telegram
  var arrayCoachSheet = [
    //['Инна', ''],
    //['СБ', ''],
    //['Аня М', ''],
    //['Вадим', ''],
    //['СМР', ''],
    //['Валерий', ''],
    //['Саша', ''],
    ['СМ', '']
    //['Ксюша', ''],
    //['Женя', ''],    
    //['Таня', ''],
    //['Гаяне', '']
  ];
  
 //Logger.log(arrayCoachSheet);  
 //Logger.log('имя тренера ' + arrayCoachSheet[0][0]);
 //Logger.log('ID чата с ним ' + arrayCoachSheet[0][1]);
  
  //цикл повторов длиной в массив тренеров
  //взять первого тренера и если есть ID,то запустить проверку долгов
  for (var a=0; a<arrayCoachSheet.length; a++){
    if (arrayCoachSheet[a][1] != 'empty'){
      //подключение к странице тренера   
      var admList = admSheet.getSheetByName(arrayCoachSheet[a][0]); 
 
      //найти последнюю строку в таблице, а потом сформировать диапазон ячеек для загрузки 
      var lastRow = admList.getLastRow();
      var rurrentRow = lastRow-1;
      //Logger.log(rurrentRow);
      var curentRange = 'H3:BQ' + rurrentRow;
      //Logger.log(curentRange);
  
      //взять массив значений цветов ячеек
      var arrayBackgrounds = admList.getRange(curentRange).getBackgrounds();  
      //Logger.log(arrayBackgrounds);
      //Logger.log(arrayBackgrounds.length);
  
      var rowDate = 1;
      var colFirstDate = 8; 
      var rowFirstName = 3;  
      var colName = 6;
      
      var arrayDebtor = [];
  
      //начало перебора массива по именам клиентов
      for (var j=0; j<arrayBackgrounds.length; j++){    
        var name = admList.getRange(rowFirstName+j, colName).getValue();
        var arrayName = name.split(" ");
        var newNameClient = arrayName[0] + ' ' + arrayName[1];  
        var countDebt = 0;  
        
        //начало перебора ячеек по датам
        for (var i=0; i<62; i++){    
          var x = rowFirstName-3;
          if (arrayBackgrounds[x+j][i] == '#ff0000'){
            countDebt++;
          } else if((arrayBackgrounds[x+j][i] == '#ea9999')||
                   (arrayBackgrounds[x+j][i] == '#e6b8af')||
                   (arrayBackgrounds[x+j][i] == '#fce5cd')||
                   (arrayBackgrounds[x+j][i] == '#e06666')){
           countDebt = 0;
           }
        }//конец перебора ячеек по датам
        //Logger.log(newNameClient + ' ' + countDebt);
  
        if (countDebt !== 0){
          arrayDebtor.push(newNameClient + ' (долг: ' + countDebt + ' шт)');
        }
 
      }//конец перебора массива по именам клиентов 
  
      var idChatWBot = arrayCoachSheet[a][1];
      //Logger.log(arrayDebtor);
  
      if(arrayDebtor.length != 0){
        sendSMS(arrayDebtor, idChatWBot);
      }
    }//конец условия проверки наличия ID
  }//конец цикла тренеров
}
      
function sendSMS(arrayDebtor, idChatWBot){
  var date = Utilities.formatDate(new Date(), 'GMT', 'dd.MM.yyyy');
  //Logger.log(date);
  var dotID = ''; //AdminRobot  
  
  var text = encodeURIComponent(date + ' Список должников: ' + arrayDebtor);
  var createLink = "https://api.telegram.org/bot" + dotID + "/sendMessage?chat_id=" + idChatWBot + "&text=" + text;  
  //Logger.log(createLink);
  var loadLink = UrlFetchApp.fetch(createLink);
}
       
       
       
       
       
       
       
       
       
       
       
       
       
       
