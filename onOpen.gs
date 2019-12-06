function onOpen() {    
  var ui = SpreadsheetApp.getUi();   
  ui.createMenu("Меню тренера")
  .addItem("Копирование данных из Google Calendar", "main")
  .addSeparator()
  .addItem("Подведение итогов месяца", "sumData") 
  .addToUi();  
}
