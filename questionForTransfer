function questionForTransfer(sheet, list, bottomPoint, analiticList){
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Сделать копирование \"итогов месяца\" в AnaliticList?", ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    transferDataToAnaliticList(sheet, list, bottomPoint, analiticList);
    Logger.log('The user clicked "Yes."');
  } else {   
    Logger.log('The user clicked "No" or the close button in the dialog\'s title bar.');
  }  
}
