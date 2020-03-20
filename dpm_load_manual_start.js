/**
* Attention! Functions use NAMED RANGES at sheet 'metadata !important':
*  - getGeneralData
*  - requestStages
*/

//Create custom menu items for manual run
function onOpen() {
  let customMenuItem = SpreadsheetApp.getUi();
  customMenuItem.createMenu('Селектор')
  .addItem('▶ Создать новый отчёт', 'makeReport')
  .addSeparator()
  .addItem('↻ Обновить текущий отчёт', 'refreshReport')
  .addToUi();
}

//public static void main(String[] args) :-D
function makeReport() {
  const SHEET_DATA = SpreadsheetApp.getActive().getSheetByName('DATA !important'); // Лист для хранения расчетной информации
  const SHEET_METADATA = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('metadata !important'); //Лист для хранения справочной информации
  const MAIN_SHEET_ID = SpreadsheetApp.openById('16iysH71qZygfc1D7jA7-V19tpzAvthrH0hFBQ_IeDHk'); //ID основного документа
  checkEmptyProd(MAIN_SHEET_ID);
}

function checkEmptyProd(SS_MAIN_DATA) {
  let indexProdCol;
  let emptyCells = [];
  let sheet = SS_MAIN_DATA.getSheetByName('Заявки в ДПМ');
  let titleData = sheet.getRange(1, 1, 1, sheet.getDataRange().getLastColumn()).getValues();
  for(let i = 0; i < titleData.length; i++){
    for(let j = 0; j < titleData[0].length; j++){
      if(titleData[i][j] == 'Продукт который будет оформлен клиенту'){
        indexProdCol = j + 1;
      }
    }
  }
  let prodValues = sheet.getRange(1, indexProdCol.toFixed(0), sheet.getLastRow(), 1).getValues();
  for(let i = 0; i < prodValues.length; i++){
    if(prodValues[i][0] == ''){
      emptyCells.push(i + 1);
    }
  }
  if(emptyCells.length > 0){
    SpreadsheetApp.getUi().alert('Отсутствует название продукта в строке: ' + emptyCells);
    return;
  }
}

