// Custom menu item
function onOpen() {
	var customMenuItem = SpreadsheetApp.getUi();
  	customMenuItem.createMenu('Селектор')
  	.addItem('▶ Создать новый отчёт', 'makeReport')
  	.addSeparator()
  	.addItem('↻ Обновить текущий отчёт', 'refreshReport')
  	.addToUi();
}