function onOpen() {

  const ui = SpreadsheetApp.getUi();

  ui.createMenu("Олимпиада")
    .addItem("Обновить id шаблонов", "idTable")
    .addItem("Создать структуру", "createStruct")
    .addToUi();
    
}
