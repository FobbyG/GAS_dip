function onOpen() {

  const ui1 = SpreadsheetApp.getUi();

  ui1.createMenu("Олимпиада")
    .addItem("Сформировать ведомость", "generateStatement")
    .addToUi();
    
  const ui2 = SpreadsheetApp.getUi();

  ui2.createMenu("Грамоты")
    .addItem("Сформировать грамоты за 6 класс", "generateDiploma('6 класс')")
    .addItem("Сформировать грамоты за 7 класс", "generateDiploma('7 класс')")
    .addItem("Сформировать грамоты за 8 класс", "generateDiploma('8 класс')")
    .addToUi();
  
}
