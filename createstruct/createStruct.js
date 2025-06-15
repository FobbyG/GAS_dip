function createStruct() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const idSheet = ss.getSheetByName("id");
  const sheetStruct = ss.getSheetByName("Создание структуры");

  if (!idSheet) {
    SpreadsheetApp.getUi().alert("Лист 'id' не найден.");
    return;
  }

  if (!sheetStruct) {
    SpreadsheetApp.getUi().alert("Лист 'Создание структуры' не найден.");
    return;
  }

  
  const year = sheetStruct.getRange("E3").getValue().toString().trim();
  if (!year) {
    throw new Error("Год не указан в ячейке E3 листа 'Создание структуры'");
  }
  
  const rootFolderName = sheetStruct.getRange("E4").getValue().toString().trim();
  if (!rootFolderName) {
    throw new Error("Корневая папка не указана в ячейке E4 листа 'Создание структуры'");
  }
  
  
  const rootFolderIterator = DriveApp.getFoldersByName(rootFolderName);
  if (!rootFolderIterator.hasNext()) {
    throw new Error("Папка '" + rootFolderName + "' не найдена");
  }
  const root = rootFolderIterator.next();
  

  const yearFolderIterator = root.getFoldersByName(year);
  if (yearFolderIterator.hasNext()) {
    throw new Error("Папка за " + year + " год уже существует");
  }
  
  const yearFolder = root.createFolder(year);
  //PropertiesService.getScriptProperties().setProperty("yearFolderId", yearFolder.getId()); пока не нужен 

  const regFolder = yearFolder.createFolder(year + " Регистрация");
  const checkFolder = yearFolder.createFolder(year +" Проверка");
  const answFolder = yearFolder.createFolder(year +" Ответы команд");
  const infFolder = yearFolder.createFolder(year +" Положения");
  
  const lastRow = idSheet.getLastRow();
  const data = idSheet.getRange(2, 5, lastRow - 1, 2).getValues(); 
  // выбираем диапазон - (начальная строка, начальный столбец, кол-во строк, кол-во столбцов)

  data.forEach(function(row) {
    const fileName = row[0];
    const fileId   = row[1];
    
    if (!fileName || !fileId) return; // пропускаем пустые записи
    
    const newName = fileName.toString().replace(/шаблон/gi, year);
    
    // Если имя файла содержит "форма" – копируем форму и создаём таблицу ответов,
    if (fileName.toString().toLowerCase().indexOf("форма") !== -1) {
      const copiedFormFile = DriveApp.getFileById(fileId).makeCopy(newName, regFolder);
      Utilities.sleep(1000); 
      
      const form = FormApp.openById(copiedFormFile.getId());
      createFormSubmitTrigger(copiedFormFile.getId());
      
      const responseSheet = SpreadsheetApp.create(year + " Ответы на форму регистраций");
      form.setDestination(FormApp.DestinationType.SPREADSHEET, responseSheet.getId());

      const sheets = responseSheet.getSheets();
      const secondSheet = sheets.find(s => s.getName() === "Лист1");
      if (secondSheet) {
        secondSheet.setName("Email-Команды");
        secondSheet.appendRow(["Email", "Название команды", "Подтверждение"]);
      }

      const responseFile = DriveApp.getFileById(responseSheet.getId());
      regFolder.addFile(responseFile);
      DriveApp.getRootFolder().removeFile(responseFile);
    }

    else if (
      fileName.toString().toLowerCase().indexOf("ведомость") !== -1 ||
      fileName.toString().toLowerCase().indexOf("грамот") !== -1
    ) {
      DriveApp.getFileById(fileId).makeCopy(newName, checkFolder);
    }
    else if (
      fileName.toString().toLowerCase().indexOf("управляющая таблица") !== -1 
    ){
      DriveApp.getFileById(fileId).makeCopy(newName, yearFolder);
    }
    // Остальные файлы копируем в папку регистрации
    else {
      DriveApp.getFileById(fileId).makeCopy(newName, regFolder);
    }
  });
  
  return "Структура для " + year + " создана, шаблоны скопированы.";
}

//собираем информацию для таблицы Email-Команды
function updateEmailListFromForm(e) { 
  const email = e.values[1];       
  const teamName = e.values[7];     

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Email-Команды");
  sheet.appendRow([email, teamName]);
}

//триггер, при котором, каждый раз при отправки формы, заполняется таблица Email-Команды 
function createFormSubmitTrigger(formId) {

  const form = FormApp.openById(formId); 
  ScriptApp.newTrigger("updateEmailListFromForm")
    .forForm(form)
    .onFormSubmit()
    .create();
}



