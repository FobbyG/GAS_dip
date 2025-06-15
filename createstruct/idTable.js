function idTable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("id");
  
  const lastRow = sheet.getLastRow();
  const folderNames = sheet.getRange("A2:A" + lastRow).getValues(); //получаем массив значений из столбца A листа id, а именно названия папок, где будем искать файлы 
  
  const output = [];
  
  for (let i = 0; i < folderNames.length; i++){

    const folderName = folderNames[i][0];

    if (folderName && folderName.toString().trim() !== "") {
    
      const folderIterator = DriveApp.getFoldersByName(folderName);
      
      if (folderIterator.hasNext()) {
        const folder = folderIterator.next();
        const fileIterator = folder.getFiles();
        
        while (fileIterator.hasNext()){
          const file = fileIterator.next();
          output.push([folderName, file.getName(), file.getId()]);// Добавляем строку: [название папки, название файла, ID файла]
        }

        if (!output.find((row) => row[0] === folderName)) {
          output.push([folderName, "Файлов не найдено", ""]);// В случае пустой папки
        }
      } else {
         output.push([folderName, "Папка не найдена", ""]);
      }
    }
  }

  sheet.getRange(2, 4, sheet.getMaxRows() - 1, 3).clearContent(); 
  
  if (output.length > 0) {
     sheet.getRange(2, 4, output.length, output[0].length).setValues(output);
  } else {
     SpreadsheetApp.getUi().alert("Нет данных для записи");
  }

  UpdateDate();
}

function UpdateDate() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetStruct = spreadsheet.getSheetByName("Создание структуры"); 

  if (!sheetStruct) {
    throw new Error("Лист 'Создание структуры' не найден");
  }
  const sheet = spreadsheet.getSheetByName("Создание структуры");
  const now = new Date();

  sheet.getRange("E7").setValue(now);
}