function idTable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("id");

  const ssName = ss.getName(); 
  const yearMatch = ssName.match(/^(\d{4})_/); // берем год из названия файла

  if (!yearMatch) {
    Logger.log("Не удалось определить год из названия файла.");
    return;
  }

  const year = yearMatch[1]; 

  const folderNames = [
    year + " Регистрация",
    year + " Проверка",
    year + " Ответы команд",
    year + " Положения"
  ];

  const output = [];

  for (let i = 0; i < folderNames.length; i++) {
    const folderName = folderNames[i];

    if (folderName && folderName.toString().trim() !== "") {
      const folderIterator = DriveApp.getFoldersByName(folderName);

      if (folderIterator.hasNext()) {
        const folder = folderIterator.next();
        const fileIterator = folder.getFiles();
        let hasFiles = false;

        while (fileIterator.hasNext()) {
          hasFiles = true;
          const file = fileIterator.next();
          output.push([folderName, file.getName(), file.getId()]);
        }

        if (!hasFiles) {
          output.push([folderName, "Файлов не найдено", ""]);
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

//запоминаем имя последнего обновления для удобства 
function UpdateDate() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Основной лист");

  if (!sheet) {
    throw new Error("Лист 'Основной лист' не найден");
  }

  const now = new Date();
  sheet.getRange("D1").setValue(now);
}
