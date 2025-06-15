function generateStatement() {
  const thisFile = SpreadsheetApp.getActiveSpreadsheet();
  const thisFileName = thisFile.getName(); 
  const yearMatch = thisFileName.match(/^(\d{4})_/); // Берем год из названия файла
  
  if (!yearMatch) {
    Logger.log("Не удалось определить год из названия файла.");
    return;
  }

  const year = yearMatch[1];
  const formResponsesName = `${year}_Ответы на форму`;

  const files = DriveApp.getFilesByName(formResponsesName);

  if (!files.hasNext()) {
    Logger.log(`Файл с ответами на форму не найден: ${formResponsesName}`);
    return;
  }

  const formFile = files.next();
  const formSpreadsheet = SpreadsheetApp.open(formFile);
  const formSheet = formSpreadsheet.getSheets()[0];
  const data = formSheet.getDataRange().getValues();

  const headers = data[0];
  const teamIndex = headers.indexOf("Название команды");
  const klIndex = headers.indexOf("Класс");

  if (teamIndex === -1 || klIndex === -1) {
    Logger.log("Не найдены столбцы 'Название команды' и 'Класс' в ответах на форму.");
    return;
  }

  const klStatement = {
    "6": [],
    "7": [],
    "8": []
  };

  for (let i = 1; i < data.length; i++) {
    const team = data[i][teamIndex];
    const klass = String(data[i][klIndex]).trim();

    if (klStatement[klass]) {
      klStatement[klass].push([team]); 
    }
  }

  Object.keys(klStatement).forEach(klass => {
    const teams = klStatement[klass];
    let list = thisFile.getSheetByName(klass);

    if (!list) {
      list = thisFile.insertSheet(klass);
    } else {
      const lastRow = list.getLastRow();
        if (lastRow >= 3) {
          list.getRange(3, 1, lastRow - 2, list.getMaxColumns()).clearContent();
        }
    }

    for (let i = 0; i < teams.length; i++) {
      const rowIndex = i + 3; 
      list.getRange(rowIndex, 1).setValue(teams[i][0]); 

      const firstFormula = `=СУММ(B${rowIndex}:K${rowIndex})`;
      list.getRange(rowIndex, 12).setFormula(firstFormula);

      const secondFormula = `=СУММ(M${rowIndex}:V${rowIndex})`;
      list.getRange(rowIndex, 23).setFormula(secondFormula);

      const totalFormula = `=СУММ(L${rowIndex}:W${rowIndex})`;
      list.getRange(rowIndex, 24).setFormula(totalFormula);

    }
  });

  SpreadsheetApp.getUi().alert("Ведомость успешно сформирована.");
}
