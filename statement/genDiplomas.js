function generateDiploma(sheetName) {
  const templateSs = SpreadsheetApp.getActiveSpreadsheet();

  const currentFileName = templateSs.getName();
  const yearMatch = currentFileName.match(/^(\d{4})_/);

  if (!yearMatch) {
    SpreadsheetApp.getUi().alert("Не удалось определить год из имени файла.");
    return;
  }
  const year = yearMatch[1];

  const managerFileName = `${year}_управляющая таблица`;

  const files = DriveApp.getFilesByName(managerFileName);

  if (!files.hasNext()) {
    SpreadsheetApp.getUi().alert(`Файл "${managerFileName}" не найден в Google Диске.`);
    return;
  }

  const managerSs = SpreadsheetApp.openById(files.next().getId());

  const idSheet = managerSs.getSheetByName("id");
  if (!idSheet) {
    SpreadsheetApp.getUi().alert("Лист 'id' не найден в управляющей таблице.");
    return;
  }

  const idData = idSheet.getDataRange().getValues();
  const header = idData[0];
  const nameCol = header.findIndex(h => h.toString().toLowerCase().includes("название"));
  const idCol = header.findIndex(h => h.toString().toLowerCase().includes("id"));

  if (nameCol === -1 || idCol === -1) {
    SpreadsheetApp.getUi().alert("В листе 'id' нет нужных столбцов (название и id).");
    return;
  }

  const docRow = idData.find(row => row[nameCol] && row[nameCol].toString().toLowerCase().includes("грамота"));
  if (!docRow) {
    SpreadsheetApp.getUi().alert("Не найден шаблон документа 'грамота' в листе 'id'.");
    return;
  }

  const templateDocId = docRow[idCol];
  const templateDoc = DocumentApp.openById(templateDocId);
  const bodyTemplate = templateDoc.getBody().getText();

  let formFileId = null;
  for (let i = 1; i < idData.length; i++) {
    const name = idData[i][nameCol];
    if (name && name.toString().toLowerCase().includes("ответы на форму")) {
      formFileId = idData[i][idCol];
      break;
    }
  }

  if (!formFileId) {
    SpreadsheetApp.getUi().alert("Не найден ID файла с ответами на форму.");
    return;
  }

  const formSs = SpreadsheetApp.openById(formFileId);
  const formSheet = formSs.getSheets()[0];
  const formData = formSheet.getDataRange().getValues();
  const formHeaders = formData[0];

  const formTeamCol = formHeaders.indexOf("Название команды");
  const participantCols = formHeaders
    .map((h, i) => ({ h, i }))
    .filter(col => col.h.includes("ФИО участника"))
    .map(col => col.i);

  const dataSheet = templateSs.getSheetByName(sheetName);
  if (!dataSheet) {
    SpreadsheetApp.getUi().alert(`Лист '${sheetName}' не найден.`);
    return;
  }
  const data = dataSheet.getDataRange().getValues();

  
  const folderName = `${sheetName} - Грамоты`;
  let folder;
  const folderSearch = DriveApp.getFoldersByName(folderName);
  folder = folderSearch.hasNext() ? folderSearch.next() : DriveApp.createFolder(folderName);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const teamName = row[0]; 
    const statusRaw = (row[24] || "").toString().toLowerCase(); // Столбец Y

    if (!statusRaw.includes("победитель") && !statusRaw.includes("лауреат")) continue;
    const status = statusRaw.includes("победитель") ? "Победитель" : "Лауреат";


    const formRow = formData.find(r => r[formTeamCol] === teamName);
    let participantsText = "";
    if (formRow) {
      const names = participantCols.map(i => formRow[i]).filter(v => v);
      if (names.length > 0) {
        participantsText = "\nУчастники команды:\n" + names.map(n => `• ${n}`).join("\n");
      }
    }

    let finalText = bodyTemplate
      .replace("{{название команды}}", teamName)
      .replace("{{статус}}", status)
      + "\n" + participantsText;


    const newDoc = DocumentApp.create(`${teamName} - ${status}`);
    newDoc.getBody().setText(finalText);
    newDoc.saveAndClose();
    DriveApp.getFileById(newDoc.getId()).moveTo(folder);
  }

  SpreadsheetApp.getUi().alert(`Грамоты созданы для листа '${sheetName}'.`);
}
