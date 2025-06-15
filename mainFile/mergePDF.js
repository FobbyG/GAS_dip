function mergePDF() {
  const yearMatch = SpreadsheetApp.getActiveSpreadsheet().getName().match(/^(\d{4})_/);

  if (!yearMatch) {
    Logger.log("Не удалось определить год из имени файла");
    return;
  }

  const year = yearMatch[1];

  const rootFolder = DriveApp.getFoldersByName(`${year} Ответы команд`);

  const processedLabelName = 'processed';

  let processedLabel = GmailApp.getUserLabelByName(processedLabelName);
  if (!processedLabel) {
    processedLabel = GmailApp.createLabel(processedLabelName);
  }

  const folders = rootFolder.getFolders();

  const props = PropertiesService.getScriptProperties();
  let pdfCounter = Number(props.getProperty('pdfCounter')) || 1;

  const maxFolders = 5; //лимит на кол-во обрабатываемых папок, чтобы не превышать лимит файлов на сервисе 

  let processedFoldersCount = 0;

  while (folders.hasNext() && processedFoldersCount < maxFolders) {
    const folder = folders.next();

    let processedFoldersIds = props.getProperty('processedFoldersIds');
    processedFoldersIds = processedFoldersIds ? processedFoldersIds.split(',') : [];

    if (processedFoldersIds.includes(folder.getId())) {
      continue;
    }

    const files = folder.getFilesByType(MimeType.PDF);
    const pdfFiles = [];
    while (files.hasNext()) {
      pdfFiles.push(files.next());
    }

    if (pdfFiles.length === 0) {
      
      processedFoldersIds.push(folder.getId());
      processedFoldersCount++;
      continue;
    }

    const outputName = `Работы на печать ${pdfCounter}`;

    mergePDFsFromFiles(pdfFiles, outputName);

    processedFoldersIds.push(folder.getId());

    props.setProperty('processedFoldersIds', processedFoldersIds.join(','));
    pdfCounter++;
    props.setProperty('pdfCounter', pdfCounter.toString());

    processedFoldersCount++;
  }

  Logger.log(`Обработано папок: ${processedFoldersCount}. Всего сгенерировано PDF: ${pdfCounter - 1}`);
}

function mergePDFsFromFiles(pdfFiles, outputName) {
  const apiKey = 'YmY4Yzc1OGUtYTJlNi00NDdhLWI5ZTgtNGE5NDU4ODYwNWI3OlRQJjd3U0lDdzZxJnFzIVRFdVp0VHluZ2FlVnZsbGdO';  
  
  const yearMatch = SpreadsheetApp.getActiveSpreadsheet().getName().match(/^(\d{4})_/);

  if (!yearMatch) {
    Logger.log("Не удалось определить год из имени файла");
    return;
  }

  const year = yearMatch[1];

  const saveFolder = DriveApp.getFoldersByName(`${year} Ответы команд`);

  const base64Files = pdfFiles.map(file => {
    const blob = file.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());
    return {
      "FileName": file.getName(),
      "FileContent": base64
    };
  });

  // формируем запрос для PDF4me Merge API
  const requestBody = {
    "ApiKey": apiKey,
    "Files": base64Files
  };

  const url = "https://api.pdf4me.com/v1/pdf/merge";

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(requestBody),
    "muteHttpExceptions": true
  };

  // делаем запрос к API
  const response = UrlFetchApp.fetch(url, options);
  
  if (response.getResponseCode() !== 200) {
    Logger.log('Ошибка API: ' + response.getContentText());
    throw new Error('Ошибка при объединении PDF: ' + response.getContentText());
  }

  // Результат — base64 PDF
  const jsonResponse = JSON.parse(response.getContentText());
  const mergedFileBase64 = jsonResponse.MergedFile;

  const mergedBlob = Utilities.newBlob(Utilities.base64Decode(mergedFileBase64), 'application/pdf', outputName + '.pdf');

  saveFolder.createFile(mergedBlob);

  Logger.log('Объединённый PDF сохранён: ' + outputName + '.pdf');
}
