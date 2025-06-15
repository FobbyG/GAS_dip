function saveFiles() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Основной лист");
  const counterCell = sheet.getRange("D5");
  let processedCount = 0;
  const maxEmails = 20; //лимит, чтобы не превышлать временное ограничение выполнения скрипта

  const yearMatch = SpreadsheetApp.getActiveSpreadsheet().getName().match(/^(\d{4})_/);

  if (!yearMatch) {
    Logger.log("Не удалось определить год из имени файла");
    return;
  }
  const year = yearMatch[1];

  const sourceLabel = GmailApp.getUserLabelByName(year + " Письма олимпиады");

  if (!sourceLabel) {
    Logger.log("Ярлык не найден: " + year + " Письма олимпиады");
    return;
  }

  const threads = sourceLabel.getThreads();
  if (!threads.length) {
    Logger.log("Нет писем с указанным ярлыком");
    return;
  }

  const doneLabel = GmailApp.getUserLabelByName("Обработано") || GmailApp.createLabel("Обработано");

  const folderIterator = DriveApp.getFoldersByName(`${year} Ответы команд`);
  
  if (!folderIterator.hasNext()) {
    Logger.log("Папка с ответами команд не найдена");
    return;
  }

  const parentFolder = folderIterator.next();

  for (let i = 0; i < threads.length && processedCount < maxEmails; i++) {
    const thread = threads[i];

    if (thread.hasLabel(doneLabel)) continue;

    const messages = thread.getMessages();
    for (let j = 0; j < messages.length; j++) {
      const message = messages[j];

      const subject = message.getSubject();
      const match = subject.match(/(\d+)\s+школа\.\s+(\d+)\s+класс\.\s+(.+)/i);

      if (!match) {
        Logger.log(`Письмо без нужного шаблона темы: ${subject}`);
        continue;
      }

      const schoolNum = match[1];
      const klass = match[2];
      const teamName = match[3].trim();
      const folderName = `${klass}.${teamName}`;

      let teamFolder;
      const existingFolders = parentFolder.getFoldersByName(folderName);
      if (existingFolders.hasNext()) {
        teamFolder = existingFolders.next();
      } else {
        teamFolder = parentFolder.createFolder(folderName);
      }

      const attachments = message.getAttachments();
      attachments.forEach((file, fileIndex) => {
        const extension = file.getName().split('.').pop();
        const newName = `${klass}.${teamName}.${fileIndex + 1}.${extension}`;
        teamFolder.createFile(file.copyBlob()).setName(newName);
      });

      thread.addLabel(doneLabel);
      processedCount++;
      break; 
    }
  }

  const previousCount = counterCell.getValue() || 0;
  counterCell.setValue(previousCount + processedCount);

  SpreadsheetApp.getUi().alert(`Обработано новых писем: ${processedCount}`);
}