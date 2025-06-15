function sortEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Основной лист");

  const ssName = sheet.getName(); 
  const yearMatch = ssName.match(/^(\d{4})_/); //берем год из названия файла

  if (!yearMatch) {
    Logger.log("Не удалось определить год из названия файла.");
    return;
  }

  const year = yearMatch[1]; 

  const startDate = sheet.getRange("G1").getValue();

  if (!startDate) {
    Logger.log("Дата начала не указана.");
    return;
  }

  const formattedDate = Utilities.formatDate(new Date(startDate), Session.getScriptTimeZone(), "yyyy/mm/dd");

  const query = `after:${formattedDate}`;
  const threads = GmailApp.search(query);

  // шаблон темы письма вида "15 школа. 8 класс. Команда "Цветочки""
  const subjectRegex = /^\d+\sшкола\.\s\d+\sкласс\.\s.+/i;

  const labelName = year + " Письма олимпиады";
  const label = GmailApp.getUserLabelByName(labelName) || GmailApp.createLabel(labelName);

  let count = 0;

  threads.forEach(thread => {
    const messages = thread.getMessages();
    for (let msg of messages) {
      const subj = msg.getSubject();
      if (subjectRegex.test(subj)) {
        thread.addLabel(label);
        count++;
        break; 
      }
    }
  });

  const currentCount = sheet.getRange("D4").getValue();

  const newCount = (typeof currentCount === "number" && !isNaN(currentCount)) ? currentCount + count : count;

  sheet.getRange("D4").setValue(newCount);
  Logger.log(`Помечено ${count} новых сообщений ярлыком "${labelName}". Итоговое значение: ${newCount}.`);
}
