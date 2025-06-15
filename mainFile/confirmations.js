function sendConfirmation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const idSheet = ss.getSheetByName("id");

  if (!idSheet) {
    SpreadsheetApp.getUi().alert("Лист 'id' не найден.");
    return;
  }

  const yearMatch = ssName.match(/^(\d{4})_/); // берем год из названия файла

  if (!yearMatch) {
    Logger.log("Не удалось определить год из названия файла.");
    return;
  }

  const year = yearMatch[1]; 

  const idData = idSheet.getRange(2, 5, idSheet.getLastRow() - 1, 2).getValues(); 
  const regRow = idData.find(row => row[0] === year + " Ответы на форму регистраций");

  if (!regRow || !regRow[1]) {
    SpreadsheetApp.getUi().alert("ID файла 'Ответы на форму регистраций' не найден на листе 'id'.");
    return;
  }

  const regSs = SpreadsheetApp.openById(regRow[1].toString().trim());
  const emailSheet = regSs.getSheetByName("Email-Команды");
  const responseSheet = regSs.getSheetByName("Ответы на форму (1)");

  if (!emailSheet || !responseSheet) {
    SpreadsheetApp.getUi().alert("Не найдены листы 'Email-Команды' или 'Ответы на форму (1)' в файле.");
    return;
  }

  const textSheet = ss.getSheetByName("Шаблон письма-подтверждения");

  if (!textSheet) {
    SpreadsheetApp.getUi().alert('Лист "Шаблон письма-подтверждения" не найден.');
    return;
  }

  const subject = textSheet.getRange("B1").getValue();
  const bodyTemplate = textSheet.getRange("B2").getValue();

  const emailData = emailSheet.getDataRange().getValues();
  const emailHeaders = emailData[0];

  const emailIndex = emailHeaders.indexOf("Email");
  const teamIndex = emailHeaders.indexOf("Название команды");
  let sentIndex = emailHeaders.indexOf("Подтверждение");

  if (sentIndex === -1) {
    emailSheet.getRange(1, emailHeaders.length + 1).setValue("Подтверждение");
    sentIndex = emailHeaders.length;
  }

  const formData = responseSheet.getDataRange().getValues();
  const formHeaders = formData[0];

  const getFormColIndex = name => formHeaders.indexOf(name);

  let sentCount = 0;
  let quotaLeft = MailApp.getRemainingDailyQuota();

  if (quotaLeft <= 0) {
    SpreadsheetApp.getUi().alert("Квота на отправку писем на сегодня исчерпана.");
    return;
  }

  for (let i = 1; i < emailData.length; i++) {
    if (quotaLeft <= 0) break;

    const row = emailData[i];
    const email = row[emailIndex];
    const team = row[teamIndex];
    const alreadySent = row[sentIndex];

    if (!email || !team || alreadySent === true) continue;

    const formRow = formData.find(r => r[getFormColIndex("Название команды")] === team);
    if (!formRow) continue;

    const school = formRow[getFormColIndex("Школа")] || "";
    const grade = formRow[getFormColIndex("Класс")] || "";
    const leaderName = formRow[getFormColIndex("Имя руководителя")] || "";
    const leaderMiddleName = formRow[getFormColIndex("Отчество руководителя")] || "";

    const participants = formHeaders
      .map((h, idx) => ({ h, idx }))
      .filter(x => x.h.includes("ФИО участника"))
      .map(x => formRow[x.idx])
      .filter(v => v && v.toString().trim() !== "")
      .join(", ");

    const htmlBody = bodyTemplate
      .replace("{{Имя}}", leaderName)
      .replace("{{Отчество}}", leaderMiddleName)
      .replace("{{Школа}}", school)
      .replace("{{Класс}}", grade)
      .replace("{{Команда}}", team)
      .replace("{{Участники}}", participants);

    // заголовки, чтобы письмо не упало в спам
    GmailApp.sendEmail(email, subject, "", {
      htmlBody: htmlBody,
      name: "Оргкомитет олимпиады",
      replyTo: "omsk.olympiad@gmail.com", 
    });

    emailSheet.getRange(i + 1, sentIndex + 1).setValue(true);
    sentCount++;
    quotaLeft--;
    Utilities.sleep(500); // пауза
  }

  SpreadsheetApp.getUi().alert(`Писем отправлено: ${sentCount}`);
}

//функция отправки тестового письма 
function sendTestEmailCon() {
  const userEmail = Session.getActiveUser().getEmail();

  if (!userEmail) {
    SpreadsheetApp.getUi().alert("Не удалось определить ваш email для теста.");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const idSheet = ss.getSheetByName("id");
  if (!idSheet) {
    SpreadsheetApp.getUi().alert("Лист 'id' не найден.");
    return;
  }

  const yearMatch = ssName.match(/^(\d{4})_/); // берем год из названия файла

  if (!yearMatch) {
    Logger.log("Не удалось определить год из названия файла.");
    return;
  }

  const year = yearMatch[1]; 

  const idData = idSheet.getRange(2, 5, idSheet.getLastRow() - 1, 2).getValues(); // E:F
  const regRow = idData.find(row => row[0] === year + "Ответы на форму регистраций");
  
  if (!regRow || !regRow[1]) {
    SpreadsheetApp.getUi().alert("ID файла 'Ответы на форму регистраций' не найден в листе 'id'.");
    return;
  }

  const regSs = SpreadsheetApp.openById(regRow[1].toString().trim());
  const emailSheet = regSs.getSheetByName("Email-Команды");
  const responseSheet = regSs.getSheetByName("Ответы на форму (1)");

  if (!emailSheet || !responseSheet) {
    SpreadsheetApp.getUi().alert('Не найдены листы "Email-Команды" или "Ответы на форму (1)"');
    return;
  }

  const textSheet = ss.getSheetByName("Шаблон письма-подтверждения");
  if (!textSheet) {
    SpreadsheetApp.getUi().alert('Лист "Шаблон письма-подтверждения" не найден.');
    return;
  }

  const subject = textSheet.getRange("B1").getValue();
  const bodyTemplate = textSheet.getRange("B2").getValue();

  const emailData = emailSheet.getDataRange().getValues();
  const formData = responseSheet.getDataRange().getValues();
  const formHeaders = formData[0];

  if (emailData.length < 2) {
    SpreadsheetApp.getUi().alert("Нет данных в 'Email-Команды' для теста.");
    return;
  }

  const testRow = emailData[1]; 
  const emailHeaders = emailData[0];
  const teamIndex = emailHeaders.indexOf("Название команды");

  const teamName = testRow[teamIndex];
  if (!teamName) {
    SpreadsheetApp.getUi().alert("В тестовой строке нет названия команды.");
    return;
  }

  const getFormColIndex = name => formHeaders.indexOf(name);
  const formRow = formData.find(r => r[getFormColIndex("Название команды")] === teamName);

  if (!formRow) {
    SpreadsheetApp.getUi().alert(`Команда "${teamName}" не найдена в листе "Ответы на форму (1)"`);
    return;
  }

  const school = formRow[getFormColIndex("Школа")] || "";
  const grade = formRow[getFormColIndex("Класс")] || "";
  const leaderName = formRow[getFormColIndex("Имя руководителя")] || "";
  const leaderMiddleName = formRow[getFormColIndex("Отчество руководителя")] || "";

  const participants = formHeaders
    .map((h, i) => ({ h, i }))
    .filter(x => x.h.includes("ФИО участника"))
    .map(x => formRow[x.i])
    .filter(v => v && v.toString().trim() !== "")
    .join(", ");

  const htmlBody = bodyTemplate
    .replace("{{Имя}}", leaderName)
    .replace("{{Отчество}}", leaderMiddleName)
    .replace("{{Школа}}", school)
    .replace("{{Класс}}", grade)
    .replace("{{Команда}}", teamName)
    .replace("{{Участники}}", participants);

  GmailApp.sendEmail(userEmail, subject + " (Тест)", " ", {
    htmlBody: htmlBody,
    name: "Оргкомитет олимпиады",
    replyTo: "omsk.olympiad@gmail.com"
  });

  SpreadsheetApp.getUi().alert("Тестовое письмо отправлено на: " + userEmail);
}
