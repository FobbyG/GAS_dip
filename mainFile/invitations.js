function sendInvitations() {
  const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const idSheet = currentSpreadsheet.getSheetByName("id");

  if (!idSheet) {
    SpreadsheetApp.getUi().alert("Лист 'id' не найден.");
    return;
  }

  const idData = idSheet.getRange(2, 5, idSheet.getLastRow() - 1, 2).getValues(); // E и F
  const dbRow = idData.find(row => row[0] === "База данных участников предыдущих лет");

  if (!dbRow || !dbRow[1]) {
    SpreadsheetApp.getUi().alert("ID файла 'База данных участников предыдущих лет' не найден на листе 'id'.");
    return;
  }

  const dbFileId = dbRow[1].toString().trim();
  const dataSpreadsheet = SpreadsheetApp.openById(dbFileId);
  const sheetData = dataSpreadsheet.getSheetByName("База данных");

  if (!sheetData) {
    SpreadsheetApp.getUi().alert("Лист 'База данных' не найден в файле.");
    return;
  }

  const sheetMessage = currentSpreadsheet.getSheetByName("Шаблон письма-приглашения");

  if (!sheetMessage) {
    SpreadsheetApp.getUi().alert("Лист 'Щаблон письма-приглашения' не найден.");
    return;
  }

  const dataRange = sheetData.getRange(2, 1, sheetData.getLastRow() - 1, 6);
  const data = dataRange.getValues();

  const subject = sheetMessage.getRange("B1").getValue();
  const bodyTemplate = sheetMessage.getRange("B2").getValue();
  const formLink = sheetMessage.getRange("B3").getValue();
  const infLink = sheetMessage.getRange("B4").getValue();

  let sentCount = 0;
  let quotaLeft = MailApp.getRemainingDailyQuota();

  if (quotaLeft <= 0) {
    SpreadsheetApp.getUi().alert("Квота на отправку писем на сегодня исчерпана.");
    return;
  }

  for (let i = 0; i < data.length; i++) {
    if (quotaLeft <= 0) break;

    const [checkbox, email, name, patronymic ] = data[i];
    if (checkbox === true || !email) continue;

    let personalizedText = bodyTemplate
      .replace("{{Имя}}", name)
      .replace("{{Отчество}}", patronymic);

    let htmlBody = `
      <p>${personalizedText}</p>
      <p><a href="${formLink}">Ссылка на регистрацию</a></p>
      <p><a href="${infLink}">Положение об олимпиаде</a></p>
    `;

    //заголовки, чтобы письмо не упало в спам
    GmailApp.sendEmail(email, subject, "", {
      htmlBody: htmlBody,
      name: "Оргкомитет олимпиады",
      replyTo: "omsk.olympiad@gmail.com", 
    });

    sheetData.getRange(i + 2, 1).setValue(true);
    sentCount++;
    quotaLeft--;

    Utilities.sleep(500); // небольшая пауза, чтобы избежать блокировки за спам
  }

  SpreadsheetApp.getUi().alert(`Рассылка завершена. Отправлено писем: ${sentCount}`);
}

//функция тестового письма
function sendTestEmailIn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const idSheet = ss.getSheetByName("id");

  if (!idSheet) {
    SpreadsheetApp.getUi().alert("Лист 'id' не найден.");
    return;
  }

  const idData = idSheet.getRange(2, 5, idSheet.getLastRow() - 1, 2).getValues(); // E и F
  const dbRow = idData.find(row => row[0] === "База данных участников предыдущих лет");

  if (!dbRow || !dbRow[1]) {
    SpreadsheetApp.getUi().alert("ID файла 'База данных участников предыдущих лет' не найден на листе 'id'.");
    return;
  }

  const dbFileId = dbRow[1].toString().trim();
  const dataSpreadsheet = SpreadsheetApp.openById(dbFileId);
  const sheetData = dataSpreadsheet.getSheetByName("База данных");

  if (!sheetData) {
    SpreadsheetApp.getUi().alert("Лист 'База данных' не найден в файле.");
    return;
  }

  const sheetMessage = ss.getSheetByName("шаблон письма-приглашения");

  if (!sheetMessage) {
    SpreadsheetApp.getUi().alert("Лист 'шаблон письма-приглашения' не найден.");
    return;
  }

  const subject = sheetMessage.getRange("B1").getValue();
  const bodyTemplate = sheetMessage.getRange("B2").getValue();
  const formLink = sheetMessage.getRange("B3").getValue();

  const firstRow = sheetData.getRange(2, 1, 1, 5).getValues()[0]; // A2:E2
  const [_, name, patronymic ] = firstRow;

  let body = bodyTemplate
    .replace("{{Имя}}", name)
    .replace("{{Отчество}}", patronymic);

  body += `\n\nСсылка на регистрацию: ${formLink}`;

  const userEmail = Session.getActiveUser().getEmail();

 GmailApp.sendEmail(userEmail, subject + " (Тест)", " ", {
    htmlBody: htmlBody,
    name: "Оргкомитет олимпиады",
    replyTo: "omsk.olympiad@gmail.com"
  });

  SpreadsheetApp.getUi().alert("Тестовое письмо отправлено на ваш адрес: " + userEmail);
}
