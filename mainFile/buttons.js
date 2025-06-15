function onOpen() {

  const ui1 = SpreadsheetApp.getUi();

  ui1.createMenu("Олимпиада")
    .addItem("Обновить id файлов", "idTable")
    .addToUi();

  const ui2 = SpreadsheetApp.getUi();

  ui2.createMenu("Рассылка")
    .addItem("Сделать рассылку писем-приглашений", "sendInvitations")
    .addItem("Отправть тестовое письмо-приглашение", "sendTestEmailIn")
    .addItem("Сделать рассылку подтверждений регистрации", "sendConfirmation")
    .addItem("Отправить тестовое подтверждение", "sendTestEmailCon")
    .addToUi();

  const ui3 = SpreadsheetApp.getUi();

  ui3.createMenu("Писма-ответы")
    .addItem("Отсортировать письма на обработку", "sortEmails")
    .addItem("Скачать файлы из отсортированных файлов", "saveFiles")
    .addToUi();
    
}
