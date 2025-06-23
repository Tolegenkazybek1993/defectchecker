function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const user = Session.getActiveUser().getEmail();
  const allowed = JSON.parse(PropertiesService.getScriptProperties().getProperty("allowedEmails") || "[]");

  if (!allowed.includes(user.toLowerCase())) {
    ui.alert("⛔ У вас нет доступа к загрузке файлов. Подсказка МКБ доступна через меню.");
  }

  ui.createMenu("🔍 Проверка")
    .addItem("📥 Загрузить файл и проверить", "showUploadDialog")
    .addToUi();
}

function showUploadDialog() {
  const html = HtmlService.createHtmlOutputFromFile("index")
    .setWidth(600)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, "Загрузка Excel-файла");
}