var COLUMN_TITLE = 0;
var COLUMN_CHANNEL = 1;
var COLUMN_LAUNCH_DATE = 2;
var COLUMN_ICON = 3;
var SUNDAY = 0;
var SATURDAY = 6;
var DAY_MSECONDS = 1000 * 60 * 60 * 24;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Slackに投稿")
    .addItem("今すぐ実行", "showExecutePrompt")
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('設定')
      .addItem("Slack API Tokenの登録", "setAPIToken"))
    .addToUi();
}


function setAPIToken() {
  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt("Slack API Tokenを入力してください", ui.ButtonSet.OK_CANCEL);
  
  switch (prompt.getSelectedButton()) {
    case ui.Button.OK:
      var str = prompt.getResponseText();
      ui.alert("登録しました");
      saveToken(str);
      break;
    default:
      break;
  }
}


function saveToken(str) {
  PropertiesService.getScriptProperties().setProperty("token", str);
}


function showExecutePrompt() {
  var output = HtmlService.createHtmlOutputFromFile('prompt')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(400)
    .setHeight(200);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.show(output);
}


function getProjectList() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var values = sheet.getSheetValues(2, 1, sheet.getLastRow(), 1);
  return values;
}


function postSingleNotificationMessage(row) {
  var notification = new Notification();
  var sheet = SpreadsheetApp.getActiveSheet();
  var values = sheet.getSheetValues(row, 1, 1, sheet.getLastColumn());
  notification.postMessage(values[0]);
}


function postAllNotificationMessage() {
  var notification = new Notification();
  if (notification.isHoliday()) {
    return;
  }

  var sheet = SpreadsheetApp.getActiveSheet();
  var values = sheet.getDataRange().getValues();
  var maxRow = sheet.getLastRow();
  
  for (var r = 1; r < maxRow; r++) {
    notification.postMessage(values[r]);
  }
}
