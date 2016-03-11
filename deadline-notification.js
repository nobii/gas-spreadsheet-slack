var COLUMN_TITLE = 0;
var COLUMN_CHANNEL = 1;
var COLUMN_LAUNCH_DATE = 2;
var COLUMN_ICON = 3;
var SUNDAY = 0;
var SATURDAY = 6;
var DAY_MSECONDS = 1000 * 60 * 60 * 24;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Slackの設定")
    .addItem("Slack API Tokenの登録", "setAPIToken")
    .addItem("今すぐ実行", "executeNow")
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


function executeNow() {
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
  return values[0]
}


function postSingleNotificationMessage(row) {
  var now = new Date();
  var workdays = loadWorkdays(now);
  if (!workdays[0]) {
    return;
  }

  var slackMessage = new SlackMessage();
  var sheet = SpreadsheetApp.getActiveSheet();
  var values = sheet.getSheetValues(row, 1, 1, sheet.getLastColumn());
  var sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();

  var title = values[0][COLUMN_TITLE];
  var channel = values[0][COLUMN_CHANNEL];
  var launchDate = values[0][COLUMN_LAUNCH_DATE];
  var icon = values[0][COLUMN_ICON];

  var message;
  var prevDate = new Date();
  prevDate.setDate(prevDate.getDate() - 1);
  if (launchDate < prevDate) {
    message = "このプロジェクトは終了しました！<" + sheetUrl + "|ここ>から削除してください";
  } else {
    var leftDays = Math.ceil((launchDate - now) / DAY_MSECONDS);
    var leftWorkDays = countWorkday(leftDays, workdays);
    message = (launchDate.getMonth() + 1) + "月" + launchDate.getDate() + "日まであと" + leftWorkDays + "営業日です";
  }

  slackMessage.postMessage(title, channel, message, icon);  
}


function postNotificationMessage() {
  var now = new Date();
  var workdays = loadWorkdays(now);
  if (!workdays[0]) {
    return;
  }
  
  var slackMessage = new SlackMessage();
  var sheet = SpreadsheetApp.getActiveSheet();
  var values = sheet.getDataRange().getValues();
  var maxRow = sheet.getLastRow();
  var sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  
  for (var r = 1; r < maxRow; r++) {
    var title = values[r][COLUMN_TITLE];
    var channel = values[r][COLUMN_CHANNEL];
    var launchDate = values[r][COLUMN_LAUNCH_DATE];
    var icon = values[r][COLUMN_ICON];
    
    var message;
    var prevDate = new Date();
    prevDate.setDate(prevDate.getDate() - 1);
    if (launchDate < prevDate) {
      message = "このプロジェクトは終了しました！<" + sheetUrl + "|ここ>から削除してください";
    } else {
      var leftDays = Math.ceil((launchDate - now) / DAY_MSECONDS);
      var leftWorkDays = countWorkday(leftDays, workdays);
      message = (launchDate.getMonth() + 1) + "月" + launchDate.getDate() + "日まであと" + leftWorkDays + "営業日です";
    }
  
    slackMessage.postMessage(title, channel, message, icon);
  }
}


function postNotificationMessage(slackMessage, workdays, values, now) {
  var title = values[COLUMN_TITLE];
  var channel = values[COLUMN_CHANNEL];
  var launchDate = values[COLUMN_LAUNCH_DATE];
  var icon = values[COLUMN_ICON];
  
  var message;
  var now = new Date();
  var prevDate = new Date();
  prevDate.setDate(prevDate.getDate() - 1);
  if (launchDate < prevDate) {
    var sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    message = "このプロジェクトは終了しました！<" + sheetUrl + "|ここ>から削除してください";
  } else {
    var leftDays = Math.ceil((launchDate - now) / DAY_MSECONDS);
    var leftWorkDays = countWorkday(leftDays, workdays);
    message = (launchDate.getMonth() + 1) + "月" + launchDate.getDate() + "日まであと" + leftWorkDays + "営業日です";
  }

  slackMessage.postMessage(title, channel, message, icon); 
}


function loadWorkdays() {
  var workdays = [];
  var calendar = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
  var now = new Date();
  var dateAfterYear = new Date();
  dateAfterYear.setYear(dateAfterYear.getFullYear() + 1);
  var holidays = calendar.getEvents(now, dateAfterYear).map(function(item) {
    var holidayDate = item.getStartTime();
    return holidayDate.getFullYear() + "/" + (holidayDate.getMonth() + 1) + "/" + holidayDate.getDate();
  }); 
  for (var i = 0; i < 365; i++) {
    var weekday = now.getDay();
    var formattedDate = now.getFullYear() + "/" + (now.getMonth() + 1) + "/" + now.getDate();
    var isHoliday = holidays.indexOf(formattedDate) != -1;
    workdays[i] = (weekday == SUNDAY || weekday == SATURDAY || isHoliday) ? 0 : 1;
    now.setDate(now.getDate() + 1);
  }
  return workdays;
}


function countWorkday(days, workdays) {
  if (days <= 0) {
    return 0;
  }
  var targetWorkdays = workdays.slice(0, days);
  var count = targetWorkdays.reduce(function(result, item, index, array) {
    return result + (index < days ? item : 0); 
  });
  return count;
}


var SlackMessage = (function() {
  function SlackMessage() {
    var prop = PropertiesService.getScriptProperties().getProperties();
    this.slackApp = SlackApp.create(prop.token);
    this.channels = this.slackApp.channelsList().channels;
  }
  
  SlackMessage.prototype.postMessage = function(title, channelName, message, icon) {
    var channelId = this.getChannelId(channelName);
    this.slackApp.postMessage(channelId,
                              message, 
                              {
                                username : title,
                                icon_emoji : ":" + icon + ":",
                              });
  };
  
  SlackMessage.prototype.getChannelId = function(channelName) {
    var channel = this.channels.filter(function(item, index) {
      return item.name === channelName;
    });
    return channel[0] ? channel[0].id : null;
  };
  
  return SlackMessage;
})();
