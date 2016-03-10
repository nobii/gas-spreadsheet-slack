var COLUMN_TITLE = 0;
var COLUMN_CHANNEL = 1;
var COLUMN_LAUNCH_DATE = 2;
var COLUMN_ICON = 3;
var SUNDAY = 0;
var SATURDAY = 6;

var workdays = [];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Slackの設定")
    .addItem("Slack API Tokenの登録", "setAPIToken")
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


function postNotificationMessage() {
  var now = new Date();
  loadWorkdays(now);
  if (!workdays[0]) {
    return;
  }
  
  var slackMessage = new SlackMessage();
  var sheet = SpreadsheetApp.getActiveSheet();
  var values = sheet.getDataRange().getValues();
  var maxRow = sheet.getLastRow();
  var sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  
  for (var r = 1; r < maxRow; r++) {
    now = new Date();
    var title = values[r][COLUMN_TITLE];
    var channel = values[r][COLUMN_CHANNEL];
    var launchDate = values[r][COLUMN_LAUNCH_DATE];
    var icon = values[r][COLUMN_ICON];
    
    if (launchDate < now) {
      var message = "このプロジェクトは終了しました！<" + sheetUrl + "|ここ>から削除してください";
      slackMessage.postMessage(title, channel, message, icon);
      continue;
    }
    
    var leftDays = Math.ceil((launchDate - now) / 86400000);
    var leftWorkDays = countWorkday(leftDays);
    var message = (launchDate.getMonth() + 1) + "月" + launchDate.getDate() + "日まであと" + leftWorkDays + "営業日です";
  
    slackMessage.postMessage(title, channel, message, icon);
  }
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
  }
  
  SlackMessage.prototype.getChannelId = function(channelName) {
    var channel = this.channels.filter(function(item, index) {
      if (item.name === channelName) return true;
    });
    return channel[0] ? channel[0].id : null;
  }
  
  return SlackMessage;
})();


function loadWorkdays(now) {
  var calendar = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
  var dateAfterYear = new Date();
  dateAfterYear.setYear(dateAfterYear.getFullYear() + 1);
  var holidays = calendar.getEvents(now, dateAfterYear).map(function(item) {
    return item.getStartTime().getFullYear() + "/" + (item.getStartTime().getMonth() + 1) + "/" + item.getStartTime().getDate();
  }); 
  for (var i = 0; i < 365; i++) {
    var weekday = now.getDay();
    var formattedDate = now.getFullYear() + "/" + (now.getMonth() + 1) + "/" + now.getDate();
    if (weekday == SUNDAY || weekday == SATURDAY || holidays.indexOf(formattedDate) != -1) {
      workdays[i] = 0;
    } else {
      workdays[i] = 1;
    }
    now.setDate(now.getDate() + 1);
  }
}


function countWorkday(days) {
  var tagetWorkdays = workdays.slice(0, days);
  var count = targetWorkdays.reduce(function(result, item, index, array) {
    return result + (index < days ? item : 0); 
  });
  return count;
}
