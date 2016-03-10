var COLUMN_TITLE = 0;
var COLUMN_CHANNEL = 1;
var COLUMN_LAUNCH_DATE = 2;
var COLUMN_ICON = 3;
var SUNDAY = 0;
var SATURDAY = 6;

var calendar = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');

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
  if (isHoliday(now)) {
    return; 
  }
  
  var slackMessage = new SlackMessage();
  var sheet = SpreadsheetApp.getActiveSheet();
  var values = sheet.getDataRange().getValues();
  var maxRow = sheet.getLastRow();
  var sheet_url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  
  for (var r = 1; r < maxRow; r++) {
    now = new Date();
    var title = values[r][COLUMN_TITLE];
    var channel = values[r][COLUMN_CHANNEL];
    var launch_date = values[r][COLUMN_LAUNCH_DATE];
    var icon = values[r][COLUMN_ICON];
    
    if (launch_date < now) {
      var message = "このプロジェクトは終了しました！<" + sheet_url + "|ここ>から削除してください";
      slackMessage.postMessage(title, channel, message, icon);
      continue;
    }
    
    var left_days = Math.ceil(countWorkday(now, launch_date));
    var message = (launch_date.getMonth() + 1) + "月" + launch_date.getDate() + "日まであと" + left_days + "営業日です";
  
    slackMessage.postMessage(title, channel, message, icon);
  }
}


var SlackMessage = (function() {
  function SlackMessage() {
    var prop = PropertiesService.getScriptProperties().getProperties();
    this.slackApp = SlackApp.create(prop.token);
    this.channels = this.slackApp.channelsList().channels;
  }
  
  SlackMessage.prototype.postMessage = function(title, channel_name, message, icon) {
    var channelId = this.getChannelId(channel_name);
    this.slackApp.postMessage(channelId,
                              message, 
                              {
                                username : title,
                                icon_emoji : ":" + icon + ":",
                              });
  }
  
  SlackMessage.prototype.getChannelId = function(channel_name) {
    var channel = this.channels.filter(function(item, index) {
      if (item.name === channel_name) return true;
    });
    return channel[0] ? channel[0].id : null;
  }
  
  return SlackMessage;
})();
 

// http://improve-future.com/google-apps-script-process-only-on-business-day.html
function isHoliday(date) {
  var weekday = date.getDay();
  if (weekday == SUNDAY || weekday == SATURDAY) {
    return true;
  }
  if (calendar.getEventsForDay(date, {max: 1}).length > 0) {
    return true;
  }
  
  return false;
}


function countWorkday(from, to) {
  var count = 0;
  while (from < to) {
    from.setDate(from.getDate() + 1);
    if (!isHoliday(from)) {
      count++;
    }
  }
  
  return count;
}