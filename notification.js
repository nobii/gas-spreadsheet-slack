var Notification = (function() {
  function Notification() {
    var now = Date();
    this.workdays = this.loadWorkdays(now);
    this.slackMessage = new SlackMessage();
  }

  Notification.prototype.postMessage = function(values) {
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
      var leftWorkDays = this.countWorkday(leftDays);
      message = (launchDate.getMonth() + 1) + "月" + launchDate.getDate() + "日まであと" + leftWorkDays + "営業日です";
    }

    this.slackMessage.postMessage(title, channel, message, icon); 
  };

  Notification.prototype.loadWorkdays = function() {
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
  };

  Notification.prototype.countWorkday = function(days) {
    if (days <= 0) {
      return 0;
    }
    var targetWorkdays = this.workdays.slice(0, days);
    var count = targetWorkdays.reduce(function(result, item, index, array) {
      return result + (index < days ? item : 0); 
    });
    return count;   
  };

  Notification.prototype.isHoliday = function() {
    return !this.workdays[0];
  };

  return Notification;
})();
