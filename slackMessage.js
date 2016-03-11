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
