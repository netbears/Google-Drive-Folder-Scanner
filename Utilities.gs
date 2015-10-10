var array = [];
var timeOut = 300;
var start = new Date();
var reachedTimeOut = false;
var curHeader;
  

function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}
