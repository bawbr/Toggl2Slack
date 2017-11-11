var vg_tokenOfToggl;
var vg_email;
var vg_wip;
var vg_tokenOfSlack;

function myFunction() {
  defineGlobalVariables();
  getNewTogglTimeEntries();
  edit();
}

function defineGlobalVariables(){
  var ss = SpreadsheetApp.openById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var sheet = ss.getSheetByName("AuthSettings");
  var range_tokenOfToggl = sheet.getRange('b1');
  var range_wip = sheet.getRange('b2');
  var range_tokenOfSlack = sheet.getRange('b3');
  
  vg_tokenOfToggl = range_tokenOfToggl.getValue();
  vg_wip = range_wip.getValue();
  vg_tokenOfSlack = range_tokenOfSlack.getValue(); 
  
}

function getNewTogglTimeEntries(){
  var dataJSON = toggl();
  var data = JSON.parse(dataJSON);
 
  var ss = SpreadsheetApp.openById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var sheet = ss.getSheetByName("TimeEntries");
  var cell = sheet.getRange('k:k');
  var cell0 = sheet.getRange('a1');
  var lastRow = sheet.getLastRow();
  var cell_id = cell.getValues();
  var ProjectName;
  
  var f = true;
  
  for(var i in data ){
    f = true;
    if (data[i].stop == undefined || data[i].stop == data[i].start ){
      f = false;
      break;
    }
    
    for (var j in cell_id ){
      if (data[i].id == cell_id[j]){
        f = false;
        break;
      } 
    }
      
    if (f == true ){
        var pid = data[i].pid;   
        if (pid != undefined ){  
          var dataProjectJSON = togglProject(pid);
          var dataProject = JSON.parse(dataProjectJSON);  
          
          if (dataProject.data.name != undefined ){
            cell0.offset(lastRow,0).setValue(dataProject.data.name);
          }
        }
        
      var h = Math.floor(data[i].duration / 3600, 0);
      var m = Math.floor((data[i].duration - h * 3600) / 60,0);
      var s = Math.floor(data[i].duration - h * 3600 - m * 60 );
      
      var startDate = new Date(getDateFromIso(data[i].start));
      var atDate = new Date(getDateFromIso(data[i].at));
      var stopDate = new Date(getDateFromIso(data[i].stop));

      cell0.offset(lastRow,1).setValue(data[i].description);
      cell0.offset(lastRow,2).setValue(h + ':' + m + ':' + s);
      cell0.offset(lastRow,3).setValue(Utilities.formatDate(startDate,'JST','yyyy/MM/dd HH:mm:ss'));
      cell0.offset(lastRow,4).setValue(Utilities.formatDate(stopDate,'JST','yyyy/MM/dd HH:mm:ss'));
      cell0.offset(lastRow,6).setValue(data[i].uid);
      cell0.offset(lastRow,7).setValue(data[i].wid);
      cell0.offset(lastRow,8).setValue(Utilities.formatDate(atDate,'JST','yyyy/MM/dd HH:mm:ss'));
      cell0.offset(lastRow,9).setValue(data[i].pid);
      cell0.offset(lastRow,10).setValue(data[i].id);
      cell0.offset(lastRow,11).setValue(data[i].billable);
      cell0.offset(lastRow,12).setValue(data[i].duronly);

      lastRow++;
    }
  }
  
}




function edit(){
  var ss = SpreadsheetApp.openById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var sheet = ss.getSheetByName("TimeEntries");
  var cell = sheet.getRange('a1');
  var col = 0;
  var today = new Date();

  var sheetdata = sheet.getRange(1,1,sheet.getLastRow(),6).getValues();
  for(var i=1; i<sheet.getLastRow(); i++ ){
    
    var cProjectName = sheetdata[i][0];
    var cTaskDescription = sheetdata[i][1];
    var cDuration = sheetdata[i][2];
    var cStart = sheetdata[i][3];
    var cStop = sheetdata[i][4];
    var cSlack = sheetdata[i][5];
    
    var message;

      if( cDuration != '' && cSlack == '' ){
          cSlack = today;
          cell.offset(i,5).setValue( cSlack );

        message =           'Project: ' + cProjectName + '\n';
        message = message + 'Task Description: ' + cTaskDescription + '\n';
        message = message + 'Duration: ' + Utilities.formatDate(cDuration,'JST','HH:mm:ss') + '\n';
        message = message + 'Start: ' + cStart + '\n';
        message = message + 'Stop: ' + cStop  + '\n';
      
        postSlackMessage(message);
        Logger.log(message);
      }
    }
}


function toggl(){
  var url = 'https://www.toggl.com/api/v8/time_entries';
  var token = vg_tokenOfToggl;
  var token64 = Utilities.base64Encode(token);
  var auth = 'Basic ' + token64;
  var wip = vg_wip;
  var is_private = true;
  var billable = true;
  var method = 'get';
  var contenttype = 'Content-Type: application/json';
  var muteHttpExceptions = true;
  
  var headers = {
    'Authorization'      : auth
  };
  
  var params = {
    'method' : method,
    'contentType' : contenttype,
    'headers' : headers,
    'muteHttpExceptions' : muteHttpExceptions
  };
  
  var response = UrlFetchApp.fetch(url, params);
  return response;
  
}

function togglProject(pid){
  var url = 'https://www.toggl.com/api/v8/projects/'+ pid;
  Logger.log(url);
  var token = vg_tokenOfToggl;
  var token64 = Utilities.base64Encode(token);
  var auth = 'Basic ' + token64;
  var wip = vg_wip;
  var is_private = true;
  var billable = true;
  var method = 'get';
  var contenttype = 'Content-Type: application/json';
  var muteHttpExceptions = true;
  
  var headers = {
    'Authorization'      : auth
  };
  
  var params = {
    'method' : method,
    'contentType' : contenttype,
    'headers' : headers,
    'muteHttpExceptions' : muteHttpExceptions
  };
  
  var response = UrlFetchApp.fetch(url, params);
  return response;
  
}

function postSlackMessage(message) {
  var url        = 'https://slack.com/api/chat.postMessage';
  var token      = vg_tokenOfSlack;
  var channel    = '#general';
  var text       = message;
  var username   = 'Toggl Update BOT';
  var parse      = 'full';
  var icon_emoji = ':dog:';
  var method     = 'post';
 
  var payload = {
        'token'      : token,
        'channel'    : channel,
        'text'       : text,
        'username'   : username,
        'parse'      : parse,
        'icon_emoji' : icon_emoji
  };
 
  var params = {
        'method' : method,
        'payload' : payload
  };
 
  var response = UrlFetchApp.fetch(url, params);
  
  
}

function getDateFromIso(string) {
  try{
    var aDate = new Date();
    var regexp = "([0-9]{4})(-([0-9]{2})(-([0-9]{2})" +
        "(T([0-9]{2}):([0-9]{2})(:([0-9]{2})(\\.([0-9]+))?)?" +
        "(Z|(([-+])([0-9]{2}):([0-9]{2})))?)?)?)?";
    var d = string.match(new RegExp(regexp));

    var offset = 0;
    var date = new Date(d[1], 0, 1);

    if (d[3]) { date.setMonth(d[3] - 1); }
    if (d[5]) { date.setDate(d[5]); }
    if (d[7]) { date.setHours(d[7]); }
    if (d[8]) { date.setMinutes(d[8]); }
    if (d[10]) { date.setSeconds(d[10]); }
    if (d[12]) { date.setMilliseconds(Number("0." + d[12]) * 1000); }
    if (d[14]) {
      offset = (Number(d[16]) * 60) + Number(d[17]);
      offset *= ((d[15] == '-') ? 1 : -1);
    }

    offset -= date.getTimezoneOffset();
    time = (Number(date) + (offset * 60 * 1000));
    return aDate.setTime(Number(time));
  } catch(e){
    return;
  }
}