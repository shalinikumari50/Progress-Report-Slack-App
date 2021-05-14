function onOpen(e) {
   SpreadsheetApp.getUi()
       .createMenu('Messager Menu')
       .addItem('Send to Recipients', 'postLoop')
       .addToUi();
 }


var COL_A = 1; //index of column in google sheets
var COL_B = 2; //index of column in google sheets

var EMAIL_COL = 0;  //index of column in google sheets

function postLoop () {
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Input")


 var rangeValues = sheet.getDataRange().getValues();

  for (i=1;i<rangeValues.length;i++){
  
    var lookupResponse = JSON.parse(UrlFetchApp.fetch(
  "https://slack.com/api/users.lookupByEmail?email="+ rangeValues[i][EMAIL_COL],
  {
    method             : 'GET',
   
    contentType        : 'application/x-www-form-urlencoded',
    
    headers            : {
      Authorization : 'Bearer ' + 'YOUR_USER_OAUTH_TOKEN',
     
    
    },
    muteHttpExceptions : true,
}).getContentText());
try{
  var channel = lookupResponse['user']['id'];}
  catch(e){
  
  var changeRange = sheet.getRange(i+1,1,1,sheet.getLastColumn());
changeRange.setBackgroundRGB(224, 102, 102);
  continue;}

    postToSlack(channel, rangeValues[i][COL_A], rangeValues[i][COL_B])
  }
}


function postToSlack(channel, data1, data2) {
  var payload = {
    'channel' : channel,
    'text' : 'SLACK_MESSAGE_GOES_HERE',
    'as_user' : true
  }
  
return UrlFetchApp.fetch(
  "https://slack.com/api/chat.postMessage",
  {
    method             : 'post',
    contentType        : 'application/json',
    headers            : {
      Authorization : 'Bearer ' + 'YOUR_USER_OAUTH_TOKEN'
    },
    payload            : JSON.stringify(payload),
    muteHttpExceptions : true,
})
}
