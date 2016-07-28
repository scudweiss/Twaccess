/**
MIT License

Copyright (c) [2016] [Samuel G. Scudere-Weiss]

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
**/

function onInstall(e){
  onOpen();
}

function onOpen(e){
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  if (e&&e.authMode == "NONE"){
    menu.addItem("Add Tokens", "firstRun").addToUi();
  } else {
    menu.addItem("Create Search", "showSidebar")
      .addItem("Delete Search", "showDeleteDialog")
      .addItem("Run Search", "twitterSearch")
      .addItem("Update Tokens", "showDialog")
      .addItem("Get Twitter Handle", "getHandle")
      .addToUi();
    PropertiesService.getDocumentProperties().setProperty('SOURCE_DATA_ID', SpreadsheetApp.getActiveSpreadsheet().getId());
    setupData();
    createAnalysis();
  }
}

function firstRun(){
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addItem("Create Search", "showSidebar")
      .addItem("Delete Search", "showDeleteDialog")
      .addItem("Run Search", "twitterSearch")
      .addItem("Update Tokens", "showDialog")
      .addToUi();
  PropertiesService.getDocumentProperties().setProperty('SOURCE_DATA_ID', SpreadsheetApp.getActiveSpreadsheet().getId());
  setupData();
  createAnalysis()
  showDialog();
}

function showDialog() {
  onOpen();
  var html = HtmlService.createHtmlOutputFromFile('AddTokens')
      .setWidth(300)
      .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, 'Please Enter Your Acess Tokens and Keys');
}

function showDeleteDialog() {
  var html = HtmlService.createHtmlOutputFromFile('DeleteSearch')
      .setWidth(290)
      .setHeight(240);
  SpreadsheetApp.getUi().showModalDialog(html, 'Plese Select the Search you Would Like to Delete');
}

function showSidebar () {
  var html = HtmlService.createHtmlOutputFromFile('CreateSearch')
      .setWidth(450)
      .setHeight(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function setupData (){
  var spreadsheet = SpreadsheetApp.getActive();
  if (spreadsheet.getSheetByName('AuthVals')==null) createAuthVals();
  if (spreadsheet.getSheetByName('SearchString')==null) createSearchPage();
  if (spreadsheet.getSheetByName('SearchLogs')==null) createSearchLogs();
}

function createSearchLogs(){
  var sheet = SpreadsheetApp.getActive().insertSheet().setName('SearchLogs');
  var titleRow = ['Search Name', 'Time Completed'];
  sheet.getRange(1, 1, 1, titleRow.length).setValues([titleRow]).setFontWeights([['bold','bold']]);
  
}

function createAuthVals (){
  var sheet = SpreadsheetApp.getActive().insertSheet().setName('AuthVals');
  sheet.getRange(1,1).setValue('Twitter Tokens').setFontWeight('bold');
  sheet.getRange(1, 2).setValue('Enter Values Below').setFontWeight('bold');
  sheet.getRange(2,1).setValue('Consumer Key').setFontWeight('bold');
  sheet.getRange(3,1).setValue('Consumer Secret').setFontWeight('bold');
  sheet.getRange(4,1).setValue('Access Token').setFontWeight('bold');
  sheet.getRange(5,1).setValue('Klout Token').setFontWeight('bold');
  sheet.protect();
}

function insertTokens (ConsumerKey, ConsumerSecret, KloutToken){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AuthVals');
  sheet.getRange(2,2).setValue(ConsumerKey);
  sheet.getRange(3,2).setValue(ConsumerSecret);
  sheet.getRange(5,2).setValue(KloutToken);
  postOAuth2Token();
}

function createSearchPage(){
  var sheet = SpreadsheetApp.getActive().insertSheet().setName('SearchString');
  sheet.getRange(1,1).setValue('Search Name').setFontWeight('bold');
  sheet.getRange(1,2).setValue('Search String').setFontWeight('bold');
  sheet.protect();
}

function newSearch (name, query){
  if(query!="&exclude=retweets&lang=en&count=100&include_entities=0"){
    SpreadsheetApp.getActive().getSheetByName('SearchString').appendRow([name, query]);
    SpreadsheetApp.getUi().alert("New Search "+name+" Added");
  } else SpreadsheetApp.getUi().alert("No search terms were added!");
  showSidebar();
}

function removeSearch (name, deleteData){
  var spreadsheet= SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName('SearchString');
  sheet.deleteRow(searchCol(sheet,name));
  sheet = spreadsheet.getSheetByName('Daily Analysis');
  sheet.deleteRow(searchCol(sheet,name));
  sheet = spreadsheet.getSheetByName('Monthly Analysis');
  sheet.deleteColumn(searchRow(sheet,name));
  if (deleteData){ 
    spreadsheet.deleteSheet(spreadsheet.getSheetByName(name));
  }
}

function searchRow (sheet, text){
  var num = sheet.getLastColumn();
  for (var i = 1; i<num; i++){
    if (sheet.getRange(1, i).getValue()==text){
      return i;
    }
  }  
  return num;
}

function searchCol (sheet, name){
  var num = sheet.getLastRow();
  for (var i = 1; i<num; i++) {
    if (sheet.getRange(i, 1).getValue()==name){ 
      return i; 
    }
  } 
  return num;
}

function twitterSearch(){
  var spreadsheet = SpreadsheetApp.openById(PropertiesService.getDocumentProperties().getProperty('SOURCE_DATA_ID'));
  for (var j=2; j<=spreadsheet.getSheetByName("SearchString").getLastRow(); j++){ 
    var firstSearch = false;
    var name = spreadsheet.getSheetByName('SearchString').getRange(j,1).getValue();
     if (spreadsheet.getSheetByName(name)==null){
        var sheet = spreadsheet.insertSheet().setName(name);
        sheet.appendRow(["UID","Status","Favorites","Retweets","Klout Score","Time","Tweet ID"]);
        sheet.setFrozenRows(1);
        spreadsheet.setColumnWidth(2, 300);
        spreadsheet.setColumnWidth(8, 300);
        firstSearch = true;
    }
    var sheet = spreadsheet.getSheetByName(name);
    var timeCol = searchRow (sheet, "Time");
    var numTweetsCol = timeCol+1;
    var lastRunCol = numTweetsCol+1;
    var SinceIdCol = lastRunCol +1;
    var averageKloutCol = SinceIdCol+1;
    var averageKlout=spreadsheet.getSheetByName(name).getRange(1,averageKloutCol).getValue();
    var options = {headers:{Authorization: "Bearer "+getToken()}};
    var url = 'https://api.twitter.com/1.1/search/tweets.json';
    var query = "q="+spreadsheet.getSheetByName('SearchString').getRange(j,2).getValue();
    var sinceId = "&since_id="+spreadsheet.getSheetByName(name).getRange(1, SinceIdCol).getValue();
    var string = "?"+query+sinceId;
    var responce = UrlFetchApp.fetch(url+string, options);
    var data = JSON.parse(responce);
    if (data.statuses.length>0){
      if (averageKlout==0) averageKlout = addDataToSheet(data,sheet);
      else averageKlout = (averageKlout+addDataToSheet(data,sheet))/2;
    }
    var sinceId = JSON.stringify(data.search_metadata.max_id);
    var noData = (sinceId==sheet.getRange(1, SinceIdCol).getValue());
    sheet.insertRowAfter(sheet.getMaxRows());
    sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()).sort([{column: timeCol, ascending: false}]);
    var date = new Date();
    spreadsheet.getSheetByName('SearchLogs').appendRow([name,date]).autoResizeColumn(1).autoResizeColumn(2);
    var lastRow = sheet.getLastRow();
    lastCol = sheet.getLastColumn();
    sheet.getRange(1, numTweetsCol).setValue(lastRow-1);
    sheet.getRange(1, lastRunCol).setValue(date);
    sheet.getRange(1, SinceIdCol).setValue(sinceId);
    sheet.getRange(1, averageKloutCol).setValue(averageKlout);
    for (var k=1; k<=lastCol; k++){
      if (k==2)sheet.getRange(1, k, lastRow).setWrap(true);
      else if (k==8) sheet.setColumnWidth(k, 100);
      else sheet.autoResizeColumn(k);
    }
    if (noData) removeDuplicates(sheet);
    sheet.hideColumns(numTweetsCol-1,6);
  }
  updateDailyAnalysis();
}

function getHandle (){
    var ui = SpreadsheetApp.getUi();
    var buttons = ui.ButtonSet.OK_CANCEL;
    var prompt = ui.prompt("Get Twitter Handle","Type the ID for the twitter handel you woulld like to find.\n Press OK to Continue", buttons);
    if (prompt.getSelectedButton()==ui.Button.OK){
      var id = prompt.getResponseText();
      var url = "https://api.twitter.com/1.1/users/show.json?user_id="+id;
      var options = {headers:{Authorization: "Bearer "+getToken()}};
      try {
        var responce = UrlFetchApp.fetch(url,options);
        var data = JSON.parse(responce);
        var name = "@"+data.screen_name;
        if (data.verified==true) var verified = "The user is verified.";
        else var verified = "The user is not verified";
        ui.alert("Get Twitter Handle: "+id, name+"\n"+verified, buttons);
      } 
      catch (e){
        ui.alert("Get Twitter Handle", id+" is not a valid UID. \n Please Try Again.", buttons);
      }
    }
}

function removeDuplicates(sheet) {
  var data = sheet.getDataRange().getValues();
  var newData = new Array();
  for(i in data){
    var row = data[i];
    var duplicate = false;
    for(j in newData){
      if(row.join() == newData[j].join()){
        duplicate = true;
      }
    }
    if(!duplicate){
      newData.push(row);
    }
  }
  sheet.clearContents();
  sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
}

function addDataToSheet(data,sheet){
  var key = SpreadsheetApp.getActive().getSheetByName('AuthVals').getRange(5,2).getValue();
  var averageKlout = 0;
  var length = data.statuses.length;
  for (var i=0; i<length; i++){
    var id = JSON.stringify(data.statuses[i].user.id);
    var text = JSON.stringify(data.statuses[i].text);
    var friends = JSON.stringify(data.statuses[i].user.friends_count);
    var followers = JSON.stringify(data.statuses[i].user.followers_count);
    var fav = JSON.stringify(data.statuses[i].favorite_count);
    var retweet = JSON.stringify(data.statuses[i].retweet_count);
    var time = JSON.stringify(data.statuses[i].created_at);
    var tweetId = JSON.stringify(data.statuses[i].id)
    var outreach = getKloutScore(id,key);
    averageKlout+=outreach;
    sheet.appendRow([id,text,fav,retweet,outreach,time,tweetId]);
  }
  return averageKlout/length;
}

function postOAuth2Token() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('AuthVals');
  if (sheet==null){
    setupSata();
  }
  var consumerKey = sheet.getRange(2,2).getValue();
  var consumerSecret = sheet.getRange(3,2).getValue();
  var encodedKeys = Utilities.base64Encode(consumerKey+':'+consumerSecret);
  var lngth = encodedKeys.length;
  var options = {
   headers: {
     Authorization: "Basic "+encodedKeys,
     "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8"    
   },
   method: "post",
   payload: "grant_type=client_credentials"
  };
  var responce = UrlFetchApp.fetch('https://api.twitter.com/oauth2/token', options);
  var data = JSON.parse(responce);
  if (data.token_type=="bearer"){
    sheet.getRange(4,2).setValue(data.access_token);
  }
}

function getToken (){
  var sheet = SpreadsheetApp.getActive().getSheetByName('AuthVals');
  if(sheet==null) setupData();
  var x = sheet.getRange(4,2).getValue();
  return x;
}

function getKloutScore (twitterID,key){
  try {
    var data = JSON.parse(UrlFetchApp.fetch('http://api.klout.com/v2/identity.json/tw/'+twitterID+'?key='+key));
    try {data = JSON.parse(UrlFetchApp.fetch('http://api.klout.com/v2/user.json/'+data.id+'/score?key='+key));}
    catch(err){data={score: 0};}
    var score = data.score;
    Utilities.sleep(150);
    return score;
  }
  catch(err){return 0;}
}

//Data Analysis

function createAnalysis(){
  createDailyAnalysis();
  createMonthlyAnalysis();
}

function createDailyAnalysis() {
  var id = PropertiesService.getDocumentProperties().getProperty('SOURCE_DATA_ID');
  var sheetName = "Daily Analysis";
  var spreadsheet = SpreadsheetApp.openById(id);
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (sheet==null){ 
    sheet=spreadsheet.insertSheet(sheetName);
    var titleRow = ["Search Name","Daily Tweets","Yesterday's Max","Total Tweets"];
    sheet.getRange(1, 1, 1, 4).setValues([titleRow]).setFontWeights([['bold','bold','bold','bold']]);
    update(true);
  }
}

function createMonthlyAnalysis() {
  var id = PropertiesService.getDocumentProperties().getProperty('SOURCE_DATA_ID');
  var sheetName = "Monthly Analysis";
  var spreadsheet = SpreadsheetApp.openById(id);
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (sheet==null){ 
    sheet=spreadsheet.insertSheet(sheetName);
    sheet.getRange(1, 1).setValue("Date").setFontWeight('bold');
  }
}

function updateMonthlyAnalysis() {
  var id = PropertiesService.getDocumentProperties().getProperty('SOURCE_DATA_ID');
  var sheetName = "Monthly Analysis";
  var spreadsheet = SpreadsheetApp.openById(id);
  var sheet = spreadsheet.getSheetByName(sheetName);
  var daySheet = spreadsheet.getSheetByName("Daily Analysis");
  var row = sheet.getLastRow()+1;
  var dayRow = daySheet.getLastRow(); 
  if (row>sheet.getMaxRows()) sheet.insertRowAfter(sheet.getMaxRows());
  var n =[""];
  var name = daySheet.getRange(2,1,dayRow-1).getValues();
  for (var i=0; i<name.length;i++) n[i] = name[i][0];
  name = [n];
  sheet.getRange(1, 2, 1, n.length).setValues(name);
  var d = new Date();
  d.setDate(d.getDate()-1);
  sheet.autoResizeColumn(1);
  var numTweets = daySheet.getRange(2,3,dayRow-1).getValues();
  var x = [d];
  for (var i=0; i<numTweets.length;i++) x[i+1] = numTweets[i][0];
  sheet.appendRow(x);
}

function dailyUpdate () {
  var id = PropertiesService.getDocumentProperties().getProperty('SOURCE_DATA_ID');
  var sheetName = "SearchLogs";
  var spreadsheet = SpreadsheetApp.openById(id);
  var sheet = spreadsheet.getSheetByName(sheetName);
  var titleRow = ['Search Name', 'Time Completed'];
  updateMonthlyAnalysis();
  sheet.clear().getRange(1, 1, 1, titleRow.length).setValues([titleRow]).setFontWeights([['bold','bold']]);
  update(true);
}

function updateDailyAnalysis(){
  update(false)
}

function update(dailyUpdate) {
  var id = PropertiesService.getDocumentProperties().getProperty('SOURCE_DATA_ID');
  var sheetName = "Daily Analysis";
  var spreadsheet = SpreadsheetApp.openById(id);
  var sheet = spreadsheet.getSheetByName(sheetName);
  var searchSheet = spreadsheet.getSheetByName("searchString");
  for (var i=2; i<=searchSheet.getLastRow(); i++){
    if (sheet.getRange(i, 3).isBlank()) sheet.getRange(i, 3).setValue(0);
    var name = searchSheet.getRange(i, 1).getValue();
    var tweetSheet = spreadsheet.getSheetByName(name);
    var timeCol = searchRow(tweetSheet, "Time");
    var totalTweets = tweetSheet.getRange(1, timeCol+1).getValue();
    var dayTweets = totalTweets-(sheet.getRange(i, 3).getValue());
    if (dailyUpdate==true) sheet.getRange(i, 3).setValue(totalTweets);
    sheet.getRange(i, 1).setValue(name);
    sheet.getRange(i, 2).setValue(dayTweets);
    sheet.getRange(i, 4).setValue(totalTweets);
  }
  for (var i=1; i<=sheet.getLastColumn(); i++) sheet.autoResizeColumn(i);
}
