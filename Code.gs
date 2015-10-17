/**
 * @OnlyCurrentDoc  Limits the script to only accessing the current spreadsheet.
 */

var SIDEBAR_TITLE = 'Get Drive Report';

function onOpen(e){
  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Drive scanner', 'showSidebar')
      .addToUi();
}

function onInstall(e) {
  onOpen(e);
}



/**
 * Opens a sidebar. The sidebar structure is described in the Sidebar.html
 * project file.
 */
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('Sidebar')
      .evaluate()
      .setTitle(SIDEBAR_TITLE)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(ui);
}

function showPicker(obj) {
  
  PropertiesService.getScriptProperties().setProperty('Name', obj.Name);
  PropertiesService.getScriptProperties().setProperty('Type', obj.Type);
  PropertiesService.getScriptProperties().setProperty('Description', obj.Description);
  PropertiesService.getScriptProperties().setProperty('Id', obj.Id);
  PropertiesService.getScriptProperties().setProperty('Url', obj.Url);
  PropertiesService.getScriptProperties().setProperty('Date created', obj['Date created']);
  PropertiesService.getScriptProperties().setProperty('Last updated', obj['Last updated']);
  PropertiesService.getScriptProperties().setProperty('Size', obj.Size);
  PropertiesService.getScriptProperties().setProperty('Owner', obj.Owner);
  PropertiesService.getScriptProperties().setProperty('Sharing Access', obj['Sharing Access']);
  PropertiesService.getScriptProperties().setProperty('Sharing Permission', obj['Sharing Permission']);
  PropertiesService.getScriptProperties().setProperty('Viewers', obj.Viewers);
  PropertiesService.getScriptProperties().setProperty('Editors', obj.Editors);
  PropertiesService.getScriptProperties().setProperty('Parent folder', obj['Parent folder']);
 
  
  var html = HtmlService.createHtmlOutputFromFile('Picker.html')
      .setWidth(900)
      .setHeight(637.5)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
   SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/**
 * Returns the value in the active text input.
 *
 * @return {object} - An object with "folderName" and "folderID" properties.
 */
function getFolderName(value) {
  var sheet;
  var folderID = value;
  var folderName = DriveApp.getFolderById(folderID).getName();
  
  //SpreadsheetApp.getUi().alert('Gathering data for folder "'+folderName+'". \nThis might take a few minutes depending on the size of the folder. \nPlease wait...');
  
  start = new Date();
  
  setCurrentHeaders([PropertiesService.getScriptProperties().getProperty('Name'),
             PropertiesService.getScriptProperties().getProperty('Type'),
             PropertiesService.getScriptProperties().getProperty('Description'),
             PropertiesService.getScriptProperties().getProperty('Id'),
             PropertiesService.getScriptProperties().getProperty('Url'),
             PropertiesService.getScriptProperties().getProperty('Date created'),
             PropertiesService.getScriptProperties().getProperty('Last updated'),
             PropertiesService.getScriptProperties().getProperty('Size'),
             PropertiesService.getScriptProperties().getProperty('Owner'),
             PropertiesService.getScriptProperties().getProperty('Sharing Access'),
             PropertiesService.getScriptProperties().getProperty('Sharing Permission'),
             PropertiesService.getScriptProperties().getProperty('Viewers'),
             PropertiesService.getScriptProperties().getProperty('Editors'),
             PropertiesService.getScriptProperties().getProperty('Parent folder')
              ]);
  
  if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Report - Google Drive Folder Scanner - '+folderName)) 
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Report - Google Drive Folder Scanner - '+folderName);
  else 
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet().setName('Report - Google Drive Folder Scanner - '+folderName);
  
  clearSheet(sheet);
  
  
  
  // Check folder for initial access
  var checkFolder = checkInitialPermission(folderID);
  
  
  // Start recursive function to get folders within folders
  getFolderDataByFolderID(folderID);
  
  // Gets the files within the current folder
  getFileDataByFolderID(folderID,DriveApp.getFolderById(folderID).getName());
   
  // Checks if script reached timeout
  if (reachedTimeOut) SpreadsheetApp.getUi().alert('Was only able to get PARTIAL data for "'+DriveApp.getFolderById(folderID).getName()+'".');
  
  // Appends headers to the sheet
  appendHeaders(curHeader,sheet);
  
  // Outputs data in the sheet (if any)
  if (array.length>0 && array[0].length>0) {
        sheet.getRange(2, 1, array.length, array[0].length).setValues(array);
        sheet.getRange(1, 1, 1, array[0].length).setFontWeight('bold');
  }
  
  
  SpreadsheetApp.getUi().alert('Data gathering has finished SUCCESSFULLY for "'+DriveApp.getFolderById(folderID).getName()+'".');
  
}

/**
 * Void - Gets folder details recursively
 *
 * @param {String} folderID - Value offered recursively depending on the folder level currently analyzing.
 */

function getFolderDataByFolderID(folderID){
  var folder,sharingAccess,sharingPermission,viewers,editors,owner,id,parent;
  
  var currentFolder = DriveApp.getFolderById(folderID);
  
  var childFolderIterator = currentFolder.getFolders();
  
  // Check inside current folder for more folders
  while (childFolderIterator.hasNext()){
    // Get initial data
    folder = childFolderIterator.next();
    id = folder.getId();
    
    // Declare array for current item to be checked
    var smallArray = [];
    
    // Check if values from CONFIG sheet are set to TRUE and fill array
    if (checkHeader(curHeader,'Name')==true) smallArray.push(folder.getName());
    if (checkHeader(curHeader,'Type')==true) smallArray.push('Google Drive folder');
    if (checkHeader(curHeader,'Description')==true) smallArray.push(folder.getDescription());
    if (checkHeader(curHeader,'Id')==true) smallArray.push(folder.getId());
    if (checkHeader(curHeader,'Url')==true) smallArray.push(folder.getUrl());
    if (checkHeader(curHeader,'Date created')==true) smallArray.push(folder.getDateCreated());
    if (checkHeader(curHeader,'Last updated')==true) smallArray.push(folder.getLastUpdated());
    if (checkHeader(curHeader,'Size')==true) smallArray.push(folder.getSize());
    if (checkHeader(curHeader,'Owner')==true) smallArray.push(folder.getOwner().getName());
    
    
    // Use try-catch method as getSharingAccess() and getSharingPermission() methods might fail.
    if (checkHeader(curHeader,'Sharing Access')==true){
      try{
        sharingAccess = folder.getSharingAccess();
        smallArray.push(sharingAccess);
      }
      catch(e){
        smallArray.push('Unable to get sharing Access');
      }
    }
    
    if (checkHeader(curHeader,'Sharing Permission')==true){
      try{
        sharingPermission = folder.getSharingPermission();
        smallArray.push(sharingPermission);
      }
      catch(e){
        smallArray.push('Unable to get sharing Permission');
      }
    }
    
    
    // Add rest of methods if set to TRUE
    if (checkHeader(curHeader,'Viewers')==true) smallArray.push(getUsers(folder.getViewers()));
    if (checkHeader(curHeader,'Editors')==true) smallArray.push(getUsers(folder.getEditors()));
    if (checkHeader(curHeader,'Parent folder')==true) smallArray.push(DriveApp.getFolderById(folderID).getName());
    
    // Push current array to global one
    array.push(smallArray);
    
    
    // Check current execution time and exit out if greather than "timeOut" value
    var now = new Date();
    if ((now.getTime() - start.getTime()) / 1000 > timeOut) {
      
      reachedTimeOut = true;
      return ;
      
    }
    
    // Check files within the current folder
    getFileDataByFolderID(id,folder.getName());
    
    // Go forward recursively
    getFolderDataByFolderID(id);
    
  }
    
}

/**
 * Void - Gets file details from the current folder
 *
 * @param {String} folderID - Value of the current folder ID being analyzed
 * @param {String} folderParent - Value of current folder name being analyzed
 */

function getFileDataByFolderID(folderID,folderParent){
  
  
  
  var file,sharingAccess,sharingPermission,viewers,editors,owner,id;
  
  var fileIterator = DriveApp.getFolderById(folderID).getFiles();
  
  while (fileIterator.hasNext()){
    file = fileIterator.next();
    id = file.getId();
    
    // Declare array for current item to be checked
    var smallArray = [];
    
    // Check if values from CONFIG sheet are set to TRUE and fill array
    if (checkHeader(curHeader,'Name')==true) smallArray.push(file.getName());
    if (checkHeader(curHeader,'Type')==true) smallArray.push(getGoogleMymeType(file.getMimeType()));
    if (checkHeader(curHeader,'Description')==true) smallArray.push(file.getDescription());
    if (checkHeader(curHeader,'Id')==true) smallArray.push(file.getId());
    if (checkHeader(curHeader,'Url')==true) smallArray.push(file.getUrl());
    if (checkHeader(curHeader,'Date created')==true) smallArray.push(file.getDateCreated());
    if (checkHeader(curHeader,'Last updated')==true) smallArray.push(file.getLastUpdated());
    if (checkHeader(curHeader,'Size')==true) smallArray.push(file.getSize());
    if (checkHeader(curHeader,'Owner')==true) smallArray.push(file.getOwner().getName());
    
    // Use try-catch method as getSharingAccess() and getSharingPermission() methods might fail. 
    if (checkHeader(curHeader,'Sharing Access')==true){
      try{
        sharingAccess = file.getSharingAccess();
        smallArray.push(sharingAccess);
      }
      catch(e){
        smallArray.push('Unable to get sharing Access');
      }
    }
    
    if (checkHeader(curHeader,'Sharing Permission')==true){
      try{
        sharingPermission = file.getSharingPermission();
        smallArray.push(sharingPermission);
      }
      catch(e){
        smallArray.push('Unable to get sharing Permission');
      }
    }
  
    // Add rest of methods if set to TRUE
    if (checkHeader(curHeader,'Viewers')==true) smallArray.push(getUsers(file.getViewers()));
    if (checkHeader(curHeader,'Editors')==true) smallArray.push(getUsers(file.getEditors()));
    if (checkHeader(curHeader,'Parent folder')==true) smallArray.push(folderParent);

    // Push current array to global one  
    array.push(smallArray);
  }
  
  
  
   // Check current execution time and exit out if greather than "timeOut" value
  var now = new Date();
  if ((now.getTime() - start.getTime()) / 1000 > timeOut) {
      
      reachedTimeOut = true;
      return ;
      
  }
  
}

/**
* Return {String} - String of email addresses
 *
 * @param {Array} arrayUsers - Gets as parameter an array of object of type "User"
 */

function getUsers(arrayUsers){
  var users='';
  if ((typeof arrayUsers)=='object' && arrayUsers!='undefined' && arrayUsers){
    for (var i in arrayUsers) users = users + arrayUsers[i].getName()+', ';
  }
  else users ='. ';
  
  return users.substring(0,users.length-2);
}
    

/**
 * Void - Sets the header of the spreadsheet according to current values set as true
 *
 * @param {Object} curHeader - Object of type "headerObj" returned by the function getCurrentHeaders()
 * @param {sheet} sheet - Current sheet to be used.
 */

function appendHeaders(curHeader,sheet){
  var smallArray = [];
  if (checkHeader(curHeader,'Name')==true) smallArray.push('Name');
  if (checkHeader(curHeader,'Type')==true) smallArray.push('Type');
  if (checkHeader(curHeader,'Description')==true) smallArray.push('Description');
  if (checkHeader(curHeader,'Id')==true) smallArray.push('Id');
  if (checkHeader(curHeader,'Url')==true) smallArray.push('Url');
  if (checkHeader(curHeader,'Date created')==true) smallArray.push('Date created');
  if (checkHeader(curHeader,'Last updated')==true) smallArray.push('Last updated');
  if (checkHeader(curHeader,'Size')==true) smallArray.push('Size');
  if (checkHeader(curHeader,'Owner')==true) smallArray.push('Owner');
  if (checkHeader(curHeader,'Sharing Access')==true) smallArray.push('Sharing Access');
  if (checkHeader(curHeader,'Sharing Permission')==true) smallArray.push('Sharing Permission');
  if (checkHeader(curHeader,'Viewers')==true) smallArray.push('Viewers');
  if (checkHeader(curHeader,'Editors')==true) smallArray.push('Editors');
  if (checkHeader(curHeader,'Parent folder')==true) smallArray.push('Parent folder');
  
  
  var array = [smallArray];
  if (array.length>0 && array[0].length>0) {
    sheet.getRange(1, 1, array.length, array[0].length).setValues(array);
  }
  
  
}

/**
 * Return {Object} headerObj - Returns an object with property names and values as setup in the CONFIG sheet - used to check what values to be outputted.
 *
 */

function setCurrentHeaders(array){
  
  var headerObj={};
  headerObj['Name']=array[0];
  headerObj['Type']=array[1];
  headerObj['Description']=array[2];
  headerObj['Id']=array[3];
  headerObj['Url']=array[4];
  headerObj['Date created']=array[5];
  headerObj['Last updated']=array[6];
  headerObj['Size']=array[7];
  headerObj['Owner']=array[8];
  headerObj['Sharing Access']=array[9];
  headerObj['Sharing Permission']=array[10];
  headerObj['Viewers']=array[11];
  headerObj['Editors']=array[12];
  headerObj['Parent folder']=array[13];
  curHeader = headerObj;
  return;
}



/**
 * Void - Clears the current sheet of all data
 *
 * @param {sheet} sheet - Current sheet to be used.
 */

function clearSheet(sheet){
  if (sheet.getMaxRows()>1)  sheet.deleteRows(1, sheet.getMaxRows()-1);
  else {
    sheet.getRange("A1:Z1").clear();
  }
}


/**
 * Void - Deletes a certain trigger
 *
 * @param {String} functionName - Function to be removed from the triggering system.
 */

function deleteTrigger(functionName){
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) {
      Logger.log('Deleting trigger');
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * Void - Created a certain trigger
 *
 * @param {String} functionName - Function to be added to the triggering system.
 * @param {Number} timeLimit - Time frame in which the trigger will start (IN SECONDS).
 */

function createTrigger(functionName,timeLimit){
  var curr = new Date();
  var seconds = curr.getSeconds()+timeLimit;
  curr.setSeconds(seconds);
  var builder = ScriptApp.newTrigger(functionName).timeBased();
    builder.at(curr);
    builder.create();
}

/**
 * Return {Boolean} - Function checks the current header property and outputs if it should be included in the report or not
 *
 * @param {Object} headerObj - Object in which to check for properties
 * @param {String} headerProp - Property name to be checked
 */

function checkHeader(headerObj,headerProp){
  if (headerObj[headerProp].toString().toLowerCase()=='true' ||
      headerObj[headerProp]==true) return true;
  else return false;
  
}

/**
 * Return {Boolean} - Function checks the current header property and outputs if it should be included in the report or not
 *
 * @param {String} folderID - FolderID to check if permissions are valid
 */

function checkInitialPermission(folderID){
  try 
    {
      var folder = DriveApp.getFolderById(folderID);
      return true;
    }
  catch(e) {
    return false;
  }
   
}

function getGoogleMymeType(type){
  switch (type){
    case 'application/vnd.google-apps.document':
      return 'Google Docs';
    case 'application/vnd.google-apps.drawing':
      return 'Google Drawing';
    case 'application/vnd.google-apps.file':
      return 'Google Drive file';
    case 'application/vnd.google-apps.form':
      return 'Google Forms';
    case 'application/vnd.google-apps.fusiontable':
      return 'Google Fusion Tables';
    case 'application/vnd.google-apps.presentation':
      return 'Google Slides';
    case 'application/vnd.google-apps.script':
      return 'Google Apps Scripts';
    case 'application/vnd.google-apps.sites':
      return 'Google Sites';
    case 'application/vnd.google-apps.spreadsheet':
      return 'Google Sheets';
    default:
      return type;
  }
}