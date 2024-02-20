/* 
get specific sheet by ID

to get sheet in the active spreadsheet
{*Library*}.getSheetById(sheetID)

to get sheet in another spreadsheet
{*Library*}.getSheetById(sheetID,spreadSheetID)
*/
function getSheetById(sheetID, spreadsheetID) {
  if (typeof spreadsheetID == "undefined") {
    return SpreadsheetApp.getActiveSpreadsheet().getSheets()
      .filter(function (s) { return s.getSheetId() === sheetID; })[0]
  }
  else {
    return SpreadsheetApp.openById(spreadsheetID).getSheets()
      .filter(function (s) { return s.getSheetId() === sheetID; })[0]
  }
}


/* 
Get all SpreadSheets from a drive folder using folder ID 
*/
function getSpreadsheetsFromDriveFolder(folderID) {
  var folder = DriveApp.getFolderById(folderID)
  Logger.log("Folder name is : " + folder.getName())

  files = folder.getFilesByType('application/vnd.google-apps.spreadsheet')

  var spreadsheets = []
  while (files.hasNext()) {
    var file = files.next()

    var properties = []
    properties.push(file.getName())
    properties.push(file.getId())
    properties.push(file.getUrl())

    spreadsheets.push(properties)
  }
  return spreadsheets;
}

/*
  give permission for a spreadsheet to access other spreadsheets links typed in a specified range
  
  Credit : This code is from this guy here 
  https://www.reddit.com/r/GoogleAppsScript/comments/rg34w0/anyone_have_a_script_to_auto_allow_links_in/ 
*/
function addImportrangePermission(spreadsheetID, sheetName, range) {
  const token = ScriptApp.getOAuthToken();

  var urls = SpreadsheetApp.openById(spreadsheetID).getSheetByName(sheetName)
    .getRange(range)
    .getValues()
    .flat()
    .filter(url => url != "")
    .map(url => {
      Logger.log(url)
      var id = /d\/(.*?)\//.exec(url)[1];
      var params = {
        url: `https://docs.google.com/spreadsheets/d/${spreadsheetID}/externaldata/addimportrangepermissions?donorDocId=${id}`,
        method: 'post',
        headers: {
          Authorization: 'Bearer ' + token,
        },
        muteHttpExceptions: true
      };

      return params;
    })

  UrlFetchApp.fetchAll(urls);
}


// عمل dropdownlist 
function dropDownList(list, sheet, cell) {
  var dropDownList = SpreadsheetApp.newDataValidation().requireValueInList(list).build()
  sheet.getRange(cell).setDataValidation(dropDownList)
}


/* 
  Open a website in a new tab 
*/
function openURL(url) {
  var html = HtmlService.createHtmlOutput('<html><script>'
    + 'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};'
    + 'var a = document.createElement("a"); a.href="' + url + '"; a.target="_blank";'
    + 'if(document.createEvent){'
    + '  var event=document.createEvent("MouseEvents");'
    + '  if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'
    + '  event.initEvent("click",true,true); a.dispatchEvent(event);'
    + '}else{ a.click() }'
    + 'close();'
    + '</script>'
    // Offer URL as clickable link in case above code fails.
    + '<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="' + url + '" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>'
    + '<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script>'
    + '</html>')
    .setWidth(90).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(html, "Opening ...");
}


/*
 Get all links in a website
*/
function getAllLinksFromExternalWebsite(url, uniqueValues = 0) {

  var response = UrlFetchApp.fetch(url);
  var htmlContent = response.getContentText();

  var links = htmlContent.match(/<a[^>]+href="([^"]+)"/g).filter(i => i.includes('https'))
  if (uniqueValues != 0) links = [...new Set(links)];    // if you want only unique values 
  // Logger.log(links)

  for (var i = 0; i < links.length; i++) {
    var pos = links[i].indexOf('https')
    links[i] = links[i].substring(pos, links[i].indexOf('"', pos))
  }
  // Logger.log(links)
  return links;
}





