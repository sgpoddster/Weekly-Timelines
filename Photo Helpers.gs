/*************************************************
 * PODDSTER — PHOTO JPG FOLDER HELPER
 **************************************************/

/**
 * Create the Photo JPG Helper menu
 * Called by the main onOpen() function in the primary script
 */
function createPhotoJpgMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📸 Photo JPG Helper')
    .addItem('Create JPG folders for PHOTO sessions', 'createJpgFoldersFromBoard')
    .addItem('🔄 Convert ARW to JPG (Active Cell Link)', 'convertArwFromActiveCell')
    .addItem('📄 Create HTML from existing JPGs (Link/ID)', 'createHtmlFromExistingJpegs')
    .addSeparator()
    .addItem('📧 Send TEST Email (Active Row)', 'sendTestEmailActiveRow')
    .addItem('📅 Test Calendar Sync', 'testCalendarSync')
    .addItem('❓ Debug Spreadsheet Layout', 'debugSpreadsheetLayout')
    .addToUi();
}

var PHOTO_JPG_SETTINGS = {
  JPG_FOLDER_NAME: 'JPG'
};

var GCP_SETTINGS = {
  // Put the URL from your Cloud Console here:
  URL: 'https://arw-convertor-900391242734.asia-southeast1.run.app',
  // Must match the token in your Python script:
  TOKEN: 'my-secure-poddster-token-123'
};

// Global timer start
var SCRIPT_START_TIME = new Date().getTime();
var MAX_EXECUTION_TIME_MS = 300 * 1000;

function isTimeUp() {
  return (new Date().getTime() - SCRIPT_START_TIME) > MAX_EXECUTION_TIME_MS;
}

/**
 * Convert ARW files from the link in the currently selected cell
 */
function convertArwFromActiveCell() {
  SCRIPT_START_TIME = new Date().getTime();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var range = sheet.getActiveRange();
  var r = range.getRow();
  var c = range.getColumn();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var ui = SpreadsheetApp.getUi();

  var richText = sheet.getRange(1, 1, lastRow, lastCol).getRichTextValues();

  var rootInfo = findRootLinkNearCell_(richText, r - 1, c - 1, lastRow, lastCol);
  if (!rootInfo || !rootInfo.url) {
    ui.alert('❌ No Drive link found near the selected cell.\n\nPlease select a cell near a Drive folder link and try again.');
    return;
  }

  var folderId = extractDriveIdFromUrl_(rootInfo.url);
  if (!folderId) {
    ui.alert('❌ Could not extract folder ID from the link.');
    return;
  }

  try {
    var initialFolder = DriveApp.getFolderById(folderId);
    var rootFolder = navigateUpToSessionRoot_(initialFolder);

    ui.alert('✅ Found session folder: "' + rootFolder.getName() + '"\n\nStarting ARW to JPG conversion...');

    var stats = { found: 0, created: 0, converted: 0, skipped: 0, timeOut: false, foldersScanned: [rootFolder.getId()] };

    Logger.log('=== Manual conversion for: ' + rootFolder.getName() + ' ===');

    processFolderRecursively_(rootFolder, stats);

    var msg = 'Conversion finished for: ' + rootFolder.getName() + '\n\n' +
      'JPG Folders created: ' + stats.created + '\n' +
      'New conversions triggered: ' + stats.converted;

    if (stats.timeOut) {
      msg += '\n\n⚠️ TIME LIMIT REACHED. Run again to finish remaining files.';
    }

    ui.alert(msg);

  } catch (e) {
    ui.alert('❌ Error processing folder:\n\n' + e.message);
    Logger.log('Error in convertArwFromActiveCell: ' + e);
  }
}

/**
 * Manually trigger HTML creation for a specific folder
 */
function createHtmlFromExistingJpegs() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Create HTML Contact Sheet', 'Please paste the Google Drive Folder Link (or ID):', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    var rawInput = response.getResponseText().trim();
    if (!rawInput) {
      ui.alert('No input provided.');
      return;
    }

    var folderId = extractDriveIdFromUrl_(rawInput);
    if (!folderId) {
      folderId = rawInput;
    }

    try {
      var folder = DriveApp.getFolderById(folderId);
      ui.alert('Found folder: "' + folder.getName() + '"\n\nGenerating HTML sheet now...');
      updateHtmlContactSheet(folder);
      ui.alert('✅ Success!\n\n"Selection_Sheet.html" has been created in:\n' + folder.getName());
    } catch (e) {
      ui.alert('❌ Error: Could not access folder.\n\nDetails: ' + e.message);
    }
  }
}

function createJpgFoldersFromBoard() {
  SCRIPT_START_TIME = new Date().getTime();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 1 || lastCol < 1) {
    SpreadsheetApp.getUi().alert('No data on this sheet.');
    return;
  }

  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var richText = sheet.getRange(1, 1, lastRow, lastCol).getRichTextValues();

  var stats = { found: 0, created: 0, converted: 0, skipped: 0, timeOut: false, foldersScanned: [] };

  Logger.log('=== Photo JPG Helper run ===');

  outerLoop:
  for (var r = 0; r < lastRow; r++) {
    for (var c = 0; c < lastCol; c++) {
      if (isTimeUp()) {
        stats.timeOut = true;
        break outerLoop;
      }

      var cellValue = values[r][c];
      if (!cellValue) continue;
      var text = String(cellValue).trim().toUpperCase();

      if (text.indexOf('PHOTO') === -1) continue;

      stats.found++;

      var rootInfo = findRootLinkNearCell_(richText, r, c, lastRow, lastCol);
      if (!rootInfo || !rootInfo.url) {
        stats.skipped++;
        continue;
      }

      var folderId = extractDriveIdFromUrl_(rootInfo.url);
      if (!folderId) {
        stats.skipped++;
        continue;
      }

      try {
        var initialFolder = DriveApp.getFolderById(folderId);
        var rootFolder = navigateUpToSessionRoot_(initialFolder);

        if (stats.foldersScanned.indexOf(rootFolder.getId()) !== -1) {
          continue;
        }
        stats.foldersScanned.push(rootFolder.getId());

        Logger.log('Scanning Session Root: ' + rootFolder.getName());

        processFolderRecursively_(rootFolder, stats);

        if (stats.timeOut) break outerLoop;

      } catch (e) {
        stats.skipped++;
        Logger.log('Error processing row ' + (r + 1) + ': ' + e);
        logErrorToSheet_('Row ' + (r+1), 'Startup', e.toString());
      }
    }
  }

  var msg = 'Helper finished.\n\n' +
    'Sessions found (Rows): ' + stats.found + '\n' +
    'Unique Folders Scanned: ' + stats.foldersScanned.length + '\n' +
    'JPG Folders ensured: ' + stats.created + '\n' +
    'New conversions triggered: ' + stats.converted + '\n' +
    'Skipped Rows (No Link): ' + stats.skipped;

  if (stats.timeOut) {
    msg += '\n\n⚠️ TIME LIMIT REACHED. Run the script again to finish processing remaining files.';
  }

  SpreadsheetApp.getUi().alert(msg);
}

function navigateUpToSessionRoot_(folder) {
  var current = folder;
  var steps = 0;
  while (steps < 3) {
    var name = current.getName();
    if (name.match(/^(EP\s?\d+|Photos?|Cam\s?[A-D]|Audio|Proxy|Footage)$/i)) {
      var parents = current.getParents();
      if (parents.hasNext()) {
        current = parents.next();
        steps++;
      } else {
        break;
      }
    } else {
      break;
    }
  }
  return current;
}

function processFolderRecursively_(folder, stats) {
  if (stats.timeOut || isTimeUp()) {
    stats.timeOut = true;
    return;
  }

  var filesIterator = folder.getFiles();
  var arwFiles = [];
  while (filesIterator.hasNext()) {
    var f = filesIterator.next();
    var fName = f.getName();
    if (fName.toLowerCase().endsWith('.arw') && !fName.startsWith('._')) {
      arwFiles.push(f);
    }
  }

  if (arwFiles.length > 0) {
    var needsProcessing = true;
    var jpgFolder = null;

    var folders = folder.getFoldersByName(PHOTO_JPG_SETTINGS.JPG_FOLDER_NAME);
    while (folders.hasNext()) {
      var potentialFolder = folders.next();
      if (!potentialFolder.isTrashed()) {
        jpgFolder = potentialFolder;
        break;
      }
    }

    if (jpgFolder) {
      var jpgCount = 0;
      var jFiles = jpgFolder.getFiles();
      while (jFiles.hasNext()) {
        var jf = jFiles.next();
        if (!jf.getName().startsWith('._') && (jf.getName().toLowerCase().endsWith('.jpg') || jf.getName().toLowerCase().endsWith('.jpeg'))) {
          jpgCount++;
        }
      }

      if (jpgCount >= arwFiles.length) {
         Logger.log('⏩ Skipping ' + folder.getName() + ' (Already has ' + jpgCount + ' JPGs)');
         needsProcessing = false;
      }
    } else {
      jpgFolder = folder.createFolder(PHOTO_JPG_SETTINGS.JPG_FOLDER_NAME);
      stats.created++;
    }

    if (needsProcessing) {
      Logger.log('📸 Processing ' + folder.getName() + ' (' + arwFiles.length + ' ARWs)');
      var convertedCount = scanAndConvertArwFiles_(arwFiles, jpgFolder, stats, folder.getName());
      stats.converted += convertedCount;

      if (!stats.timeOut) {
        updateHtmlContactSheet(jpgFolder);
      }
    }
  }

  var subFoldersIt = folder.getFolders();
  var subFolders = [];
  while (subFoldersIt.hasNext()) subFolders.push(subFoldersIt.next());

  subFolders.sort(function(a, b) { return a.getName().localeCompare(b.getName()); });

  for (var i = 0; i < subFolders.length; i++) {
    if (stats.timeOut || isTimeUp()) {
      stats.timeOut = true;
      return;
    }

    var subFolder = subFolders[i];
    var subName = subFolder.getName();

    if (subName === PHOTO_JPG_SETTINGS.JPG_FOLDER_NAME) continue;

    if (subName.match(/^(Audio|Cam\s?[A-D]|Proxy|Footage|Project|Render|Export)/i)) {
      continue;
    }

    processFolderRecursively_(subFolder, stats);
  }
}

/**
 * Returns Detailed Status & Retries on Failure
 */
function scanAndConvertArwFiles_(arwFiles, destFolder, stats, sessionName) {
  var triggeredCount = 0;
  var failedFiles = [];

  var existingJpgs = {};
  var jpgFiles = destFolder.getFiles();
  while (jpgFiles.hasNext()) {
    var jName = jpgFiles.next().getName();
    if (!jName.startsWith('._')) {
      existingJpgs[jName] = true;
    }
  }

  // --- PASS 1: Main Attempt ---
  for (var i = 0; i < arwFiles.length; i++) {
    if (isTimeUp()) {
      stats.timeOut = true;
      break;
    }

    var file = arwFiles[i];
    var name = file.getName();

    if (name.startsWith('._')) continue;

    var baseName = name.substring(0, name.lastIndexOf('.'));
    var expectedJpgName = baseName + ".jpg";

    if (existingJpgs[expectedJpgName]) continue;

    Logger.log('      >> Sending conversion request for: ' + name);

    var result = callConversionCloudFunction_(file.getId(), destFolder.getId(), name);

    if (result.success) {
      triggeredCount++;
    } else {
      Logger.log('      ⚠️ Attempt 1 Failed. Queueing for immediate retry.');
      logErrorToSheet_(sessionName, name, "Attempt 1: " + result.error);
      failedFiles.push(file);
    }
  }

  // --- PASS 2: Immediate Retry (Warm Server) ---
  if (failedFiles.length > 0 && !stats.timeOut) {
    Logger.log('🔄 RETRYING ' + failedFiles.length + ' failed files now...');
    Utilities.sleep(2000);

    for (var k = 0; k < failedFiles.length; k++) {
      if (isTimeUp()) {
        stats.timeOut = true;
        break;
      }

      var rFile = failedFiles[k];
      var rName = rFile.getName();
      Logger.log('      >> Retry Request for: ' + rName);

      var rResult = callConversionCloudFunction_(rFile.getId(), destFolder.getId(), rName);

      if (rResult.success) {
        Logger.log('      ✅ Retry Successful!');
        triggeredCount++;
      } else {
        Logger.log('      ❌ Retry Failed. Logging error.');
        logErrorToSheet_(sessionName, rName, "FINAL FAILURE: " + rResult.error);
      }
    }
  }

  return triggeredCount;
}

/**
 * Logs errors to a tab named "⚠️ Error Log"
 */
function logErrorToSheet_(session, file, errorMsg) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "⚠️ Error Log";
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(["Timestamp", "Session / Folder", "File Name", "Error Message"]);
      sheet.setFrozenRows(1);
      sheet.getRange("1:1").setFontWeight("bold");
    }

    sheet.appendRow([new Date(), session, file, errorMsg]);
  } catch(e) {
    Logger.log("Failed to log error to sheet: " + e);
  }
}

/**
 * Returns Object {success: boolean, error: string}
 */
function callConversionCloudFunction_(fileId, destFolderId, filename) {
  if (GCP_SETTINGS.URL.indexOf('YOUR_PROJECT') !== -1) {
    return { success: false, error: 'Cloud Function URL not configured' };
  }

  var payload = {
    token: GCP_SETTINGS.TOKEN,
    file_id: fileId,
    dest_folder_id: destFolderId,
    filename: filename
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(GCP_SETTINGS.URL, options);
    var code = response.getResponseCode();

    if (code == 200) {
      Logger.log('      ✅ Success');
      return { success: true, error: null };
    } else {
      var err = response.getContentText();
      Logger.log('      ❌ Error (' + code + '): ' + err);
      return { success: false, error: err };
    }
  } catch (e) {
    Logger.log('      ❌ Network Error: ' + e);
    return { success: false, error: e.toString() };
  }
}

function updateHtmlContactSheet(jpgFolder) {
  var files = jpgFolder.getFiles();
  var fileList = [];
  while (files.hasNext()) {
    var file = files.next();
    var name = file.getName();
    if (name.startsWith('._')) continue;
    if (name.toLowerCase().endsWith('.jpg') || name.toLowerCase().endsWith('.jpeg')) {
      fileList.push(file);
    }
  }
  fileList.sort(function(a, b) { return a.getName().localeCompare(b.getName()); });
  if (fileList.length === 0) return;

  var sessionName = "Session";
  try {
    var photosFolder = jpgFolder.getParents().next();
    var parentName = photosFolder.getName();
    if (parentName.match(/^(Photos?|Cam\s?[A-D])$/i)) {
       photosFolder = photosFolder.getParents().next();
       parentName = photosFolder.getName();
    }
    if (parentName.match(/^EP\s?\d+/i)) {
       var grandParent = photosFolder.getParents().next();
       sessionName = grandParent.getName() + " - " + parentName;
    } else {
       sessionName = parentName;
    }
  } catch (e) {
    sessionName = jpgFolder.getName();
  }

  Logger.log('Generating HTML for ' + sessionName + ' (' + fileList.length + ' images)');

  var fontStack = "font-family: 'Lato', Helvetica, Arial, sans-serif;";
  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>' +
    '@import url("https://fonts.googleapis.com/css2?family=Lato:wght@400;700&display=swap");' +
    'body { margin: 0; padding: 20px; background: #e0e0e0; text-align: center; ' + fontStack + ' }' +
    '.container { background: white; padding: 40px 20px; border-radius: 8px; display: inline-block; max-width: 1040px; width: 100%; box-sizing: border-box; }' +
    'h2 { color: #333; margin-top: 0; margin-bottom: 10px; }' +
    'p.instructions { color: #666; font-size: 15px; margin-bottom: 40px; line-height: 1.5; }' +
    '.photo-block { margin-bottom: 40px; border-bottom: 1px solid #eee; padding-bottom: 20px; }' +
    '.filename { color: #444; font-family: monospace; font-size: 18px; margin-bottom: 10px; font-weight: bold; }' +
    'img { width: 100%; height: auto; border-radius: 4px; display: block; }' +
    '.footer { color: #888; font-size: 12px; margin-top: 20px; }' +
    '</style></head><body>';

  html += '<div class="container">';
  html += '<h2>Poddster Session Photos: ' + sessionName + '</h2>';
  html += '<p class="instructions">' +
          'Here are the photographs from your session, please pick one and reply to this email with the file name and we\'ll edit and get it back to you! (Click any image to inlarge)' +
          '</p>';

  for (var i = 0; i < fileList.length; i++) {
    var file = fileList[i];
    var fname = file.getName();
    var blob = file.getBlob();
    var b64 = Utilities.base64Encode(blob.getBytes());
    var type = blob.getContentType();
    var dataUrl = 'data:' + type + ';base64,' + b64;

    html += '<div class="photo-block">';
    html +=   '<div class="filename">' + fname + '</div>';
    html +=   '<img src="' + dataUrl + '" alt="' + fname + '">';
    html += '</div>';
  }

  html += '<div class="footer">Generated by Poddster Automation</div>';
  html += '</div></body></html>';

  var htmlFileName = 'Selection_Sheet.html';
  var existing = jpgFolder.getFilesByName(htmlFileName);
  while (existing.hasNext()) {
    existing.next().setTrashed(true);
  }
  jpgFolder.createFile(htmlFileName, html, MimeType.HTML);
}

function findRootLinkNearCell_(richText, rowIndex, colIndex, lastRow, lastCol) {
  var maxRowsUp = 12;
  var maxColOffset = 3;
  var startRow = Math.max(0, rowIndex - maxRowsUp);
  var startCol = Math.max(0, colIndex - maxColOffset);
  var endCol = Math.min(lastCol - 1, colIndex + maxColOffset);

  for (var rr = rowIndex - 1; rr >= startRow; rr--) {
    for (var cc = startCol; cc <= endCol; cc++) {
      var rt = richText[rr][cc];
      var url = getLinkUrlFromRichText_(rt);
      if (url) return { url: url, row: rr, col: cc };
    }
  }
  return null;
}

function getLinkUrlFromRichText_(rt) {
  if (!rt) return null;
  var direct = rt.getLinkUrl && rt.getLinkUrl();
  if (direct) return direct;
  var runs = rt.getRuns ? rt.getRuns() : [];
  for (var i = 0; i < runs.length; i++) {
    var url = runs[i].getLinkUrl();
    if (url) return url;
  }
  return null;
}

function extractDriveIdFromUrl_(url) {
  if (!url) return null;
  var m = url.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (m && m[1]) return m[1];
  m = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m && m[1]) return m[1];
  m = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m && m[1]) return m[1];
  return null;
}

/**
 * DEBUG TOOL
 */
function debugSpreadsheetLayout() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  SpreadsheetApp.getUi().alert("Sheet: " + sheet.getName() + "\nLast Row: " + sheet.getLastRow());
}

function sendTestEmailActiveRow() {
  var recipient = "ben@poddster.com";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var range = sheet.getActiveRange();
  var r = range.getRow();
  var c = range.getColumn();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  var richText = sheet.getRange(1, 1, lastRow, lastCol).getRichTextValues();
  Logger.log('Preparing inline email for Row ' + r);

  var rootInfo = findRootLinkNearCell_(richText, r, c, lastRow, lastCol);
  if (!rootInfo || !rootInfo.url) {
    SpreadsheetApp.getUi().alert('Could not find a Drive link near this cell.');
    return;
  }

  var folderId = extractDriveIdFromUrl_(rootInfo.url);
  try {
    var rootFolder = DriveApp.getFolderById(folderId);
    var realRoot = navigateUpToSessionRoot_(rootFolder);
    var jpgFolder = null;

    var folderStack = [realRoot];
    while(folderStack.length > 0) {
      var curr = folderStack.shift();
      if (curr.getName() === PHOTO_JPG_SETTINGS.JPG_FOLDER_NAME) {
        jpgFolder = curr;
        break;
      }
      var sub = curr.getFolders();
      while(sub.hasNext()) folderStack.push(sub.next());
    }

    if (!jpgFolder) {
        SpreadsheetApp.getUi().alert('No "JPG" folder found in this session tree.');
        return;
    }

    var files = jpgFolder.getFiles();
    var inlineImages = {};
    var fileList = [];

    while (files.hasNext()) {
      var file = files.next();
      var name = file.getName();
      if (name.toLowerCase().endsWith('.jpg') || name.toLowerCase().endsWith('.jpeg')) {
        fileList.push(file);
      }
    }
    fileList.sort(function(a, b) { return a.getName().localeCompare(b.getName()); });

    if (fileList.length === 0) {
      SpreadsheetApp.getUi().alert('No JPGs found to embed.');
      return;
    }

    var fontStack = "font-family: 'Lato', Helvetica, Arial, sans-serif;";
    var htmlBody = '<div style="' + fontStack + ' background: #e0e0e0; padding: 20px; text-align: center;">';
    htmlBody += '<div style="background: white; padding: 40px 20px; border-radius: 8px; display: inline-block; max-width: 1040px;">';
    htmlBody += '<h2 style="color: #333; margin-top: 0; margin-bottom: 10px;">Poddster Session Photos: ' + realRoot.getName() + '</h2>';
    htmlBody += '<p style="color: #666; font-size: 15px; margin-bottom: 40px; line-height: 1.5;">' +
                'Hi there! Here are the photographs from your session, please pick one and reply to this email with the file name and we will edit it and get it back to you!' +
                '</p>';

    for (var i = 0; i < fileList.length; i++) {
      var file = fileList[i];
      var name = file.getName();
      if (name.startsWith('._')) continue;

      inlineImages[name] = file.getBlob();
      htmlBody += '<div style="margin-bottom: 40px; border-bottom: 1px solid #eee; padding-bottom: 20px;">';
      htmlBody +=   '<h3 style="color: #444; font-family: monospace; font-size: 18px; margin-bottom: 10px;">' + name + '</h3>';
      htmlBody +=   '<img src="cid:' + name + '" style="width: 100%; height: auto; border-radius: 4px;" />';
      htmlBody += '</div>';
    }
    htmlBody += '<p style="color: #888; font-size: 12px; margin-top: 20px;">Generated by Poddster Automation</p>';
    htmlBody += '</div></div>';

    var subject = "Photo Selection: " + realRoot.getName();

    GmailApp.sendEmail(recipient, subject, "Please view this email in a generic HTML client.", {
      htmlBody: htmlBody,
      inlineImages: inlineImages,
      name: 'Poddster Automation'
    });

    SpreadsheetApp.getUi().alert('✅ Inline Email sent to ' + recipient);

  } catch (e) {
    SpreadsheetApp.getUi().alert('Error: ' + e.message);
    Logger.log(e);
  }
}
