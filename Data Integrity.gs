/***************************************************************
* PODDSTER — SESSION DATA INTEGRITY CHECKER (ONE FILE)
* - Robust merged-header handling (scans left)
* - UK/SG date format support (DD/MM/YYYY)
* - Popup Integrity Summary (filter: all / issues)
* - Email version includes "Only rows with issues" + "All rows"
* - Remote handling:
*    • If "remote" appears in session block: expected camera/audio counts reduced by 1
*    • Remote files? checks for top-level folder containing "record" OR "remote"
*    • If NOT remote: Remote files? shows "No (not remote)" and is treated as ✅ good
* - LIVE handling:
*    • If a camera set path contains "live" (case-insensitive), skip size-variation comparison
* - Ignore macOS sidecar files starting with "._" (audio + video + duplicates)
* - NEW: Timestamp drift checks
*    • Video timestamps OK? checks within each camera folder (direct files)
*    • Audio timestamps OK? checks within each take folder (or Audio root if no takes)
*    • Same type should be created within 5 minutes (Drive file created time)
***************************************************************/


/** Master onOpen for Poddster workbook */




/**************** TIMESTAMP DRIFT CONFIG + HELPERS ****************/
var TIMESTAMP_MAX_DRIFT_MINUTES = 5;
var TIMESTAMP_MAX_DRIFT_MS = TIMESTAMP_MAX_DRIFT_MINUTES * 60 * 1000;


function diffMinutes_(ms) {
 return ms / 60000;
}


/**
* Compute creation-time drift for relevant files directly inside a folder.
* - relevantRe: regex for file extension/type
* - ignores macOS "._" sidecars
* - returns { ok:boolean, count:number, min:Date|null, max:Date|null, driftMs:number }
*/
function folderCreatedTimeDrift_(folder, relevantRe) {
 var min = null;
 var max = null;
 var count = 0;


 var it = folder.getFiles();
 while (it.hasNext()) {
   var f = it.next();
   var name = f.getName();
   if (name.indexOf('._') === 0) continue;


   var lower = name.toLowerCase();
   if (!relevantRe.test(lower)) continue;


   var created = f.getLastUpdated();
   if (!created) continue;


   count++;
   if (!min || created < min) min = created;
   if (!max || created > max) max = created;
 }


 if (count <= 1) {
   return { ok: true, count: count, min: min, max: max, driftMs: 0 };
 }


 var driftMs = max.getTime() - min.getTime();
 return { ok: driftMs <= TIMESTAMP_MAX_DRIFT_MS, count: count, min: min, max: max, driftMs: driftMs };
}


function fmtDrift_(driftMs) {
 return diffMinutes_(driftMs).toFixed(1) + ' min';
}




/*************** DATE HELPERS ****************/
// NOTE: pad2_ is defined in Code Scheduling.gs to avoid duplicate declaration

function dateToKey_(d) {
 return d.getFullYear() + '-' + pad2_(d.getMonth() + 1) + '-' + pad2_(d.getDate());
}


/**
* Find the latest date key (yyyy-MM-dd) present on THIS sheet that is < todayKey.
* Returns null if none found.
*/
function findLatestSessionDateKeyOnSheet_(sheet) {
 var lastRow = sheet.getLastRow();
 var lastCol = sheet.getLastColumn();
 if (lastRow < 2 || lastCol < 2) return null;


 var range    = sheet.getRange(1, 1, lastRow, lastCol);
 var grid     = range.getDisplayValues();
 var rich     = range.getRichTextValues();
 var formulas = range.getFormulas();


 var targetCols = [];
 for (var c = 0; c < lastCol; c++) targetCols.push(c);


 var sessions = extractSessions_(grid, rich, formulas, targetCols);


 var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
 var todayKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');


 var keys = [];
 sessions.forEach(function (s) {
   var k = headerTextToKey_(s.dateDisplay);
   if (k && k < todayKey && keys.indexOf(k) === -1) keys.push(k);
 });


 if (!keys.length) return null;
 keys.sort(); // yyyy-MM-dd sorts correctly as strings
 return keys[keys.length - 1];
}


/**
* Header parser:
* 1. "Monday, 24 November 2025"
* 2. "24/11/2025" (UK/SG)
* 3. "24 November" (adds current year)
*/
function headerTextToKey_(txt) {
 if (!txt) return null;
 txt = String(txt).trim();
 if (!txt) return null;


 // 1) DD/MM/YYYY or DD.MM.YYYY
 var dmy = txt.match(/^(\d{1,2})[\/\.](\d{1,2})[\/\.](\d{4})$/);
 if (dmy) {
   var dObj = new Date(dmy[3], dmy[2] - 1, dmy[1]);
   return dateToKey_(dObj);
 }


 // 2) Text format: strip weekday, commas
 var cleanTxt = txt.replace(/^[A-Za-z]+,?\s*/, '');
 if (!/\d{4}/.test(cleanTxt)) {
   var currentYear = new Date().getFullYear();
   cleanTxt = cleanTxt + ' ' + currentYear;
 }


 var d = new Date(cleanTxt);
 if (isNaN(d.getTime())) return null;


 return dateToKey_(d);
}




/*************** RUNNERS ***************/
function runSanityToday() {
 runSanityForSheetDate_(SpreadsheetApp.getActiveSheet(), "Today's Data", 0);
}


function runSanityYesterday() {
 // Try real "yesterday" first
 var sheet = SpreadsheetApp.getActiveSheet();


 var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
 var now = new Date();
 var yTime = new Date(now.getTime() - 1 * 24 * 60 * 60 * 1000);
 var yKey  = Utilities.formatDate(yTime, tz, 'yyyy-MM-dd');


 // Pull sessions for yesterday; if none, fall back to latest available date on THIS sheet
 var lastRow = sheet.getLastRow();
 var lastCol = sheet.getLastColumn();
 if (lastRow < 2 || lastCol < 2) {
   SpreadsheetApp.getUi().alert('Sheet looks empty.');
   return;
 }


 var range    = sheet.getRange(1, 1, lastRow, lastCol);
 var grid     = range.getDisplayValues();
 var rich     = range.getRichTextValues();
 var formulas = range.getFormulas();


 var targetCols = [];
 for (var c = 0; c < lastCol; c++) targetCols.push(c);


 var allSessions = extractSessions_(grid, rich, formulas, targetCols);


 function sessionsForKey_(key) {
   return allSessions.filter(function (sess) {
     var k = headerTextToKey_(sess.dateDisplay);
     return k && k === key;
   });
 }


 var sessions = sessionsForKey_(yKey);


 var label;
 var targetKey;
 if (!sessions.length) {
   var latest = findLatestSessionDateKeyOnSheet_(sheet);
   if (!latest) {
     SpreadsheetApp.getUi().alert(
       '❌ No sessions found for Yesterday’s Data.\n\n' +
       '• Script looked for: ' + yKey + '\n' +
       '• Found in sheet: None'
     );
     return;
   }
   sessions  = sessionsForKey_(latest);
   targetKey = latest;
   label     = "Latest Available Data";
 } else {
   targetKey = yKey;
   label     = "Yesterday's Data";
 }


 var results = sessions.map(processSession_);
 var header  = sanityHeader_();
 writeSanitySheet_(header, results);
 showSanityPopup_(results, header, 'Integrity Summary — ' + label + ' (' + targetKey + ')');
}


function runSanityWholeTab() {
 runSanityForSheet_(SpreadsheetApp.getActiveSheet(), 'Entire Week');
}


/**
* Date-filtered run (used by Today)
*/
function runSanityForSheetDate_(sheet, label, offsetDays) {
 var lastRow = sheet.getLastRow();
 var lastCol = sheet.getLastColumn();
 if (lastRow < 2 || lastCol < 2) {
   SpreadsheetApp.getUi().alert('Sheet looks empty.');
   return;
 }


 var range    = sheet.getRange(1, 1, lastRow, lastCol);
 var grid     = range.getDisplayValues();
 var rich     = range.getRichTextValues();
 var formulas = range.getFormulas();


 var targetCols = [];
 for (var c = 0; c < lastCol; c++) targetCols.push(c);


 var allSessions = extractSessions_(grid, rich, formulas, targetCols);


 var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
 var now = new Date();
 var targetTime = new Date(now.getTime() + offsetDays * 24 * 60 * 60 * 1000);
 var targetKey = Utilities.formatDate(targetTime, tz, 'yyyy-MM-dd');


 var sessions = allSessions.filter(function (sess) {
   var k = headerTextToKey_(sess.dateDisplay);
   return k && k === targetKey;
 });


 if (!sessions.length) {
   var foundKeys = [];
   allSessions.forEach(function (s) {
     var k = headerTextToKey_(s.dateDisplay);
     if (k && foundKeys.indexOf(k) === -1) foundKeys.push(k);
   });


   SpreadsheetApp.getUi().alert(
     '❌ No sessions found for ' + label + '.\n\n' +
     '• Script looked for: ' + targetKey + '\n' +
     '• Found in sheet: ' + (foundKeys.join(', ') || 'None')
   );
   return;
 }


 var results = sessions.map(processSession_);
 var header  = sanityHeader_();
 writeSanitySheet_(header, results);
 showSanityPopup_(results, header, 'Integrity Summary — ' + label + ' (' + targetKey + ')');
}


/**
* Whole-tab run (no date filter)
*/
function runSanityForSheet_(sheet, label) {
 var lastRow = sheet.getLastRow();
 var lastCol = sheet.getLastColumn();
 if (lastRow < 2 || lastCol < 2) {
   SpreadsheetApp.getUi().alert('Sheet looks empty.');
   return;
 }


 var range    = sheet.getRange(1, 1, lastRow, lastCol);
 var grid     = range.getDisplayValues();
 var rich     = range.getRichTextValues();
 var formulas = range.getFormulas();


 var targetCols = [];
 for (var c = 0; c < lastCol; c++) targetCols.push(c);


 var sessions = extractSessions_(grid, rich, formulas, targetCols);
 if (!sessions.length) {
   SpreadsheetApp.getUi().alert('No sessions found for ' + label + '.');
   return;
 }


 var results = sessions.map(processSession_);
 var header  = sanityHeader_();
 writeSanitySheet_(header, results);
 showSanityPopup_(results, header, 'Integrity Summary — ' + label);
}




/*************** SESSION EXTRACTION (MERGED HEADERS) ***************/
function extractSessions_(grid, rich, formulas, targetCols) {
 var sessions = [];
 var rows = grid.length;
 var timeRe = /\b\d{1,2}:\d{2}\s*[-–]\s*\d{1,2}:\d{2}\b/; // 11:00 - 12:00


 for (var r = 1; r < rows; r++) {
   for (var i = 0; i < targetCols.length; i++) {
     var c    = targetCols[i];
     var cell = grid[r][c] || '';
     if (!timeRe.test(cell)) continue;


     var clientData = getClientDeepScan_(grid, rich, formulas, r, c);
     if (!clientData.text || !clientData.url) continue;


     var timeMatch   = cell.match(timeRe);
     var time        = timeMatch ? timeMatch[0] : '';


     // MERGED HEADER FIX: look left up to 5 columns
     var dateDisplay = '';
     if (grid[1]) {
       for (var k = 0; k <= 5; k++) {
         if (c - k < 0) break;
         var val = grid[1][c - k];
         if (val && String(val).trim() !== '') {
           dateDisplay = val;
           break;
         }
       }
     }


     var room   = safeRoom_(grid, r + 1, c);
     var seats  = safeSeats_(grid, r + 1, c);
     var remote = detectRemoteFlag_(grid, r, c);


     sessions.push({
       dateDisplay: dateDisplay,
       time:        time,
       room:        room,
       seats:       seats,
       remote:      remote,
       clientText:  clientData.text,
       clientUrl:   clientData.url,
       driveId:     parseDriveId_(clientData.url)
     });
   }
 }
 return sessions;
}


// Grab raw client text + link from the left cell
function getClientDeepScan_(grid, rich, formulas, r, c) {
 var targetC = c - 1;
 if (targetC < 0) return { text: '', url: '' };


 var text = (grid[r][targetC] || '').trim();
 var url  = '';


 var rt = rich[r][targetC];


 // 1) Simple link
 if (rt && rt.getLinkUrl && rt.getLinkUrl()) {
   url = rt.getLinkUrl();
 }


 // 2) Multi-run rich text (smart chips etc.)
 if (!url && rt && rt.getRuns) {
   var runs = rt.getRuns();
   for (var i = 0; i < runs.length; i++) {
     if (runs[i].getLinkUrl && runs[i].getLinkUrl()) {
       url = runs[i].getLinkUrl();
       break;
     }
   }
 }


 // 3) HYPERLINK() formula
 if (!url) {
   var form = formulas[r][targetC];
   if (form && /HYPERLINK/i.test(form)) {
     var m = form.match(/HYPERLINK\s*\(\s*"([^"]+)"/i);
     if (m && m[1]) url = m[1];
   }
 }


 return { text: text, url: url };
}


function safeRoom_(grid, r, c) {
 if (r >= grid.length) return '';
 for (var dc = -2; dc <= 2; dc++) {
   var col = c + dc;
   if (col < 0 || col >= grid[0].length) continue;
   var val = (grid[r][col] || '').trim();
   if (val && isNaN(val)) return val;
 }
 return '';
}


function safeSeats_(grid, r, c) {
 if (r >= grid.length) return 0;
 for (var dc = -2; dc <= 2; dc++) {
   var col = c + dc;
   if (col < 0 || col >= grid[0].length) continue;
   var val = (grid[r][col] || '').trim();
   if (/^\d+$/.test(val)) return Number(val);
 }
 return 0;
}


/**
* Detect "remote" in the session block:
* scan a few rows below the time cell, around the same columns,
* for any cell containing the word "remote".
*/
function detectRemoteFlag_(grid, timeRow, timeCol) {
 var rows = grid.length;
 var cols = grid[0].length;


 for (var r = timeRow + 1; r < Math.min(rows, timeRow + 6); r++) {
   for (var dc = -2; dc <= 2; dc++) {
     var c = timeCol + dc;
     if (c < 0 || c >= cols) continue;
     var val = (grid[r][c] || '').toString().toLowerCase();
     if (val.indexOf('remote') !== -1) {
       return true;
     }
   }
 }
 return false;
}




/*************** PROCESS ONE SESSION ***************/
function processSession_(sess) {
 var displayName = sess.clientText;


 // Enrich from Drive hierarchy if we have an ID
 if (sess.driveId) {
   var details = getDriveDetails_(sess.driveId);
   if (details.realName) displayName = details.realName;
 }


 var rootObj = probeDriveSafe_(sess.driveId);
 var root    = rootObj.folder || null;


 var video = checkVideo_(root, sess.seats, sess.remote);
 var audio = checkAudio_(root, sess.seats, sess.remote);
 var dups  = checkDuplicateCameraFiles_(root);


 var notes = [];


 // --- REMOTE NOTES ---
 if (sess.remote) {
   if (video.remoteFilesStatus === 'Yes') {
     notes.push('Remote: recording folder found');
   } else if (video.remoteFilesStatus === 'No') {
     notes.push('Remote: recording folder missing');
   } else {
     notes.push('Remote: recording status unknown');
   }
 } else {
   notes.push('Remote: not expected');
 }


 // --- AUDIO NOTES ---
 if (audio.notes && audio.notes.length) {
   notes = notes.concat(audio.notes);
 }


 // --- VIDEO NOTES ---
 if (video.smallFiles && video.smallFileList && video.smallFileList.length) {
   notes.push('Small video files (<100 MB): ' + video.smallFileList.join('; '));
 }


 if (video.cameraCountIssues && video.cameraCountIssues.length) {
   notes.push('Camera count issues: ' + video.cameraCountIssues.join(' | '));
 }


 if (video.variationDetails && video.variationDetails.length) {
   notes.push('Video size variation by part: ' + video.variationDetails.join(' | '));
 } else if (video.variation && video.debug) {
   notes.push('Video size variation – totals: ' + video.debug.replace(/^Video totals –\s*/, ''));
 }


 // --- TIMESTAMP NOTES ---
 if (!video.videoTimestampsOK && video.videoTimestampIssues.length) {
   notes.push('Video timestamp drift (> ' + TIMESTAMP_MAX_DRIFT_MINUTES + ' min): ' + video.videoTimestampIssues.join(' | '));
 }
 if (!audio.audioTimestampsOK && audio.audioTimestampIssues.length) {
   notes.push('Audio timestamp drift (> ' + TIMESTAMP_MAX_DRIFT_MINUTES + ' min): ' + audio.audioTimestampIssues.join(' | '));
 }


 if (video.debug) {
   notes.push(video.debug);
 }


 return [
   sess.dateDisplay,                      // 0
   sess.time,                             // 1
   sess.room,                             // 2
   displayName,                           // 3
   sess.clientUrl,                        // 4
   video.emptyFolders,                    // 5
   video.camerasPresent,                  // 6
   video.cameraCountOK,                   // 7
   dups ? 'Yes' : 'No',                   // 8  Duplicate camera files?
   video.smallFiles ? 'Yes' : 'No',       // 9  Small video files?
   video.variation ? 'Yes' : 'No',        // 10 File size variation?
   audio.expected,                        // 11 Expected Audio Files
   audio.actual,                          // 12 Actual audio files
   audio.matchOK ? 'Yes' : 'No',          // 13 Files Match?
   audio.stereoOK ? 'Yes' : 'No',         // 14 Stereo Mix OK?
   video.remoteFilesStatus || '',         // 15 Remote files?
   video.videoTimestampsOK ? 'Yes' : 'No',// 16 Video timestamps OK?
   audio.audioTimestampsOK ? 'Yes' : 'No',// 17 Audio timestamps OK?
   notes.join(' | ')                      // 18 Notes
 ];
}




/*************** DRIVE HELPERS ***************/
function getDriveDetails_(id) {
 try {
   var f = DriveApp.getFolderById(id);
   var current = f;
   var clientName = '';


   var clientPattern = / \([A-Za-z0-9]{3}\)$/;  // e.g. "Kevin Folonier (H8Q)"
   var banned = {
     'drive': true,
     'my drive': true,
     'poddster client raw files': true
   };


   for (var hops = 0; hops < 6; hops++) {
     var parents = current.getParents();
     if (!parents.hasNext()) break;
     var p = parents.next();
     var name = p.getName();
     var lower = name.toLowerCase();


     if (clientPattern.test(name) && !banned[lower]) {
       clientName = name;
       break;
     }


     current = p;
   }


   return { realName: clientName, folder: f };
 } catch (e) {
   return {};
 }
}


function probeDriveSafe_(id) {
 if (!id) return {};
 try { return { folder: DriveApp.getFolderById(id) }; } catch (e) {}
 try { return { file: DriveApp.getFileById(id) }; } catch (e2) {}
 return {};
}


function parseDriveId_(url) {
 if (!url) return '';
 var m = url.match(/[-\w]{25,}/);
 return m ? m[0] : '';
}




/*************** VIDEO / CAMERA CHECKS ***************/
function checkVideo_(root, seats, isRemote) {
 var res = {
   emptyFolders:        0,
   camerasPresent:      '',
   cameraCountOK:       'Yes',
   smallFiles:          false,
   variation:           false,
   debug:               '',
   smallFileList:       [],
   cameraCountIssues:   [],
   variationDetails:    [],
   remoteFilesStatus:   '',
   videoTimestampsOK:   true,   // NEW
   videoTimestampIssues:[]      // NEW
 };
 if (!root) {
   res.cameraCountOK = 'No';
   res.videoTimestampsOK = false;
   return res;
 }


 // --- Remote Recording folder check (top level) ---
 if (isRemote) {
   var top = root.getFolders();
   var hasRemoteFolder = false;
   while (top.hasNext()) {
     var tf  = top.next();
     var n   = tf.getName().toLowerCase();
    if (
 n.indexOf('record') !== -1 ||
 n.indexOf('remote') !== -1 ||
 n.indexOf('riverside') !== -1
) {
 hasRemoteFolder = true;
 break;
}
   }
   res.remoteFilesStatus = hasRemoteFolder ? 'Yes' : 'No';
 } else {
   // Not a remote session – we DON'T expect a remote-recording folder
   res.remoteFilesStatus = 'No (not remote)'; // ✅ good (handled in popup/email)
 }


 var cams = [];
 var empty = 0;


 var overallTotals = {};


 var SMALL_FILE_MB        = 100;
 var BYTES_MB             = 1024 * 1024;
 var BYTES_GB             = 1024 * 1024 * 1024;
 var BASE_VAR_THRESHOLD   = 0.05;
 var SMALL_SET_THRESHOLD  = 0.20;


 function sumCameraBytes_(folder) {
   var total = 0;
   (function recurse(cf) {
     var itF = cf.getFiles();
     while (itF.hasNext()) {
       var f2    = itF.next();
       var name  = f2.getName();
       var lower = name.toLowerCase();
       if (name.indexOf('._') === 0) continue;


       if (/\.(mp4|mov|mxf)$/i.test(lower)) {
         total += f2.getSize();
       }
     }
     var itSub = cf.getFolders();
     while (itSub.hasNext()) recurse(itSub.next());
   })(folder);
   return total;
 }


 function analyseCameraSet_(parentFolder, parentLabel, cameraFolders) {
   if (!cameraFolders.length) return;


   var setTotalsBytes = {};


   cameraFolders.forEach(function (info) {
     var camName = info.name;
     var cf      = info.folder;


     if (cams.indexOf(camName) === -1) cams.push(camName);


     var bytes = sumCameraBytes_(cf);
     setTotalsBytes[camName] = bytes;


     overallTotals[camName] = (overallTotals[camName] || 0) + bytes;
   });


   var camNames = Object.keys(setTotalsBytes);
   if (!camNames.length) return;


   // Camera-count check for THIS part
   var needed;
   if (seats <= 1) {
     needed = 2;
   } else {
     needed = 3;
   }
   if (isRemote && needed > 0) needed--;


   if (camNames.length < needed) {
     res.cameraCountOK = 'No';
     res.cameraCountIssues.push(
       parentLabel + ': only ' + camNames.length +
       ' camera(s) – ' + camNames.join(', ') +
       ' (expected ' + needed + ')'
     );
   }


   // Timestamp drift check per camera folder (direct files only in that camera folder)
   cameraFolders.forEach(function(info) {
     var drift = folderCreatedTimeDrift_(info.folder, /\.(mp4|mov|mxf)$/i);
     if (!drift.ok) {
       res.videoTimestampsOK = false;
       res.videoTimestampIssues.push(
         parentLabel + '/' + info.name + ' drift ' + fmtDrift_(drift.driftMs) + ' (n=' + drift.count + ')'
       );
     }
   });


   // LIVE SET HANDLING: if path contains "live", skip variation comparison
   var isLiveSet = parentLabel.toLowerCase().indexOf('live') !== -1;


   // Size-variation check for THIS part (skip if LIVE)
   if (!isLiveSet && camNames.length >= 2) {
     var bytesArr = camNames.map(function (k) { return setTotalsBytes[k]; });
     var max = Math.max.apply(null, bytesArr);
     var min = Math.min.apply(null, bytesArr);


     if (max > 0) {
       var isSmallSet = bytesArr.some(function (b) {
         return b < 5 * BYTES_GB;
       });
       var threshold = isSmallSet ? SMALL_SET_THRESHOLD : BASE_VAR_THRESHOLD;


       var ratioDiff = (max - min) / max;
       if (ratioDiff > threshold) {
         res.variation = true;


         var partSummary = camNames.sort().map(function (cam) {
           var gb = setTotalsBytes[cam] / BYTES_GB;
           return cam + ': ' + gb.toFixed(2) + ' GB';
         }).join(', ');


         res.variationDetails.push(parentLabel + ' – ' + partSummary);
       }
     }
   }
 }


 (function walk(f, path) {
   path = path || f.getName();


   var hasFile = false;
   var files   = f.getFiles();
   while (files.hasNext()) {
     hasFile = true;
     var file   = files.next();
     var name   = file.getName();
     var lower  = name.toLowerCase();


     if (name.indexOf('._') === 0) continue;


     var sizeBytes = file.getSize();


     if (/\.(mp4|mov|mxf)$/i.test(lower)) {
       var sizeMB = sizeBytes / BYTES_MB;
       if (sizeMB < SMALL_FILE_MB) {
         res.smallFiles = true;
         res.smallFileList.push(
           path + '/' + name + ' (' + sizeMB.toFixed(1) + ' MB)'
         );
       }
     }
   }


   var subs   = f.getFolders();
   var hasSub = false;


   var cameraSubs = [];
   var otherSubs  = [];


   while (subs.hasNext()) {
     hasSub = true;
     var s = subs.next();
     var n = s.getName();


     if (/^camera\s*[A-Z]$/i.test(n) || /^cam\s*[A-Z]$/i.test(n) || /^[A-D]$/i.test(n)) {
       cameraSubs.push({ name: n, folder: s });
     } else {
       otherSubs.push(s);
     }
   }


   if (cameraSubs.length) {
     analyseCameraSet_(f, path, cameraSubs);
   }


   cameraSubs.forEach(function (info) {
     var subFolder = info.folder;
     walk(subFolder, path + '/' + subFolder.getName());
   });
   otherSubs.forEach(function (subFolder) {
     walk(subFolder, path + '/' + subFolder.getName());
   });


   if (!hasFile && !hasSub) empty++;
 })(root, root.getName());


 res.emptyFolders   = empty;
 res.camerasPresent = cams.join(', ');


 if (!cams.length) {
   res.cameraCountOK = 'No';
   res.videoTimestampsOK = false;
   res.videoTimestampIssues.push('No camera folders found');
 }


 if (Object.keys(overallTotals).length) {
   var parts = [];
   Object.keys(overallTotals).sort().forEach(function (cam) {
     var gb = overallTotals[cam] / BYTES_GB;
     parts.push(cam + ': ' + gb.toFixed(2) + ' GB');
   });
   res.debug = 'Video totals – ' + parts.join(', ');
 }


 return res;
}




/**
* Duplicate camera files:
* - Only considers video files (.mp4, .mov, .mxf)
* - Ignores index files starting with "._"
* - Flags “Yes” if the same filename appears more than once
*/
function checkDuplicateCameraFiles_(root) {
 if (!root) return false;


 var seen = {};
 var found = false;


 (function walk(f) {
   if (found) return;


   var it = f.getFiles();
   while (it.hasNext()) {
     var file  = it.next();
     var name  = file.getName();
     var lower = name.toLowerCase();


     if (name.indexOf('._') === 0) continue;


     if (!(/\.(mp4|mov|mxf)$/i.test(lower))) continue;


     if (seen[lower]) {
       found = true;
       return;
     }
     seen[lower] = true;
   }


   var subs = f.getFolders();
   while (subs.hasNext() && !found) {
     walk(subs.next());
   }
 })(root);


 return found;
}




/*************** AUDIO CHECKS ***************/
function checkAudio_(root, seats, isRemote) {
 var res = {
   expected: 0,
   actual:   0,
   matchOK:  false,
   stereoOK: false,
   notes:    [],
   audioTimestampsOK:    true,   // NEW
   audioTimestampIssues: []      // NEW
 };


 if (!root) {
   res.notes.push('No folder');
   res.audioTimestampsOK = false;
   return res;
 }


 var audioRoot = null;
 (function findAudio(f) {
   if (f.getName().toLowerCase() === 'audio') {
     audioRoot = f;
     return;
   }
   var subs = f.getFolders();
   while (subs.hasNext() && !audioRoot) {
     findAudio(subs.next());
   }
 })(root);


 if (!audioRoot) {
   res.notes.push('No Audio folder');
   res.audioTimestampsOK = false;
   return res;
 }


 var totalExp  = 0;
 var totalAct  = 0;
 var hasAnyWavs      = false;
 var hasAnyStereoWav = false;


 function processTake(tf) {
   var wavs = [];
   var it = tf.getFiles();
   while (it.hasNext()) {
     var file    = it.next();
     var rawName = file.getName();


     if (rawName.indexOf('._') === 0) continue;


     var name = rawName.toLowerCase();
     if (name.endsWith('.wav')) {
       wavs.push(file);
       hasAnyWavs = true;
       if (/stereo|mix/.test(name)) {
         hasAnyStereoWav = true;
       }
     }
   }


   var exp = (seats || 0) + 1;
   if (isRemote && exp > 0) exp--;


   var act = wavs.length;
   totalExp += exp;
   totalAct += act;


   // Timestamp drift check inside this take folder (direct wav files)
   var drift = folderCreatedTimeDrift_(tf, /\.wav$/i);
   if (!drift.ok) {
     res.audioTimestampsOK = false;
     res.audioTimestampIssues.push(
       'Audio/' + tf.getName() + ' drift ' + fmtDrift_(drift.driftMs) + ' (n=' + drift.count + ')'
     );
   }
 }


 var takes = audioRoot.getFolders();
 var hasTakes = false;
 while (takes.hasNext()) {
   hasTakes = true;
   processTake(takes.next());
 }
 if (!hasTakes) {
   processTake(audioRoot);
 }


 res.expected = totalExp;
 res.actual   = totalAct;
 res.matchOK  = (totalExp === totalAct);
 res.stereoOK = hasAnyWavs && hasAnyStereoWav;


 if (!res.stereoOK) {
   res.notes.push('Audio: no Stereo/Mix file found');
 }


 // If there were no wavs at all, timestamps are not meaningful
 if (!hasAnyWavs) {
   res.audioTimestampsOK = false;
   res.audioTimestampIssues.push('No WAV files found');
 }


 return res;
}




/*************** SHEET HEADER & WRITER ***************/
function sanityHeader_() {
 return [
   'Date',
   'Time',
   'Room',
   'Client',
   'Link',
   'Empty Folders?',
   'Cameras Present?',
   'Enough Cameras?',
   'Duplicate camera files?',
   'Small video files?',
   'File size variation?',
   'Expected Audio Files',
   'Actual audio files?',
   'Files Match?',
   'Stereo Mix OK?',
   'Remote files?',
   'Video timestamps OK?',
   'Audio timestamps OK?',
   'Notes'
 ];
}


function writeSanitySheet_(header, rows) {
 var ss = SpreadsheetApp.getActive();
 var sh = ss.getSheetByName('Sanity Checks');
 if (!sh) sh = ss.insertSheet('Sanity Checks');
 sh.clearContents();
 sh.getRange(1,1,1,header.length).setValues([header]);
 if (rows.length) {
   sh.getRange(2,1,rows.length,header.length).setValues(rows);
 }
 sh.getDataRange().setHorizontalAlignment('left');
}




/*************** POPUP (Integrity Summary) ***************/
function showSanityPopup_(rows, header, title) {
 var idxLink   = header.indexOf('Link');
 var idxClient = header.indexOf('Client');


 var displayCols = [];
 for (var i = 0; i < header.length; i++) {
   if (i === idxLink) continue;
   displayCols.push(i);
 }


 function esc(s) {
   return String(s == null ? '' : s)
     .replace(/&/g,'&amp;')
     .replace(/</g,'&lt;')
     .replace(/>/g,'&gt;');
 }


 // Yes = good
 var yesGoodCols = {
   'Enough Cameras?': true,
   'Files Match?': true,
   'Stereo Mix OK?': true,
   'Video timestamps OK?': true,
   'Audio timestamps OK?': true
 };


 // No = good
 var noGoodCols = {
   'Small video files?': true,
   'File size variation?': true,
   'Duplicate camera files?': true
 };


 function classify(colIndex, value) {
   var label = header[colIndex];
   var v = String(value || '').toLowerCase();


   if (label === 'Empty Folders?') {
     var n = Number(value);
     if (!isFinite(n)) return 'neutral';
     return n === 0 ? 'good' : 'bad';
   }


   if (label === 'Remote files?') {
     if (v.indexOf('not remote') !== -1 || v.indexOf('not expected') !== -1) {
       return 'good';
     }
     if (v.indexOf('yes') === 0) return 'good';
     if (v === 'no')             return 'bad';
     return 'neutral';
   }


   if (yesGoodCols[label]) {
     if (v === 'yes') return 'good';
     if (v === 'no')  return 'bad';
     return 'neutral';
   }


   if (noGoodCols[label]) {
     if (v === 'no')  return 'good';
     if (v === 'yes') return 'bad';
     return 'neutral';
   }


   return 'neutral';
 }


 function rowHasIssue(row) {
   for (var j = 0; j < header.length; j++) {
     if (classify(j, row[j]) === 'bad') return true;
   }
   return false;
 }


 function decorate(label, value) {
   var v = String(value || '');
   var vLower = v.toLowerCase();


   if (label === 'Empty Folders?') {
     var n = Number(value);
     if (!isFinite(n)) return v;
     if (n === 0) return '✅ 0';
     return '❌ ' + n;
   }


   if (label === 'Remote files?') {
     if (vLower.indexOf('not remote') !== -1 || vLower.indexOf('not expected') !== -1) {
       return '✅ No';
     }
     if (vLower.indexOf('yes') === 0) return '✅ Yes';
     if (vLower === 'no')             return '❌ No';
     return v;
   }


   if (yesGoodCols[label]) {
     if (vLower === 'yes') return '✅ Yes';
     if (vLower === 'no')  return '❌ No';
     return v;
   }


   if (noGoodCols[label]) {
     if (vLower === 'no')  return '✅ No';
     if (vLower === 'yes') return '❌ Yes';
     return v;
   }


   return v;
 }


 var html = '<!DOCTYPE html><html><head><base target="_blank">' +
   '<style>' +
   'body{font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Helvetica,Arial,sans-serif;padding:16px;color:#222;}' +
   'h1{margin:0 0 8px 0;font-size:22px;}' +
   'h2{margin:0 0 16px 0;font-size:16px;font-weight:500;color:#555;}' +
   'table{width:100%;border-collapse:collapse;font-size:13px;margin-top:8px;}' +
   'th{position:sticky;top:0;background:#f8f8f8;border-bottom:1px solid #ddd;padding:6px 8px;text-align:left;white-space:nowrap;}' +
   'td{padding:5px 8px;border-bottom:1px solid #eee;white-space:nowrap;}' +
   '.good{background:#e6f4ea;color:#137333;}' +
   '.bad{background:#fce8e6;color:#c5221f;}' +
   'a{color:#1a73e8;text-decoration:none;font-weight:500;}' +
   'a:hover{text-decoration:underline;}' +
   '.rowHover:hover{background:#fff8e1 !important;}' +
   '</style></head><body>' +
   '<h1>Integrity Summary</h1>' +
   '<h2>' + esc(title) + '</h2>' +
   '<div style="margin-bottom:8px;font-size:13px;">' +
   '<label style="margin-right:16px;"><input type="radio" name="filterMode" value="all" checked> Show all</label>' +
   '<label><input type="radio" name="filterMode" value="issues"> Only rows with issues</label>' +
   '</div>' +
   '<div style="overflow-x:auto;border:1px solid #ccc;border-radius:6px;">' +
   '<table id="sanityTable"><thead><tr>';


 for (var d = 0; d < displayCols.length; d++) {
   html += '<th>' + esc(header[displayCols[d]]) + '</th>';
 }
 html += '</tr></thead><tbody>';


 for (var r = 0; r < rows.length; r++) {
   var row = rows[r];
   var hasIssue = rowHasIssue(row);
   html += '<tr class="rowHover sanity-row" data-has-issue="' + (hasIssue ? '1' : '0') + '">';


   for (var k = 0; k < displayCols.length; k++) {
     var colIndex = displayCols[k];
     var raw      = row[colIndex];
     var label    = header[colIndex];


     var status   = classify(colIndex, raw);
     var classAttr = status === 'good' ? 'good' : (status === 'bad' ? 'bad' : '');


     var displayValue = decorate(label, raw);
     var cellHtml;


     if (colIndex === idxClient && idxLink >= 0) {
       var url = row[idxLink];
       var labelText = esc(displayValue);
       if (url) cellHtml = '<a href="' + esc(url) + '">' + labelText + '</a>';
       else     cellHtml = labelText;
     } else {
       cellHtml = esc(displayValue);
     }


     html += '<td class="' + classAttr + '">' + cellHtml + '</td>';
   }


   html += '</tr>';
 }


 html += '</tbody></table></div>' +
   '<script>' +
   'const radios=document.getElementsByName("filterMode");' +
   'for(let i=0;i<radios.length;i++){' +
     'radios[i].addEventListener("change",function(){' +
       'const mode=this.value;' +
       'const rows=document.querySelectorAll("#sanityTable tbody tr.sanity-row");' +
       'rows.forEach(function(row){' +
         'const issue=row.getAttribute("data-has-issue")==="1";' +
         'row.style.display=(mode==="issues" && !issue)?"none":"";' +
       '});' +
     '});' +
   '}' +
   '</script></body></html>';


 var out = HtmlService.createHtmlOutput(html)
   .setWidth(1100)
   .setHeight(700);
 SpreadsheetApp.getUi().showModalDialog(out, 'Integrity Summary');
}




/*************** EMAIL SUPPORT ***************/


/**
* Find latest available date key across ALL sheets (strictly before today).
*/
function findLatestSessionDateKeyAcrossSpreadsheet_() {
 var ss = SpreadsheetApp.getActive();
 var tz = ss.getSpreadsheetTimeZone();
 var todayKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');


 var keys = [];


 ss.getSheets().forEach(function (sheet) {
   var lastRow = sheet.getLastRow();
   var lastCol = sheet.getLastColumn();
   if (lastRow < 2 || lastCol < 2) return;


   var range    = sheet.getRange(1, 1, lastRow, lastCol);
   var grid     = range.getDisplayValues();
   var rich     = range.getRichTextValues();
   var formulas = range.getFormulas();


   var targetCols = [];
   for (var c = 0; c < lastCol; c++) targetCols.push(c);


   var allSessions = extractSessions_(grid, rich, formulas, targetCols);
   allSessions.forEach(function (s) {
     var k = headerTextToKey_(s.dateDisplay);
     if (k && k < todayKey && keys.indexOf(k) === -1) keys.push(k);
   });
 });


 if (!keys.length) return null;
 keys.sort();
 return keys[keys.length - 1];
}


/**
* Build the data set for emailing for a specific date key (yyyy-MM-dd)
*/
function getSanityDataForDateKey_(sheet, targetKey) {
 var lastRow = sheet.getLastRow();
 var lastCol = sheet.getLastColumn();
 if (lastRow < 2 || lastCol < 2) {
   return { hasData: false, header: sanityHeader_(), rows: [] };
 }


 var range    = sheet.getRange(1, 1, lastRow, lastCol);
 var grid     = range.getDisplayValues();
 var rich     = range.getRichTextValues();
 var formulas = range.getFormulas();


 var targetCols = [];
 for (var c = 0; c < lastCol; c++) targetCols.push(c);


 var allSessions = extractSessions_(grid, rich, formulas, targetCols);


 var sessions = allSessions.filter(function (sess) {
   var k = headerTextToKey_(sess.dateDisplay);
   return k && k === targetKey;
 });


 var header = sanityHeader_();
 var rows   = sessions.map(processSession_);


 return {
   hasData: rows.length > 0,
   header: header,
   rows: rows
 };
}


/**
* Build HTML email body:
* - Includes "Only rows with issues" table + "All rows" table (email-safe)
*/
function buildSanityEmailHtml_(rows, header, title) {
 var yesGoodCols = {
   'Enough Cameras?': true,
   'Files Match?': true,
   'Stereo Mix OK?': true,
   'Video timestamps OK?': true,
   'Audio timestamps OK?': true
 };
 var noGoodCols = {
   'Small video files?': true,
   'File size variation?': true,
   'Duplicate camera files?': true
 };


 function esc(s) {
   return String(s == null ? '' : s)
     .replace(/&/g,'&amp;')
     .replace(/</g,'&lt;')
     .replace(/>/g,'&gt;');
 }


 function classify(label, value) {
   var v = String(value || '').toLowerCase();


   if (label === 'Empty Folders?') {
     var n = Number(value);
     if (!isFinite(n)) return 'neutral';
     return n === 0 ? 'good' : 'bad';
   }


   if (label === 'Remote files?') {
     if (v.indexOf('not remote') !== -1 || v.indexOf('not expected') !== -1) return 'good';
     if (v.indexOf('yes') === 0) return 'good';
     if (v === 'no') return 'bad';
     return 'neutral';
   }


   if (yesGoodCols[label]) {
     if (v === 'yes') return 'good';
     if (v === 'no')  return 'bad';
     return 'neutral';
   }


   if (noGoodCols[label]) {
     if (v === 'no')  return 'good';
     if (v === 'yes') return 'bad';
     return 'neutral';
   }


   return 'neutral';
 }


 function rowHasIssue_(row) {
   for (var i = 0; i < header.length; i++) {
     if (classify(header[i], row[i]) === 'bad') return true;
   }
   return false;
 }


 function decorate(label, value) {
   var v = String(value || '');
   var vLower = v.toLowerCase();


   if (label === 'Empty Folders?') {
     var n = Number(value);
     if (!isFinite(n)) return v;
     if (n === 0) return '✅ 0';
     return '❌ ' + n;
   }


   if (label === 'Remote files?') {
     if (vLower.indexOf('not remote') !== -1 || vLower.indexOf('not expected') !== -1) return '✅ No';
     if (vLower.indexOf('yes') === 0) return '✅ Yes';
     if (vLower === 'no') return '❌ No';
     return v;
   }


   if (yesGoodCols[label]) {
     if (vLower === 'yes') return '✅ Yes';
     if (vLower === 'no')  return '❌ No';
     return v;
   }


   if (noGoodCols[label]) {
     if (vLower === 'no')  return '✅ No';
     if (vLower === 'yes') return '❌ Yes';
     return v;
   }


   return v;
 }


 var idxLink   = header.indexOf('Link');
 var idxClient = header.indexOf('Client');


 var displayCols = [];
 for (var i = 0; i < header.length; i++) {
   if (i === idxLink) continue;
   displayCols.push(i);
 }


 var issuesRows = rows.filter(rowHasIssue_);
 var totalCount = rows.length;
 var issueCount = issuesRows.length;


 function renderTable_(tableRows) {
   var t = '<table><thead><tr>';
   for (var d = 0; d < displayCols.length; d++) {
     t += '<th>' + esc(header[displayCols[d]]) + '</th>';
   }
   t += '</tr></thead><tbody>';


   if (!tableRows.length) {
     t += '<tr><td colspan="' + displayCols.length + '" style="padding:10px 6px;color:#555;">✅ No issues found.</td></tr>';
     t += '</tbody></table>';
     return t;
   }


   for (var r = 0; r < tableRows.length; r++) {
     var row = tableRows[r];
     t += '<tr>';
     for (var k = 0; k < displayCols.length; k++) {
       var colIndex = displayCols[k];
       var raw      = row[colIndex];
       var label    = header[colIndex];


       var status = classify(label, raw);
       var cls    = status === 'good' ? 'good' : (status === 'bad' ? 'bad' : '');


       var displayValue = decorate(label, raw);


       var cellHtml;
       if (colIndex === idxClient && idxLink >= 0 && row[idxLink]) {
         cellHtml = '<a href="' + esc(row[idxLink]) + '">' + esc(displayValue) + '</a>';
       } else {
         cellHtml = esc(displayValue);
       }


       t += '<td class="' + cls + '">' + cellHtml + '</td>';
     }
     t += '</tr>';
   }


   t += '</tbody></table>';
   return t;
 }


 var html = ''
   + '<html><head><meta charset="UTF-8"><style>'
   + 'body{font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Helvetica,Arial,sans-serif;color:#222;}'
   + 'h1{font-size:20px;margin:0 0 4px 0;}'
   + 'h2{font-size:14px;margin:0 0 12px 0;color:#555;}'
   + '.meta{font-size:12px;color:#555;margin:0 0 14px 0;}'
   + '.pill{display:inline-block;padding:2px 8px;border-radius:999px;font-size:12px;margin-right:8px;}'
   + '.pillBad{background:#fce8e6;color:#c5221f;}'
   + '.pillGood{background:#e6f4ea;color:#137333;}'
   + 'table{width:100%;border-collapse:collapse;font-size:12px;margin-top:6px;}'
   + 'th,td{padding:4px 6px;border-bottom:1px solid #eee;text-align:left;white-space:nowrap;}'
   + '.good{background:#e6f4ea;color:#137333;}'
   + '.bad{background:#fce8e6;color:#c5221f;}'
   + 'a{color:#1a73e8;text-decoration:none;}'
   + '.sectionTitle{font-size:13px;margin:16px 0 6px 0;color:#333;font-weight:700;}'
   + '</style></head><body>'
   + '<h1>Integrity Summary</h1>'
   + '<h2>' + esc(title) + '</h2>'
   + '<p class="meta">'
   +   '<span class="pill ' + (issueCount ? 'pillBad' : 'pillGood') + '">'
   +     (issueCount ? ('Issues: ' + issueCount) : 'Issues: 0')
   +   '</span>'
   +   '<span class="pill pillGood">Total sessions: ' + totalCount + '</span>'
   + '</p>'
   + '<div class="sectionTitle">Only rows with issues</div>'
   + renderTable_(issuesRows)
   + '<div class="sectionTitle">All rows</div>'
   + renderTable_(rows)
   + '</body></html>';


 return html;
}


/**
* Time-driven trigger target: run daily at 12:00.
* Finds the MOST RECENT date key before today that exists across ANY sheet,
* then emails combined results for that date.
*/
function dailySanityEmail() {
 var ss  = SpreadsheetApp.getActive();
 var tz  = ss.getSpreadsheetTimeZone();
 var todayKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');


 var allDateKeys = [];


 // Discover all date keys present in any sheet
 ss.getSheets().forEach(function (sheet) {
   var lastRow = sheet.getLastRow();
   var lastCol = sheet.getLastColumn();
   if (lastRow < 2 || lastCol < 2) return;


   var range    = sheet.getRange(1, 1, lastRow, lastCol);
   var grid     = range.getDisplayValues();
   var rich     = range.getRichTextValues();
   var formulas = range.getFormulas();


   var targetCols = [];
   for (var c = 0; c < lastCol; c++) targetCols.push(c);


   var allSessions = extractSessions_(grid, rich, formulas, targetCols);
   allSessions.forEach(function (s) {
     var k = headerTextToKey_(s.dateDisplay);
     if (k && allDateKeys.indexOf(k) === -1) allDateKeys.push(k);
   });
 });


 if (!allDateKeys.length) {
   MailApp.sendEmail({
     to: 'ben@poddster.com',
     subject: 'Poddster Sanity Checks – No data available',
     htmlBody: '<p>No sessions found on any sheet. There are no date keys to report on.</p>'
   });
   return;
 }


 allDateKeys.sort();


 var candidateKeys = allDateKeys.filter(function (k) { return k < todayKey; });


 if (!candidateKeys.length) {
   MailApp.sendEmail({
     to: 'ben@poddster.com',
     subject: 'Poddster Sanity Checks – No past data available',
     htmlBody:
       '<p>No session dates were found before today (' + todayKey + ').</p>' +
       '<p>Available date keys: ' + allDateKeys.join(', ') + '</p>'
   });
   return;
 }


 var targetKey = candidateKeys[candidateKeys.length - 1];


 var header  = sanityHeader_();
 var allRows = [];


 ss.getSheets().forEach(function (sheet) {
   var data = getSanityDataForDateKey_(sheet, targetKey);
   if (data.rows && data.rows.length) allRows = allRows.concat(data.rows);
 });


 var subject = 'Poddster Sanity Checks – ' + targetKey;


 if (!allRows.length) {
   MailApp.sendEmail({
     to: 'ben@poddster.com',
     subject: subject,
     htmlBody:
       '<p>No sessions found for Automatic – Latest Available Data (' + targetKey + ').</p>' +
       '<p>Available date keys on any sheet: ' + allDateKeys.join(', ') + '</p>'
   });
   return;
 }


 var html = buildSanityEmailHtml_(allRows, header, "Automatic – Latest Available Data (" + targetKey + ")");


 MailApp.sendEmail({
   to: 'ben@poddster.com',
   subject: subject,
   htmlBody: html
 });
}


/**
* Menu item: send test email for yesterday immediately
* - Tries strict yesterday first, else falls back to latest available date across spreadsheet
*/
function sendSanityTestEmail() {
 var ss = SpreadsheetApp.getActive();


 var tz = ss.getSpreadsheetTimeZone();
 var now = new Date();
 var yTime = new Date(now.getTime() - 1 * 24 * 60 * 60 * 1000);
 var yKey  = Utilities.formatDate(yTime, tz, 'yyyy-MM-dd');


 var header  = sanityHeader_();
 var allRows = [];


 ss.getSheets().forEach(function (sheet) {
   var dataY = getSanityDataForDateKey_(sheet, yKey);
   if (dataY.rows && dataY.rows.length) allRows = allRows.concat(dataY.rows);
 });


 var targetKey = yKey;
 var label     = "Test – Yesterday's Data";


 if (!allRows.length) {
   var latest = findLatestSessionDateKeyAcrossSpreadsheet_();
   if (!latest) {
     MailApp.sendEmail({
       to: 'ben@poddster.com',
       subject: 'Poddster Sanity Checks – No data available',
       htmlBody: '<p>No sessions found on any sheet (no prior date keys before today).</p>'
     });
     return;
   }


   allRows = [];
   ss.getSheets().forEach(function (sheet) {
     var dataL = getSanityDataForDateKey_(sheet, latest);
     if (dataL.rows && dataL.rows.length) allRows = allRows.concat(dataL.rows);
   });


   targetKey = latest;
   label     = "Test – Latest Available Data";
 }


 var subject = 'Poddster Sanity Checks – ' + targetKey;


 var html = buildSanityEmailHtml_(allRows, header, label + ' (' + targetKey + ')');


 MailApp.sendEmail({
   to: 'ben@poddster.com',
   subject: subject,
   htmlBody: html
 });
}