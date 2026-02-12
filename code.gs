*************************************************
* PODDSTER — MENUS
**************************************************/

/* ==================================================
* CONSTANTS
* ================================================== */

const DAY_NAMES = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];


function onOpen(e) {
 var ui = SpreadsheetApp.getUi();


 // --- Assign Operators / Editor Assigner menu ---
 var assignMenu = ui.createMenu('👩‍🎨 Assign Operators');
 assignMenu
   .addItem('Assign Editors to TBC on Active Sheet', 'assignEditorsOnActiveSheet')
   .addItem('Undo (replace names back to TBC on Active Sheet)', 'undoAssignedOnActiveSheet')
   .addItem('Undo for Day…', 'undoAssignedForDayOnActiveSheet')
   .addItem('Undo for Day of Active Cell', 'undoAssignedForActiveCellDay')
   .addSeparator()
   .addItem('Show Hours Summary (Active Sheet)', 'showHoursSummaryDialogOnActiveSheet')
   .addItem('Show Break Plan (Active Sheet)', 'showBreakPlanDialogOnActiveSheet')
   .addItem('Show Studio & Set Usage (Active Sheet)', 'showStudioSetUsageDialogOnActiveSheet')
   .addItem('Write Studio Usage Row (Weekly)', 'writeStudioUsageRow')
   .addItem('Write Set Usage Row (Weekly)', 'writeSetUsageRow')
   .addItem('Studio & Set Usage — Monthly (all tabs)', 'showStudioSetUsageMonthlyAllTabs')
   .addToUi();


 // --- Timelines menu ---
 ui.createMenu('⏱️ Timelines')
   .addItem('Show me the room timelines!', 'openTimeline')
   .addItem('Show me the operator timelines!', 'openOperatorTimeline')
   .addToUi();


 // --- Session Data Integrity Checker menu ---
 ui.createMenu('🎬 Data Checker')
   .addItem("Check Today's Data", 'runSanityToday')
   .addItem("Check Yesterday's Data", 'runSanityYesterday')
   .addItem('Check Entire Tab', 'runSanityWholeTab')
   .addSeparator()
   .addItem('Send Test Email', 'sendSanityTestEmail')
   .addToUi();


 // --- Photo JPG Helper menu ---
 ui.createMenu('📸 Photo JPG Helper')
   .addItem('Create JPG folders for PHOTO sessions', 'createJpgFoldersFromBoard')
   .addItem('📄 Create HTML from existing JPGs (Link/ID)', 'createHtmlFromExistingJpegs')
   .addSeparator()
   .addItem('📧 Send TEST Email (Active Row)', 'sendTestEmailActiveRow')
   .addItem('📅 Test Calendar Sync', 'testCalendarSync')
   .addItem('❓ Debug Spreadsheet Layout', 'debugSpreadsheetLayout')
   .addToUi();
}


/* -------------------- Dialogs -------------------- */


function openTimeline() {
 const html = HtmlService.createTemplateFromFile('timeline by room')
   .evaluate()
   .setWidth(1100)
   .setHeight(760);
 SpreadsheetApp.getUi().showModalDialog(html, 'Weekly Timeline');
}


function openOperatorTimeline() {
 const html = HtmlService.createTemplateFromFile('timeline by operator')
   .evaluate()
   .setWidth(1100)
   .setHeight(760);
 SpreadsheetApp.getUi().showModalDialog(html, 'Operator Timeline');
}


/* ==================================================
* SHARED HELPERS – day headers / week anchor / etc.
* ================================================== */


/**
* Find the row that contains weekday headers and return
* both that row index and the [day, startCol, endCol] segments.
*
* Accepts cells like "Monday, 24 November 2025" — we match by
* checking whether the cell *starts with* a day name.
*/
function detectDaySegments_(values, maxScanRows) {
 maxScanRows = maxScanRows || 10;
 const rows = values.length;
 const cols = values[0].length;


 let headerRowIdx = -1;


 // 1) Locate the header row
 outer:
 for (let r = 0; r < Math.min(rows, maxScanRows); r++) {
   for (let c = 0; c < cols; c++) {
     const cell = String(values[r][c] || '').trim();
     if (!cell) continue;
     const lc = cell.toLowerCase();
     for (let d = 0; d < DAY_NAMES.length; d++) {
       const dn = DAY_NAMES[d].toLowerCase();
       if (lc.startsWith(dn)) {
         headerRowIdx = r;
         break outer;
       }
     }
   }
 }


 if (headerRowIdx === -1) {
   throw new Error('Could not find a row with day names in the first few rows.');
 }


 // 2) Build day segments from that header row
 const headerRow = values[headerRowIdx];
 const dayCols = [];


 for (let c = 0; c < cols; c++) {
   const cell = String(headerRow[c] || '').trim();
   if (!cell) continue;
   const lc = cell.toLowerCase();
   for (let d = 0; d < DAY_NAMES.length; d++) {
     const dn = DAY_NAMES[d].toLowerCase();
     if (lc.startsWith(dn)) {
       dayCols.push({ day: DAY_NAMES[d], startCol: c });
       break;
     }
   }
 }


 if (!dayCols.length) {
   throw new Error('No weekday headers detected on the header row.');
 }


 for (let i = 0; i < dayCols.length; i++) {
   dayCols[i].endCol = (i < dayCols.length - 1) ? dayCols[i + 1].startCol : cols;
 }


 return { headerRowIdx, dayCols };
}


/**
* Parse the "6th - 12th October" style header in A1 to get
* the Monday date for that week. Fallback = current week.
*/
function parseWeekHeader(txt) {
 if (!txt) return null;
 const now = new Date();
 const yr = now.getFullYear();


 const mDash = String(txt).match(
   /(\d{1,2})(?:st|nd|rd|th)?\s*[-–—]\s*\d{1,2}(?:st|nd|rd|th)?\s+([A-Za-z]+)(?:\s+(\d{4}))?/
 );
 const mOne = String(txt).match(
   /(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+)(?:\s+(\d{4}))?/
 );


 let monthName, day1, year = yr;
 if (mDash) {
   day1 = parseInt(mDash[1], 10);
   monthName = mDash[2];
   year = mDash[3] ? parseInt(mDash[3], 10) : yr;
 } else if (mOne) {
   day1 = parseInt(mOne[1], 10);
   monthName = mOne[2];
   year = mOne[3] ? parseInt(mOne[3], 10) : yr;
 } else {
   return null;
 }


 const monthIndex = monthNameToIndex(monthName);
 if (monthIndex == null) return null;


 const d = new Date(year, monthIndex, day1);
 const dow = d.getDay(); // 0 = Sun
 const delta = (dow === 0) ? -6 : (1 - dow);
 const monday = new Date(d);
 monday.setDate(d.getDate() + delta);
 monday.setHours(0, 0, 0, 0);


 return { mondayDate: monday, year };
}


function fallbackWeekInfo() {
 const d = new Date();
 const dow = d.getDay();
 const delta = (dow === 0) ? -6 : (1 - dow);
 const monday = new Date(d);
 monday.setDate(d.getDate() + delta);
 monday.setHours(0, 0, 0, 0);
 return { mondayDate: monday, year: monday.getFullYear() };
}


function monthNameToIndex(name) {
 const map = {
   january: 0, february: 1, march: 2, april: 3, may: 4, june: 5,
   july: 6, august: 7, september: 8, october: 9, november: 10, december: 11
 };
 const key = String(name || '').toLowerCase();
 return (key in map) ? map[key] : null;
}


/* ==================================================
* WEEKLY TIMELINE (STUDIOS / SETS)
* ================================================== */


function getWeeklyTimelinePayload() {
 const sh = SpreadsheetApp.getActiveSheet();
 const rng = sh.getDataRange();
 const values = rng.getDisplayValues();
 if (!values.length || !values[0].length) throw new Error('Sheet looks empty.');


 // Detect header row + day segments using robust helper
 const { headerRowIdx, dayCols } = detectDaySegments_(values, 10);


 // Week anchor from A1
 const a1 = (values[0][0] || '').toString();
 const weekInfo = parseWeekHeader(a1) || fallbackWeekInfo();


 // Map day name -> actual Date in that week (Monday baseline)
 const mondayIndex = DAY_NAMES.indexOf('Monday');
 const dayToDate = {};
 for (const seg of dayCols) {
   const base = new Date(weekInfo.mondayDate);
   base.setDate(base.getDate() + (DAY_NAMES.indexOf(seg.day) - mondayIndex));
   dayToDate[seg.day] = base;
 }


 const nameTimeRegex = /name\s*\/\s*time/i;
 const timeRe = /^\s*(\d{1,2}):(\d{2})\s*[-–—]\s*(\d{1,2}):(\d{2})\s*$/;


 function findRowIndex(regex, startRow) {
   for (let rr = startRow; rr < Math.min(values.length, startRow + 6); rr++) {
     const left = values[rr]
       .slice(0, Math.min(6, values[rr].length))
       .join(' ')
       .toLowerCase();
     if (regex.test(left)) return rr;
   }
   return null;
 }


 // From the "Room/Seats" row extract a likely [room,set]/[seats] near the time triplet.
 function pickRoomAndSeats(rowVals, seg, baseCol) {
   if (!rowVals) return { room: '', seats: '' };
   const wStart = Math.max(seg.startCol, baseCol - 2);
   const wEnd   = Math.min(seg.endCol,   baseCol + 4);
   let room = '', seats = '';
   for (let c = wStart; c < wEnd - 1; c++) {
     const a = String(rowVals[c]   || '').trim();
     const b = String(rowVals[c+1] || '').trim();
     const aNum = /^\d+(\.\d+)?$/.test(a);
     const bNum = /^\d+(\.\d+)?$/.test(b);
     if (!room  && !aNum && a) room = a;                // text like Exec / Iris / etc.
     if (!seats && (bNum || /^\d+$/.test(b))) seats = b;
     if (room && seats) break;
   }
   return { room, seats };
 }


 const rows = [];


 // Scan each "Name / Time" block row
 for (let r = headerRowIdx + 1; r < values.length; r++) {
   const rowHasNameTime = values[r]
     .slice(0, Math.min(6, values[r].length))
     .some(v => nameTimeRegex.test(String(v || '')));
   if (!rowHasNameTime) continue;


   const roomRowIdx = findRowIndex(/\broom\b.*\bseats\b/i, r + 1);


   // For each day segment, scan for time strings and read client/host to the left
   for (const seg of dayCols) {
     const segVals = values[r]
       .slice(seg.startCol, seg.endCol)
       .map(s => String(s).trim());


     for (let t = 0; t < segVals.length; t++) {
       const m = segVals[t].match(timeRe);
       if (!m) continue;


       if (t - 2 < 0) continue;
       const left2 = String(segVals[t - 2] || '').trim(); // host
       const left1 = String(segVals[t - 1] || '').trim(); // client


       let client = left1, host = left2;
       if (left2.length > left1.length) { client = left2; host = left1; }
       if (!client && !host) continue;


       const baseCol  = seg.startCol + (t - 2);
       const baseDate = dayToDate[seg.day];


       const shh = Math.min(23, parseInt(m[1], 10));
       const smm = Math.min(59, parseInt(m[2], 10));
       const ehh = Math.min(23, parseInt(m[3], 10));
       const emm = Math.min(59, parseInt(m[4], 10));


       const start = new Date(baseDate); start.setHours(shh, smm, 0, 0);
       const end   = new Date(baseDate); end.setHours(ehh, emm, 0, 0);
       if (end <= start) continue;


       const rs = (roomRowIdx != null)
         ? pickRoomAndSeats(values[roomRowIdx], seg, baseCol)
         : { room: '', seats: '' };


       rows.push({
         group: `${seg.day} — ${rs.room || ''}`,
         label: client || '(TBC)',
         host:  host   || '',
         room:  rs.room || '',
         seats: rs.seats || '',
         start,
         end,
         startMs: start.getTime(),
         endMs:   end.getTime()
       });
     }
   }
 }


 return {
   title: sh.getRange(1, 1).getDisplayValue() || 'Week',
   sheetName: sh.getName(),
   items: rows
 };
}


function getWeeklyTimelinePayloadSafe() {
 const p = getWeeklyTimelinePayload();
 const items = (p.items || [])
   .map(it => ({
     group:  String(it.group  || ''),
     label:  String(it.label  || ''),
     host:   String(it.host   || ''),
     room:   String(it.room   || ''),
     seats:  String(it.seats  || ''),
     startMs: Number(it.startMs),
     endMs:   Number(it.endMs)
   }))
   .filter(x => isFinite(x.startMs) && isFinite(x.endMs) && x.endMs > x.startMs);


 return {
   title: String(p.title || ''),
   sheetName: String(p.sheetName || ''),
   items
 };
}


/* ==================================================
* OPERATOR TIMELINE (Naz / Syaz / Sufi)
* ================================================== */


function getOperatorTimelinePayload() {
 const sh = SpreadsheetApp.getActiveSheet();
 const rng = sh.getDataRange();
 const values = rng.getDisplayValues();
 if (!values.length || !values[0].length) throw new Error('Sheet looks empty.');


 // Day segments (same as weekly)
 const { headerRowIdx, dayCols } = detectDaySegments_(values, 10);


 // Week anchor
 const a1 = (values[0][0] || '').toString();
 const weekInfo = parseWeekHeader(a1) || fallbackWeekInfo();
 const mondayIndex = DAY_NAMES.indexOf('Monday');
 const dayToDate = {};
 for (const seg of dayCols) {
   const base = new Date(weekInfo.mondayDate);
   base.setDate(base.getDate() + (DAY_NAMES.indexOf(seg.day) - mondayIndex));
   dayToDate[seg.day] = base;
 }


 const nameTimeRegex = /name\s*\/\s*time/i;
 const timeRe = /^\s*(\d{1,2}):(\d{2})\s*[-–—]\s*(\d{1,2}):(\d{2})\s*$/;


 function normOp(x) {
   const v = String(x || '').trim().toLowerCase().replace(/[^\p{L}\p{N}]+/gu, '');
   if (v.startsWith('naz'))  return 'Naz';
   if (v.startsWith('syaz')) return 'Syaz';
   if (v.startsWith('sufi')) return 'Sufi';
   return v ? v.charAt(0).toUpperCase() + v.slice(1) : '';
 }


 // Read operator from the first column of the day segment; walk upward for merged cells
 function getOperatorFromSegment(rowIdx, seg) {
   const c = seg.startCol;
   for (let r = rowIdx; r >= Math.max(headerRowIdx + 1, rowIdx - 10); r--) {
     const raw = values[r] && values[r][c];
     if (raw && String(raw).trim()) return normOp(raw);
   }
   return '';
 }


 function findRowIndex(regex, startRow) {
   for (let rr = startRow; rr < Math.min(values.length, startRow + 6); rr++) {
     const left = values[rr]
       .slice(0, Math.min(6, values[rr].length))
       .join(' ')
       .toLowerCase();
     if (regex.test(left)) return rr;
   }
   return null;
 }


 function pickRoomAndSeats(rowVals, seg, baseCol) {
   if (!rowVals) return { room: '', seats: '' };
   const wStart = Math.max(seg.startCol, baseCol - 2);
   const wEnd   = Math.min(seg.endCol,   baseCol + 4);
   let room = '', seats = '';
   for (let c = wStart; c < wEnd - 1; c++) {
     const a = String(rowVals[c]   || '').trim();
     const b = String(rowVals[c+1] || '').trim();
     const aNum = /^\d+(\.\d+)?$/.test(a);
     const bNum = /^\d+(\.\d+)?$/.test(b);
     if (!room  && !aNum && a) room = a;
     if (!seats && (bNum || /^\d+$/.test(b))) seats = b;
     if (room && seats) break;
   }
   return { room, seats };
 }


 const rows = [];


 // Walk each "Name/Time" row
 for (let r = headerRowIdx + 1; r < values.length; r++) {
   const rowHasNameTime = values[r]
     .slice(0, Math.min(6, values[r].length))
     .some(v => nameTimeRegex.test(String(v || '')));
   if (!rowHasNameTime) continue;


   const roomRowIdx = findRowIndex(/\broom\b.*\bseats\b/i, r + 1);


   for (const seg of dayCols) {
     const segVals = values[r]
       .slice(seg.startCol, seg.endCol)
       .map(s => String(s).trim());


     for (let t = 0; t < segVals.length; t++) {
       const m = segVals[t].match(timeRe);
       if (!m) continue;


       if (t - 2 < 0) continue;
       const left2 = String(segVals[t - 2] || '').trim();
       const left1 = String(segVals[t - 1] || '').trim();
       let client = left1, host = left2;
       if (left2.length > left1.length) { client = left2; host = left1; }
       if (!client && !host) continue;


       const baseCol  = seg.startCol + (t - 2);
       const baseDate = dayToDate[seg.day];


       const shh = Math.min(23, parseInt(m[1], 10));
       const smm = Math.min(59, parseInt(m[2], 10));
       const ehh = Math.min(23, parseInt(m[3], 10));
       const emm = Math.min(59, parseInt(m[4], 10));


       const start = new Date(baseDate); start.setHours(shh, smm, 0, 0);
       const end   = new Date(baseDate); end.setHours(ehh, emm, 0, 0);
       if (end <= start) continue;


       const operator = getOperatorFromSegment(r, seg) || 'Unassigned';
       const rs = (roomRowIdx != null)
         ? pickRoomAndSeats(values[roomRowIdx], seg, baseCol)
         : { room: '', seats: '' };


       rows.push({
         group:    `${seg.day} — ${operator}`,
         operator: operator,
         label:    client || '(TBC)',
         host:     host   || '',
         room:     rs.room || '',
         seats:    rs.seats || '',
         start,
         end,
         startMs: start.getTime(),
         endMs:   end.getTime()
       });
     }
   }
 }


 return {
   title: sh.getRange(1, 1).getDisplayValue() || 'Week',
   sheetName: sh.getName(),
   items: rows
 };
}


function getOperatorTimelinePayloadSafe() {
 const p = getOperatorTimelinePayload();
 const items = (p.items || [])
   .map(it => ({
     group:    String(it.group    || ''),
     operator: String(it.operator || ''),
     label:    String(it.label    || ''),
     host:     String(it.host     || ''),
     room:     String(it.room     || ''),
     seats:    String(it.seats    || ''),
     startMs:  Number(it.startMs),
     endMs:    Number(it.endMs)
   }))
   .filter(x => isFinite(x.startMs) && isFinite(x.endMs) && x.endMs > x.startMs);


 return {
   title: String(p.title || ''),
   sheetName: String(p.sheetName || ''),
   items
 };
}