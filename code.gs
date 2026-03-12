/**************************************************
* PODDSTER — MENUS
***************************************************/

/* ==================================================
* CONSTANTS
* ================================================== */

const DAY_NAMES = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];


// NOTE: This onOpen is disabled - the master onOpen is in Code Scheduling.gs
// function onOpen(e) {
//  // This function has been moved to Code Scheduling.gs to avoid conflicts
// }


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


/* ==================================================
* DAILY STUDIO & SET USAGE TRACKING
* ================================================== */

// NOTE: Constants STUDIO_ORDER, SET_ORDER, SET_TO_STUDIO, and UNKNOWN_STUDIO
// are defined in "Code Scheduling.gs" to avoid duplicate declarations

/**
 * Write daily studio usage for the active sheet
 */
function writeDailyStudioUsage() {
  const dailyStats = _getDailyStatsForActiveSheet_();

  dailyStats.forEach(dayStat => {
    _upsertDailyUsageRow_({
      sheetName: 'Studio Usage (Daily)',
      label: dayStat.dateLabel,
      labels: STUDIO_ORDER,
      rowBuilder: function(name) {
        const rec = dayStat.byStudio[name] || { hours: 0, count: 0 };
        const pct = dayStat.totalHours > 0 ? (rec.hours / dayStat.totalHours) : 0;
        return [rec.hours, pct, rec.count];
      }
    });
  });

  SpreadsheetApp.getUi().alert('Daily studio usage written for ' + dailyStats.length + ' day(s).');
}


/**
 * Write daily set usage for the active sheet
 */
function writeDailySetUsage() {
  const dailyStats = _getDailyStatsForActiveSheet_();

  dailyStats.forEach(dayStat => {
    _upsertDailyUsageRow_({
      sheetName: 'Set Usage (Daily)',
      label: dayStat.dateLabel,
      labels: SET_ORDER,
      rowBuilder: function(name) {
        const rec = dayStat.bySet[name] || { hours: 0, count: 0 };
        const pct = dayStat.totalHours > 0 ? (rec.hours / dayStat.totalHours) : 0;
        return [rec.hours, pct, rec.count];
      }
    });
  });

  SpreadsheetApp.getUi().alert('Daily set usage written for ' + dailyStats.length + ' day(s).');
}

// NOTE: backfillDailyStudioUsage() and backfillDailySetUsage() have been moved
// to "Code Timelines - Daily Usage.gs" to avoid duplicate function definitions


/**
 * Get daily stats for the active sheet - breaks down by day
 */
function _getDailyStatsForActiveSheet_() {
  const sh = SpreadsheetApp.getActiveSheet();
  return _getDailyStatsForSheet_(sh);
}


/**
 * Get daily stats for a specific sheet - breaks down by day
 */
function _getDailyStatsForSheet_(sheet) {
  const rng = sheet.getDataRange();
  const values = rng.getDisplayValues();

  if (!values.length || !values[0].length) {
    throw new Error('Sheet looks empty.');
  }

  // Detect day segments
  const { headerRowIdx, dayCols } = detectDaySegments_(values, 10);

  // Parse dates from row 2 (the header row with dates)
  const dayDates = {};
  dayCols.forEach(seg => {
    const cellText = values[headerRowIdx][seg.startCol] || '';
    const parsedDate = _parseDateFromDayHeader_(cellText);
    dayDates[seg.day] = parsedDate;
  });

  // Get timeline payload to extract session data
  const payload = _getTimelinePayloadForSheet_(sheet);
  const items = (payload && payload.items) ? payload.items : [];

  if (!items.length) {
    return [];
  }

  // Group items by day
  const byDay = {};
  items.forEach(it => {
    const dayName = _extractDayFromGroup_(it.group);
    if (!byDay[dayName]) {
      byDay[dayName] = {
        dayName: dayName,
        date: dayDates[dayName] || new Date(),
        items: []
      };
    }
    byDay[dayName].items.push(it);
  });

  // Calculate stats per day
  const dailyStats = [];

  Object.keys(byDay).forEach(dayName => {
    const dayData = byDay[dayName];
    let totalHours = 0;
    const byStudio = {};
    const bySet = {};

    dayData.items.forEach(it => {
      const ms = Math.max(0, Number(it.endMs || 0) - Number(it.startMs || 0));
      if (!ms) return;

      const hours = ms / 3600000;
      const setName = String(it.room || '').trim();
      if (!setName) return;

      const studioName = SET_TO_STUDIO[setName.toLowerCase()] || UNKNOWN_STUDIO;
      totalHours += hours;

      if (!bySet[setName]) bySet[setName] = { hours: 0, count: 0 };
      bySet[setName].hours += hours;
      bySet[setName].count += 1;

      if (!byStudio[studioName]) byStudio[studioName] = { hours: 0, count: 0 };
      byStudio[studioName].hours += hours;
      byStudio[studioName].count += 1;
    });

    if (totalHours > 0) {
      dailyStats.push({
        dayName: dayName,
        date: dayData.date,
        dateLabel: _formatDateLabel_(dayData.date),
        totalHours: totalHours,
        byStudio: byStudio,
        bySet: bySet
      });
    }
  });

  return dailyStats;
}


/**
 * Parse date from day header cell (e.g., "Monday, 9 February 2026")
 */
function _parseDateFromDayHeader_(cellText) {
  const text = String(cellText || '').trim();

  // Try to match "Monday, 9 February 2026" or "Monday, 9 February"
  const match = text.match(/,\s*(\d{1,2})\s+([A-Za-z]+)(?:\s+(\d{4}))?/);

  if (match) {
    const day = parseInt(match[1], 10);
    const monthName = match[2];
    const year = match[3] ? parseInt(match[3], 10) : new Date().getFullYear();

    const monthIndex = monthNameToIndex(monthName);
    if (monthIndex != null) {
      return new Date(year, monthIndex, day);
    }
  }

  // Fallback: return today
  return new Date();
}


/**
 * Format date as YYYY-MM-DD for use as row label
 */
function _formatDateLabel_(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}


/**
 * Extract day name from group string like "Monday — Iris"
 */
function _extractDayFromGroup_(groupStr) {
  const parts = String(groupStr || '').split('—');
  return parts[0] ? parts[0].trim() : '';
}


/**
 * Get timeline payload for a specific sheet (not just active)
 */
function _getTimelinePayloadForSheet_(sheet) {
  const rng = sheet.getDataRange();
  const values = rng.getDisplayValues();
  const rawValues = rng.getValues(); // Raw values give us actual Date objects from Sheets cells

  if (!values.length || !values[0].length) {
    throw new Error('Sheet looks empty.');
  }

  const { headerRowIdx, dayCols } = detectDaySegments_(values, 10);

  // Build dayToDate map - try three methods in order of reliability:
  // 1. Raw Date cell from row 2 (post-Nov 2025 sheets have actual date cells)
  // 2. Text parsing of row 2 display value (e.g. "Monday, 9 March 2026")
  // 3. A1 week header fallback (pre-Nov 2025 sheets have "Monday" text in row 2,
  //    but A1 has parseable text like "9th Mar - 15th Mar")
  // Never fall back to fallbackWeekInfo() as that returns the current week which
  // causes all sessions to be wrongly attributed to the current week's dates.
  const dayToDate = {};

  for (const seg of dayCols) {
    // Method 1: raw Date object from Sheets cell
    const rawCell = rawValues[headerRowIdx][seg.startCol];
    if (rawCell instanceof Date && !isNaN(rawCell.getTime())) {
      const d = new Date(rawCell);
      d.setHours(0, 0, 0, 0);
      dayToDate[seg.day] = d;
      continue;
    }
    // Method 2: parse display text e.g. "Monday, 9 March 2026"
    const cellText = values[headerRowIdx][seg.startCol];
    const parsed = _parseDateFromRow2Cell_(cellText);
    if (parsed) {
      dayToDate[seg.day] = parsed;
    }
    // Method 3: A1 fallback handled below after the loop
  }

  // Method 3: if any days are still missing dates, try A1 week header
  // (needed for old sheets where row 2 just says "Monday", "Tuesday" etc.)
  const missingDays = dayCols.filter(seg => !dayToDate[seg.day]);
  if (missingDays.length > 0) {
    const a1 = (values[0][0] || '').toString();
    const weekInfo = parseWeekHeader(a1);
    if (weekInfo) {
      const mondayIndex = DAY_NAMES.indexOf('Monday');
      missingDays.forEach(seg => {
        const base = new Date(weekInfo.mondayDate);
        base.setDate(base.getDate() + (DAY_NAMES.indexOf(seg.day) - mondayIndex));
        dayToDate[seg.day] = base;
      });
    } else {
      // All three methods failed - log and skip these days entirely
      missingDays.forEach(seg => {
        Logger.log('WARNING: Cannot determine date for ' + seg.day + ' on sheet "' + sheet.getName() + '" - sessions for this day will be skipped.');
      });
    }
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

  function pickRoomAndSeats(rowVals, seg, baseCol) {
    if (!rowVals) return { room: '', seats: '' };
    const wStart = Math.max(seg.startCol, baseCol - 2);
    const wEnd = Math.min(seg.endCol, baseCol + 4);
    let room = '', seats = '';
    for (let c = wStart; c < wEnd - 1; c++) {
      const a = String(rowVals[c] || '').trim();
      const b = String(rowVals[c + 1] || '').trim();
      const aNum = /^\d+(\.\d+)?$/.test(a);
      const bNum = /^\d+(\.\d+)?$/.test(b);
      if (!room && !aNum && a) room = a;
      if (!seats && (bNum || /^\d+$/.test(b))) seats = b;
      if (room && seats) break;
    }
    return { room, seats };
  }

  const rows = [];

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

        const baseCol = seg.startCol + (t - 2);
        const baseDate = dayToDate[seg.day];

        const shh = Math.min(23, parseInt(m[1], 10));
        const smm = Math.min(59, parseInt(m[2], 10));
        const ehh = Math.min(23, parseInt(m[3], 10));
        const emm = Math.min(59, parseInt(m[4], 10));

        const start = new Date(baseDate); start.setHours(shh, smm, 0, 0);
        const end = new Date(baseDate); end.setHours(ehh, emm, 0, 0);
        if (end <= start) continue;

        const rs = (roomRowIdx != null)
          ? pickRoomAndSeats(values[roomRowIdx], seg, baseCol)
          : { room: '', seats: '' };

        rows.push({
          group: `${seg.day} — ${rs.room || ''}`,
          label: client || '(TBC)',
          host: host || '',
          room: rs.room || '',
          seats: rs.seats || '',
          start,
          end,
          startMs: start.getTime(),
          endMs: end.getTime(),
          rowIndex: r,                    // row in sheet values array (0-based)
          colIndex: seg.startCol + t      // absolute col of time cell (0-based)
        });
      }
    }
  }

  return {
    title: sheet.getRange(1, 1).getDisplayValue() || 'Week',
    sheetName: sheet.getName(),
    items: rows
  };
}


/**
 * Upsert a daily usage row (by date label in column A)
 */
function _upsertDailyUsageRow_(opts) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(opts.sheetName);
  if (!sh) sh = ss.insertSheet(opts.sheetName);

  // Build desired header
  const header = ['Date'];
  opts.labels.forEach(function(lbl) {
    header.push(lbl + ' (h)', lbl + ' (%)', lbl + ' (#)');
  });

  // Ensure header exists
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else {
    const maxCols = Math.max(sh.getLastColumn(), header.length);
    const currentHeader = sh.getRange(1, 1, 1, maxCols).getValues()[0];
    let mismatch = false;
    for (let i = 0; i < header.length; i++) {
      if ((currentHeader[i] || '') !== header[i]) { mismatch = true; break; }
    }
    if (mismatch) {
      sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  }

  // Build the row to write
  const row = [opts.label];
  opts.labels.forEach(function(lbl) {
    const parts = opts.rowBuilder(lbl) || [0, 0, 0];
    row.push(Number(parts[0] || 0), Number(parts[1] || 0), Number(parts[2] || 0));
  });

  // Upsert by date label in column A
  let lastRow = sh.getLastRow();
  const dataRows = Math.max(0, lastRow - 1);
  let targetRow = -1;

  if (dataRows > 0) {
    const colA = sh.getRange(2, 1, dataRows, 1).getDisplayValues(); // Must use getDisplayValues - getValues() returns Date objects which don't match strings
    const want = String(opts.label || '').trim();
    for (let r = 0; r < dataRows; r++) {
      if (String(colA[r][0] || '').trim() === want) {
        targetRow = r + 2;
        break;
      }
    }
  }

  if (targetRow > 0) {
    sh.getRange(targetRow, 1, 1, row.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }

  // Formatting
  lastRow = sh.getLastRow();
  const dataRowsAfter = Math.max(0, lastRow - 1);
  const totalCols = row.length;

  if (dataRowsAfter > 0) {
    for (let c = 2; c <= totalCols; c += 3) {
      sh.getRange(2, c, dataRowsAfter, 1).setNumberFormat('0.00');   // Hours
    }
    for (let c2 = 3; c2 <= totalCols; c2 += 3) {
      sh.getRange(2, c2, dataRowsAfter, 1).setNumberFormat('0.00%'); // Percent
    }
    for (let c3 = 4; c3 <= totalCols; c3 += 3) {
      sh.getRange(2, c3, dataRowsAfter, 1).setNumberFormat('0');     // Count
    }
  }

  sh.autoResizeColumns(1, totalCols);
}