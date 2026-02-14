/**************************************************
 * DAILY USAGE TRACKING FOR TIMELINES
 * Copy this entire file into your Google Apps Script project
 **************************************************/

/* ==================================================
 * CONSTANTS (Add these if not already in your project)
 * ================================================== */

// If DAY_NAMES is not already defined in another file, uncomment this:
// const DAY_NAMES = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];

// Canonical orders for stable columns
const STUDIO_ORDER_DAILY = ['Studio 1', 'Studio 2', 'Studio 3', 'Studio 4', 'Other'];
const SET_ORDER_DAILY = ['Iris', 'Club', 'Nest', 'Exec', 'Nova', 'Soho', 'Other'];

// Set → Studio mapping (case-insensitive)
const SET_TO_STUDIO_DAILY = {
  'iris': 'Studio 2',
  'club': 'Studio 2',
  'nova': 'Studio 3',
  'nest': 'Studio 1',
  'exec': 'Studio 1',
  'soho': 'Studio 4'
};

const UNKNOWN_STUDIO_DAILY = 'Other';


/* ==================================================
 * PUBLIC MENU FUNCTIONS
 * ================================================== */

/**
 * Write daily studio usage for the active sheet
 */
function writeDailyStudioUsage() {
  const dailyStats = _getDailyStatsForActiveSheet_();

  dailyStats.forEach(dayStat => {
    _upsertDailyUsageRow_({
      sheetName: 'Studio Usage (Daily)',
      label: dayStat.dateLabel,
      labels: STUDIO_ORDER_DAILY,
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
      labels: SET_ORDER_DAILY,
      rowBuilder: function(name) {
        const rec = dayStat.bySet[name] || { hours: 0, count: 0 };
        const pct = dayStat.totalHours > 0 ? (rec.hours / dayStat.totalHours) : 0;
        return [rec.hours, pct, rec.count];
      }
    });
  });

  SpreadsheetApp.getUi().alert('Daily set usage written for ' + dailyStats.length + ' day(s).');
}


/**
 * Backfill daily studio usage from beginning of year to today across all sheets
 */
function backfillDailyStudioUsage() {
  const ss = SpreadsheetApp.getActive();
  const sheets = ss.getSheets();
  const startOfYear = new Date(new Date().getFullYear(), 0, 1); // Jan 1 of current year
  const today = new Date();
  today.setHours(23, 59, 59, 999);

  // Accumulate all data by date first
  const aggregatedByDate = {}; // { 'YYYY-MM-DD': { totalHours, byStudio } }

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();

    // Skip usage/summary sheets
    if (sheetName.indexOf('Usage') >= 0) return;
    if (sheetName.indexOf('Working Hours') >= 0) return;
    if (sheetName.indexOf('Break Plan') >= 0) return;
    if (sheetName.indexOf('Hours Summary') >= 0) return;

    try {
      const dailyStats = _getDailyStatsForSheet_(sheet);

      dailyStats.forEach(dayStat => {
        const dayDate = new Date(dayStat.date);

        // Only process days within range
        if (dayDate >= startOfYear && dayDate <= today) {
          const dateLabel = dayStat.dateLabel;

          // Initialize if first time seeing this date
          if (!aggregatedByDate[dateLabel]) {
            aggregatedByDate[dateLabel] = {
              totalHours: 0,
              byStudio: {}
            };
          }

          // Accumulate hours and counts
          aggregatedByDate[dateLabel].totalHours += dayStat.totalHours;

          Object.keys(dayStat.byStudio).forEach(studioName => {
            if (!aggregatedByDate[dateLabel].byStudio[studioName]) {
              aggregatedByDate[dateLabel].byStudio[studioName] = { hours: 0, count: 0 };
            }
            aggregatedByDate[dateLabel].byStudio[studioName].hours += dayStat.byStudio[studioName].hours;
            aggregatedByDate[dateLabel].byStudio[studioName].count += dayStat.byStudio[studioName].count;
          });
        }
      });
    } catch (e) {
      Logger.log('Skipping sheet ' + sheet.getName() + ': ' + e.message);
    }
  });

  // Now write aggregated data
  const dates = Object.keys(aggregatedByDate).sort();
  dates.forEach(dateLabel => {
    const dayData = aggregatedByDate[dateLabel];

    _upsertDailyUsageRow_({
      sheetName: 'Studio Usage (Daily)',
      label: dateLabel,
      labels: STUDIO_ORDER_DAILY,
      rowBuilder: function(name) {
        const rec = dayData.byStudio[name] || { hours: 0, count: 0 };
        const pct = dayData.totalHours > 0 ? (rec.hours / dayData.totalHours) : 0;
        return [rec.hours, pct, rec.count];
      }
    });
  });

  SpreadsheetApp.getUi().alert('Backfill complete! Processed ' + dates.length + ' unique date(s) from ' + startOfYear.toLocaleDateString() + ' to ' + today.toLocaleDateString());
}


/**
 * Backfill daily set usage from beginning of year to today across all sheets
 */
function backfillDailySetUsage() {
  const ss = SpreadsheetApp.getActive();
  const sheets = ss.getSheets();
  const startOfYear = new Date(new Date().getFullYear(), 0, 1);
  const today = new Date();
  today.setHours(23, 59, 59, 999);

  // Accumulate all data by date first
  const aggregatedByDate = {}; // { 'YYYY-MM-DD': { totalHours, bySet } }

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();

    // Skip usage/summary sheets
    if (sheetName.indexOf('Usage') >= 0) return;
    if (sheetName.indexOf('Working Hours') >= 0) return;
    if (sheetName.indexOf('Break Plan') >= 0) return;
    if (sheetName.indexOf('Hours Summary') >= 0) return;

    try {
      const dailyStats = _getDailyStatsForSheet_(sheet);

      dailyStats.forEach(dayStat => {
        const dayDate = new Date(dayStat.date);

        // Only process days within range
        if (dayDate >= startOfYear && dayDate <= today) {
          const dateLabel = dayStat.dateLabel;

          // Initialize if first time seeing this date
          if (!aggregatedByDate[dateLabel]) {
            aggregatedByDate[dateLabel] = {
              totalHours: 0,
              bySet: {}
            };
          }

          // Accumulate hours and counts
          aggregatedByDate[dateLabel].totalHours += dayStat.totalHours;

          Object.keys(dayStat.bySet).forEach(setName => {
            if (!aggregatedByDate[dateLabel].bySet[setName]) {
              aggregatedByDate[dateLabel].bySet[setName] = { hours: 0, count: 0 };
            }
            aggregatedByDate[dateLabel].bySet[setName].hours += dayStat.bySet[setName].hours;
            aggregatedByDate[dateLabel].bySet[setName].count += dayStat.bySet[setName].count;
          });
        }
      });
    } catch (e) {
      Logger.log('Skipping sheet ' + sheet.getName() + ': ' + e.message);
    }
  });

  // Now write aggregated data
  const dates = Object.keys(aggregatedByDate).sort();
  dates.forEach(dateLabel => {
    const dayData = aggregatedByDate[dateLabel];

    _upsertDailyUsageRow_({
      sheetName: 'Set Usage (Daily)',
      label: dateLabel,
      labels: SET_ORDER_DAILY,
      rowBuilder: function(name) {
        const rec = dayData.bySet[name] || { hours: 0, count: 0 };
        const pct = dayData.totalHours > 0 ? (rec.hours / dayData.totalHours) : 0;
        return [rec.hours, pct, rec.count];
      }
    });
  });

  SpreadsheetApp.getUi().alert('Backfill complete! Processed ' + dates.length + ' unique date(s) from ' + startOfYear.toLocaleDateString() + ' to ' + today.toLocaleDateString());
}


/* ==================================================
 * INTERNAL HELPER FUNCTIONS
 * ================================================== */

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

  // Detect day segments (this function should exist in your Code Timelines.gs)
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

      const studioName = SET_TO_STUDIO_DAILY[setName.toLowerCase()] || UNKNOWN_STUDIO_DAILY;
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
 * NOTE: This is a simplified version - you may need to use getWeeklyTimelinePayload instead
 */
function _getTimelinePayloadForSheet_(sheet) {
  const rng = sheet.getDataRange();
  const values = rng.getDisplayValues();

  if (!values.length || !values[0].length) {
    throw new Error('Sheet looks empty.');
  }

  // Use DAY_NAMES constant (should be defined in your main Code Timelines.gs)
  const DAY_NAMES_LOCAL = (typeof DAY_NAMES !== 'undefined') ? DAY_NAMES : ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];

  const { headerRowIdx, dayCols } = detectDaySegments_(values, 10);

  const a1 = (values[0][0] || '').toString();
  const weekInfo = parseWeekHeader(a1) || fallbackWeekInfo();

  const mondayIndex = DAY_NAMES_LOCAL.indexOf('Monday');
  const dayToDate = {};
  for (const seg of dayCols) {
    const base = new Date(weekInfo.mondayDate);
    base.setDate(base.getDate() + (DAY_NAMES_LOCAL.indexOf(seg.day) - mondayIndex));
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
          endMs: end.getTime()
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
    const colA = sh.getRange(2, 1, dataRows, 1).getValues();
    const want = String(opts.label || '').trim().toLowerCase();
    for (let r = 0; r < dataRows; r++) {
      if (String(colA[r][0] || '').trim().toLowerCase() === want) {
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
