/**************************************************
 * DAILY USAGE TRACKING FOR TIMELINES
 * Copy this entire file into your Google Apps Script project
 **************************************************/

/* ==================================================
 * CONSTANTS (Add these if not already in your project)
 * ================================================== */

// DAY_NAMES constant - required for day detection
// Note: DAY_NAMES is already defined in code.gs, so we don't redefine it here
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
 * Write daily studio usage for YESTERDAY only
 */
function writeDailyStudioUsage() {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  yesterday.setHours(0, 0, 0, 0);

  const stats = _getStatsForSpecificDate_(yesterday);

  if (stats && stats.totalHours > 0) {
    _upsertDailyUsageRow_({
      sheetName: 'Studio Usage (Daily)',
      label: stats.dateLabel,
      labels: STUDIO_ORDER_DAILY,
      rowBuilder: function(name) {
        const rec = stats.byStudio[name] || { hours: 0, count: 0 };
        const pct = stats.totalHours > 0 ? (rec.hours / stats.totalHours) : 0;
        return [rec.hours, pct, rec.count];
      }
    });

    SpreadsheetApp.getUi().alert('Daily studio usage written for ' + stats.dateLabel + ' (' + stats.totalHours.toFixed(2) + ' hours)');
  } else {
    SpreadsheetApp.getUi().alert('No bookings found for ' + _formatDateLabel_(yesterday));
  }
}


/**
 * Write daily set usage for YESTERDAY only
 */
function writeDailySetUsage() {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  yesterday.setHours(0, 0, 0, 0);

  const stats = _getStatsForSpecificDate_(yesterday);

  if (stats && stats.totalHours > 0) {
    _upsertDailyUsageRow_({
      sheetName: 'Set Usage (Daily)',
      label: stats.dateLabel,
      labels: SET_ORDER_DAILY,
      rowBuilder: function(name) {
        const rec = stats.bySet[name] || { hours: 0, count: 0 };
        const pct = stats.totalHours > 0 ? (rec.hours / stats.totalHours) : 0;
        return [rec.hours, pct, rec.count];
      }
    });

    SpreadsheetApp.getUi().alert('Daily set usage written for ' + stats.dateLabel + ' (' + stats.totalHours.toFixed(2) + ' hours)');
  } else {
    SpreadsheetApp.getUi().alert('No bookings found for ' + _formatDateLabel_(yesterday));
  }
}


/**
 * Backfill daily studio usage - ONE MONTH ONLY (to avoid timeout)
 * User will be prompted to select which month
 */
function backfillDailyStudioUsage() {
  const ui = SpreadsheetApp.getUi();

  // Prompt for month
  const response = ui.prompt(
    'Backfill Studio Usage - Select Month',
    'Enter month to backfill (1-12) or "all" for current month:\nExample: 1 for January, 2 for February, etc.',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText().trim().toLowerCase();
  const today = new Date();
  let startDate, endDate;

  if (input === 'all' || input === '') {
    // Current month only
    startDate = new Date(today.getFullYear(), today.getMonth(), 1);
    endDate = new Date(today.getFullYear(), today.getMonth() + 1, 0);
  } else {
    const monthNum = parseInt(input, 10);
    if (isNaN(monthNum) || monthNum < 1 || monthNum > 12) {
      ui.alert('Invalid month number. Please enter 1-12.');
      return;
    }
    startDate = new Date(today.getFullYear(), monthNum - 1, 1);
    endDate = new Date(today.getFullYear(), monthNum, 0);
  }

  // Don't process future dates or today (today isn't finished yet)
  const yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);

  if (endDate > yesterday) endDate = new Date(yesterday);

  // Process each date one by one
  const dates = [];
  const current = new Date(startDate);

  while (current <= endDate) {
    const stats = _getStatsForSpecificDate_(new Date(current));

    if (stats && stats.totalHours > 0) {
      _upsertDailyUsageRow_({
        sheetName: 'Studio Usage (Daily)',
        label: stats.dateLabel,
        labels: STUDIO_ORDER_DAILY,
        rowBuilder: function(name) {
          const rec = stats.byStudio[name] || { hours: 0, count: 0 };
          const pct = stats.totalHours > 0 ? (rec.hours / stats.totalHours) : 0;
          return [rec.hours, pct, rec.count];
        }
      });
      dates.push(stats.dateLabel);
    }

    current.setDate(current.getDate() + 1);
  }

  const monthName = startDate.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
  SpreadsheetApp.getUi().alert('Backfill complete for ' + monthName + '!\nProcessed ' + dates.length + ' date(s) with bookings.');
}


/**
 * Backfill daily set usage - ONE MONTH ONLY (to avoid timeout)
 * User will be prompted to select which month
 */
function backfillDailySetUsage() {
  const ui = SpreadsheetApp.getUi();

  // Prompt for month
  const response = ui.prompt(
    'Backfill Set Usage - Select Month',
    'Enter month to backfill (1-12) or "all" for current month:\nExample: 1 for January, 2 for February, etc.',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText().trim().toLowerCase();
  const today = new Date();
  let startDate, endDate;

  if (input === 'all' || input === '') {
    // Current month only
    startDate = new Date(today.getFullYear(), today.getMonth(), 1);
    endDate = new Date(today.getFullYear(), today.getMonth() + 1, 0);
  } else {
    const monthNum = parseInt(input, 10);
    if (isNaN(monthNum) || monthNum < 1 || monthNum > 12) {
      ui.alert('Invalid month number. Please enter 1-12.');
      return;
    }
    startDate = new Date(today.getFullYear(), monthNum - 1, 1);
    endDate = new Date(today.getFullYear(), monthNum, 0);
  }

  // Don't process future dates or today (today isn't finished yet)
  const yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);

  if (endDate > yesterday) endDate = new Date(yesterday);

  const dates = [];
  const current = new Date(startDate);

  while (current <= endDate) {
    const stats = _getStatsForSpecificDate_(new Date(current));

    if (stats && stats.totalHours > 0) {
      _upsertDailyUsageRow_({
        sheetName: 'Set Usage (Daily)',
        label: stats.dateLabel,
        labels: SET_ORDER_DAILY,
        rowBuilder: function(name) {
          const rec = stats.bySet[name] || { hours: 0, count: 0 };
          const pct = stats.totalHours > 0 ? (rec.hours / stats.totalHours) : 0;
          return [rec.hours, pct, rec.count];
        }
      });
      dates.push(stats.dateLabel);
    }

    current.setDate(current.getDate() + 1);
  }

  const monthName = startDate.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
  SpreadsheetApp.getUi().alert('Backfill complete for ' + monthName + '!\nProcessed ' + dates.length + ' date(s) with bookings.');
}


/**
 * Show monthly studio usage summary (last complete month)
 */
function showMonthlyStudioSummary() {
  const html = _buildMonthlySummaryHtml_('studio');
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(1100).setHeight(750),
    'Monthly Studio Usage'
  );
}


/**
 * Show monthly set usage summary (last complete month)
 */
function showMonthlySetSummary() {
  const html = _buildMonthlySummaryHtml_('set');
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(1100).setHeight(750),
    'Monthly Set Usage'
  );
}


/* ==================================================
 * AUTO-TRIGGER FUNCTIONS
 * ================================================== */

/**
 * Install daily trigger to auto-update yesterday's data at 2am
 * Run this once to set up automatic daily updates
 */
function installDailyTrigger() {
  // Delete existing triggers first
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'autoUpdateYesterdayData') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new trigger for 2am daily
  ScriptApp.newTrigger('autoUpdateYesterdayData')
    .timeBased()
    .atHour(2)
    .everyDays(1)
    .create();

  SpreadsheetApp.getUi().alert('Daily auto-update installed! Will run at 2am every day.');
}


/**
 * Auto-update function called by trigger
 */
function autoUpdateYesterdayData() {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  yesterday.setHours(0, 0, 0, 0);

  const stats = _getStatsForSpecificDate_(yesterday);

  if (stats && stats.totalHours > 0) {
    // Update studio usage
    _upsertDailyUsageRow_({
      sheetName: 'Studio Usage (Daily)',
      label: stats.dateLabel,
      labels: STUDIO_ORDER_DAILY,
      rowBuilder: function(name) {
        const rec = stats.byStudio[name] || { hours: 0, count: 0 };
        const pct = stats.totalHours > 0 ? (rec.hours / stats.totalHours) : 0;
        return [rec.hours, pct, rec.count];
      }
    });

    // Update set usage
    _upsertDailyUsageRow_({
      sheetName: 'Set Usage (Daily)',
      label: stats.dateLabel,
      labels: SET_ORDER_DAILY,
      rowBuilder: function(name) {
        const rec = stats.bySet[name] || { hours: 0, count: 0 };
        const pct = stats.totalHours > 0 ? (rec.hours / stats.totalHours) : 0;
        return [rec.hours, pct, rec.count];
      }
    });

    Logger.log('Auto-updated daily usage for ' + stats.dateLabel);
  }
}


/* ==================================================
 * INTERNAL HELPER FUNCTIONS
 * ================================================== */

/**
 * Get stats for a specific date across all sheets (avoiding duplicates)
 */
function _getStatsForSpecificDate_(targetDate) {
  const ss = SpreadsheetApp.getActive();
  const sheets = ss.getSheets();
  const targetDateStr = _formatDateLabel_(targetDate);

  let totalHours = 0;
  const byStudio = {};
  const bySet = {};
  const seenSessions = {}; // Track unique sessions to avoid duplicates
  let sheetsProcessed = 0;
  let itemsFound = 0;

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();

    // Skip non-timeline sheets
    if (sheetName.indexOf('Usage') >= 0) return;
    if (sheetName.indexOf('Working Hours') >= 0) return;
    if (sheetName.indexOf('Break Plan') >= 0) return;
    if (sheetName.indexOf('Hours Summary') >= 0) return;
    if (sheetName.indexOf('Photo') >= 0) return;
    if (sheetName.indexOf('Integrity') >= 0) return;

    // Skip archived tabs from 2025 or earlier (old date format causes parsing issues)
    // Pattern matches: 2015-2025 (4-digit) or '15-'25 (2-digit abbreviations)
    if (/\b20(1[0-9]|2[0-5])\b|\b'(0[0-9]|1[0-9]|2[0-5])\b/.test(sheetName)) {
      return;  // Skip silently - these are archived tabs
    }

    // Skip tabs with month names but no year ONLY if month is 3+ months in the past
    // Example: "22nd - 28th Sep" in February 2026 → skip (Sep was 5 months ago)
    // Example: "16th - 22nd Feb" in February 2026 → keep (current month)
    const hasYear = /\b20\d{2}\b|\b'\d{2}\b/.test(sheetName);
    const monthMatch = sheetName.match(/\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)/i);

    if (monthMatch && !hasYear) {
      const monthName = monthMatch[1].toLowerCase();
      const monthIndex = {
        jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
        jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11
      }[monthName];

      const now = new Date();
      const currentMonth = now.getMonth();

      // Calculate how many months in the past this is
      let monthsAgo = currentMonth - monthIndex;
      if (monthsAgo < 0) monthsAgo += 12; // Handle year wrap (e.g., Feb seeing Dec)

      if (monthsAgo >= 3) {
        return;  // Skip - tab is from 3+ months ago, likely archived
      }
    }

    try {
      sheetsProcessed++;
      const payload = _getTimelinePayloadForSheet_(sheet);
      const items = (payload && payload.items) ? payload.items : [];

      items.forEach(it => {
        const sessionDate = new Date(it.startMs);
        sessionDate.setHours(0, 0, 0, 0);
        const sessionDateStr = _formatDateLabel_(sessionDate);

        // Only process items for our target date
        if (sessionDateStr !== targetDateStr) return;

        itemsFound++;
        Logger.log('  Found session on sheet "' + sheetName + '" for ' + targetDateStr + ': ' + it.label + ' in ' + it.room + ' (' + it.group + ')');

        // Check if this session is crossed out (strikethrough)
        const isCrossedOut = _isSessionCrossedOut_(sheet, it);
        if (isCrossedOut) return;

        // Create unique session ID to avoid counting duplicates
        // Normalize label and room to catch case/whitespace variations
        const normalizedLabel = String(it.label || '').trim().toLowerCase();
        const normalizedRoom = String(it.room || '').trim().toLowerCase();
        const sessionId = `${sessionDateStr}-${it.startMs}-${it.endMs}-${normalizedRoom}-${normalizedLabel}`;
        if (seenSessions[sessionId]) return; // Skip duplicate
        seenSessions[sessionId] = true;

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
    } catch (e) {
      Logger.log('Error processing sheet ' + sheetName + ' for date ' + targetDateStr + ': ' + e.message);
    }
  });

  Logger.log('Processed ' + sheetsProcessed + ' sheets for ' + targetDateStr + ', found ' + itemsFound + ' items, total hours: ' + totalHours);

  if (totalHours === 0) return null;

  return {
    dateLabel: targetDateStr,
    date: targetDate,
    totalHours: totalHours,
    byStudio: byStudio,
    bySet: bySet
  };
}


/**
 * Check if a session is crossed out (has strikethrough formatting)
 * Checks the specific row where the session was parsed (name + time row)
 */
function _isSessionCrossedOut_(sheet, item) {
  try {
    // If we have row/column info from parsing, check that specific row
    if (item.rowIndex != null && item.colIndex != null) {
      const formats = sheet.getDataRange().getFontLines();
      const r = item.rowIndex;

      // Check strikethrough on cells in the same row near the time
      // Check the name cell (2 cells left of time) and the time cell itself
      const checkCols = [item.colIndex - 2, item.colIndex - 1, item.colIndex];

      for (let i = 0; i < checkCols.length; i++) {
        const c = checkCols[i];
        if (c >= 0 && c < formats[r].length) {
          if (formats[r][c] === 'line-through') {
            Logger.log('  → Session "' + item.label + '" is struck through (row ' + (r+1) + ', col ' + (c+1) + ') - SKIPPING');
            return true;
          }
        }
      }
      return false;
    }

    // Fallback: search entire sheet (old behavior)
    const values = sheet.getDataRange().getDisplayValues();
    const formats = sheet.getDataRange().getFontLines();
    const searchLabel = String(item.label || '').trim().toLowerCase();
    if (!searchLabel) return false;

    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        const cellValue = String(values[r][c] || '').trim().toLowerCase();
        if (cellValue === searchLabel) {
          if (formats[r][c] === 'line-through') {
            Logger.log('  → Session "' + item.label + '" is struck through - SKIPPING');
            return true;
          }
        }
      }
    }
  } catch (e) {
    Logger.log('Error checking strikethrough for "' + item.label + '": ' + e.message);
  }
  return false;
}


/**
 * DEBUG: Test function to see what data exists for a specific date
 */
function debugDateLookup() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Debug Date Lookup', 'Enter date to debug (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const dateStr = response.getResponseText().trim();
  const parts = dateStr.split('-');
  const testDate = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));

  Logger.clear();
  Logger.log('=== DEBUG LOOKUP FOR ' + dateStr + ' ===');

  const stats = _getStatsForSpecificDate_(testDate);

  if (stats) {
    Logger.log('\nFOUND DATA:');
    Logger.log('Total hours: ' + stats.totalHours);
    Logger.log('Studios: ' + JSON.stringify(stats.byStudio));
    Logger.log('Sets: ' + JSON.stringify(stats.bySet));
  } else {
    Logger.log('No data found for this date');
  }

  ui.alert('Check the execution log (Extensions > Apps Script > Executions) for detailed results');
}


/**
 * Build monthly summary HTML popup
 */
function _buildMonthlySummaryHtml_(type) {
  // Get last complete month
  const now = new Date();
  const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const monthStart = new Date(lastMonth.getFullYear(), lastMonth.getMonth(), 1);
  const monthEnd = new Date(lastMonth.getFullYear(), lastMonth.getMonth() + 1, 0);

  const monthName = monthStart.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });

  // Aggregate data for the month
  let totalHours = 0;
  const byCategory = {}; // studio or set

  const current = new Date(monthStart);
  while (current <= monthEnd) {
    const stats = _getStatsForSpecificDate_(new Date(current));

    if (stats) {
      totalHours += stats.totalHours;
      const source = (type === 'studio') ? stats.byStudio : stats.bySet;

      Object.keys(source).forEach(name => {
        if (!byCategory[name]) byCategory[name] = { hours: 0, count: 0 };
        byCategory[name].hours += source[name].hours;
        byCategory[name].count += source[name].count;
      });
    }

    current.setDate(current.getDate() + 1);
  }

  // Calculate working days in month
  const workingDays = _countWorkingDays_(monthStart, monthEnd);

  // Build HTML
  function esc(s) {
    s = (s == null) ? '' : String(s);
    return s.replace(/[&<>"']/g, function(ch) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch];
    });
  }

  const css = `
    <style>
      body { font-family: Arial, sans-serif; margin: 0; background: #f8f9fa; }
      .wrap { padding: 24px; max-width: 900px; margin: 0 auto; }
      .hdr { text-align: center; margin-bottom: 24px; }
      .chip {
        background: linear-gradient(135deg, #1a73e8, #6ea8fe);
        color: #fff;
        padding: 12px 24px;
        border-radius: 999px;
        font-weight: 600;
        font-size: 18px;
        box-shadow: 0 4px 12px rgba(26, 115, 232, 0.3);
        display: inline-block;
      }
      .summary-cards {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 16px;
        margin-bottom: 24px;
      }
      .card {
        background: white;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
        text-align: center;
      }
      .card-label {
        font-size: 14px;
        color: #5f6368;
        margin-bottom: 8px;
      }
      .card-value {
        font-size: 32px;
        font-weight: 700;
        color: #1a73e8;
      }
      .card-unit {
        font-size: 16px;
        color: #80868b;
        margin-left: 4px;
      }
      .table-card {
        background: white;
        border-radius: 12px;
        padding: 24px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
      }
      table {
        width: 100%;
        border-collapse: collapse;
      }
      th, td {
        padding: 12px;
        text-align: left;
        border-bottom: 1px solid #e8eaed;
      }
      th {
        background: linear-gradient(90deg, #1a73e8, #6ea8fe);
        color: white;
        font-weight: 600;
      }
      tr:hover td {
        background: #f8f9fa;
      }
      .bar-cell {
        width: 200px;
      }
      .bar-container {
        background: #e8eaed;
        border-radius: 4px;
        height: 20px;
        overflow: hidden;
      }
      .bar-fill {
        background: linear-gradient(90deg, #34a853, #7bc96f);
        height: 100%;
        transition: width 0.3s ease;
      }
      .actions {
        text-align: center;
        margin-top: 24px;
      }
      button {
        padding: 10px 24px;
        border: none;
        border-radius: 8px;
        background: #1a73e8;
        color: white;
        cursor: pointer;
        font-size: 14px;
        font-weight: 500;
      }
      button:hover {
        background: #1557b0;
      }
    </style>
  `;

  // Summary cards
  let summaryHtml = '<div class="summary-cards">';
  summaryHtml += '<div class="card">';
  summaryHtml += '<div class="card-label">Total Hours</div>';
  summaryHtml += '<div class="card-value">' + totalHours.toFixed(1) + '<span class="card-unit">h</span></div>';
  summaryHtml += '</div>';

  summaryHtml += '<div class="card">';
  summaryHtml += '<div class="card-label">Working Days</div>';
  summaryHtml += '<div class="card-value">' + workingDays + '<span class="card-unit">days</span></div>';
  summaryHtml += '</div>';

  summaryHtml += '<div class="card">';
  summaryHtml += '<div class="card-label">Avg Hours/Day</div>';
  summaryHtml += '<div class="card-value">' + (totalHours / workingDays).toFixed(1) + '<span class="card-unit">h</span></div>';
  summaryHtml += '</div>';

  const totalSessions = Object.values(byCategory).reduce((sum, cat) => sum + cat.count, 0);
  summaryHtml += '<div class="card">';
  summaryHtml += '<div class="card-label">Total Sessions</div>';
  summaryHtml += '<div class="card-value">' + totalSessions + '</div>';
  summaryHtml += '</div>';
  summaryHtml += '</div>';

  // Table
  const rows = Object.keys(byCategory).map(name => ({
    name: name,
    hours: byCategory[name].hours,
    count: byCategory[name].count,
    pct: totalHours > 0 ? (byCategory[name].hours / totalHours) * 100 : 0
  })).sort((a, b) => b.hours - a.hours);

  let tableHtml = '<div class="table-card"><table><thead><tr>';
  tableHtml += '<th>' + (type === 'studio' ? 'Studio' : 'Set') + '</th>';
  tableHtml += '<th>Hours</th>';
  tableHtml += '<th>Sessions</th>';
  tableHtml += '<th>% of Total</th>';
  tableHtml += '<th class="bar-cell">Usage</th>';
  tableHtml += '</tr></thead><tbody>';

  rows.forEach(row => {
    tableHtml += '<tr>';
    tableHtml += '<td><strong>' + esc(row.name) + '</strong></td>';
    tableHtml += '<td>' + row.hours.toFixed(2) + 'h</td>';
    tableHtml += '<td>' + row.count + '</td>';
    tableHtml += '<td>' + row.pct.toFixed(1) + '%</td>';
    tableHtml += '<td class="bar-cell">';
    tableHtml += '<div class="bar-container">';
    tableHtml += '<div class="bar-fill" style="width: ' + row.pct + '%"></div>';
    tableHtml += '</div>';
    tableHtml += '</td>';
    tableHtml += '</tr>';
  });

  tableHtml += '</tbody></table></div>';

  const html = `
    <html>
      <head>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        ${css}
      </head>
      <body>
        <div class="wrap">
          <div class="hdr">
            <div class="chip">${monthName} ${type === 'studio' ? 'Studio' : 'Set'} Usage</div>
          </div>
          ${summaryHtml}
          ${tableHtml}
          <div class="actions">
            <button onclick="google.script.host.close()">Close</button>
          </div>
        </div>
      </body>
    </html>
  `;

  return html;
}


/**
 * Count working days (Mon-Fri) in a date range
 */
function _countWorkingDays_(startDate, endDate) {
  let count = 0;
  const current = new Date(startDate);
  while (current <= endDate) {
    const day = current.getDay();
    if (day !== 0 && day !== 6) { // Not Sunday or Saturday
      count++;
    }
    current.setDate(current.getDate() + 1);
  }
  return count;
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

  if (!values.length || !values[0].length) {
    throw new Error('Sheet looks empty.');
  }

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
    // Only process rows where the FIRST cell contains "Name/Time"
    // This prevents processing duplicate rows for the same session
    const firstCell = String(values[r][0] || '').trim().toLowerCase();
    const isNameTimeRow = /^name\s*\/\s*time/i.test(firstCell);
    if (!isNameTimeRow) continue;

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
          rowIndex: r,  // Store row index for strikethrough checking
          colIndex: seg.startCol + t  // Store column index where time was found
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
    const colA = sh.getRange(2, 1, dataRows, 1).getDisplayValues(); // Use getDisplayValues instead of getValues
    const want = String(opts.label || '').trim();
    for (let r = 0; r < dataRows; r++) {
      const existing = String(colA[r][0] || '').trim();
      if (existing === want) {
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


/* ==================================================
 * REQUIRED HELPER FUNCTIONS FROM CODE TIMELINES.GS
 * (These must be included if not already available)
 * ================================================== */

/**
 * Find the row that contains weekday headers and return both that row index
 * and the [day, startCol, endCol] segments.
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
 * Parse the "6th - 12th October" style header in A1 to get the Monday date for that week.
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

  const mIdx = monthNameToIndex(monthName);
  if (mIdx == null) return null;

  const candidate = new Date(year, mIdx, day1);
  const dow = candidate.getDay();
  const mondayOffset = (dow === 0) ? -6 : (1 - dow);
  const monday = new Date(candidate);
  monday.setDate(monday.getDate() + mondayOffset);

  return { mondayDate: monday, weekStartDay: day1, monthName: monthName };
}


/**
 * Fallback week info = current week
 */
function fallbackWeekInfo() {
  const now = new Date();
  const dow = now.getDay();
  const mondayOffset = (dow === 0) ? -6 : (1 - dow);
  const monday = new Date(now);
  monday.setDate(monday.getDate() + mondayOffset);
  return { mondayDate: monday, weekStartDay: monday.getDate(), monthName: '' };
}


/**
 * Convert month name to month index (0-11)
 */
function monthNameToIndex(monthName) {
  const mn = String(monthName || '').toLowerCase().trim();
  const map = {
    'january': 0, 'jan': 0,
    'february': 1, 'feb': 1,
    'march': 2, 'mar': 2,
    'april': 3, 'apr': 3,
    'may': 4,
    'june': 5, 'jun': 5,
    'july': 6, 'jul': 6,
    'august': 7, 'aug': 7,
    'september': 8, 'sep': 8, 'sept': 8,
    'october': 9, 'oct': 9,
    'november': 10, 'nov': 10,
    'december': 11, 'dec': 11
  };
  return (mn in map) ? map[mn] : null;
}
