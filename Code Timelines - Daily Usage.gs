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
 * DASHBOARD LAYOUT CONSTANTS
 * All hidden data tables start at DASH_DATA_ROW (row 1).
 * Pre-allocate DASH_MAX_MONTHS rows so charts never need range updates —
 * just rewrite the data and charts auto-refresh.
 * ================================================== */

var DASH_DATA_ROW   = 1;   // All hidden data tables start here
var DASH_MAX_MONTHS = 60;  // 5 years of monthly data headroom

// ── Hidden data columns (1-based) ──────────────────────────────────────────
// Individual chart data (5 cols each: Month|Hrs|HrsLabel|Pct|PctLabel)
// Studios (excl. Other): Studio 1=18, Studio 2=23, Studio 3=28, Studio 4=33
var DASH_IV_S_COL_MAP = { 'Studio 1': 18, 'Studio 2': 23, 'Studio 3': 28, 'Studio 4': 33 };
// Sets (excl. Other): Iris=38, Club=43, Nest=48, Exec=53, Nova=58, Soho=63
var DASH_IV_P_COL_MAP = { 'Iris': 38, 'Club': 43, 'Nest': 48, 'Exec': 53, 'Nova': 58, 'Soho': 63 };

// ── Fixed visual chart positions (rows/cols in Dashboard sheet) ─────────────
// Individual charts: ~550px tall → ~30 rows; gap 2
var DASH_IV_CHART_H   = 30;
var DASH_IV_GAP       = 2;
var DASH_IV_ROW_BLOCK = DASH_IV_CHART_H + DASH_IV_GAP;  // 32

// Fixed row anchors — individual charts start near the top (no overview section)
// 4 studio charts → 2 pairs → 2×32 = 64 rows
var DASH_IV_S_SECTION_ROW     = 4;
var DASH_IV_S_CHART_START_ROW = 5;                                    // rows 5-68
// 6 set charts → 3 pairs → 3×32 = 96 rows
var DASH_IV_P_SECTION_ROW     = DASH_IV_S_CHART_START_ROW + 2 * DASH_IV_ROW_BLOCK;  // 69
var DASH_IV_P_CHART_START_ROW = DASH_IV_P_SECTION_ROW + 1;                          // 70

var DASH_LEFT_CHART_COL  = 1;
var DASH_RIGHT_CHART_COL = 9;

// ── Summary charts (total hours/month + hours/week) ─────────────────────────
// Placed below the individual set charts
// 6 set charts → 3 pairs → 3×32 = 96 rows; +2 gap gives section header row
var DASH_SUM_SECTION_ROW  = DASH_IV_P_CHART_START_ROW + 3 * DASH_IV_ROW_BLOCK + 2;  // 168
var DASH_SUM_CHART_ROW    = DASH_SUM_SECTION_ROW + 1;                                // 169
// Hidden columns: monthly = 2 cols, weekly = 2 cols
var DASH_SUM_MONTHLY_COL  = 68;   // cols 68-69: Month | Total Hours
var DASH_SUM_WEEKLY_COL   = 70;   // cols 70-71: Week  | Total Hours
var DASH_SUM_MAX_WEEKS    = 300;  // pre-allocate ~6 years of weekly rows


/* ==================================================
 * PUBLIC MENU FUNCTIONS
 * ================================================== */

// REMOVED: writeDailyStudioUsage() and writeDailySetUsage()
// These functions are no longer needed - use autoUpdateYesterdayData() instead
// which is triggered automatically at 2am daily


/**
 * Remove duplicate date rows from Studio Usage (Daily) and Set Usage (Daily) sheets
 * Keeps the LAST occurrence of each date (most recently written = most accurate)
 */
function deduplicateDailyUsageSheets() {
  const ui = SpreadsheetApp.getUi();
  const sheetNames = ['Studio Usage (Daily)', 'Set Usage (Daily)'];
  const ss = SpreadsheetApp.getActive();
  let report = '';

  sheetNames.forEach(function(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      report += sheetName + ': not found\n';
      return;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      report += sheetName + ': no data rows\n';
      return;
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();

    // Find which rows to delete (keep last occurrence of each date)
    const seen = {};
    const rowsToDelete = [];

    // First pass: find last occurrence of each date
    for (let r = 0; r < data.length; r++) {
      const dateStr = String(data[r][0] || '').trim();
      if (!dateStr) continue;
      seen[dateStr] = r + 2; // 1-indexed, +1 for header
    }

    // Second pass: mark earlier duplicates for deletion
    const lastOccurrence = new Set(Object.values(seen));
    for (let r = 0; r < data.length; r++) {
      const dateStr = String(data[r][0] || '').trim();
      if (!dateStr) continue;
      const rowNum = r + 2;
      if (!lastOccurrence.has(rowNum)) {
        rowsToDelete.push(rowNum);
      }
    }

    // Delete rows in reverse order so indices don't shift
    rowsToDelete.sort((a, b) => b - a);
    rowsToDelete.forEach(function(rowNum) {
      sheet.deleteRow(rowNum);
    });

    report += sheetName + ': removed ' + rowsToDelete.length + ' duplicate rows\n';
  });

  ui.alert('Deduplication Complete\n\n' + report + '\nRun "Backfill" to refresh the cleaned data.');
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
    'Enter month to backfill:\n• MM/YYYY for a specific year (e.g. 9/2025 for Sep 2025)\n• 1-12 for a month in the current year\n• "all" for the current month',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText().trim().toLowerCase();
  const today = new Date();
  let startDate, endDate;

  if (input === 'all' || input === '') {
    startDate = new Date(today.getFullYear(), today.getMonth(), 1);
    endDate = new Date(today.getFullYear(), today.getMonth() + 1, 0);
  } else if (input.indexOf('/') !== -1) {
    const parts = input.split('/');
    const monthNum = parseInt(parts[0], 10);
    const year = parseInt(parts[1], 10);
    if (isNaN(monthNum) || monthNum < 1 || monthNum > 12 || isNaN(year) || year < 2020 || year > today.getFullYear()) {
      ui.alert('Invalid format. Use MM/YYYY, e.g. 9/2025 for September 2025.');
      return;
    }
    startDate = new Date(year, monthNum - 1, 1);
    endDate = new Date(year, monthNum, 0);
  } else {
    const monthNum = parseInt(input, 10);
    if (isNaN(monthNum) || monthNum < 1 || monthNum > 12) {
      ui.alert('Invalid month. Enter 1-12, MM/YYYY, or "all".');
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
    'Enter month to backfill:\n• MM/YYYY for a specific year (e.g. 9/2025 for Sep 2025)\n• 1-12 for a month in the current year\n• "all" for the current month',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText().trim().toLowerCase();
  const today = new Date();
  let startDate, endDate;

  if (input === 'all' || input === '') {
    startDate = new Date(today.getFullYear(), today.getMonth(), 1);
    endDate = new Date(today.getFullYear(), today.getMonth() + 1, 0);
  } else if (input.indexOf('/') !== -1) {
    const parts = input.split('/');
    const monthNum = parseInt(parts[0], 10);
    const year = parseInt(parts[1], 10);
    if (isNaN(monthNum) || monthNum < 1 || monthNum > 12 || isNaN(year) || year < 2020 || year > today.getFullYear()) {
      ui.alert('Invalid format. Use MM/YYYY, e.g. 9/2025 for September 2025.');
      return;
    }
    startDate = new Date(year, monthNum - 1, 1);
    endDate = new Date(year, monthNum, 0);
  } else {
    const monthNum = parseInt(input, 10);
    if (isNaN(monthNum) || monthNum < 1 || monthNum > 12) {
      ui.alert('Invalid month. Enter 1-12, MM/YYYY, or "all".');
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
  const html = _buildMonthlySummaryHtml_('studio', false);
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(1100).setHeight(750),
    'Monthly Studio Usage'
  );
}


/**
 * Show monthly set usage summary (last complete month)
 */
function showMonthlySetSummary() {
  const html = _buildMonthlySummaryHtml_('set', false);
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(1100).setHeight(750),
    'Monthly Set Usage'
  );
}


/**
 * Show monthly studio usage summary EXCLUDING "Other" category
 * Percentages calculated based only on known studios (Studio 1-4)
 */
function showMonthlyStudioSummaryExclOther() {
  const html = _buildMonthlySummaryHtml_('studio', true);
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(1100).setHeight(750),
    'Monthly Studio Usage (Excl. Other)'
  );
}


/**
 * Show monthly set usage summary EXCLUDING "Other" category
 * Percentages calculated based only on known sets (Iris, Club, Nest, Exec, Nova, Soho)
 */
function showMonthlySetSummaryExclOther() {
  const html = _buildMonthlySummaryHtml_('set', true);
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(1100).setHeight(750),
    'Monthly Set Usage (Excl. Other)'
  );
}

/**
 * Send test monthly summary email to ben@poddster.com
 */
function sendTestMonthlySummaryEmail() {
  sendMonthlySummaryEmail();
  SpreadsheetApp.getUi().alert('Test email sent to ben@poddster.com');
}

/**
 * Send monthly summary email with all 4 summaries
 * This should be triggered on the 2nd of each month
 */
function sendMonthlySummaryEmail() {
  const recipient = 'ben@poddster.com';

  // Get last complete month for subject line
  const now = new Date();
  const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const monthName = lastMonth.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });

  const subject = 'Monthly Usage Summary - ' + monthName;

  // Get all 4 summaries
  const studioHtml = _buildMonthlySummaryHtml_('studio', false);
  const studioExclHtml = _buildMonthlySummaryHtml_('studio', true);
  const setHtml = _buildMonthlySummaryHtml_('set', false);
  const setExclHtml = _buildMonthlySummaryHtml_('set', true);

  // Extract CSS from the first summary (all have same CSS)
  const css = _extractCss_(studioHtml);

  // Build email with all styling preserved
  let emailBody = '<html><head><meta charset="UTF-8">';
  emailBody += css; // Include all the CSS
  emailBody += '</head><body style="font-family: Arial, sans-serif; background: #f8f9fa; padding: 20px;">';

  // Add header
  emailBody += '<div style="max-width: 1200px; margin: 0 auto; background: white; padding: 30px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">';
  emailBody += '<h1 style="color: #1a73e8; text-align: center; margin-bottom: 30px;">Monthly Usage Summary - ' + monthName + '</h1>';

  // Add all 4 summaries with their content
  emailBody += _extractBodyContent_(studioHtml);
  emailBody += '<hr style="margin: 40px 0; border: none; border-top: 2px solid #e8eaed;">';

  emailBody += _extractBodyContent_(studioExclHtml);
  emailBody += '<hr style="margin: 40px 0; border: none; border-top: 2px solid #e8eaed;">';

  emailBody += _extractBodyContent_(setHtml);
  emailBody += '<hr style="margin: 40px 0; border: none; border-top: 2px solid #e8eaed;">';

  emailBody += _extractBodyContent_(setExclHtml);

  emailBody += '</div></body></html>';

  // Send email
  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    htmlBody: emailBody
  });
}

/**
 * Helper function to extract CSS from HTML
 */
function _extractCss_(html) {
  // Extract <style>...</style> tags
  const styleMatch = html.match(/<style[^>]*>([\s\S]*?)<\/style>/i);
  if (styleMatch && styleMatch[0]) {
    return styleMatch[0];
  }
  return '';
}

/**
 * Helper function to extract body content from HTML popup
 * Extracts the complete .wrap div (including the div itself) without the Close button
 */
function _extractBodyContent_(html) {
  // Extract the complete .wrap div up to (but not including) the actions div
  const wrapMatch = html.match(/(<div class="wrap">[\s\S]*?)<div class="actions">/i);
  if (wrapMatch && wrapMatch[1]) {
    // Return the .wrap div with closing tag
    return wrapMatch[1] + '</div>';
  }
  // Fallback: extract everything inside <body>...</body>
  const bodyMatch = html.match(/<body[^>]*>([\s\S]*)<\/body>/i);
  if (bodyMatch && bodyMatch[1]) {
    return bodyMatch[1];
  }
  return html;
}

/**
 * Install monthly email trigger for the 2nd of each month at 6am
 */
function installMonthlyEmailTrigger() {
  // Delete existing monthly email triggers first
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendMonthlySummaryEmail') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new trigger for 2nd of month at 6am
  ScriptApp.newTrigger('sendMonthlySummaryEmail')
    .timeBased()
    .onMonthDay(2)
    .atHour(6)
    .create();

  SpreadsheetApp.getUi().alert('Monthly email trigger installed!\n\nWill send summary email on the 2nd of each month at 6am to ben@poddster.com');
}


/* ==================================================
 * DASHBOARD FUNCTIONS
 * ================================================== */

/**
 * SETUP: Full dashboard rebuild — creates all charts from scratch.
 *
 * Run this from the Apps Script editor (not the menu) since it wipes all
 * existing charts and rebuilds layout. After running:
 *   • Right-click each individual chart → Edit chart → Customise → Series
 *     → tick "Data labels". Charts are persistent so you only do this once.
 *
 * Day-to-day data refreshes use updateDashboardData() instead.
 */
function setupDashboard() {
  const ss = SpreadsheetApp.getActive();

  let dashboard = ss.getSheetByName('Dashboard');
  if (!dashboard) {
    dashboard = ss.insertSheet('Dashboard');
  } else {
    dashboard.getCharts().forEach(function(c) { dashboard.removeChart(c); });
    dashboard.clear();
  }
  dashboard.setTabColor('#1a73e8');

  dashboard.getRange('A1').setValue('Studio & Set Usage Dashboard')
    .setFontSize(16).setFontWeight('bold');

  const monthlyData = _getMonthlyAggregatedData_();

  // Write all data to hidden columns first
  _writeDashboardData_(dashboard, monthlyData);

  // Section headers at fixed rows
  _writeSectionHeader_(dashboard, DASH_IV_S_SECTION_ROW, '🏢 Individual Studios');
  _writeSectionHeader_(dashboard, DASH_IV_P_SECTION_ROW, '🎬 Individual Sets');
  _writeSectionHeader_(dashboard, DASH_SUM_SECTION_ROW,  '📈 Summary Totals');

  // Pre-allocated range size (header + 60 months)
  var totalRows = DASH_MAX_MONTHS + 1;

  // ── Individual charts (2-column grid, fixed positions) ─────────────
  var studioLabels = Object.keys(DASH_IV_S_COL_MAP);  // Studio 1-4
  var setLabels    = Object.keys(DASH_IV_P_COL_MAP);   // 6 sets

  _insertIndivCharts_(dashboard, studioLabels, DASH_IV_S_CHART_START_ROW,
    DASH_IV_S_COL_MAP, _getStudioColors_(studioLabels), totalRows);

  _insertIndivCharts_(dashboard, setLabels, DASH_IV_P_CHART_START_ROW,
    DASH_IV_P_COL_MAP, _getSetColors_(setLabels), totalRows);

  // ── Summary charts ────────────────────────────────────────────────
  _insertSummaryCharts_(dashboard);

  SpreadsheetApp.getUi().alert(
    'Dashboard set up successfully! ✅\n\n' +
    '💡 To enable data labels on the individual charts:\n' +
    '   Right-click a chart → Edit chart → Customise → Series → tick "Data labels"\n\n' +
    'You only need to do this once per chart — labels persist through future data updates.'
  );
}

/**
 * DATA UPDATE: Rewrites chart data without touching any charts.
 *
 * Run this on the 1st of each month (or via menu) to add the latest month.
 * Charts auto-refresh. Any chart customisations (data labels, colours etc.)
 * are preserved because no charts are created or removed.
 */
function updateDashboardData() {
  const ss = SpreadsheetApp.getActive();
  var dashboard = ss.getSheetByName('Dashboard');
  if (!dashboard) {
    SpreadsheetApp.getUi().alert(
      'Dashboard sheet not found.\n\nRun setupDashboard() from the Apps Script editor first.'
    );
    return;
  }

  const monthlyData = _getMonthlyAggregatedData_();
  _writeDashboardData_(dashboard, monthlyData);

  SpreadsheetApp.getUi().alert('Dashboard data updated! Charts will refresh automatically.');
}

/**
 * Write ALL data to the fixed hidden column positions.
 * Called by both setupDashboard() and updateDashboardData().
 */
function _writeDashboardData_(dashboard, monthlyData) {
  // Individual chart blocks
  _writeIndivBlock_(dashboard, monthlyData.months, monthlyData.studios, DASH_IV_S_COL_MAP);
  _writeIndivBlock_(dashboard, monthlyData.months, monthlyData.sets,    DASH_IV_P_COL_MAP);

  // Summary totals (monthly + weekly)
  _writeSummaryData_(dashboard, monthlyData);

  // Update "last updated" line
  dashboard.getRange('A2').setValue(
    'Last updated: ' + new Date().toLocaleDateString('en-US', {
      month: 'long', day: 'numeric', year: 'numeric'
    })
  );
}

/**
 * Write a multi-series overview data table to hidden columns.
 * Always uses the canonical label order so the column mapping never changes.
 */
function _writeOverviewBlock_(dashboard, months, data, startCol, metric, canonicalLabels) {
  var numCols = canonicalLabels.length + 1;  // Month + one per label
  var n       = months.length;

  // Header (always overwrite)
  dashboard.getRange(DASH_DATA_ROW, startCol, 1, numCols)
    .setValues([['Month'].concat(canonicalLabels)])
    .setFontWeight('bold');

  // Format: month as text, data as 1 dp
  dashboard.getRange(DASH_DATA_ROW + 1, startCol,     DASH_MAX_MONTHS, 1).setNumberFormat('@');
  dashboard.getRange(DASH_DATA_ROW + 1, startCol + 1, DASH_MAX_MONTHS, canonicalLabels.length)
    .setNumberFormat('0.0');

  // Clear previous data rows then write fresh
  dashboard.getRange(DASH_DATA_ROW + 1, startCol, DASH_MAX_MONTHS, numCols).clearContent();

  if (n > 0) {
    var rows = months.map(function(m) {
      var mk  = m.year + '-' + String(m.month).padStart(2, '0');
      var row = [m.name];
      canonicalLabels.forEach(function(lbl) {
        var val = (metric === 'pct')
          ? (data.percentages[lbl] ? data.percentages[lbl][mk] || 0 : 0)
          : (data.hours[lbl]       ? data.hours[lbl][mk]       || 0 : 0);
        row.push(Math.round(val * 10) / 10);
      });
      return row;
    });
    dashboard.getRange(DASH_DATA_ROW + 1, startCol, n, numCols).setValues(rows);
  }
}

/**
 * Write individual chart data (5 cols each) to hidden columns.
 * colMap: { label: startCol } — fixed column positions per label.
 */
function _writeIndivBlock_(dashboard, months, data, colMap) {
  var n = months.length;
  Object.keys(colMap).forEach(function(label) {
    var dataCol = colMap[label];

    // Header
    dashboard.getRange(DASH_DATA_ROW, dataCol, 1, 5)
      .setValues([['Month', 'Hours', 'HoursLabel', 'Pct', 'PctLabel']])
      .setFontWeight('bold');

    // Number formats (set once; persist even for empty rows)
    dashboard.getRange(DASH_DATA_ROW + 1, dataCol,     DASH_MAX_MONTHS, 1).setNumberFormat('@');
    dashboard.getRange(DASH_DATA_ROW + 1, dataCol + 1, DASH_MAX_MONTHS, 1).setNumberFormat('0.0');
    dashboard.getRange(DASH_DATA_ROW + 1, dataCol + 2, DASH_MAX_MONTHS, 1).setNumberFormat('@');
    dashboard.getRange(DASH_DATA_ROW + 1, dataCol + 3, DASH_MAX_MONTHS, 1).setNumberFormat('0.0');
    dashboard.getRange(DASH_DATA_ROW + 1, dataCol + 4, DASH_MAX_MONTHS, 1).setNumberFormat('@');

    // Clear then write
    dashboard.getRange(DASH_DATA_ROW + 1, dataCol, DASH_MAX_MONTHS, 5).clearContent();

    if (n > 0 && data.hours[label]) {
      var rows = months.map(function(m) {
        var mk  = m.year + '-' + String(m.month).padStart(2, '0');
        var hrs = Math.round((data.hours[label][mk]       || 0) * 10) / 10;
        var pct = Math.round((data.percentages[label][mk] || 0) * 10) / 10;
        return [m.name, hrs, hrs + 'h', pct, pct + '%'];
      });
      dashboard.getRange(DASH_DATA_ROW + 1, dataCol, n, 5).setValues(rows);
    }
  });
}

// Data collection started October 2025 — always chart from here
var DASHBOARD_START_YEAR  = 2025;
var DASHBOARD_START_MONTH = 9; // 0-based: 9 = October

/**
 * Get monthly aggregated data from Daily Usage sheets.
 * Always starts from October 2025 up to last complete month.
 */
function _getMonthlyAggregatedData_() {
  const ss = SpreadsheetApp.getActive();
  const studioSheet = ss.getSheetByName('Studio Usage (Daily)');
  const setSheet = ss.getSheetByName('Set Usage (Daily)');

  if (!studioSheet || !setSheet) {
    throw new Error('Daily Usage sheets not found. Please run backfill first.');
  }

  // Build month list from Oct 2025 to last complete month
  const now = new Date();
  const lastCompleteMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const months = [];
  const cursor = new Date(DASHBOARD_START_YEAR, DASHBOARD_START_MONTH, 1);
  while (cursor <= lastCompleteMonth) {
    months.push({
      date:  new Date(cursor),
      // "Oct-2025" format — dash prevents Sheets auto-parsing as a date
      name:  cursor.toLocaleDateString('en-GB', { month: 'short', year: 'numeric' }).replace(' ', '-'),
      year:  cursor.getFullYear(),
      month: cursor.getMonth()
    });
    cursor.setMonth(cursor.getMonth() + 1);
  }

  const studios = _aggregateByMonth_(studioSheet, months);
  const sets    = _aggregateByMonth_(setSheet,    months);

  return { studios, sets, months };
}

/**
 * Aggregate usage data by month from a Daily Usage sheet.
 * Header format: Date | Label (h) | Label (%) | Label (#) | ...
 */
function _aggregateByMonth_(sheet, months) {
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return { labels: [], percentages: {}, hours: {} };

  const header = data[0];

  // Find "(h)" columns — these hold the hours values
  const labels = [];
  const colMap = {};
  for (let c = 1; c < header.length; c++) {
    const h = String(header[c]).trim();
    if (h.endsWith('(h)')) {
      const label = h.replace(/\s*\(h\)\s*$/, '').trim();
      if (label && label.toLowerCase() !== 'total') {
        labels.push(label);
        colMap[label] = c;
      }
    }
  }

  // Accumulate raw hours by month key (YYYY-MM)
  const hours = {};
  labels.forEach(function(label) { hours[label] = {}; });

  for (let r = 1; r < data.length; r++) {
    const dateStr = String(data[r][0]).trim();
    if (!dateStr) continue;
    const date = new Date(dateStr);
    if (isNaN(date.getTime())) continue;

    const mk = date.getFullYear() + '-' + String(date.getMonth()).padStart(2, '0');
    labels.forEach(function(label) {
      const h = parseFloat(data[r][colMap[label]]) || 0;
      hours[label][mk] = (hours[label][mk] || 0) + h;
    });
  }

  // Calculate percentages per month
  const percentages = {};
  labels.forEach(function(label) { percentages[label] = {}; });

  months.forEach(function(m) {
    const mk = m.year + '-' + String(m.month).padStart(2, '0');
    var total = 0;
    labels.forEach(function(label) { total += hours[label][mk] || 0; });
    if (total > 0) {
      labels.forEach(function(label) {
        percentages[label][mk] = (hours[label][mk] || 0) / total * 100;
      });
    }
  });

  return { labels: labels, percentages: percentages, hours: hours };
}

/* ==================================================
 * SUMMARY CHART FUNCTIONS (total hours/month + hours/week)
 * ================================================== */

/**
 * Add the two summary charts to an existing dashboard.
 * Run this once from the menu — no full rebuild required.
 */
function addSummaryCharts() {
  var ss = SpreadsheetApp.getActive();
  var dashboard = ss.getSheetByName('Dashboard');
  if (!dashboard) {
    SpreadsheetApp.getUi().alert('Dashboard sheet not found. Run setupDashboard() first.');
    return;
  }
  var monthlyData = _getMonthlyAggregatedData_();
  _writeSummaryData_(dashboard, monthlyData);
  _writeSectionHeader_(dashboard, DASH_SUM_SECTION_ROW, '📈 Summary Totals');
  _insertSummaryCharts_(dashboard);
  SpreadsheetApp.getUi().alert('Summary charts added! ✅');
}

/**
 * Write monthly totals and weekly totals to the hidden summary columns.
 * Called by _writeDashboardData_() so runs on every data refresh.
 */
function _writeSummaryData_(dashboard, monthlyData) {
  var months = monthlyData.months;
  var n      = months.length;

  // ── Monthly totals (sum of ALL studio hours per month) ─────────────
  var mCol = DASH_SUM_MONTHLY_COL;
  dashboard.getRange(DASH_DATA_ROW, mCol, 1, 2)
    .setValues([['Month', 'Total Hours']]).setFontWeight('bold');
  dashboard.getRange(DASH_DATA_ROW + 1, mCol,     DASH_MAX_MONTHS, 1).setNumberFormat('@');
  dashboard.getRange(DASH_DATA_ROW + 1, mCol + 1, DASH_MAX_MONTHS, 1).setNumberFormat('0.0');
  dashboard.getRange(DASH_DATA_ROW + 1, mCol, DASH_MAX_MONTHS, 2).clearContent();

  if (n > 0) {
    var mRows = months.map(function(m) {
      var mk    = m.year + '-' + String(m.month).padStart(2, '0');
      var total = 0;
      monthlyData.studios.labels.forEach(function(lbl) {
        total += monthlyData.studios.hours[lbl][mk] || 0;
      });
      return [m.name, Math.round(total * 10) / 10];
    });
    dashboard.getRange(DASH_DATA_ROW + 1, mCol, n, 2).setValues(mRows);
  }

  // ── Weekly totals (from daily usage sheet) ──────────────────────────
  var wCol    = DASH_SUM_WEEKLY_COL;
  var wkRows  = _getWeeklyHoursRows_();
  var wn      = wkRows.length;

  dashboard.getRange(DASH_DATA_ROW, wCol, 1, 2)
    .setValues([['Week (w/c)', 'Total Hours']]).setFontWeight('bold');
  dashboard.getRange(DASH_DATA_ROW + 1, wCol,     DASH_SUM_MAX_WEEKS, 1).setNumberFormat('@');
  dashboard.getRange(DASH_DATA_ROW + 1, wCol + 1, DASH_SUM_MAX_WEEKS, 1).setNumberFormat('0.0');
  dashboard.getRange(DASH_DATA_ROW + 1, wCol, DASH_SUM_MAX_WEEKS, 2).clearContent();

  if (wn > 0) {
    dashboard.getRange(DASH_DATA_ROW + 1, wCol, wn, 2).setValues(wkRows);
  }
}

/**
 * Read Studio Usage (Daily), group daily totals by week (Monday = week start).
 * Returns [[weekLabel, totalHours], ...] sorted oldest → newest.
 */
function _getWeeklyHoursRows_() {
  var ss     = SpreadsheetApp.getActive();
  var sheet  = ss.getSheetByName('Studio Usage (Daily)');
  if (!sheet) return [];

  var data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return [];

  // Find all (h) columns, excluding Total
  var hCols  = [];
  var header = data[0];
  for (var c = 1; c < header.length; c++) {
    var h = String(header[c]).trim();
    if (h.endsWith('(h)') && h.replace(/\s*\(h\)\s*$/, '').trim().toLowerCase() !== 'total') {
      hCols.push(c);
    }
  }

  var MON_ABBR   = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var weekTotals = {};
  var weekOrder  = [];

  for (var r = 1; r < data.length; r++) {
    var dateStr = String(data[r][0]).trim();
    if (!dateStr) continue;
    var date = new Date(dateStr);
    if (isNaN(date.getTime())) continue;

    // Find the Monday of this week
    var dow    = date.getDay();                  // 0=Sun … 6=Sat
    var toMon  = (dow === 0) ? -6 : 1 - dow;
    var mon    = new Date(date);
    mon.setDate(date.getDate() + toMon);

    // Sort key: YYYY-MM-DD of that Monday
    var sk = mon.getFullYear() + '-'
           + String(mon.getMonth() + 1).padStart(2, '0') + '-'
           + String(mon.getDate()).padStart(2, '0');

    var dayTotal = 0;
    hCols.forEach(function(c) { dayTotal += parseFloat(data[r][c]) || 0; });

    if (!weekTotals[sk]) { weekTotals[sk] = 0; weekOrder.push(sk); }
    weekTotals[sk] += dayTotal;
  }

  weekOrder.sort();
  return weekOrder.map(function(sk) {
    var d     = new Date(sk);
    var label = d.getDate() + '-' + MON_ABBR[d.getMonth()] + '-' + String(d.getFullYear()).slice(2);
    return [label, Math.round(weekTotals[sk] * 10) / 10];
  });
}

/**
 * Insert the two summary COLUMN charts at DASH_SUM_CHART_ROW.
 * Monthly chart on the left, weekly on the right.
 */
function _insertSummaryCharts_(dashboard) {
  // Monthly totals
  dashboard.insertChart(
    dashboard.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dashboard.getRange(DASH_DATA_ROW, DASH_SUM_MONTHLY_COL, DASH_MAX_MONTHS + 1, 2))
      .setNumHeaders(1)
      .setPosition(DASH_SUM_CHART_ROW, DASH_LEFT_CHART_COL, 0, 0)
      .setOption('title', 'Total Hours per Month')
      .setOption('titleTextStyle', { fontSize: 13, bold: true })
      .setOption('height', 350)
      .setOption('width', 750)
      .setOption('legend', { position: 'none' })
      .setOption('colors', ['#1a73e8'])
      .setOption('vAxis', { title: 'Hours', minValue: 0 })
      .setOption('hAxis', { slantedText: true, slantedTextAngle: 45 })
      .build()
  );

  // Weekly totals
  dashboard.insertChart(
    dashboard.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dashboard.getRange(DASH_DATA_ROW, DASH_SUM_WEEKLY_COL, DASH_SUM_MAX_WEEKS + 1, 2))
      .setNumHeaders(1)
      .setPosition(DASH_SUM_CHART_ROW, DASH_RIGHT_CHART_COL, 0, 0)
      .setOption('title', 'Total Hours per Week')
      .setOption('titleTextStyle', { fontSize: 13, bold: true })
      .setOption('height', 350)
      .setOption('width', 750)
      .setOption('legend', { position: 'none' })
      .setOption('colors', ['#34a853'])
      .setOption('vAxis', { title: 'Hours', minValue: 0 })
      .setOption('hAxis', { slantedText: true, slantedTextAngle: 45 })
      .build()
  );
}

// ── Dead code below (overview charts removed Mar 2026) ──────────────────────

/**
 * Insert a full-width overview COLUMN chart at a fixed row position.
 * Data is already written in the hidden columns — this just creates the chart.
 */
function _insertOverviewChart_(sheet, dataRange, chartRow, chartCol, title, metric, colors) {
  var vAxisOptions = (metric === 'pct')
    ? { title: '% of total hours', minValue: 0, maxValue: 100 }
    : { title: 'Hours', minValue: 0 };

  sheet.insertChart(
    sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dataRange)
      .setNumHeaders(1)
      .setPosition(chartRow, chartCol, 0, 0)
      .setOption('title', title)
      .setOption('titleTextStyle', { fontSize: 14, bold: true })
      .setOption('height', 380)
      .setOption('width', 900)
      .setOption('isStacked', false)
      .setOption('legend', { position: 'right' })
      .setOption('vAxis', vAxisOptions)
      .setOption('hAxis', { title: 'Month', slantedText: true, slantedTextAngle: 45 })
      .setOption('colors', colors)
      .build()
  );
}

/**
 * Insert individual dual-axis COLUMN charts in a 2-column grid at fixed positions.
 * Each chart: Hours (left axis) + % of total (right axis).
 * Data is already written in the hidden columns — this just creates the charts.
 * totalRows: header + DASH_MAX_MONTHS (pre-allocated so range never needs updating).
 */
function _insertIndivCharts_(sheet, labels, startRow, colMap, colors, totalRows) {
  labels.forEach(function(label, idx) {
    var dataCol    = colMap[label];
    if (!dataCol) return;

    var isLeft     = (idx % 2 === 0);
    var chartRow   = startRow + Math.floor(idx / 2) * DASH_IV_ROW_BLOCK;
    var chartCol   = isLeft ? DASH_LEFT_CHART_COL : DASH_RIGHT_CHART_COL;
    var solidColor = colors[idx] || '#4285f4';
    var lightColor = _lightenColor_(solidColor, 0.45);
    var dataRange  = sheet.getRange(DASH_DATA_ROW, dataCol, totalRows, 5);

    sheet.insertChart(
      sheet.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(dataRange)
        .setNumHeaders(1)
        .setPosition(chartRow, chartCol, 0, 0)
        .setOption('title', label + ' — Hours & % of Total')
        .setOption('titleTextStyle', { fontSize: 13, bold: true, color: solidColor })
        .setOption('height', 550)
        .setOption('width', 800)
        .setOption('legend', { position: 'bottom' })
        .setOption('colors', [solidColor, lightColor])
        .setOption('annotations', {
          alwaysOutside: true,
          textStyle: { fontSize: 10, bold: true, color: '#333333', auraColor: 'none' },
          stem: { color: 'transparent', length: 4 }
        })
        .setOption('series', {
          0: { targetAxisIndex: 0 },
          1: { targetAxisIndex: 1 }
        })
        .setOption('vAxes', {
          0: { title: 'Hours', minValue: 0 },
          1: { title: '% of total', minValue: 0, maxValue: 100 }
        })
        .setOption('hAxis', { slantedText: true, slantedTextAngle: 45 })
        .build()
    );
  });
}

/**
 * Write a bold section heading into a single merged cell row.
 */
function _writeSectionHeader_(sheet, row, text) {
  if (row < 1) return;
  sheet.getRange(row, 1, 1, 12)
    .merge()
    .setValue(text)
    .setFontSize(13)
    .setFontWeight('bold')
    .setBackground('#e8f0fe')
    .setFontColor('#1a73e8');
}

/**
 * Lighten a hex colour towards white by `factor` (0 = unchanged, 1 = white).
 */
function _lightenColor_(hex, factor) {
  var r = parseInt(hex.slice(1, 3), 16);
  var g = parseInt(hex.slice(3, 5), 16);
  var b = parseInt(hex.slice(5, 7), 16);
  r = Math.round(r + (255 - r) * factor);
  g = Math.round(g + (255 - g) * factor);
  b = Math.round(b + (255 - b) * factor);
  return '#' + ('0' + r.toString(16)).slice(-2)
             + ('0' + g.toString(16)).slice(-2)
             + ('0' + b.toString(16)).slice(-2);
}

// _createIndividualCharts_ removed — replaced by _insertIndivCharts_() above.
// This stub prevents "not defined" errors if old triggers still reference it.
function _createIndividualCharts_(sheet, months, data, startRow, colors, labels) {
  if (!labels || labels.length === 0) return startRow;

  var n          = months.length;
  var DATA_ROWS  = n + 1;   // header + one row per month (for range ref only)
  var CHART_ROWS = 38;      // rows a 550px-tall chart occupies (~21px/row)
  var GAP        = 1;
  var ROW_BLOCK  = CHART_ROWS + GAP;  // no visible data table rows in the block

  // Chart anchor columns (1-based)
  var LEFT_CHART_COL  = 1;
  var RIGHT_CHART_COL = 9;

  // Hidden data columns (far right, off visible dashboard area)
  // 5 columns each: Month | Hours | HoursLabel | Pct | PctLabel
  var LEFT_DATA_COL  = 18;
  var RIGHT_DATA_COL = 24;

  var leftRow  = startRow;
  var rightRow = startRow;

  labels.forEach(function(label, idx) {
    var isLeft   = (idx % 2 === 0);
    var chartRow = isLeft ? leftRow  : rightRow;
    var chartCol = isLeft ? LEFT_CHART_COL  : RIGHT_CHART_COL;
    var dataRow  = isLeft ? leftRow  : rightRow;
    var dataCol  = isLeft ? LEFT_DATA_COL   : RIGHT_DATA_COL;

    var solidColor = colors[idx] || '#4285f4';
    var lightColor = _lightenColor_(solidColor, 0.45);

    // ------------------------------------------------------------------
    // Data table (5 cols, hidden far right):
    //   Month | Hours | "X.Xh" | % | "X.X%"
    //
    // The string columns immediately following each numeric column are
    // treated by Google Charts as annotation (data label) columns —
    // labels appear above bars automatically, no manual steps needed.
    // ------------------------------------------------------------------
    sheet.getRange(dataRow, dataCol, 1, 5)
      .setValues([['Month', 'Hours', 'HoursLabel', 'Pct', 'PctLabel']])
      .setFontWeight('bold');

    // Force text format BEFORE writing so Sheets doesn't coerce strings to numbers
    sheet.getRange(dataRow + 1, dataCol,     n, 1).setNumberFormat('@');   // Month
    sheet.getRange(dataRow + 1, dataCol + 1, n, 1).setNumberFormat('0.0'); // Hours
    sheet.getRange(dataRow + 1, dataCol + 2, n, 1).setNumberFormat('@');   // HoursLabel
    sheet.getRange(dataRow + 1, dataCol + 3, n, 1).setNumberFormat('0.0'); // Pct
    sheet.getRange(dataRow + 1, dataCol + 4, n, 1).setNumberFormat('@');   // PctLabel

    var rows = months.map(function(m) {
      var mk  = m.year + '-' + String(m.month).padStart(2, '0');
      var hrs = Math.round((data.hours[label][mk]       || 0) * 10) / 10;
      var pct = Math.round((data.percentages[label][mk] || 0) * 10) / 10;
      return [m.name, hrs, hrs + 'h', pct, pct + '%'];
    });
    sheet.getRange(dataRow + 1, dataCol, n, 5).setValues(rows);

    // ------------------------------------------------------------------
    // Dual-axis grouped COLUMN chart
    //   Series 0 (Hours)       → left  Y-axis, auto-labelled from col 3
    //   Series 1 (% of total)  → right Y-axis, auto-labelled from col 5
    // ------------------------------------------------------------------
    var dataRange = sheet.getRange(dataRow, dataCol, DATA_ROWS, 5);

    var chart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dataRange)
      .setNumHeaders(1)
      .setPosition(chartRow, chartCol, 0, 0)
      .setOption('title', label + ' — Hours & % of Total')
      .setOption('titleTextStyle', { fontSize: 13, bold: true, color: solidColor })
      .setOption('height', 550)
      .setOption('width', 800)
      .setOption('legend', { position: 'bottom' })
      .setOption('colors', [solidColor, lightColor])
      // Chart-level annotation styling applies to all annotation columns
      .setOption('annotations', {
        alwaysOutside: true,
        textStyle: { fontSize: 10, bold: true, color: '#333333', auraColor: 'none' },
        stem: { color: 'transparent', length: 4 }
      })
      .setOption('series', {
        0: { targetAxisIndex: 0 },
        1: { targetAxisIndex: 1 }
      })
      .setOption('vAxes', {
        0: { title: 'Hours', minValue: 0 },
        1: { title: '% of total', minValue: 0, maxValue: 100 }
      })
      .setOption('hAxis', { slantedText: true, slantedTextAngle: 45 })
      .build();

    sheet.insertChart(chart);

    if (isLeft) {
      rightRow = leftRow;
    } else {
      leftRow  += ROW_BLOCK;
      rightRow += ROW_BLOCK;
    }
  });

  if (labels.length % 2 !== 0) leftRow += ROW_BLOCK;

  return Math.max(leftRow, rightRow) + 2;
}

/**
 * Get consistent colors for studios
 */
function _getStudioColors_(labels) {
  const colorMap = {
    'Studio 1': '#34a853',  // Green
    'Studio 2': '#4285f4',  // Blue
    'Studio 3': '#fbbc04',  // Yellow
    'Studio 4': '#ea4335',  // Red
    'Other': '#9e9e9e'      // Gray
  };

  return labels.map(label => colorMap[label] || '#666666');
}

/**
 * Get consistent colors for sets
 */
function _getSetColors_(labels) {
  const colorMap = {
    'Iris': '#9c27b0',     // Purple
    'Club': '#00bcd4',     // Cyan
    'Nest': '#4caf50',     // Green
    'Exec': '#ff9800',     // Orange
    'Nova': '#f44336',     // Red
    'Soho': '#3f51b5',     // Indigo
    'Other': '#9e9e9e'     // Gray
  };

  return labels.map(label => colorMap[label] || '#666666');
}

/**
 * Install monthly dashboard data-update trigger for 1st of month at 6am.
 * This triggers updateDashboardData() which refreshes data without
 * touching charts — preserving any manual customisations (e.g. data labels).
 */
function installDashboardTrigger() {
  // Remove any old triggers for both the old and new handler names
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    var fn = trigger.getHandlerFunction();
    if (fn === 'updateDashboard' || fn === 'updateDashboardData') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('updateDashboardData')
    .timeBased()
    .onMonthDay(1)
    .atHour(6)
    .create();

  SpreadsheetApp.getUi().alert(
    'Dashboard trigger installed!\n\nWill refresh dashboard data on the 1st of each month at 6am.\n' +
    'Charts are not rebuilt — all your customisations (data labels etc.) are preserved.'
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

  // Refresh dashboard data every night so weekly/monthly charts stay current.
  // Runs silently (no UI alert) since this is a background trigger.
  var dashboard = SpreadsheetApp.getActive().getSheetByName('Dashboard');
  if (dashboard) {
    try {
      var monthlyData = _getMonthlyAggregatedData_();
      _writeDashboardData_(dashboard, monthlyData);
      Logger.log('Dashboard data refreshed by daily trigger.');
    } catch (e) {
      Logger.log('Dashboard refresh skipped: ' + e.message);
    }
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

    // Skip non-timeline sheets (system/utility sheets)
    if (sheetName.indexOf('Usage') >= 0) return;
    if (sheetName.indexOf('Working Hours') >= 0) return;
    if (sheetName.indexOf('Break Plan') >= 0) return;
    if (sheetName.indexOf('Hours Summary') >= 0) return;
    if (sheetName.indexOf('Photo') >= 0) return;
    if (sheetName.indexOf('Integrity') >= 0) return;
    if (sheetName.indexOf('Chart') >= 0) return;
    if (sheetName.indexOf('Error Log') >= 0) return;

    // NO TAB NAME FILTERING - dates are parsed from Row 2 only
    // The _parseDateFromRow2Cell_ function will determine if sessions belong to target date

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
        const roomName = String(it.room || '').trim();
        if (!roomName) return;

        // Map room to studio (Event → Other, Iris → Studio 2, etc.)
        const studioName = SET_TO_STUDIO_DAILY[roomName.toLowerCase()] || UNKNOWN_STUDIO_DAILY;

        // Map room to set category (Event → Other, unknown rooms → Other)
        // Known sets: Iris, Club, Nest, Exec, Nova, Soho
        const knownSets = ['iris', 'club', 'nest', 'exec', 'nova', 'soho'];
        const setName = knownSets.includes(roomName.toLowerCase()) ? roomName : 'Other';

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
 * @param {string} type - 'studio' or 'set'
 * @param {boolean} excludeOther - If true, exclude "Other" category from totals and percentages
 */
function _buildMonthlySummaryHtml_(type, excludeOther) {
  excludeOther = excludeOther || false;

  // Get last complete month
  const now = new Date();
  const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const monthStart = new Date(lastMonth.getFullYear(), lastMonth.getMonth(), 1);
  const monthEnd = new Date(lastMonth.getFullYear(), lastMonth.getMonth() + 1, 0);

  const monthName = monthStart.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });

  // READ FROM DAILY USAGE SHEETS (much faster than re-parsing timeline sheets!)
  const sheetName = (type === 'studio') ? 'Studio Usage (Daily)' : 'Set Usage (Daily)';
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('Sheet "' + sheetName + '" not found. Please run backfill first.');
  }

  const data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) {
    throw new Error('No data found in "' + sheetName + '". Please run backfill first.');
  }

  // Parse header to find column positions
  const header = data[0];
  const colMap = {};
  for (let c = 0; c < header.length; c++) {
    colMap[header[c]] = c;
  }

  // Aggregate data for the month.
  // Use getFullYear()/getMonth() key matching (same as dashboard) rather than
  // date-range comparison — new Date("YYYY-MM-DD") parses as UTC midnight which
  // causes the last day of the month to fail a <= monthEnd check in UTC+8/UTC+9.
  const targetMk = lastMonth.getFullYear() + '-' + String(lastMonth.getMonth()).padStart(2, '0');

  let totalHours = 0;
  let totalHoursExclOther = 0; // Total excluding "Other"
  const byCategory = {}; // studio or set
  const labels = (type === 'studio') ? STUDIO_ORDER_DAILY : SET_ORDER_DAILY;

  // Process each row
  for (let r = 1; r < data.length; r++) {
    const dateStr = data[r][0]; // First column is date (YYYY-MM-DD)
    if (!dateStr) continue;

    // Match on year+month using local time (getMonth() = 0-based) to avoid
    // timezone edge cases on the first/last day of the month.
    const rowDate = new Date(dateStr);
    const rowMk = rowDate.getFullYear() + '-' + String(rowDate.getMonth()).padStart(2, '0');
    if (rowMk === targetMk) {
      // Sum up hours and counts for each category
      labels.forEach(label => {
        const hoursCol = colMap[label + ' (h)'];
        const countCol = colMap[label + ' (#)'];

        if (hoursCol !== undefined && countCol !== undefined) {
          const hours = parseFloat(data[r][hoursCol]) || 0;
          const count = parseInt(data[r][countCol]) || 0;

          if (!byCategory[label]) byCategory[label] = { hours: 0, count: 0 };
          byCategory[label].hours += hours;
          byCategory[label].count += count;
          totalHours += hours;

          // Track total excluding "Other"
          if (label !== 'Other') {
            totalHoursExclOther += hours;
          }
        }
      });
    }
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
        display: -webkit-box;
        display: -ms-flexbox;
        display: flex;
        -webkit-box-orient: horizontal;
        -webkit-box-direction: normal;
        -ms-flex-direction: row;
        flex-direction: row;
        -ms-flex-wrap: wrap;
        flex-wrap: wrap;
        gap: 16px;
        margin-bottom: 24px;
      }
      .card {
        background: white;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
        text-align: center;
        -webkit-box-flex: 1;
        -ms-flex: 1 1 200px;
        flex: 1 1 200px;
        min-width: 200px;
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

  // Determine which total to use for calculations
  const hoursForCalc = excludeOther ? totalHoursExclOther : totalHours;

  // Summary cards
  let summaryHtml = '<div class="summary-cards">';
  summaryHtml += '<div class="card">';
  summaryHtml += '<div class="card-label">Total Hours' + (excludeOther ? ' (Excl. Other)' : '') + '</div>';
  summaryHtml += '<div class="card-value">' + hoursForCalc.toFixed(1) + '<span class="card-unit">h</span></div>';
  summaryHtml += '</div>';

  summaryHtml += '<div class="card">';
  summaryHtml += '<div class="card-label">Working Days</div>';
  summaryHtml += '<div class="card-value">' + workingDays + '<span class="card-unit">days</span></div>';
  summaryHtml += '</div>';

  summaryHtml += '<div class="card">';
  summaryHtml += '<div class="card-label">Avg Hours/Day</div>';
  summaryHtml += '<div class="card-value">' + (hoursForCalc / workingDays).toFixed(1) + '<span class="card-unit">h</span></div>';
  summaryHtml += '</div>';

  // Calculate total sessions - exclude "Other" if needed
  let totalSessions = 0;
  Object.keys(byCategory).forEach(function(label) {
    if (excludeOther && label === 'Other') return;
    totalSessions += byCategory[label].count;
  });

  summaryHtml += '<div class="card">';
  summaryHtml += '<div class="card-label">Total Sessions</div>';
  summaryHtml += '<div class="card-value">' + totalSessions + '</div>';
  summaryHtml += '</div>';
  summaryHtml += '</div>';

  // Table - filter out "Other" if excludeOther is true
  let categoryKeys = Object.keys(byCategory);
  if (excludeOther) {
    categoryKeys = categoryKeys.filter(name => name !== 'Other');
  }

  const rows = categoryKeys.map(name => ({
    name: name,
    hours: byCategory[name].hours,
    count: byCategory[name].count,
    pct: hoursForCalc > 0 ? (byCategory[name].hours / hoursForCalc) * 100 : 0
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

  const titleSuffix = excludeOther ? ' (Excl. Other)' : '';
  const html = `
    <html>
      <head>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        ${css}
      </head>
      <body>
        <div class="wrap">
          <div class="hdr">
            <div class="chip">${monthName} ${type === 'studio' ? 'Studio' : 'Set'} Usage${titleSuffix}</div>
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


/* ==================================================
 * NOTE: Timeline helper functions are defined in code.gs
 *
 * The following functions are used by this file but
 * defined in code.gs to avoid duplicate declarations:
 *
 * - detectDaySegments_(values, maxScanRows)
 * - parseWeekHeader(txt)
 * - fallbackWeekInfo()
 * - monthNameToIndex(monthName)
 * - _parseDateFromDayHeader_(cellText)
 * - _formatDateLabel_(date)
 * - _extractDayFromGroup_(groupStr)
 * - _getTimelinePayloadForSheet_(sheet)
 * - _upsertDailyUsageRow_(opts)
 * ================================================== */
