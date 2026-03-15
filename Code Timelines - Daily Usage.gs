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
 * Update dashboard with charts - run on 1st of each month
 */
function updateDashboard() {
  const ss = SpreadsheetApp.getActive();

  // Get or create Dashboard sheet
  let dashboard = ss.getSheetByName('Dashboard');
  if (!dashboard) {
    dashboard = ss.insertSheet('Dashboard');
  } else {
    // Remove all existing charts before clearing data
    dashboard.getCharts().forEach(function(c) { dashboard.removeChart(c); });
    dashboard.clear();
  }

  // Set up dashboard
  dashboard.setTabColor('#1a73e8'); // Blue color

  // Add title
  dashboard.getRange('A1').setValue('Studio & Set Usage Dashboard');
  dashboard.getRange('A1').setFontSize(16).setFontWeight('bold');
  dashboard.getRange('A2').setValue('Last updated: ' + new Date().toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' }));

  // Build monthly aggregation data
  const monthlyData = _getMonthlyAggregatedData_();

  const months = monthlyData.months;

  // Layout: each chart block = data table (numMonths+1 rows) + 1 spacer + chart (20 rows) + 2 spacers
  // Row anchors (1-based): title=1, subtitle=2, gap=3
  // Studios % : data at row 4,  chart at row 4+numMonths+2
  // Studios h : data at row 4+numMonths+2+20+2, chart after that
  // Sets %    : ...
  // Sets h    : ...
  const n = months.length;
  const CHART_H = 22;  // rows a chart occupies
  const BLOCK   = (n + 1) + 1 + CHART_H + 2; // data + gap + chart + gap

  // ── Section 1: Overview — all studios together ───────────────────
  var R = 4;
  _writeSectionHeader_(dashboard, R - 1, '📊 Studio Overview');
  _createTrendBlock_(dashboard, months, monthlyData.studios, R,
    'Studio Usage — % of Total Hours', 'pct', _getStudioColors_(monthlyData.studios.labels));
  R += BLOCK;
  _createTrendBlock_(dashboard, months, monthlyData.studios, R,
    'Studio Usage — Hours per Month', 'h', _getStudioColors_(monthlyData.studios.labels));
  R += BLOCK;

  // ── Section 2: Overview — all sets together ──────────────────────
  _writeSectionHeader_(dashboard, R - 1, '📊 Set Overview');
  _createTrendBlock_(dashboard, months, monthlyData.sets, R,
    'Set Usage — % of Total Hours', 'pct', _getSetColors_(monthlyData.sets.labels));
  R += BLOCK;
  _createTrendBlock_(dashboard, months, monthlyData.sets, R,
    'Set Usage — Hours per Month', 'h', _getSetColors_(monthlyData.sets.labels));
  R += BLOCK;

  // ── Section 3: Individual studio charts (2-column grid) ──────────
  _writeSectionHeader_(dashboard, R - 1, '🏢 Individual Studios');
  R = _createIndividualCharts_(dashboard, months, monthlyData.studios, R,
    _getStudioColors_(monthlyData.studios.labels), monthlyData.studios.labels);
  R += 2;

  // ── Section 4: Individual set charts (2-column grid) ─────────────
  _writeSectionHeader_(dashboard, R - 1, '🎬 Individual Sets');
  _createIndividualCharts_(dashboard, months, monthlyData.sets, R,
    _getSetColors_(monthlyData.sets.labels), monthlyData.sets.labels);

  SpreadsheetApp.getUi().alert('Dashboard updated successfully!');
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

/**
 * Write a data table + chart for one trend block.
 *
 * @param {Sheet}  sheet    - Dashboard sheet
 * @param {Array}  months   - Array of month objects from _getMonthlyAggregatedData_
 * @param {Object} data     - { labels, percentages, hours } from _aggregateByMonth_
 * @param {number} startRow - 1-based row where the data table starts
 * @param {string} title    - Chart title
 * @param {string} metric   - 'pct' for percentage, 'h' for hours
 * @param {Array}  colors   - Array of hex colour strings
 */
function _createTrendBlock_(sheet, months, data, startRow, title, metric, colors) {
  var labels = data.labels;
  if (!labels || labels.length === 0) return;

  var numCols = labels.length + 1; // Month col + one col per label
  var numRows = months.length + 1; // Header row + one row per month

  // ------------------------------------------------------------------
  // 1. Write header row
  // ------------------------------------------------------------------
  var headerVals = [['Month'].concat(labels)];
  sheet.getRange(startRow, 1, 1, numCols).setValues(headerVals).setFontWeight('bold');

  // ------------------------------------------------------------------
  // 2. Force month column to TEXT so Sheets doesn't parse it as a date
  // ------------------------------------------------------------------
  sheet.getRange(startRow + 1, 1, months.length, 1).setNumberFormat('@');

  // ------------------------------------------------------------------
  // 3. Write data rows
  // ------------------------------------------------------------------
  var dataRows = months.map(function(m) {
    var mk = m.year + '-' + String(m.month).padStart(2, '0');
    var row = [m.name];
    labels.forEach(function(label) {
      if (metric === 'pct') {
        row.push(Math.round((data.percentages[label][mk] || 0) * 10) / 10);
      } else {
        row.push(Math.round((data.hours[label][mk] || 0) * 100) / 100);
      }
    });
    return row;
  });
  sheet.getRange(startRow + 1, 1, months.length, numCols).setValues(dataRows);

  // Format numeric columns to 1 decimal place
  sheet.getRange(startRow + 1, 2, months.length, labels.length)
       .setNumberFormat('0.0');

  // ------------------------------------------------------------------
  // 4. Create chart from the data table
  // ------------------------------------------------------------------
  var tableRange   = sheet.getRange(startRow, 1, numRows, numCols);
  var chartRow     = startRow + numRows + 1;   // one blank row gap
  var vAxisOptions = metric === 'pct'
    ? { title: '% of total hours', minValue: 0, maxValue: 100 }
    : { title: 'Hours', minValue: 0 };

  var chartBuilder = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(tableRange)
    .setNumHeaders(1)
    .setPosition(chartRow, 1, 0, 0)
    .setOption('title', title)
    .setOption('titleTextStyle', { fontSize: 14, bold: true })
    .setOption('height', 380)
    .setOption('width', 900)
    .setOption('isStacked', false)
    .setOption('legend', { position: 'right' })
    .setOption('vAxis', vAxisOptions)
    .setOption('hAxis', { title: 'Month', slantedText: true, slantedTextAngle: 45 })
    .setOption('colors', colors);

  sheet.insertChart(chartBuilder.build());
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
 * Create one COLUMN chart per label arranged in a 2-column grid.
 * Each chart has its own small data table (Month | Hours | Annotation)
 * so that Google Charts shows the value on top of each bar.
 * Returns the next available row after all charts.
 */
function _createIndividualCharts_(sheet, months, data, startRow, colors, labels) {
  if (!labels || labels.length === 0) return startRow;

  var n          = months.length;
  var DATA_ROWS  = n + 1;        // header + one row per month
  var CHART_ROWS = 22;           // chart height in rows
  var GAP        = 2;
  var ROW_BLOCK  = DATA_ROWS + 1 + CHART_ROWS + GAP;  // full block height per row of charts

  var LEFT_COL  = 1;
  var RIGHT_COL = 8;  // cols 8-10 for right-hand tables/charts

  var leftRow  = startRow;
  var rightRow = startRow;

  labels.forEach(function(label, idx) {
    var isLeft   = (idx % 2 === 0);
    var tableRow = isLeft ? leftRow : rightRow;
    var tableCol = isLeft ? LEFT_COL : RIGHT_COL;

    // ------------------------------------------------------------------
    // Data table: 3 columns — Month (text) | Hours (num) | Annotation (str)
    // Google Charts treats a string column immediately after a number column
    // in the same range as an annotation → value appears above the bar.
    // ------------------------------------------------------------------
    var labelColor = colors[idx] || '#4285f4';

    // Header row (annotation header = '' so no legend entry for it)
    sheet.getRange(tableRow, tableCol, 1, 3)
      .setValues([['Month', label, '']])
      .setFontWeight('bold');

    // Force month column to plain text so Sheets doesn't parse dates.
    // Force annotation column to text BEFORE writing — otherwise Sheets
    // coerces "69.0" → 69.0 (number) and the chart plots it as a second
    // data series instead of using it as a bar label.
    sheet.getRange(tableRow + 1, tableCol, n, 1).setNumberFormat('@');
    sheet.getRange(tableRow + 1, tableCol + 2, n, 1).setNumberFormat('@');

    // Data rows: [month-label, hours-number, "X.Xh" annotation string]
    // The 'h' suffix is a second guard so Sheets can never parse the
    // annotation as a number even if the format slips back to Automatic.
    var rows = months.map(function(m) {
      var mk  = m.year + '-' + String(m.month).padStart(2, '0');
      var val = Math.round((data.hours[label][mk] || 0) * 10) / 10;
      return [m.name, val, val.toFixed(1) + 'h'];
    });
    sheet.getRange(tableRow + 1, tableCol, n, 3).setValues(rows);
    sheet.getRange(tableRow + 1, tableCol + 1, n, 1).setNumberFormat('0.0');

    // ------------------------------------------------------------------
    // Column chart using the full 3-column range
    // ------------------------------------------------------------------
    var chartRow  = tableRow + DATA_ROWS + 1;
    var fullRange = sheet.getRange(tableRow, tableCol, DATA_ROWS, 3);

    var chart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(fullRange)
      .setNumHeaders(1)
      .setPosition(chartRow, tableCol, 0, 0)
      .setOption('title', label + ' — Hours per Month')
      .setOption('titleTextStyle', { fontSize: 12, bold: true, color: labelColor })
      .setOption('height', 320)
      .setOption('width', 460)
      .setOption('legend', { position: 'none' })
      .setOption('colors', [labelColor])
      .setOption('vAxis', { title: 'Hours', minValue: 0 })
      .setOption('hAxis', { slantedText: true, slantedTextAngle: 45 })
      .setOption('annotations', {
        alwaysOutside: false,
        textStyle: { fontSize: 10, bold: true, color: '#333333' },
        stem: { color: 'transparent' }
      })
      .build();

    sheet.insertChart(chart);

    // Advance rows: after filling the right slot, move both cursors down one block
    if (isLeft) {
      // right-side row will be set on next iteration (same block start)
      rightRow = leftRow;
    } else {
      leftRow  += ROW_BLOCK;
      rightRow += ROW_BLOCK;
    }
  });

  // If odd number of labels, left side still needs to advance
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
 * Install monthly dashboard update trigger for 1st of month at 6am
 */
function installDashboardTrigger() {
  // Delete existing dashboard triggers first
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'updateDashboard') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new trigger for 1st of month at 6am
  ScriptApp.newTrigger('updateDashboard')
    .timeBased()
    .onMonthDay(1)
    .atHour(6)
    .create();

  SpreadsheetApp.getUi().alert('Dashboard trigger installed!\n\nWill update dashboard on the 1st of each month at 6am');
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

  // Aggregate data for the month
  let totalHours = 0;
  let totalHoursExclOther = 0; // Total excluding "Other"
  const byCategory = {}; // studio or set
  const labels = (type === 'studio') ? STUDIO_ORDER_DAILY : SET_ORDER_DAILY;

  // Process each row
  for (let r = 1; r < data.length; r++) {
    const dateStr = data[r][0]; // First column is date (YYYY-MM-DD)
    if (!dateStr) continue;

    // Check if date is within the target month
    const rowDate = new Date(dateStr);
    if (rowDate >= monthStart && rowDate <= monthEnd) {
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
