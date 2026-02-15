/** Editor Assigner v1.3 (popup break plan, grouped by day)
 * - Reads "Working Hours" (A=Name, B=Start, C=End) from DISPLAY values
 * - Active weekly tab: assigns every "TBC" using time two cells to the RIGHT (e.g., "11:00 - 12:00")
 * - Respects shifts, NO BACK-TO-BACK (top rule), optional buffer minutes
 * - Enforces a required daily break on long shifts (placed closest to mid-shift)
 * - Menu: Assign • Undo • Hours Summary • Break Plan (popup)
 */

/* ==================================================
 * CONSTANTS (shared with daily usage tracking)
 * ================================================== */

// Canonical orders for stable columns
const STUDIO_ORDER = ['Studio 1', 'Studio 2', 'Studio 3', 'Studio 4', 'Other'];
const SET_ORDER = ['Iris', 'Club', 'Nest', 'Exec', 'Nova', 'Soho', 'Other'];

// Set → Studio mapping (case-insensitive)
const SET_TO_STUDIO = {
  'iris': 'Studio 2',
  'club': 'Studio 2',
  'nova': 'Studio 3',
  'nest': 'Studio 1',
  'exec': 'Studio 1',
  'soho': 'Studio 4'
};

const UNKNOWN_STUDIO = 'Other';


/** Master onOpen for Poddster workbook */
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();

  // --- Assign Operators / Editor Assigner menu ---
  var assignMenu = ui.createMenu('🎨 Assign Operators');
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
    .addSeparator()
    .addItem('🔄 Backfill Daily Studio Usage', 'backfillDailyStudioUsage')
    .addItem('🔄 Backfill Daily Set Usage', 'backfillDailySetUsage')
    .addSeparator()
    .addItem('📊 Monthly Studio Summary', 'showMonthlyStudioSummary')
    .addItem('📊 Monthly Set Summary', 'showMonthlySetSummary')
    .addItem('📊 Monthly Studio Summary (Excl. Other)', 'showMonthlyStudioSummaryExclOther')
    .addItem('📊 Monthly Set Summary (Excl. Other)', 'showMonthlySetSummaryExclOther')
    .addSeparator()
    .addItem('📧 Send Test Monthly Summary Email', 'sendTestMonthlySummaryEmail')
    .addSeparator()
    .addItem('📈 Update Dashboard', 'updateDashboard')
    .addSeparator()
    .addItem('⚙️ Install Daily Auto-Update (2am)', 'installDailyTrigger')
    .addItem('⚙️ Install Monthly Email (2nd of month, 6am)', 'installMonthlyEmailTrigger')
    .addItem('⚙️ Install Dashboard Update (1st of month, 6am)', 'installDashboardTrigger')
    .addToUi();

  // --- Timelines menu ---
  ui.createMenu('🕐 Timelines')
    .addItem('Show me the room timelines!', 'openTimeline')
    .addItem('Show me the operator timelines!', 'openOperatorTimeline')
    .addToUi();

  // --- Session Data Integrity Checker menu ---
  ui.createMenu('🎬 Poddster Session Data Integrity Checker')
    .addItem("Check Today's Data", 'runSanityToday')
    .addItem("Check Yesterday's Data", 'runSanityYesterday')
    .addItem('Check Entire Tab', 'runSanityWholeTab')
    .addSeparator()
    .addItem('Send Test Email', 'sendSanityTestEmail')
    .addToUi();

  // --- Photo JPG Helper menu ---
  ui.createMenu('📸 Photo JPG Helper')
    .addItem('Create JPG folders for PHOTO sessions', 'createJpgFoldersFromBoard')
    .addToUi();
}

const CFG = {
  WORKING_HOURS_SHEET: 'Working Hours',
  COL_WH: { NAME: 1, START: 2, END: 3 }, // A/B/C on Working Hours
  TIME_SCAN_MAX_OFFSET: 2,
  COLORIZE: true,
  PERSON_COLORS: ['#E6F4EA','#E8F0FE','#FFF4E5','#FDE7E9','#E7F5FE','#EAF7E7'],
  IGNORE_ZERO_TIMES: true,
  DAY_HEADER_SEARCH_ROWS: 200,
  // Prefer a 1h break near mid-shift when choosing who gets a slot
  PRIORITIZE_MID_BREAK: false,
  // Weekend controls
  INCLUDE_WEEKENDS_IN_CAPACITY: false,
  SHOW_WEEKENDS_IN_TABLES: true,
  // Back-to-back rule (highest priority)
  AVOID_BACK_TO_BACK: true,
  AVOID_BACK_TO_BACK_STRICT: true, // if only back-to-back choices exist, leave TBC unassigned
  MIN_GAP_MINUTES: 0,              // >0 to require a buffer (e.g., 5 minutes)
  // Required daily break (mid-shift placement)
  REQUIRE_BREAK: true,                     // enforce a daily break
  BREAK_DURATION_MIN: 60,                  // change to 90 for 1.5h
  APPLY_BREAK_WHEN_SHIFT_MIN_AT_LEAST: 540, // only enforce break if shift >= 9h (540 min)
  // Soft mixing preference (spread shoots across operators in the same day)
  MIX_OPERATORS_SOFT: true,   // turn off if you don't want this behaviour
  CLUSTER_WINDOW_MIN: 60,     // treat slots within this many minutes as "back-to-back-ish"
};

/* =========================
 * MAIN – assign whole sheet
 * ========================= */
function assignEditorsOnActiveSheet() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  var values = sheet.getDataRange().getValues();
  var people = loadPeople_();
  if (people.length === 0) throw new Error('No editors found on "' + CFG.WORKING_HOURS_SHEET + '".');

  // usage minutes per person
  var usedTotal = {};
  for (var i = 0; i < people.length; i++) usedTotal[people[i].name] = 0;

  // overlap windows per day/section context (e.g., each day column)
  var windowsByContext = {};
  for (var j = 0; j < people.length; j++) windowsByContext[people[j].name] = {};

  // colors
  var colorMap = {};
  for (var k = 0; k < people.length; k++) {
    var nm = people[k].name;
    colorMap[nm] = CFG.PERSON_COLORS[k] || hashPastel_(nm);
  }

  var assignedCount = 0;
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[0].length; c++) {
      var cellRaw = values[r][c];
      var txt = String(cellRaw != null ? cellRaw : '').trim().toLowerCase();
      if (txt !== 'tbc') continue;

      var timeInfo = findTimeOnRow_(values[r], c);
      if (!timeInfo) continue;
      var startMin = timeInfo.startMin, endMin = timeInfo.endMin;
      if (CFG.IGNORE_ZERO_TIMES && startMin === 0 && endMin === 0) continue;
      if (endMin <= startMin) continue;

      var contextKey = findContextKey_(sheet, r, c, values);
      var picked = pickEditor_(people, startMin, endMin, usedTotal, windowsByContext, contextKey);
      if (!picked) continue;

      sheet.getRange(r + 1, c + 1).setValue(picked.name);
      if (CFG.COLORIZE && colorMap[picked.name]) {
        sheet.getRange(r + 1, c + 1).setBackground(colorMap[picked.name]);
      }
      usedTotal[picked.name] += (endMin - startMin);
      if (!windowsByContext[picked.name][contextKey]) windowsByContext[picked.name][contextKey] = [];
      windowsByContext[picked.name][contextKey].push({ s: startMin, e: endMin });
      assignedCount++;
    }
  }
  SpreadsheetApp.getUi().alert('Assigned ' + assignedCount + ' slot(s) on "' + sheet.getName() + '".');
}

/* ==========================
 * Undo – revert names to TBC
 * ========================== */
function undoAssignedOnActiveSheet() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  var values = sheet.getDataRange().getValues();
  var people = loadPeople_();
  var names = {};
  for (var i = 0; i < people.length; i++) names[people[i].name.toLowerCase()] = true;

  var reverted = 0;
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[0].length; c++) {
      var v = String(values[r][c] != null ? values[r][c] : '').trim().toLowerCase();
      if (!names[v]) continue;
      var t = findTimeOnRow_(values[r], c);
      if (!t) continue;
      sheet.getRange(r + 1, c + 1).setValue('TBC').setBackground(null);
      reverted++;
    }
  }
  SpreadsheetApp.getUi().alert('Reverted ' + reverted + ' cell(s) to "TBC" on "' + sheet.getName() + '".');
}

function showHoursSummaryDialogOnActiveSheet() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  var values = sheet.getDataRange().getValues();
  var people = loadPeople_();
  if (people.length === 0) {
    SpreadsheetApp.getUi().alert('No editors found on "' + CFG.WORKING_HOURS_SHEET + '".');
    return;
  }

  // ---- config for overhead (setup + teardown) ----
  var SHOOT_OVERHEAD_MIN = 30; // 0.5h per distinct shoot (contiguous block)

  // Maps and shift lengths
  var namesLow = {}, prettyName = {}, shiftLenMin = {};
  for (var i = 0; i < people.length; i++) {
    var nm = people[i].name;
    namesLow[nm.toLowerCase()] = true;
    prettyName[nm.toLowerCase()] = nm;
    shiftLenMin[nm] = (people[i].endMin - people[i].startMin);
  }

  // Used minutes + per-day minutes + per-day windows (for shoot counting)
  var totalMin = {}, perDayMin = {}, perDayWins = {};
  for (var j = 0; j < people.length; j++) {
    var n = people[j].name;
    totalMin[n] = 0;
    perDayMin[n] = {};
    perDayWins[n] = {};
  }

  // Collect assignments and day labels
  var daySeenRaw = {};
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[0].length; c++) {
      var cell = String(values[r][c] != null ? values[r][c] : '').trim();
      var key = cell.toLowerCase();
      if (!namesLow[key]) continue;
      var ti = findTimeOnRow_(values[r], c);
      if (!ti) continue;
      if (CFG.IGNORE_ZERO_TIMES && ti.startMin === 0 && ti.endMin === 0) continue;
      if (ti.endMin <= ti.startMin) continue;
      var minutes = ti.endMin - ti.startMin;
      var ctx = findContextKey_(sheet, r, c, values);
      var dayLabel = extractDayLabel_(ctx);
      var nm = prettyName[key];
      totalMin[nm] += minutes;
      if (!perDayMin[nm][dayLabel]) perDayMin[nm][dayLabel] = 0;
      perDayMin[nm][dayLabel] += minutes;
      if (!perDayWins[nm][dayLabel]) perDayWins[nm][dayLabel] = [];
      perDayWins[nm][dayLabel].push({ s: ti.startMin, e: ti.endMin });
      daySeenRaw[dayLabel] = true;
    }
  }

  // Build ordered day lists
  var days = [];
  var seenLC = {};
  for (var k in daySeenRaw) {
    if (!isRealDayLabel_(k)) continue;
    var lc = String(k).toLowerCase();
    if (!seenLC[lc]) { days.push(k); seenLC[lc] = true; }
  }
  if (days.length === 0) days = detectWeekdayHeadersInTopRows_(values);
  days.sort(function(a,b){ return dayOrder_(a) - dayOrder_(b); });

  var daysForCapacity = CFG.INCLUDE_WEEKENDS_IN_CAPACITY
    ? days.slice()
    : days.filter(function(d){ return !isWeekendLabel_(d); });
  var daysForTables = CFG.SHOW_WEEKENDS_IN_TABLES
    ? days.slice()
    : days.filter(function(d){ return !isWeekendLabel_(d); });

  // ---- Overhead calculation: per contiguous block per person/day ----
  function countShoots(winArr){
    if (!winArr || !winArr.length) return 0;
    winArr = winArr.slice().sort(function(a,b){ return a.s - b.s; });
    var merged = [{ s: winArr[0].s, e: winArr[0].e }];
    for (var i = 1; i < winArr.length; i++) {
      var cur = winArr[i], last = merged[merged.length-1];
      if (cur.s <= last.e) { // overlap or touch → same shoot
        if (cur.e > last.e) last.e = cur.e;
      } else {
        merged.push({ s: cur.s, e: cur.e });
      }
    }
    return merged.length;
  }

  var perDayOverMin = {}, overheadTotalMin = {};
  for (var p = 0; p < people.length; p++) {
    var nmP = people[p].name;
    perDayOverMin[nmP] = {};
    overheadTotalMin[nmP] = 0;
    for (var d = 0; d < days.length; d++) {
      var dayLbl = days[d];
      var shoots = countShoots(perDayWins[nmP][dayLbl] || []);
      var over = shoots * SHOOT_OVERHEAD_MIN;
      perDayOverMin[nmP][dayLbl] = over;
      overheadTotalMin[nmP] += over;
    }
  }

  // Break-aware capacity math (same as before)
  var requireBreak = !!CFG.REQUIRE_BREAK;
  var reqBreakMin  = requireBreak ? Number(CFG.BREAK_DURATION_MIN || 0) : 0;
  var threshold    = Number(CFG.APPLY_BREAK_WHEN_SHIFT_MIN_AT_LEAST || 0);

  function netPerDay(nm) {
    var raw = shiftLenMin[nm] || 0;
    return (requireBreak && raw >= threshold) ? Math.max(0, raw - reqBreakMin) : raw;
  }

  var weeklyCapMin = {};
  for (var i2 = 0; i2 < people.length; i2++) {
    var nm2 = people[i2].name;
    weeklyCapMin[nm2] = netPerDay(nm2) * (daysForCapacity.length || 0);
  }

  // ------- HTML build (same style as Break Plan) -------
  function esc(s){ s=(s==null)?'':String(s); return s.replace(/[&<>"']/g, function(ch){return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]);}); }

  var css =
    '<style>body{font-family:Arial,sans-serif;margin:0}.wrap{padding:18px}'
    +'.hdr{display:flex;justify-content:center;margin:0 0 14px}'
    +'.chip{background:linear-gradient(135deg,#1a73e8,#6ea8fe);color:#fff;padding:10px 16px;border-radius:999px;font-weight:600;box-shadow:0 2px 6px rgba(0,0,0,.12)}'
    +'.card{background:#fff;border:1px solid #e5e7eb;border-radius:12px;padding:12px;margin-bottom:14px}'
    +'.card.hilite{background:#f8fafc}'
    +'.card h3{margin:4px 0 8px;font-size:14px;color:#202124}'
    +'.tbl{border-collapse:collapse;width:100%;font-size:13px}'
    +'.tbl th,.tbl td{border:1px solid #e0e0e0;padding:8px}'
    +'.tbl th{background:#2f70df;color:#fff;text-align:left}'
    +'.tbl .subhdr th{background:#eef2ff;color:#1f2937}'
    +'.tbl tr:nth-child(even) td{background:#fafafa}'
    +'.actions{text-align:right;margin-top:8px}'
    +'.actions button{padding:8px 12px;border:none;border-radius:8px;background:#1a73e8;color:#fff;cursor:pointer}'
    +'</style>';

  // Table 1: Summary (Used includes overhead)
  var t1 = '<div class="card hilite"><table class="tbl"><thead><tr>'
    + '<th>Name</th><th>Shift (h/day)</th><th>Days</th><th>Capacity (h)</th><th>Used (h)</th><th>Utilization %</th>'
    + '</tr></thead><tbody>';
  for (var x = 0; x < people.length; x++) {
    var nm3   = people[x].name;
    var shiftH= (shiftLenMin[nm3] || 0) / 60;
    var daysCnt = daysForCapacity.length || 0;
    var capH  = (weeklyCapMin[nm3] || 0) / 60;
    var usedH = ( (totalMin[nm3] || 0) + (overheadTotalMin[nm3] || 0) ) / 60;  // <-- includes overhead
    var util  = capH > 0 ? (usedH / capH) * 100 : 0;
    t1 += '<tr>'
       + '<td>'+esc(nm3)+'</td>'
       + '<td>'+shiftH.toFixed(2)+'</td>'
       + '<td>'+daysCnt+'</td>'
       + '<td>'+capH.toFixed(2)+'</td>'
       + '<td>'+usedH.toFixed(2)+'</td>'
       + '<td>'+util.toFixed(2)+'</td>'
       + '</tr>';
  }
  t1 += '</tbody></table></div>';

  // Table 2: Per-day hours (raw scheduled, no overhead)
  var t2 = '<div class="card"><h3>Per-day hours</h3><table class="tbl"><thead><tr class="subhdr"><th>Name</th>';
  for (var d0 = 0; d0 < daysForTables.length; d0++) t2 += '<th>'+esc(daysForTables[d0])+'</th>';
  t2 += '</tr></thead><tbody>';
  for (var y = 0; y < people.length; y++) {
    var nm4 = people[y].name;
    t2 += '<tr><td>'+esc(nm4)+'</td>';
    for (var d1 = 0; d1 < daysForTables.length; d1++) {
      var lbl = daysForTables[d1];
      var mins = (perDayMin[nm4] && perDayMin[nm4][lbl]) ? perDayMin[nm4][lbl] : 0;
      t2 += '<td>'+(mins/60).toFixed(2)+'</td>';
    }
    t2 += '</tr>';
  }
  t2 += '</tbody></table></div>';

  // Table 3: Per-day utilization % (uses overhead)
  var t3 = '<div class="card"><h3>Per-day utilization %</h3><table class="tbl"><thead><tr class="subhdr"><th>Name</th>';
  for (var d2 = 0; d2 < daysForTables.length; d2++) t3 += '<th>'+esc(daysForTables[d2])+'</th>';
  t3 += '</tr></thead><tbody>';
  for (var z = 0; z < people.length; z++) {
    var nm5 = people[z].name;
    t3 += '<tr><td>'+esc(nm5)+'</td>';
    var denom = netPerDay(nm5) || 0;
    for (var d3 = 0; d3 < daysForTables.length; d3++) {
      var lbl2 = daysForTables[d3];
      var mins2 = ((perDayMin[nm5] && perDayMin[nm5][lbl2]) ? perDayMin[nm5][lbl2] : 0)
                + ((perDayOverMin[nm5] && perDayOverMin[nm5][lbl2]) ? perDayOverMin[nm5][lbl2] : 0); // <-- add overhead
      var pct = denom > 0 ? (mins2 / denom) * 100 : 0;
      t3 += '<td>'+pct.toFixed(2)+'</td>';
    }
    t3 += '</tr>';
  }
  t3 += '</tbody></table></div>';

  var html =
    '<html><head><meta name="viewport" content="width=device-width, initial-scale=1">'+css+'</head>'
    + '<body><div class="wrap">'
    + '<div class="hdr"><div class="chip">Hours Summary</div></div>'
    + t1 + t2 + t3
    + '<div class="actions"><button onclick="google.script.host.close()">Close</button></div>'
    + '</div></body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(1100).setHeight(750),
    'Hours Summary'
  );
}

// Back-compat alias if an older menu still points here.
function generateBreakPlanOnActiveSheet(){ return showBreakPlanDialogOnActiveSheet(); }

/* -----------------------------
 * Choose a concrete break slot (midpoint-optimised)
 * -----------------------------
 * Returns { possible:bool, s:number, e:number, dist:number }
 */
function computeBreakPlacement_(person, wins, reqBreakMin) {
  reqBreakMin = Number(reqBreakMin || 0);
  // Build merged occupied windows within the shift
  wins = (wins || []).slice().sort(function(a,b){ return a.s - b.s; });
  var merged = [];
  for (var i = 0; i < wins.length; i++) {
    var cur = wins[i];
    if (!merged.length) merged.push({ s: cur.s, e: cur.e });
    else {
      var last = merged[merged.length - 1];
      if (cur.s <= last.e) { if (cur.e > last.e) last.e = cur.e; }
      else merged.push({ s: cur.s, e: cur.e });
    }
  }

  // Free gaps
  var gaps = [];
  var prev = person.startMin;
  for (var j = 0; j < merged.length; j++) {
    var occ = merged[j];
    if (occ.s > prev) gaps.push({ s: prev, e: occ.s });
    if (occ.e > prev) prev = occ.e;
  }
  if (person.endMin > prev) gaps.push({ s: prev, e: person.endMin });

  var mid = (person.startMin + person.endMin) / 2;
  if (!reqBreakMin) {
    var anyGap = gaps.length ? gaps[0] : { s: person.startMin, e: person.endMin };
    var start = Math.max(anyGap.s, Math.min(anyGap.e, Math.floor(mid)));
    return { possible: true, s: start, e: start, dist: 0 };
  }

  function candidateIn(a, b) {
    if (b - a < reqBreakMin) return null;
    var idealStart = Math.floor(mid - reqBreakMin / 2);
    var start = Math.max(a, Math.min(b - reqBreakMin, idealStart));
    var center = start + reqBreakMin / 2;
    var dist = Math.abs(center - mid);
    return { s: start, e: start + reqBreakMin, dist: dist };
  }

  var best = null;
  for (var g = 0; g < gaps.length; g++) {
    var c = candidateIn(gaps[g].s, gaps[g].e);
    if (!c) continue;
    if (!best || c.dist < best.dist || (c.dist === best.dist && c.s < best.s)) best = c;
  }
  return best ? { possible: true, s: best.s, e: best.e, dist: best.dist } : { possible: false };
}

/* ================
 * Assign heuristics
 * ================ */
function pickEditor_(people, startMin, endMin, usedTotal, windowsByContext, contextKey) {
  var GAP = Number(CFG.MIN_GAP_MINUTES || 0);
  var preferMid = !!CFG.PRIORITIZE_MID_BREAK;
  var requireBreakGlobal = !!CFG.REQUIRE_BREAK;
  var reqBreakMin = requireBreakGlobal ? Number(CFG.BREAK_DURATION_MIN || 0) : 0;
  var threshold = Number(CFG.APPLY_BREAK_WHEN_SHIFT_MIN_AT_LEAST || 0);
  var mixSoft = !!CFG.MIX_OPERATORS_SOFT;
  var CLUST = Number(CFG.CLUSTER_WINDOW_MIN || 0);

  function overlaps(win) { return (startMin < win.e && endMin > win.s); }
  function adjacent(win) {
    if (GAP > 0) {
      var startTooClose = (startMin >= win.e) && (startMin < win.e + GAP);
      var endTooClose   = (win.s   >= endMin) && (win.s   < endMin + GAP);
      return startTooClose || endTooClose;
    }
    return (startMin === win.e || endMin === win.s);
  }

  function clusterScore(wins) {
    if (!mixSoft || CLUST <= 0 || !wins || !wins.length) return 0;
    var score = 0;
    for (var i = 0; i < wins.length; i++) {
      var w = wins[i];
      var gapL = startMin - w.e;  // how long since their last ended before this
      var gapR = w.s - endMin;    // how long until their next starts after this
      if (gapL >= 0 && gapL < CLUST) score++;       // near a slot before
      if (gapR >= 0 && gapR < CLUST) score++;       // near a slot after
    }
    return score; // 0, 1 or 2 typically
  }

  var good = [], adj = [];
  for (var i = 0; i < people.length; i++) {
    var p = people[i];
    // must be within shift
    if (startMin < p.startMin || endMin > p.endMin) continue;
    var wins = windowsByContext[p.name][contextKey] || [];
    var hasOverlap = false, hasAdjacency = false;
    for (var j = 0; j < wins.length; j++) {
      var w = wins[j];
      if (overlaps(w)) { hasOverlap = true; break; }
      if (CFG.AVOID_BACK_TO_BACK && adjacent(w)) { hasAdjacency = true; }
    }
    if (hasOverlap) continue;

    // Break feasibility + mid distance (for PRIORITIZE_MID_BREAK)
    var enforceBreak = requireBreakGlobal && ((p.endMin - p.startMin) >= threshold);
    var midDist = null;
    if (enforceBreak) {
      var winsPlus = wins.slice();
      winsPlus.push({ s: startMin, e: endMin });
      var res = computeBreakPlacement_(p, winsPlus, reqBreakMin);
      if (!res || !res.possible) continue;         // no full break possible → reject
      midDist = (res.dist != null) ? res.dist : 0; // minutes from mid-shift
    }

    // Soft-mix metrics
    var softCluster = clusterScore(wins); // prefer smaller
    var shootsToday = wins.length;        // prefer fewer

    // Fairness figures
    var remaining = (p.endMin - p.startMin) - (usedTotal[p.name] || 0);
    var used = (usedTotal[p.name] || 0);

    var cand = {
      person: p,
      hasAdjacency: hasAdjacency,
      enforceBreak: enforceBreak,
      midDist: midDist,
      softCluster: softCluster,
      shootsToday: shootsToday,
      remaining: remaining,
      used: used
    };
    if (CFG.AVOID_BACK_TO_BACK && hasAdjacency) adj.push(cand);
    else good.push(cand);
  }

  var pool = good.length ? good : (CFG.AVOID_BACK_TO_BACK_STRICT ? [] : adj);
  if (!pool.length) return null;

  pool.sort(function(a, b) {
    // 1) break-midpoint preference (if enabled)
    if (preferMid) {
      if (a.enforceBreak !== b.enforceBreak) return a.enforceBreak ? -1 : 1;
      if (a.enforceBreak && b.enforceBreak) {
        var dA = (a.midDist == null ? 1e9 : a.midDist);
        var dB = (b.midDist == null ? 1e9 : b.midDist);
        if (dA !== dB) return dA - dB; // closer to mid is better
      }
    }
    // 2) soft mixing (spread operators within the day)
    if (mixSoft) {
      if (a.softCluster !== b.softCluster) return a.softCluster - b.softCluster; // fewer near-neighbours
      if (a.shootsToday !== b.shootsToday) return a.shootsToday - b.shootsToday; // fewer shoots today
    }
    // 3) original fairness
    if (b.remaining !== a.remaining) return b.remaining - a.remaining;
    if (a.used !== b.used) return a.used - b.used;
    return a.person.name.localeCompare(b.person.name);
  });
  return pool[0].person;
}

/* ================
 * Helper functions
 * ================ */
function loadPeople_() {
  var sh = getWorkingHoursSheet_();
  if (!sh) throw new Error('Sheet matching "' + CFG.WORKING_HOURS_SHEET + '" not found.');
  var rng = sh.getDataRange();
  var vals = rng.getValues();
  var disp = rng.getDisplayValues();
  if (vals.length < 2) return [];
  var people = [];
  for (var i = 1; i < vals.length; i++) {
    var name = String((disp[i][CFG.COL_WH.NAME - 1] || vals[i][CFG.COL_WH.NAME - 1]) || '').trim();
    if (!name) continue;
    var startTxt = String(disp[i][CFG.COL_WH.START - 1] || '').trim();
    var endTxt   = String(disp[i][CFG.COL_WH.END   - 1] || '').trim();
    var startMin = toMinutesFromText_(startTxt);
    var endMin   = toMinutesFromText_(endTxt);
    // support combined "09:00 - 18:00" in Start if End empty
    if ((isNaN(startMin) || isNaN(endMin)) && startTxt && !endTxt) {
      var comb = parseRangeFromAny_(startTxt);
      if (comb) { startMin = comb.startMin; endMin = comb.endMin; }
    }
    if (isNaN(startMin) || isNaN(endMin)) continue;
    people.push({ name: name, startMin: startMin, endMin: endMin });
  }
  return people;
}

function getWorkingHoursSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(CFG.WORKING_HOURS_SHEET);
  if (sh) return sh;
  var target = CFG.WORKING_HOURS_SHEET.toLowerCase().replace(/\s+/g,'').trim();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var nm = sheets[i].getName().toLowerCase().replace(/\s+/g,'').trim();
    if (nm === target) return sheets[i];
  }
  for (var j = 0; j < sheets.length; j++) {
    if (/working\s*hours?/i.test(sheets[j].getName())) return sheets[j];
  }
  return null;
}

function toMinutesFromText_(s) {
  if (!s) return NaN;
  s = String(s).trim();
  var m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (m) {
    var h = +m[1], mi = +m[2];
    if (h <= 23 && mi <= 59) return h * 60 + mi;
  }
  m = s.match(/^(\d{1,2})(?::(\d{2}))?\s*(am|pm)$/i);
  if (m) {
    var hh = +m[1], mm = m[2] ? +m[2] : 0, ap = m[3].toLowerCase();
    if (ap === 'pm' && hh < 12) hh += 12;
    if (ap === 'am' && hh === 12) hh = 0;
    if (hh <= 23 && mm <= 59) return hh * 60 + mm;
  }
  var range = parseRangeFromAny_(s);
  if (range) return range.startMin;
  return NaN;
}

function parseRangeFromAny_(val) {
  var text = String(val).replace(/[–—−]/g, '-');
  var m = text.match(/(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})/);
  if (!m) return null;
  var sH = +m[1], sM = +m[2], eH = +m[3], eM = +m[4];
  if (sH > 23 || eH > 23 || sM > 59 || eM > 59) return null;
  return { startMin: sH * 60 + sM, endMin: eH * 60 + eM };
}

function findTimeOnRow_(rowArr, tbcCol) {
  var cols = rowArr.length;
  function tryParse(col) {
    if (col < 0 || col >= cols) return null;
    var raw = String(rowArr[col] != null ? rowArr[col] : '').trim();
    if (!raw) return null;
    var norm = raw.replace(/[–—−]/g, '-');
    var m = norm.match(/(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})/);
    if (!m) return null;
    var sH = +m[1], sM = +m[2], eH = +m[3], eM = +m[4];
    if (sH > 23 || eH > 23 || sM > 59 || eM > 59) return null;
    return { startMin: sH * 60 + sM, endMin: eH * 60 + eM };
  }
  var hit = tryParse(tbcCol + 2); if (hit) return hit;
  hit = tryParse(tbcCol + 1); if (hit) return hit;
  for (var d = 1; d <= CFG.TIME_SCAN_MAX_OFFSET; d++) {
    var right = tryParse(tbcCol + d); if (right) return right;
    var left  = tryParse(tbcCol - d); if (left)  return left;
  }
  return tryParse(tbcCol);
}

function findContextKey_(sheet, rowIdx0, colIdx0, values) {
  // 1) Prefer scanning UP in the SAME COLUMN for a nearby day header.
  for (var r = rowIdx0; r >= 0 && (rowIdx0 - r) <= CFG.DAY_HEADER_SEARCH_ROWS; r--) {
    var raw = String(values[r][colIdx0] != null ? values[r][colIdx0] : '').trim();
    if (raw && isDayHeader_(raw)) return 'day:' + raw;
  }
  // 2) Small horizontal look-back only (max 8 columns) to avoid crossing day blocks.
  var leftLimit = Math.max(0, colIdx0 - 8);
  for (var c = colIdx0; c >= leftLimit; c--) {
    var raw2 = String(values[rowIdx0][c] != null ? values[rowIdx0][c] : '').trim();
    if (raw2 && isDayHeader_(raw2)) return 'day:' + raw2;
  }
  // 3) Fallback: tag by column.
  return 'col:' + colIdx0;
}

function isDayHeader_(text) {
  var t = String(text || '').toLowerCase();
  var days = ['mon','tue','wed','thu','fri','sat','sun','monday','tuesday','wednesday','thursday','friday','saturday','sunday'];
  for (var i = 0; i < days.length; i++) if (t.indexOf(days[i]) !== -1) return true;
  if (/\b(\d{1,2})\s*(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)\b/i.test(text)) return true;
  if (/\b\d{1,2}[\/\-]\d{1,2}\b/.test(text)) return true;
  return false;
}

function isRealDayLabel_(label) {
  if (!label) return false;
  var t = String(label).trim();
  var lc = t.toLowerCase();
  if (lc.indexOf('col:') === 0) return false;
  if (/^\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2}$/.test(lc)) return false;
  var days = ['mon','tue','wed','thu','fri','sat','sun','monday','tuesday','wednesday','thursday','friday','saturday','sunday'];
  for (var i = 0; i < days.length; i++) if (lc.indexOf(days[i]) !== -1) return true;
  if (/\b(\d{1,2})\s*(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)\b/i.test(t)) return true;
  if (t.indexOf(':') === -1 && /\b\d{1,2}[\/\-]\d{1,2}\b/.test(t)) return true;
  return false;
}

function detectWeekdayHeadersInTopRows_(values) {
  var maxHeaderRows = 12;
  var found = [];
  var seenLC = {};
  var rows = Math.min(values.length, maxHeaderRows);
  for (var r = 0; r < rows; r++) {
    for (var c = 0; c < values[0].length; c++) {
      var raw = String(values[r][c] != null ? values[r][c] : '').trim();
      if (!raw) continue;
      if (isRealDayLabel_(raw)) {
        var lc = raw.toLowerCase();
        if (!seenLC[lc]) { found.push(raw); seenLC[lc] = true; }
      }
    }
  }
  found.sort(function(a,b){ return dayOrder_(a) - dayOrder_(b); });
  return found;
}

function extractDayLabel_(contextKey) {
  if (!contextKey) return 'Day';
  if (contextKey.indexOf('day:') === 0) return contextKey.substring(4).trim();
  return contextKey; // e.g., fallback "col:12"
}

function dayOrder_(label) {
  var t = String(label || '').toLowerCase();
  if (t.indexOf('mon') > -1) return 1;
  if (t.indexOf('tue') > -1) return 2;
  if (t.indexOf('wed') > -1) return 3;
  if (t.indexOf('thu') > -1) return 4;
  if (t.indexOf('fri') > -1) return 5;
  if (t.indexOf('sat') > -1) return 6;
  if (t.indexOf('sun') > -1) return 7;
  return 99;
}

function isWeekendLabel_(label) {
  var lc = String(label || '').toLowerCase();
  return (lc.indexOf('sat') !== -1) || (lc.indexOf('sun') !== -1);
}

/* ---- tiny helpers ---- */
function __formatWH_() {
  var people = loadPeople_();
  var out = [];
  for (var i = 0; i < people.length; i++) {
    var p = people[i];
    out.push(
      p.name + ': ' +
      pad2_(Math.floor(p.startMin/60)) + ':' + pad2_(p.startMin%60) + ' - ' +
      pad2_((Math.floor(p.endMin/60))) + ':' + pad2_(p.endMin%60)
    );
  }
  return out;
}

function pad2_(n){ n = Number(n)||0; return n<10 ? '0'+n : String(n); }
function round2(n){ return Math.round((Number(n)||0)*100)/100; }

function findPersonCaseInsensitive_(people, s) {
  var low = String(s || '').toLowerCase();
  for (var i = 0; i < people.length; i++) if (people[i].name.toLowerCase() === low) return people[i];
  return null;
}

function hashPastel_(name) {
  var h = 0;
  for (var i = 0; i < name.length; i++) h = (h * 31 + name.charCodeAt(i)) >>> 0;
  var hue = h % 360; return hsvToHex_(hue, 0.25, 1.0);
}

function hsvToHex_(h, s, v) {
  var c=v*s, x=c*(1-Math.abs((h/60)%2-1)), m=v-c; var r=0,g=0,b=0;
  if (h<60){r=c;g=x;} else if (h<120){r=x;g=c;} else if (h<180){g=c;b=x;}
  else if (h<240){g=x;b=c;} else if (h<300){r=x;b=c;} else {r=c;b=x;}
  function to255(n){ return Math.round((n+m)*255); }
  return '#' + [to255(r),to255(g),to255(b)].map(function(n){return n.toString(16).padStart(2,'0')}).join('');
}

function toHHMM_(m) {
  m = Number(m)||0;
  var h = Math.floor(m/60), mi = m % 60;
  return pad2_(h) + ':' + pad2_(mi);
}

// 1 if a full break fits somewhere in the shift (with this assignment), else 0.
function breakFeasibleScore_(person, existingWins, startMin, endMin) {
  var req = (CFG.REQUIRE_BREAK ? Number(CFG.BREAK_DURATION_MIN || 0) : 0);
  if (!req) return 1;
  var threshold = Number(CFG.APPLY_BREAK_WHEN_SHIFT_MIN_AT_LEAST || 0);
  var shiftLen = person.endMin - person.startMin;
  if (shiftLen < threshold) return 1;
  var wins = (existingWins || []).slice();
  wins.push({ s: startMin, e: endMin });
  wins.sort(function(a,b){ return a.s - b.s; });
  var merged = [];
  for (var i = 0; i < wins.length; i++) {
    var cur = wins[i];
    if (!merged.length) merged.push({ s: cur.s, e: cur.e });
    else {
      var last = merged[merged.length - 1];
      if (cur.s <= last.e) { if (cur.e > last.e) last.e = cur.e; }
      else merged.push({ s: cur.s, e: cur.e });
    }
  }
  var prev = person.startMin;
  for (var k = 0; k < merged.length; k++) {
    var occ = merged[k];
    if (occ.s > prev && (occ.s - prev) >= req) return 1;
    if (occ.e > prev) prev = occ.e;
  }
  if (person.endMin > prev && (person.endMin - prev) >= req) return 1;
  return 0;
}

// --- Day matching helpers ---
function _weekdayIndex_(s) {
  var t = String(s||'').toLowerCase();
  if (t.indexOf('mon')>-1) return 1;
  if (t.indexOf('tue')>-1) return 2;
  if (t.indexOf('wed')>-1) return 3;
  if (t.indexOf('thu')>-1) return 4;
  if (t.indexOf('fri')>-1) return 5;
  if (t.indexOf('sat')>-1) return 6;
  if (t.indexOf('sun')>-1) return 7;
  return 0;
}

function _dayLabelsEquivalent_(a, b) {
  var ia = _weekdayIndex_(a), ib = _weekdayIndex_(b);
  if (ia && ib) return ia === ib; // same weekday, ignore extra date text
  return String(a||'').trim().toLowerCase() === String(b||'').trim().toLowerCase();
}

// Core worker: undo only cells that belong to the given day label
function _undoForDayCore_(sheet, values, targetDayLabel) {
  var people = loadPeople_();
  var names = {};
  for (var i = 0; i < people.length; i++) names[people[i].name.toLowerCase()] = true;
  var reverted = 0;
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[0].length; c++) {
      var v = String(values[r][c] != null ? values[r][c] : '').trim();
      if (!names[v.toLowerCase()]) continue;
      var ctx = findContextKey_(sheet, r, c, values);
      var dayLabel = extractDayLabel_(ctx);
      if (!isRealDayLabel_(dayLabel)) continue;
      if (_dayLabelsEquivalent_(dayLabel, targetDayLabel)) {
        sheet.getRange(r + 1, c + 1).setValue('TBC').setBackground(null);
        reverted++;
      }
    }
  }
  return reverted;
}

/* ====== Pretty output formatting for summary sheets ====== */
function beautifySheet_(sh) {
  var bands = sh.getBandings();
  for (var i = 0; i < bands.length; i++) bands[i].remove();
  var rng = sh.getDataRange();
  var band = rng.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  band.setHeaderRowColor('#F1F3F4');
  band.setFirstRowColor('#FFFFFF');
  band.setSecondRowColor('#F8F9FA');
  sh.getRange(1,1,1,Math.max(1, rng.getNumColumns())).setFontWeight('bold');
  var lastCol = rng.getNumColumns();
  var header = sh.getRange(1,1,1,lastCol).getDisplayValues()[0];
  for (var c = 1; c <= lastCol; c++) {
    var head = (header[c-1] || '').toString().toLowerCase();
    if (head.indexOf('%') !== -1 || head.indexOf('utilization') !== -1) {
      sh.getRange(2,c,Math.max(0, rng.getNumRows()-1),1).setNumberFormat('0.00%');
    } else if (head.indexOf('(h)') !== -1 || head.indexOf('hours') !== -1) {
      sh.getRange(2,c,Math.max(0, rng.getNumRows()-1),1).setNumberFormat('0.00');
    }
  }
  sh.autoResizeColumns(1, lastCol);
}

function showBreakPlanDialogOnActiveSheet() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  var values = sheet.getDataRange().getValues();
  var people = loadPeople_();
  if (!people.length) { SpreadsheetApp.getUi().alert('No editors found on "'+CFG.WORKING_HOURS_SHEET+'".'); return; }

  // Detect day headers from the top of the sheet
  var days = detectWeekdayHeadersInTopRows_(values);
  if (!days.length) { SpreadsheetApp.getUi().alert('No weekday/date headers found in the top rows.'); return; }
  days.sort(function(a,b){ return dayOrder_(a) - dayOrder_(b); });

  // Build occupied windows by person/day from existing assignments
  var winsByPersonDay = {};
  for (var i = 0; i < people.length; i++) winsByPersonDay[people[i].name] = {};
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[0].length; c++) {
      var cell = String(values[r][c] != null ? values[r][c] : '').trim();
      var match = findPersonCaseInsensitive_(people, cell);
      if (!match) continue;
      var ti = findTimeOnRow_(values[r], c);
      if (!ti) continue;
      if (CFG.IGNORE_ZERO_TIMES && ti.startMin === 0 && ti.endMin === 0) continue;
      if (ti.endMin <= ti.startMin) continue;
      var ctx = findContextKey_(sheet, r, c, values);
      var dayLabel = extractDayLabel_(ctx);
      if (!isRealDayLabel_(dayLabel)) continue;
      if (!winsByPersonDay[match.name][dayLabel]) winsByPersonDay[match.name][dayLabel] = [];
      winsByPersonDay[match.name][dayLabel].push({ s: ti.startMin, e: ti.endMin });
    }
  }

  // Settings for break placement
  var requireBreak = !!CFG.REQUIRE_BREAK;
  var reqBreakMin  = requireBreak ? Number(CFG.BREAK_DURATION_MIN || 0) : 0;
  var threshold    = Number(CFG.APPLY_BREAK_WHEN_SHIFT_MIN_AT_LEAST || 0);
  var peopleSorted = people.slice().sort(function(a,b){ return a.name.localeCompare(b.name); });

  // --- HTML helpers/styles (same look as the previous inbuilt popup) ---
  function esc(s){ s=(s==null)?'':String(s); return s.replace(/[&<>"']/g, function(ch){return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]);}); }

  var css =
    '<style>'
    +'body{font-family:Arial,Helvetica,sans-serif;margin:0;background:#fff}'
    +'.wrap{padding:18px}'
    +'.hdr{display:flex;justify-content:center;margin:0 0 16px}'
    +'.chip{background:linear-gradient(135deg,#1a73e8,#6ea8fe);color:#fff;padding:10px 16px;border-radius:999px;font-weight:600;letter-spacing:.2px;box-shadow:0 2px 6px rgba(0,0,0,.12)}'
    +'table{border-collapse:separate;border-spacing:0;width:100%;font-size:13px;border:1px solid #e0e0e0;border-radius:10px;overflow:hidden}'
    +'thead th{background:linear-gradient(90deg,#1a73e8,#6ea8fe);color:#fff;padding:10px 12px;text-align:left;font-weight:600}'
    +'tbody td{padding:8px 12px;border-top:1px solid #eaecef}'
    +'tbody tr:nth-child(even) td{background:#fafbff}'
    +'.day-row td{background:#eef3ff;color:#174ea6;font-weight:600;border-top:2px solid #d4e2ff}'
    +'.num{text-align:right}'
    +'.spacer td{padding:6px;background:#fff;border:none}'
    +'.actions{text-align:right;margin-top:12px}'
    +'button{padding:8px 12px;border:0;border-radius:8px;background:#1a73e8;color:#fff;cursor:pointer}'
    +'button:hover{filter:brightness(1.05)}'
    +'</style>';

  // Build grouped rows by day (one day header, then one row per editor)
  var bodyHtml = '';
  for (var d = 0; d < days.length; d++) {
    var dayLabel = days[d];
    // Day header (single full-width row)
    bodyHtml += '<tr class="day-row"><td colspan="5">'+esc(dayLabel)+'</td></tr>';
    // Editor rows for this day
    for (var p = 0; p < peopleSorted.length; p++) {
      var person = peopleSorted[p];
      var shiftLen = person.endMin - person.startMin;
      var enforce = (requireBreak && shiftLen >= threshold);
      var wins = (winsByPersonDay[person.name][dayLabel] || []).slice();
      var res = computeBreakPlacement_(person, wins, reqBreakMin);
      var shiftTxt = toHHMM_(person.startMin) + ' - ' + toHHMM_(person.endMin);
      var breakTxt = (res && res.possible) ? (toHHMM_(res.s) + ' - ' + toHHMM_(res.e)) : '—';
      var distMin  = (res && res.possible && res.dist != null) ? Math.round(res.dist) : '';
      var reason = !enforce ? 'Break not required (short shift)'
                  : (res && res.possible ? 'Placed nearest to midpoint'
                  : 'No gap long enough for break');
      bodyHtml += '<tr>'
        + '<td>'+esc(person.name)+'</td>'
        + '<td>'+esc(shiftTxt)+'</td>'
        + '<td>'+esc(breakTxt)+'</td>'
        + '<td class="num">'+esc(distMin)+'</td>'
        + '<td>'+esc(reason)+'</td>'
        + '</tr>';
    }
    // spacer between days
    bodyHtml += '<tr class="spacer"><td colspan="5"></td></tr>';
  }

  var html =
    '<html><head><meta name="viewport" content="width=device-width, initial-scale=1">'+css+'</head>'
    + '<body><div class="wrap">'
    +   '<div class="hdr"><div class="chip">Break Plan</div></div>'
    +   '<table>'
    +     '<thead><tr>'
    +       '<th>Name</th><th>Shift</th><th>1 hour break</th><th>Distance to Mid (min)</th><th>Reason</th>'
    +     '</tr></thead>'
    +     '<tbody>'+ bodyHtml +'</tbody>'
    +   '</table>'
    +   '<div class="actions"><button onclick="google.script.host.close()">Close</button></div>'
    + '</div></body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(1100).setHeight(750),
    'Break Plan'
  );
}

// (Optional) keep this alias if any old menu item still calls the sheet version
function generateBreakPlanOnActiveSheet(){ return showBreakPlanDialogOnActiveSheet(); }

function showStudioSetUsageDialogOnActiveSheet() {
  // Pull the same source your Operator Timeline uses
  var items = [];
  try {
    if (typeof getOperatorTimelinePayloadSafe === 'function') {
      var payload = getOperatorTimelinePayloadSafe();
      items = (payload && payload.items) ? payload.items : [];
    }
  } catch (e) {
    items = [];
  }
  if (!items.length) {
    SpreadsheetApp.getUi().alert("No shoots found for the active sheet (couldn't read timeline payload).");
    return;
  }

  // Aggregate
  var byStudio = {}; // { name: {hours, count} }
  var bySet    = {};
  var totalHours = 0;
  items.forEach(function (it) {
    var ms = Math.max(0, Number(it.endMs || 0) - Number(it.startMs || 0));
    if (!ms) return;
    var hours = ms / 3600000;
    // In your payload, the set name shows up in "room" (e.g., Iris/Nest/Nova/etc.)
    var setName = String(it.room || '').trim();
    var studioName = SET_TO_STUDIO[setName.toLowerCase()] || UNKNOWN_STUDIO;
    totalHours += hours;
    if (!bySet[setName]) bySet[setName] = { hours: 0, count: 0 };
    bySet[setName].hours += hours;
    bySet[setName].count += 1;
    if (!byStudio[studioName]) byStudio[studioName] = { hours: 0, count: 0 };
    byStudio[studioName].hours += hours;
    byStudio[studioName].count += 1;
  });

  function toRows(map) {
    var arr = Object.keys(map).map(function (k) {
      return { name: k, hours: map[k].hours, count: map[k].count };
    });
    arr.sort(function (a, b) { return b.hours - a.hours; });
    return arr;
  }

  var studios = toRows(byStudio);
  var sets    = toRows(bySet);

  function pct(h) { return totalHours > 0 ? (h / totalHours) * 100 : 0; }
  function esc(s){ s=(s==null)?'':String(s); return s.replace(/[&<>"']/g, function(ch){return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]);}); }

  // ---- Styling to match your other popups ----
  var css =
    '<style>'
    + 'body{font-family:Arial,sans-serif;margin:0}.wrap{padding:18px}'
    + '.hdr{display:flex;justify-content:center;margin:0 0 14px}'
    + '.chip{background:linear-gradient(135deg,#1a73e8,#6ea8fe);color:#fff;padding:10px 16px;border-radius:999px;font-weight:600;box-shadow:0 2px 6px rgba(0,0,0,.12)}'
    + '.cap{margin:18px 0 8px;font-size:13px;font-weight:700;color:#344054;background:#f4f6fb;border:1px solid #e7ecfb;padding:8px 10px;border-radius:8px}'
    + 'table.tbl{border-collapse:collapse;width:100%;font-size:13px;background:#fff;border:1px solid #e0e0e0;border-radius:10px;overflow:hidden}'
    + 'table.tbl thead th{background:linear-gradient(180deg,#3b82f6,#2b6de8);color:#fff;padding:10px;border-right:1px solid rgba(255,255,255,.25);text-align:left}'
    + 'table.tbl thead th:last-child{border-right:none}'
    + 'table.tbl tbody td{border-top:1px solid #eef2f7;padding:9px 10px;text-align:left}'
    + 'table.tbl tbody tr:nth-child(even){background:#fafbfd}'
    + '.actions{text-align:right;margin-top:10px}'
    + 'button{padding:8px 12px;border:none;border-radius:8px;background:#1a73e8;color:#fff;cursor:pointer}'
    + '</style>';

  function tableBlock(title, rows, labelHeader) {
    var html = '<div class="cap">'+esc(title)+'</div>';
    html += '<table class="tbl"><thead><tr>'
      + '<th>'+esc(labelHeader)+'</th>'
      + '<th>Hours</th>'
      + '<th>% of total</th>'
      + '<th>Shoots</th>'
      + '</tr></thead><tbody>';
    rows.forEach(function (r) {
      html += '<tr>'
        + '<td>'+esc(r.name)+'</td>'
        + '<td>'+ (r.hours).toFixed(2) +'</td>'
        + '<td>'+ (pct(r.hours)).toFixed(2) +'</td>'
        + '<td>'+ r.count +'</td>'
        + '</tr>';
    });
    html += '</tbody></table>';
    return html;
  }

  var html =
    '<html><head><meta name="viewport" content="width=device-width, initial-scale=1">'+css+'</head>'
    + '<body><div class="wrap">'
    + '<div class="hdr"><div class="chip">Studio &amp; Set Usage</div></div>'
    + tableBlock('By studio', studios, 'Studio')
    + tableBlock('By set', sets, 'Set')
    + '<div class="actions"><button onclick="google.script.host.close()">Close</button></div>'
    + '</div></body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(1100).setHeight(720),
    'Studio & Set Usage'
  );
}

// ===============================
// Day-Scoped Undo (copy/paste me)
// ===============================
// Menu action: undo ONLY the day that contains the current active cell
function undoAssignedForActiveCellDay() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  var rng = sheet.getActiveCell();
  if (!rng) { SpreadsheetApp.getUi().alert('Select a cell inside the day you want to undo.'); return; }
  var r0 = rng.getRow() - 1, c0 = rng.getColumn() - 1;
  var values = sheet.getDataRange().getValues();
  var ctx = findContextKey_(sheet, r0, c0, values);
  var dayLabel = extractDayLabel_(ctx);
  if (!isRealDayLabel_(dayLabel)) {
    SpreadsheetApp.getUi().alert('Could not detect a day header near the active cell.');
    return;
  }
  var count = _undoForDayCore_(sheet, values, dayLabel);
  SpreadsheetApp.getUi().alert('Reverted ' + count + ' cell(s) to "TBC" for ' + dayLabel + ' on "' + sheet.getName() + '".');
}

// Menu action: prompt for a day (Mon/Tue/... or the exact header text)
function undoAssignedForDayOnActiveSheet() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  var values = sheet.getDataRange().getValues();
  var days = detectWeekdayHeadersInTopRows_(values);
  if (!days.length) {
    SpreadsheetApp.getUi().alert('Could not find weekday/date headers in the top rows.');
    return;
  }
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Undo for Day', 'Type the day to undo (e.g., Mon, Tuesday, or full header). Available: ' + days.join(', '), ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var typed = (res.getResponseText() || '').trim();
  if (!typed) { ui.alert('No day entered.'); return; }
  // Map short forms like "Tue" to the actual detected header
  var picked = typed, typedIdx = _weekdayIndex_(typed);
  if (typedIdx) {
    for (var i = 0; i < days.length; i++) {
      if (_weekdayIndex_(days[i]) === typedIdx) { picked = days[i]; break; }
    }
  }
  var count = _undoForDayCore_(sheet, values, picked);
  ui.alert('Reverted ' + count + ' cell(s) to "TBC" for ' + picked + ' on "' + sheet.getName() + '".');
}

// ================================
// Studio/Set usage → weekly sheets
// ================================
// --- Public menu actions ---
function writeStudioUsageRow() {
  var stats = _getStudioSetStatsForActiveWeek_();
  _upsertUsageRow_({
    sheetName: 'Studio Usage (Weekly)',
    label: stats.weekLabel,
    labels: STUDIO_ORDER,
    // values per label: [hours, pct-as-ratio, count]
    rowBuilder: function(name) {
      var rec = stats.byStudio[name] || { hours:0, count:0 };
      var pct = stats.totalHours > 0 ? (rec.hours / stats.totalHours) : 0;
      return [rec.hours, pct, rec.count];
    }
  });
  SpreadsheetApp.getUi().alert('Studio usage written for "'+stats.weekLabel+'".');
}

function writeSetUsageRow() {
  var stats = _getStudioSetStatsForActiveWeek_();
  _upsertUsageRow_({
    sheetName: 'Set Usage (Weekly)',
    label: stats.weekLabel,
    labels: SET_ORDER,
    rowBuilder: function(name) {
      var rec = stats.bySet[name] || { hours:0, count:0 };
      var pct = stats.totalHours > 0 ? (rec.hours / stats.totalHours) : 0;
      return [rec.hours, pct, rec.count];
    }
  });
  SpreadsheetApp.getUi().alert('Set usage written for "'+stats.weekLabel+'".');
}

// --- Core stats from your timeline payload ---
function _getStudioSetStatsForActiveWeek_() {
  var ss = SpreadsheetApp.getActive();
  var activeName = ss.getActiveSheet().getName();
  var payload = null, items = [];
  try {
    if (typeof getOperatorTimelinePayloadSafe === 'function') {
      payload = getOperatorTimelinePayloadSafe();
      items = (payload && payload.items) ? payload.items : [];
    }
  } catch (e) { items = []; }
  if (!items.length) throw new Error('No shoots found (timeline payload was empty).');

  var totalHours = 0;
  var byStudio = {}; // { name: {hours, count} }
  var bySet    = {}; // { name: {hours, count} }
  items.forEach(function (it) {
    var ms = Math.max(0, Number(it.endMs || 0) - Number(it.startMs || 0));
    if (!ms) return;
    var hours = ms / 3600000;
    var setName = String(it.room || '').trim();
    var studioName = SET_TO_STUDIO[setName.toLowerCase()] || UNKNOWN_STUDIO;
    totalHours += hours;
    if (!bySet[setName]) bySet[setName] = { hours: 0, count: 0 };
    bySet[setName].hours += hours; bySet[setName].count += 1;
    if (!byStudio[studioName]) byStudio[studioName] = { hours: 0, count: 0 };
    byStudio[studioName].hours += hours; byStudio[studioName].count += 1;
  });
  var weekLabel = _deriveWeekLabel_(payload, activeName);
  return { weekLabel: weekLabel, totalHours: totalHours, byStudio: byStudio, bySet: bySet };
}

function _deriveWeekLabel_(payload, fallback) {
  var tit = (payload && payload.title) ? String(payload.title).trim() : '';
  var sname = String(fallback || '').trim();
  // Prefer a human-friendly week range in payload.title; otherwise use the sheet name.
  var looksWeek = /\b\d{1,2}(?:st|nd|rd|th)?\s*-\s*\d{1,2}(?:st|nd|rd|th)?\s+(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|january|february|march|april|may|june|july|august|september|october|november|december)\b/i.test(tit);
  return looksWeek ? tit : (sname || 'This Week');
}

function _upsertUsageRow_(opts) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(opts.sheetName);
  if (!sh) sh = ss.insertSheet(opts.sheetName);
  // Build desired header
  var header = ['Week'];
  opts.labels.forEach(function(lbl){
    header.push(lbl+' (h)', lbl+' (%)', lbl+' (#)');
  });
  // Ensure header exists
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else {
    var maxCols = Math.max(sh.getLastColumn(), header.length);
    var currentHeader = sh.getRange(1,1,1,maxCols).getValues()[0];
    var mismatch = false;
    for (var i=0;i<header.length;i++){
      if ((currentHeader[i]||'') !== header[i]) { mismatch = true; break; }
    }
    if (mismatch) {
      sh.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  }
  // Build the row to write
  var row = [opts.label];
  opts.labels.forEach(function(lbl){
    var parts = opts.rowBuilder(lbl) || [0,0,0]; // [hours, pct-ratio, count]
    row.push(Number(parts[0]||0), Number(parts[1]||0), Number(parts[2]||0));
  });
  // Upsert by week label in column A
  var lastRow = sh.getLastRow();
  var dataRows = Math.max(0, lastRow - 1);
  var targetRow = -1;
  if (dataRows > 0) {
    var colA = sh.getRange(2,1,dataRows,1).getValues();
    var want = String(opts.label||'').trim().toLowerCase();
    for (var r=0; r<dataRows; r++) {
      if (String(colA[r][0]||'').trim().toLowerCase() === want) { targetRow = r + 2; break; }
    }
  }
  if (targetRow > 0) {
    sh.getRange(targetRow, 1, 1, row.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }
  // Formatting — only if there is at least one data row
  lastRow = sh.getLastRow();
  dataRows = Math.max(0, lastRow - 1);
  var totalCols = row.length;
  if (dataRows > 0) {
    for (var c=2; c<=totalCols; c+=3) {
      sh.getRange(2, c, dataRows, 1).setNumberFormat('0.00');   // Hours
    }
    for (var c2=3; c2<=totalCols; c2+=3) {
      sh.getRange(2, c2, dataRows, 1).setNumberFormat('0.00%'); // Percent (ratio)
    }
    for (var c3=4; c3<=totalCols; c3+=3) {
      sh.getRange(2, c3, dataRows, 1).setNumberFormat('0');     // # Shoots
    }
  }
  sh.autoResizeColumns(1, totalCols);
}

/***** Monthly Studio/Set usage across ALL tabs (no name clashes) *****/
function showStudioSetUsageMonthlyAllTabs() {
  var ss   = SpreadsheetApp.getActive();
  var prev = ss.getActiveSheet();
  var allItems = [];
  var sheets = ss.getSheets();
  try {
    for (var i = 0; i < sheets.length; i++) {
      var sh = sheets[i];
      var nm = sh.getName();
      if (/working\s*hours?/i.test(nm)) continue;
      if (/usage|break plan|editor hours|hours summary|studio/i.test(nm)) continue;
      try {
        ss.setActiveSheet(sh);
        var payload = (typeof getOperatorTimelinePayloadSafe === 'function')
          ? getOperatorTimelinePayloadSafe()
          : null;
        var items = (payload && payload.items) ? payload.items : [];
        if (items && items.length) {
          for (var j = 0; j < items.length; j++) allItems.push(items[j]);
        }
      } catch (e) { /* ignore */ }
    }
  } finally {
    if (prev) ss.setActiveSheet(prev);
  }
  if (!allItems.length) {
    SpreadsheetApp.getUi().alert("Couldn't find any timeline items in the workbook.");
    return;
  }
  var agg = aggregateStudioSetByMonth_(allItems);
  var html = buildMonthlyUsageHtml_(agg);
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(1100).setHeight(750),
    'Studio & Set Usage — Monthly'
  );
}

function aggregateStudioSetByMonth_(items) {
  var months = {};
  var studioNames = {};
  var setNames = {};
  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    var ms = Math.max(0, Number(it.endMs || 0) - Number(it.startMs || 0));
    if (!ms) continue;
    var dt = new Date(Number(it.startMs || 0));
    var key = dt.getFullYear() + '-' + pad2_(dt.getMonth() + 1);
    if (!months[key]) months[key] = { totalH: 0, studio: {}, set: {} };
    var hours = ms / 3600000;
    var setNameRaw = String(it.room || '').trim();
    var setKey = setNameRaw || 'Unknown';
    var studioKey = SET_TO_STUDIO[(setNameRaw || '').toLowerCase()] || UNKNOWN_STUDIO;
    months[key].totalH += hours;
    if (!months[key].set[setKey]) months[key].set[setKey] = { h: 0, c: 0 };
    months[key].set[setKey].h += hours;
    months[key].set[setKey].c += 1;
    setNames[setKey] = true;
    if (!months[key].studio[studioKey]) months[key].studio[studioKey] = { h: 0, c: 0 };
    months[key].studio[studioKey].h += hours;
    months[key].studio[studioKey].c += 1;
    studioNames[studioKey] = true;
  }
  var monthKeys = Object.keys(months).sort();
  var studioCols = Object.keys(studioNames).sort();
  var setCols    = Object.keys(setNames).sort();
  return { months: months, monthKeys: monthKeys, studioCols: studioCols, setCols: setCols };
}

function buildMonthlyUsageHtml_(agg) {
  var css =
    '<style>'
    +'body{font-family:Arial,sans-serif;margin:0}.wrap{padding:18px}'
    +'.hdr{display:flex;justify-content:center;margin:0 0 14px}'
    +'.chip{background:linear-gradient(135deg,#1a73e8,#6ea8fe);color:#fff;padding:10px 16px;border-radius:999px;font-weight:600;box-shadow:0 2px 6px rgba(0,0,0,.12)}'
    +'.cap{margin:14px 0 8px;padding:8px 10px;background:#f4f6fb;border:1px solid #e7ecfb;border-radius:8px;font-weight:700;color:#344054}'
    +'table{border-collapse:collapse;width:100%;font-size:13px;background:#fff;border:1px solid #e0e0e0;border-radius:10px;overflow:hidden}'
    +'th,td{border-top:1px solid #eef2f7;padding:9px 10px;text-align:left}'
    +'thead th{background:linear-gradient(180deg,#3b82f6,#2b6de8);color:#fff;border-right:1px solid rgba(255,255,255,.25)}'
    +'thead th:last-child{border-right:none}'
    +'tbody tr:nth-child(even){background:#fafbfd}'
    +'.small{color:#667085;font-size:12px}'
    +'</style>';
  function tableStudios() {
    var cols = agg.studioCols.slice();
    var h = '<div class="cap">By studio (monthly)</div><table><thead><tr><th>Month</th>';
    cols.forEach(function(n){ h += '<th>'+n+'<br><span class="small">h / shoots</span></th>'; });
    h += '<th>Total hours</th></tr></thead><tbody>';
    agg.monthKeys.forEach(function(mk){
      var row = '<tr><td>'+mk+'</td>';
      var block = agg.months[mk];
      cols.forEach(function(n){
        var cell = block.studio[n] || {h:0,c:0};
        row += '<td>'+cell.h.toFixed(2)+' / '+cell.c+'</td>';
      });
      row += '<td>'+block.totalH.toFixed(2)+'</td></tr>';
      h += row;
    });
    h += '</tbody></table>';
    return h;
  }
  function tableSets() {
    var cols = agg.setCols.slice();
    var h = '<div class="cap">By set (monthly)</div><table><thead><tr><th>Month</th>';
    cols.forEach(function(n){ h += '<th>'+n+'<br><span class="small">h / shoots</span></th>'; });
    h += '<th>Total hours</th></tr></thead><tbody>';
    agg.monthKeys.forEach(function(mk){
      var row = '<tr><td>'+mk+'</td>';
      var block = agg.months[mk];
      cols.forEach(function(n){
        var cell = block.set[n] || {h:0,c:0};
        row += '<td>'+cell.h.toFixed(2)+' / '+cell.c+'</td>';
      });
      row += '<td>'+block.totalH.toFixed(2)+'</td></tr>';
      h += row;
    });
    h += '</tbody></table>';
    return h;
  }
  return '<html><head><meta name="viewport" content="width=device-width, initial-scale=1">'+css+'</head>'
    + '<body><div class="wrap">'
    + '<div class="hdr"><div class="chip">Studio &amp; Set Usage — Monthly</div></div>'
    + tableStudios()
    + tableSets()
    + '</div></body></html>';
}
