# Changelog

## 2026-03-15 - Dashboard Refinements

### 32. **Add Summary Charts: Total Hours per Month & per Week** (Commit: TBD)
- **New charts:** Two simple single-series column charts appended below the individual set/studio charts:
  1. **Total Hours per Month** — sums all studio hours per month, blue bars
  2. **Total Hours per Week** — groups daily usage by week (Monday = week start), green bars
- **One-time setup:** Run `addSummaryCharts()` from the menu (`➕ Add Summary Charts`) — no full dashboard rebuild required
- **Data updates:** `updateDashboardData()` automatically refreshes both charts' data going forward
- **Hidden columns:** Monthly totals at cols 68-69, weekly totals at cols 70-71 (300 pre-allocated rows ≈ 6 years of weekly data)
- **New functions:**
  - `addSummaryCharts()` — one-shot chart creator, in menu
  - `_writeSummaryData_()` — writes both hidden data tables; called by `_writeDashboardData_()`
  - `_getWeeklyHoursRows_()` — reads Studio Usage (Daily), groups by ISO week, returns sorted rows
  - `_insertSummaryCharts_()` — creates the two charts; also called by `setupDashboard()`
- **Files Modified:** `Code Timelines - Daily Usage.gs`, `Code Scheduling.gs`

---

### 31. **Remove Overview Charts; Individual Charts Only** (Commit: TBD)
- **Change:** Removed the four full-width overview charts (Studio % / Hours, Set % / Hours). Only the individual per-studio and per-set dual-axis charts remain.
- **Why:** Overview charts not needed — individual breakdowns are sufficient.
- **Layout:** Individual studio charts now start at row 5 (was row 103); sets at row 70 (was row 168). Much more compact.
- **Hidden columns compacted:** Studio data now cols 18-37; set data cols 38-63. Former overview range (18-45) freed up.
- **`_insertOverviewChart_()` and `_writeOverviewBlock_()`** kept as dead code but no longer called.
- **No re-setup required** — overview charts deleted manually from sheet by user.
- **Files Modified:** `Code Timelines - Daily Usage.gs`

---

## 2026-03-15 - Persistent Dashboard Architecture

### Summary
Major dashboard redesign: charts are now permanent and never rebuilt during routine data updates. Manual chart customisations (data labels, colours, etc.) survive indefinitely. The monthly trigger now only refreshes data, leaving all chart settings intact.

---

### 30. **Persistent Dashboard — Charts Never Rebuilt on Data Update** (Commit: 1c47641)
- **Problem:** `updateDashboard()` deleted and recreated all charts on every run, resetting any manual customisations (e.g. data labels the user had turned on) each time
- **Solution:** Split dashboard into two functions with distinct responsibilities:
  - `setupDashboard()` — one-time full rebuild, **script-editor only** (not in menu). Wipes and recreates all charts. Run once to set up; set data labels manually on each chart (right-click → Edit chart → Customise → Series → Data labels). Labels persist permanently.
  - `updateDashboardData()` — lightweight data refresh, **in menu**. Rewrites hidden data tables only; never touches charts. Charts auto-refresh via Google Sheets' native data-binding.
- **New Architecture — Hidden Data Tables:**
  All data lives in pre-allocated hidden columns (cols 18–91), off the visible dashboard area. Charts reference these fixed ranges (header + 60 pre-allocated rows = 5 years of headroom). Row positions never shift, so no chart range updates are needed after setup.
  ```
  Cols 18-23   Overview Studio %      (Month + Studio 1/2/3/4/Other)
  Cols 24-29   Overview Studio Hours  (same structure)
  Cols 30-37   Overview Set %         (Month + Iris/Club/Nest/Exec/Nova/Soho/Other)
  Cols 38-45   Overview Set Hours     (same structure)
  Cols 46-65   Individual studio charts  (Studio 1-4, 5 cols each)
  Cols 66-91   Individual set charts     (Iris/Club/Nest/Exec/Nova/Soho, 5 cols each)
  ```
- **Fixed Visual Layout:**
  Chart positions are derived from named constants (e.g. `DASH_OV_S_PCT_CHART_ROW = 5`) so the layout is deterministic regardless of how many months of data exist.
- **New / Renamed Functions:**
  - `setupDashboard()` — replaces `updateDashboard()` (full rebuild, hidden from menu)
  - `updateDashboardData()` — new lightweight data-only update (in menu)
  - `_writeDashboardData_()` — shared helper writing all hidden data tables
  - `_writeOverviewBlock_()` — writes one multi-series overview data block
  - `_writeIndivBlock_()` — writes individual chart data (5 cols per label)
  - `_insertOverviewChart_()` — replaces `_createTrendBlock_()` (chart creation only)
  - `_insertIndivCharts_()` — replaces `_createIndividualCharts_()` (chart creation only)
  - `DASH_*` constants — 20+ layout constants defining columns, rows, sizes
- **Removed Functions:** `_createTrendBlock_()`, `_createIndividualCharts_()` (stubs left to prevent trigger errors)
- **Menu Change:** `📈 Update Dashboard` → `📊 Update Dashboard Data` (calls `updateDashboardData`)
- **Trigger:** `installDashboardTrigger()` now installs `updateDashboardData` (not `setupDashboard`)
- **Files Modified:** `Code Timelines - Daily Usage.gs`, `Code Scheduling.gs`

---

## 2026-03-12 - Project Consolidation & Critical Bug Fixes

### Summary
Major restructuring session: consolidated multiple separate Google Apps Script projects into one, fixed persistent menu failures, and fixed critical date detection bugs causing inflated usage statistics.

---

### 29. **Fix Date Detection for Old and New Sheet Formats** (Commit: 75082ea)
- **Problem:** Backfill only found a fraction of expected dates (e.g. 11 dates for all of January)
- **Root Cause (discovered via spreadsheet analysis):**
  - Sheets before Nov 24 2025 store row 2 as plain text "Monday", "Tuesday" etc.
  - Sheets from Nov 24 2025 onwards store row 2 as actual Google Sheets date cells
  - Old code fell back to `fallbackWeekInfo()` (current week) when A1 couldn't be parsed
  - Sessions from those weeks were wrongly attributed to the current week's dates
- **Fix:** Three-method date detection in `_getTimelinePayloadForSheet_()`:
  1. Raw Date object from row 2 cell (post-Nov sheets — most reliable)
  2. Text parsing of row 2 display value e.g. "Monday, 9 March 2026"
  3. A1 week header text fallback e.g. "9th Mar - 15th Mar" (pre-Nov sheets)
  - If all three fail, day is skipped with a warning — never uses `fallbackWeekInfo()`
- **Also identified:** Several Saturday cells in January have data entry errors (wrong dates or text "Saturday") — these must be manually corrected in the spreadsheet
- **Files Modified:** `code.gs`

---

### 28. **Fix Duplicate Rows Causing Inflated Monthly Statistics** (Commit: 94ca6e7)
- **Problem:** Monthly Studio/Set summaries showing impossible values (e.g. 2258h total, 112.9h/day average for February)
- **Root Cause:** `_upsertDailyUsageRow_()` used `getValues()` to read column A for date matching. Google Sheets auto-converts "2026-02-01" strings to Date objects, so the string comparison always failed — every write **appended** a new row instead of updating the existing one, doubling all data
- **Fix:** Changed `getValues()` to `getDisplayValues()` for the upsert date lookup
- **Also Added:** `deduplicateDailyUsageSheets()` — removes duplicate date rows from Studio Usage (Daily) and Set Usage (Daily), accessible via menu
- **Files Modified:** `code.gs`, `Code Scheduling.gs`, `Code Timelines - Daily Usage.gs`

---

### 27. **Add Scheduled Automatic ARW to JPG Conversion** (Commit: 4385445)
- **Problem:** Google Apps Script sending hourly failure emails: "Script function not found: scheduledAutoConverter"
- **Root Cause:** A time-based trigger was installed pointing to a function that no longer existed
- **Fix:** Implemented `scheduledAutoConverter()` function that:
  - Scans ALL sheets in the workbook for PHOTO sessions
  - Processes any unprocessed ARW files (skips sessions already converted)
  - Respects the 5-minute execution time limit
  - Logs detailed statistics and errors
  - Skips utility sheets (Dashboard, Studio Usage, Working Hours, etc.)
- **Also Added:** `installPhotoAutoConverterTrigger()` — removes old trigger, installs new hourly trigger
- **Menu Item:** 📸 Photo JPG Helper → ⚙️ Install Hourly Auto-Converter
- **Action Required:** Delete the old orphaned trigger in Apps Script Triggers panel, then run Install Hourly Auto-Converter from the menu
- **Files Modified:** `Photo Helpers.gs`

---

### 26. **Project Consolidation — Add Missing Files** (Commit: e96e4d3)
- **Problem:** Google Apps Script project had files that existed only online, not locally, causing out-of-sync issues
- **Fix:** Added two missing files to the local project:
  - `Data Integrity.gs` — copied from SG Weekly Board Sanity Check project, with duplicate `onOpen()` and `pad2_()` removed
  - `Photo Helpers.gs` — copied from Photo ARW Parser project
- **Restored:** Full menu system in `Code Scheduling.gs` (previously commented out due to missing functions)
- **Files Added:** `Data Integrity.gs`, `Photo Helpers.gs`
- **Files Modified:** `Code Scheduling.gs`

---

### 25. **Remove Duplicate Constants and Functions** (Commits: 538e0b4, 4f91dbf)
- **Problem:** Menus not appearing — "Identifier 'STUDIO_ORDER' has already been declared" error
- **Root Cause:** Google Apps Script loads all .gs files into the same global scope. Multiple files declared the same constants and functions
- **Duplicates Removed from `code.gs`:**
  - Constants: `STUDIO_ORDER`, `SET_ORDER`, `SET_TO_STUDIO`, `UNKNOWN_STUDIO`
- **Duplicates Removed from `Code Timelines - Daily Usage.gs`** (9 functions):
  - `detectDaySegments_`, `parseWeekHeader`, `fallbackWeekInfo`, `monthNameToIndex`
  - `_parseDateFromDayHeader_`, `_formatDateLabel_`, `_extractDayFromGroup_`
  - `_getTimelinePayloadForSheet_`, `_upsertDailyUsageRow_`
- **Files Modified:** `code.gs`, `Code Timelines - Daily Usage.gs`

---

### 24. **Fix Menu System — Multiple Syntax and Structural Errors** (Commits: 7fcb463, 2aaac14, d400012, 5d16845, 4131dae, dd51221)
- **Problem:** Menus disappeared completely after code updates
- **Causes and Fixes:**
  - Apostrophe inside single-quoted strings on lines 1006 and 1309 of `Code Scheduling.gs` — changed to double quotes
  - UTF-8 BOM at start of `code.gs` — removed
  - Duplicate `generateBreakPlanOnActiveSheet()` function in `Code Scheduling.gs` — removed
  - Missing opening `/*` on line 1 of `code.gs` (introduced when removing BOM) — restored
  - Extra `/` on closing comment line (introduced in previous fix) — corrected
- **Files Modified:** `code.gs`, `Code Scheduling.gs`

---

### 23. **Add Module Documentation for Combined Project** (Commit: 40e9554)
- **Added:** `docs/modules.md` documenting all 5 modules, their functions, dependencies, and origins
- **Updated:** `README.md` with new project structure, full feature list, and reference to module docs
- **Files Added:** `docs/modules.md`
- **Files Modified:** `README.md`

---

## 2026-02-15 - Dashboard & Analytics

#### 22. **Add Dashboard with Usage Trend Charts** (Commit: fc881ee)
- **New Feature:** Dashboard sheet with visual charts for tracking studio/set usage trends
- **Charts Created:**
  1. **Studio Usage Trends** - Grouped bar chart showing % usage over last 6 months
  2. **Set Usage Trends** - Grouped bar chart showing % usage over last 6 months
  3. **Current Month Studio Comparison** - Bar chart comparing hours for last complete month
- **Key Features:**
  - Consistent colors for each studio/set across all charts for easy tracking
  - Studio colors: Studio 1 (Green), Studio 2 (Blue), Studio 3 (Yellow), Studio 4 (Red)
  - Set colors: Iris (Purple), Club (Cyan), Nest (Green), Exec (Orange), Nova (Red), Soho (Indigo)
  - Reads from Daily Usage sheets for fast performance
  - Shows last 6 complete months of data
  - Percentage-based trends for easy comparison across months
  - Automatic update trigger for 1st of month at 6am
- **New Functions:**
  - `updateDashboard()` - Main function to rebuild dashboard with all 3 charts
  - `_getMonthlyAggregatedData_()` - Aggregates 6 months of data from Daily sheets
  - `_aggregateByMonth_()` - Calculates monthly totals and percentages
  - `_createStudioTrendsChart_()` - Creates grouped bar chart for studio trends
  - `_createSetTrendsChart_()` - Creates grouped bar chart for set trends
  - `_createCurrentMonthComparisonChart_()` - Creates comparison bar chart
  - `_getStudioColors_()` - Returns consistent color array for studios
  - `_getSetColors_()` - Returns consistent color array for sets
  - `installDashboardTrigger()` - Sets up monthly auto-update on 1st at 6am
- **Menu Items:**
  - 📈 Update Dashboard - Run immediately to create/update dashboard
  - ⚙️ Install Dashboard Update (1st of month, 6am) - Auto-update trigger
- **Files Modified:** `Code Timelines - Daily Usage.gs` lines 340-640; `Code Scheduling.gs` lines 57-63

#### 21. **Replace CSS Grid with Flexbox for Email Compatibility** (Commit: 56cd676)
- **Problem:** Summary cards in email were stacking vertically instead of horizontally
- **Root Cause:** CSS Grid not supported in most email clients (Gmail, Outlook, etc.)
- **Solution:** Replaced `display: grid` with `display: flex` for `.summary-cards`
- **Changes:**
  - Added vendor prefixes for maximum compatibility (-webkit-box, -ms-flexbox)
  - Added `flex: 1 1 200px` to `.card` for equal sizing with 200px minimum
  - Added `min-width: 200px` to prevent cards from getting too narrow
- **Impact:** 4 summary cards now display in horizontal row in email clients
- **Files Modified:** `Code Timelines - Daily Usage.gs` lines 699-722

#### 20. **Fix Email Layout to Match Popup Exactly** (Commit: 0587486)
- **Problem:** Email missing card layout with shadows and proper spacing
- **Solution:** Modified `_extractBodyContent_()` to extract complete `.wrap` div including the div tag itself
- **Impact:** Email displays with proper card layout identical to popup:
  - 4 summary cards in horizontal row
  - Cards with white background, shadows, and rounded corners
  - Proper padding and margins matching popup
- **Files Modified:** `Code Timelines - Daily Usage.gs` lines 304-317

#### 19. **Fix Email Formatting to Match Popup Styling** (Commit: 6b123bb)
- **Problem:** Monthly summary emails were displaying as plain text without styling
- **Solution:** Extract and include CSS from summary popups in email HTML
- **New Functions:**
  - `_extractCss_()` - Extracts `<style>` tags from HTML for email inclusion
  - Enhanced `_extractBodyContent_()` - Properly extracts .wrap content without Close button
- **Impact:** Email now looks identical to popup windows with all styling:
  - Gradient backgrounds and blue header chips
  - Cards with shadows and proper spacing
  - Color-coded table headers
  - Green usage bars with proper width percentages
  - Professional typography matching popup design
- **Files Modified:** `Code Timelines - Daily Usage.gs` lines 234-298

#### 18. **Fix Session Count Exclusion and Add Monthly Email Summary** (Commit: eaa3095)
- **Problem:** Total Sessions count was showing all sessions (107) even in exclusion mode
- **Solution:** Filter session count to exclude "Other" sessions when `excludeOther=true`
- **New Features:**
  - Monthly summary email functionality sends all 4 summaries to ben@poddster.com
  - Automatic trigger for 2nd of each month at 6am
  - Test email function for immediate testing
- **New Functions:**
  - `sendMonthlySummaryEmail()` - Main email function with all 4 summaries
  - `sendTestMonthlySummaryEmail()` - Test version for immediate sending
  - `installMonthlyEmailTrigger()` - Sets up monthly email automation
  - `_extractBodyContent_()` - Helper to extract HTML body content for email
- **Email Contents:** (in order)
  1. Studio Usage (with Other)
  2. Studio Usage (Excl. Other)
  3. Set Usage (with Other)
  4. Set Usage (Excl. Other)
- **Menu Items Added:**
  - 📧 Send Test Monthly Summary Email
  - ⚙️ Install Monthly Email (2nd of month, 6am)
- **Confirmed:** Month logic shows last complete month (Feb in March, etc.)
- **Files Modified:** `Code Timelines - Daily Usage.gs` lines 226-318; `Code Scheduling.gs` lines 54-59

#### 17. **Add Menu Emojis to Match Existing Style** (Commit: a703855)
- **Changes:**
  - Added 🎨 emoji to "Assign Operators" menu
  - Added 🕐 emoji to "Timelines" menu
  - Matches existing emoji style (🎬 for Sanity Checker, 📸 for Photo Helper)
- **Files Modified:** `Code Scheduling.gs` lines 35, 61

#### 16. **Add Monthly Summary Exclusion Mode and Menu Cleanup** (Commit: 8acad69)
- **Changes:**
  - Removed "Write Daily Studio Usage (Yesterday)" and "Write Daily Set Usage (Yesterday)" menu items and functions (replaced by `autoUpdateYesterdayData` which runs at 2am daily)
  - Updated backfill menu text to remove "(Year to Date)" suffix (it prompts for month selection, not year)
  - Added 2 new monthly summary functions that exclude "Other" category from calculations
- **New Functions:**
  - `showMonthlyStudioSummaryExclOther()` - Studio summary excluding Event/Other hours
  - `showMonthlySetSummaryExclOther()` - Set summary excluding Event/Other hours
- **Enhanced:** `_buildMonthlySummaryHtml_(type, excludeOther)` now supports exclusion mode:
  - Filters out "Other" category from table display when `excludeOther=true`
  - Calculates percentages based only on known studios/sets (not including Other)
  - Updates popup title and labels to show "(Excl. Other)" when appropriate
  - Uses `totalHoursExclOther` for average hours/day calculation
- **Menu Structure:**
  - 📊 Monthly Studio Summary (includes all categories)
  - 📊 Monthly Set Summary (includes all categories)
  - 📊 Monthly Studio Summary (Excl. Other) (only known studios)
  - 📊 Monthly Set Summary (Excl. Other) (only known sets)
- **Files Modified:** `Code Timelines - Daily Usage.gs` lines 195-228, 487-750; `Code Scheduling.gs` lines 47-58

### Major Changes

#### 1. **CRITICAL: Parse Dates from Row 2 Only** (Commit: 5c0f615)
- **Problem:** Date calculation was parsing tab names (e.g., "16th - 22nd Feb") which caused incorrect year attribution (2025 tabs being counted as 2026)
- **Solution:** Completely rewrote date parsing to ONLY use Row 2 cell values (e.g., "Monday, 16 February 2026")
- **Impact:** Eliminates all false date matches between 2025 and 2026 tabs
- **New Function:** `_parseDateFromRow2Cell_()` - parses actual dates from Row 2 format
- **Files Modified:** `Code Timelines - Daily Usage.gs`

#### 2. **Fix Strikethrough Detection** (Commits: a2d0ff0, 4ac9993)
- **Problem:** Strikethrough detection was either too aggressive (marking all sessions as cancelled) or checking wrong cells
- **Solution:** Store row/column info during parsing and check strikethrough on the specific row where name+time appear together
- **Impact:** Correctly identifies only cancelled sessions with strikethrough
- **Files Modified:** `Code Timelines - Daily Usage.gs` lines 448-480, 903-915

#### 3. **Fix Duplicate Session Parsing** (Commit: 012ad77)
- **Problem:** Multiple rows containing "Name/Time" were being processed, creating duplicate session entries
- **Solution:** Only process rows where FIRST cell (column A) contains "Name/Time"
- **Impact:** Each physical session counted exactly once
- **Files Modified:** `Code Timelines - Daily Usage.gs` lines 852-856

#### 4. **Enhanced Session Deduplication** (Commit: 012ad77)
- **Problem:** Case/whitespace variations in labels weren't being caught as duplicates
- **Solution:** Normalize labels and room names to lowercase before creating session IDs
- **Impact:** Prevents counting "Esther Lussier" and "esther lussier" as separate sessions
- **Files Modified:** `Code Timelines - Daily Usage.gs` lines 401-407

#### 5. **Removed Tab Name Filtering** (Commits: f6d3f2d, 64f932b, 5c0f615)
- **Problem:** Complex month-based filtering was unreliable and caused false positives/negatives
- **Solution:** Removed all tab name date filtering - Row 2 dates are the source of truth
- **Impact:** Simpler, more reliable logic with no edge cases
- **Files Modified:** `Code Timelines - Daily Usage.gs` lines 344-360

#### 6. **Fix Backfill to Exclude Today** (Commit: 0abbec0)
- **Problem:** Backfill was processing today's date with incomplete data
- **Solution:** Changed backfill to only process dates up to YESTERDAY
- **Impact:** Daily usage only shows completed days
- **Files Modified:** `Code Timelines - Daily Usage.gs` lines 127-131, 196-200

#### 7. **Skip Archived Tabs (Initial Attempt)** (Commits: cd69729, 5dd4c81)
- **Problem:** Old 2025 tabs without year markers were being processed
- **Solution:** Initially tried month-based filtering (later replaced with Row 2 parsing)
- **Note:** Superseded by commit 5c0f615 which eliminates need for tab filtering
- **Files Modified:** `Code Timelines - Daily Usage.gs`

#### 8. **Add Debug Logging** (Commit: 355d12f)
- **Problem:** Difficult to diagnose which sheets were contributing session data
- **Solution:** Added detailed logging showing which sessions are found and skipped
- **Impact:** Easier troubleshooting with execution logs
- **New Function:** `debugDateLookup()` - manual debug tool
- **Files Modified:** `Code Timelines - Daily Usage.gs` lines 365, 477-502

#### 9. **Comment Out DAY_NAMES Constant** (Commit: a503720)
- **Problem:** DAY_NAMES defined in multiple files caused script conflicts and menu disappearance
- **Solution:** Commented out DAY_NAMES in Daily Usage file (uses global from code.gs)
- **Impact:** Menus work correctly, no constant redefinition errors
- **Files Modified:** `Code Timelines - Daily Usage.gs` line 11

### Bug Fixes

- **Fixed Feb 15 phantom data:** Sunday sessions from 2025 tabs were being attributed to Feb 15, 2026
- **Fixed duplicate session counts:** Studio 1 showing 8 hours instead of 4 hours (2x duplication)
- **Fixed cancelled session inclusion:** Nest showing 3 hours instead of 1 hour (2 cancelled sessions counted)
- **Fixed menu disappearance:** Multiple onOpen functions conflicting
- **Fixed backfill timeouts:** Changed to process one month at a time
- **Fixed duplicate rows in output:** Changed from getValues() to getDisplayValues() for date comparison
- **Fixed missing dates:** Added required helper functions to Daily Usage file

### Breaking Changes

**IMPORTANT:** The date parsing logic has been completely rewritten. All dates are now extracted from Row 2 ONLY.

**Required Row 2 Format:** `"Weekday, DD Month YYYY"` (e.g., "Monday, 16 February 2026")

**No longer supported:**
- Tab name date calculation
- Week header parsing for date inference
- Month-based filtering logic

### Migration Guide

1. **Update Google Apps Script:**
   - Copy entire `Code Timelines - Daily Usage.gs` file to your online Apps Script project
   - Ensure `DAY_NAMES` constant is commented out (line 11)
   - Save and refresh your Google Sheet

2. **Clear Existing Data:**
   - Delete all rows in "Studio Usage (Daily)" sheet (keep header)
   - Delete all rows in "Set Usage (Daily)" sheet (keep header)

3. **Verify Row 2 Format:**
   - Check that all weekly timeline tabs have Row 2 with format: "Weekday, DD Month YYYY"
   - Example: "Monday, 16 February 2026" (not just "Monday" or "16 Feb")

4. **Run Backfill:**
   - Menu: Usage Tracking > 🔄 Backfill Daily Studio Usage (Year to Date)
   - Enter month number (1-12) when prompted
   - Repeat for Set Usage if needed

5. **Verify Results:**
   - Check execution logs (Extensions > Apps Script > Executions)
   - Verify totals match manual counts from weekly sheets
   - Confirmed cancelled sessions (with strikethrough) are excluded

### Testing Performed

- ✅ Feb 9 totals verified (Studio 1: 4h, Nest: 1h)
- ✅ Strikethrough detection working (Damini Chawla, Michael Velten, CLW - Kristina excluded)
- ✅ No duplicate sessions counted
- ✅ 2025 tabs not processed (dates don't match 2026 dates from Row 2)
- ✅ Backfill processes correct date ranges
- ✅ Menus working correctly

### Known Limitations

- Requires Row 2 to have full date format: "Weekday, DD Month YYYY"
- Old tabs with different date formats in Row 2 will fail to parse (by design)
- Performance: Processing all sheets takes time (consider using backfill monthly)

### Files Changed

- `Code Timelines - Daily Usage.gs` - 200+ lines modified
- `CHANGELOG.md` - Created this file

### Commit History (15 commits)

1. `5c0f615` - CRITICAL: Parse dates from Row 2 only, not tab names
2. `64f932b` - Fix filter to allow current year tabs without explicit year markers
3. `f6d3f2d` - Strict filter: only process tabs with '2026' or '26 in name
4. `4ac9993` - Fix strikethrough detection to check actual session row
5. `a2d0ff0` - Fix overly aggressive strikethrough detection - use exact match only
6. `012ad77` - Fix duplicate sessions and strikethrough detection
7. `cd69729` - Fix overly aggressive filter - only skip tabs 3+ months old
8. `5dd4c81` - Skip archived tabs from 2025 and earlier to prevent date parsing issues
9. `355d12f` - Add debug logging to track which sheets contribute session data
10. `0abbec0` - Exclude today from backfill - only process up to yesterday
11. `cd69729` - Fix overly aggressive filter - only skip tabs 3+ months old
12. `7d8986c` - Fix duplicate rows bug in upsert function
13. `0c31fab` - Fix backfill timeout - process one month at a time
14. `a503720` - Comment out DAY_NAMES to avoid conflict with code.gs
15. `24b2957` - Add Code Scheduling.gs reference file with updated menu

### Contributors

- Claude Code (AI Assistant)
- Ben Draycott-Jones (User/Tester)

### Next Steps

- Consider applying same fixes to `code.gs` for consistency with Operator Timeline feature
- Monitor daily auto-update trigger (2am) to ensure it runs correctly
- Add monthly summary reports as needed

---

**Last Updated:** February 15, 2026
**Version:** 2.0 (Daily Usage Tracking - Row 2 Date Parsing)
