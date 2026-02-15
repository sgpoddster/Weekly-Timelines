# Changelog - Daily Usage Tracking Fixes

## 2026-02-15 - Critical Fixes for Daily Usage Tracking

### Latest Updates

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
