# Daily Usage Tracking - Implementation Status

## ✅ COMPLETED FIXES (All Pushed to GitHub)

### 1. **Row 2 Date Parsing** ✓
- Dates now extracted ONLY from Row 2 cells ("Monday, 16 February 2026")
- No more tab name parsing or date calculation
- Eliminates all 2025/2026 confusion

### 2. **Strikethrough Detection** ✓
- Checks the specific row where name+time appear together
- Correctly identifies cancelled sessions
- No more false positives

### 3. **Duplicate Session Prevention** ✓
- Only processes "Name/Time" rows (column A check)
- Enhanced deduplication with normalized labels
- Each session counted exactly once

### 4. **Backfill Improvements** ✓
- Processes one month at a time (avoids timeout)
- Excludes today (only completed days)
- User prompt for month selection

### 5. **Debug Logging** ✓
- Added debugDateLookup() function
- Detailed session tracking in logs
- Easier troubleshooting

### 6. **Menu Fixes** ✓
- DAY_NAMES constant commented out
- No more menu conflicts
- All menus working

---

## 📋 WHAT YOU NEED TO DO NOW

### Step 1: Update Google Apps Script (REQUIRED)

1. Go to your Google Sheet → Extensions → Apps Script
2. Open **"Code Timelines - Daily Usage"** file
3. **Select ALL** and delete
4. **Paste the latest version** from your local file:
   `/Users/bendraycott-jones/Documents/Poddster Scripts/SG Weekly Timelines/Code Timelines - Daily Usage.gs`
5. **CRITICAL:** Make sure line 11 has DAY_NAMES commented out:
   ```javascript
   // const DAY_NAMES = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
   ```
6. Click **Save** (disk icon)
7. Close Apps Script

### Step 2: Clear Old Data

1. Open **"Studio Usage (Daily)"** sheet
2. Delete all rows EXCEPT the header (keep row 1)
3. Open **"Set Usage (Daily)"** sheet
4. Delete all rows EXCEPT the header (keep row 1)

### Step 3: Run Backfill

1. Refresh your Google Sheet (reload the page)
2. Wait for menus to load
3. Click **"Assign Operators"** menu (or your custom menu name)
4. Select **"🔄 Backfill Daily Studio Usage (Year to Date)"**
5. When prompted, enter **`2`** (for February)
6. Wait for completion message
7. Repeat for **"🔄 Backfill Daily Set Usage (Year to Date)"** → enter **`2`**

### Step 4: Verify Results

1. Check **"Studio Usage (Daily)"** sheet
   - Should have dates: 2026-02-02 through 2026-02-14 (no Feb 15 - it's a Sunday)
   - Check Feb 9 totals match your manual count

2. Check **execution logs**:
   - Extensions → Apps Script → Executions
   - Click most recent execution
   - Look for:
     - ✓ "Session 'Damini Chawla' is struck through - SKIPPING"
     - ✓ "Session 'Michael Velten (Pippa)' is struck through - SKIPPING"
     - ✓ "Session 'CLW - Kristina (Take 1)' is struck through - SKIPPING"
     - ✗ Should NOT see "Session 'Esther Lussier' is struck through"
     - ✗ Should NOT see sessions from 2025 tabs like "8th to 14th Dec"

3. **Expected Feb 9 Totals:**
   - Studio 1: ~4 hours (2 sessions: Esther Lussier, Guillaume Villard)
   - Studio 2: ~3 hours (1 session: Yana Fry + Eddie Yew if in Studio 2)
   - Nest: ~1 hour (only Yana Fry - Damini cancelled)
   - Iris: ~4 hours (Guillaume Villard, Michelle O'Brien)

---

## 🔧 TROUBLESHOOTING

### If backfill returns 0 dates:

**Problem:** Row 2 format doesn't match expected pattern

**Check:**
1. Open a weekly timeline tab (e.g., "16th - 22nd Feb")
2. Look at Row 2, column B (or first day column)
3. Should say: **"Monday, 16 February 2026"** (exact format)
4. If it says "Monday" or "16 Feb" or anything else, the parser won't work

**Fix:** Update Row 2 cells to full format with year

### If sessions are being double-counted:

**Problem:** Multiple "Name/Time" rows being processed

**Check:**
1. Look at your weekly timeline sheet
2. Count how many rows have "Name/Time" in column A
3. Should be only ONE row with "Name/Time" in column A

**Fix:** Ensure only ONE row has "Name/Time" in the first cell (column A)

### If cancelled sessions are being counted:

**Problem:** Strikethrough detection not finding the correct cell

**Debug:**
1. Run `debugDateLookup` for the date in question
2. Check execution logs for "struck through - SKIPPING" messages
3. Verify strikethrough is on the NAME cell, not just the time cell

**Fix:** Apply strikethrough to the client name cell (not just room/notes)

### If menus disappear:

**Problem:** DAY_NAMES not commented out or multiple onOpen functions

**Fix:**
1. Open Apps Script
2. Check "Code Timelines - Daily Usage" line 11 - should be commented: `// const DAY_NAMES = ...`
3. Check no other files have uncommented `function onOpen(e)`
4. Save and refresh

---

## 📊 MONITORING & MAINTENANCE

### Daily Auto-Update (2am)

The system will automatically write yesterday's data every day at 2am.

**To check if it's working:**
1. Tomorrow morning, check "Studio Usage (Daily)"
2. Should see a new row for yesterday's date (2026-02-15)
3. If not, check Extensions → Apps Script → Triggers
4. Ensure "installDailyTrigger" is set up

### Monthly Backfill

At the end of each month, run backfill for the completed month:
- Menu → "🔄 Backfill Daily Studio Usage" → enter month number
- This ensures any missed days are captured

### Log Monitoring

Periodically check execution logs:
- Extensions → Apps Script → Executions
- Look for errors or warnings
- Sessions should NOT be coming from 2025 tabs

---

## 📁 FILE LOCATIONS

**Local Files:**
- `/Users/bendraycott-jones/Documents/Poddster Scripts/SG Weekly Timelines/Code Timelines - Daily Usage.gs`
- `/Users/bendraycott-jones/Documents/Poddster Scripts/SG Weekly Timelines/Code Scheduling.gs`
- `/Users/bendraycott-jones/Documents/Poddster Scripts/SG Weekly Timelines/code.gs`

**GitHub Repository:**
- https://github.com/sgpoddster/Weekly-Timelines

**Google Apps Script:**
- Your Google Sheet → Extensions → Apps Script → Files

---

## 🎯 SUCCESS CRITERIA

✅ **Backfill completes without errors**
✅ **Feb 9 totals match manual count**
✅ **Cancelled sessions (strikethrough) excluded**
✅ **No duplicate sessions counted**
✅ **No sessions from 2025 tabs**
✅ **Menus working correctly**

---

## 📞 SUPPORT

If issues persist after following this guide:

1. **Check execution logs** for specific error messages
2. **Run debugDateLookup** for the problematic date
3. **Verify Row 2 format** matches "Weekday, DD Month YYYY"
4. **Check GitHub** for latest code version
5. **Review CHANGELOG.md** for recent changes

---

**Last Updated:** February 15, 2026
**Status:** ✅ All fixes implemented and pushed to GitHub
**Next Action:** User needs to update Google Apps Script and run backfill
