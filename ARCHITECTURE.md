# Architecture Documentation

## System Overview

The Weekly Timelines system is a Google Apps Script application that integrates with Google Sheets to provide production scheduling, visualization, and analytics for podcast recording sessions.

This is a **combined project** integrating five separate modules into a single Google Apps Script deployment. See `docs/modules.md` for detailed module documentation.

## File Structure

```
SG Weekly Timelines/
├── code.gs                           # Core timeline parsing + shared helper functions
├── Code Scheduling.gs                # Master onOpen() + operator assignment
├── Code Timelines - Daily Usage.gs   # Daily/monthly usage tracking + dashboard
├── Data Integrity.gs                 # Session data validation and checks
├── Photo Helpers.gs                  # ARW to JPG conversion and management
├── timeline by room.html             # Room timeline visualization UI
├── timeline by operator.html         # Operator timeline visualization UI
├── docs/
│   └── modules.md                    # Detailed per-module documentation
├── README.md
└── CHANGELOG.md
```

> **Important:** All .gs files share the same global scope in Google Apps Script.
> Constants and functions must only be declared **once** across all files.
> `onOpen()` must only exist in `Code Scheduling.gs`.

## Architecture Diagram

```
┌─────────────────────────────────────────────────────────────────────────┐
│                          Google Sheets UI                               │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐  ┌────────────┐ │
│  │ 🎨 Assign    │  │ 🕐 Timelines │  │ 🎬 Data      │  │ 📸 Photo   │ │
│  │  Operators   │  │    Menu      │  │   Checker    │  │   Helper   │ │
│  └──────────────┘  └──────────────┘  └──────────────┘  └────────────┘ │
└────────────────────────┬────────────────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────────────────┐
│                    Google Apps Script Layer                             │
│  ┌───────────────────────────────────────────────────────────────────┐ │
│  │  Code Scheduling.gs — Menu System & Operator Assignment           │ │
│  │  • onOpen() — Master menu builder (calls all module menus)        │ │
│  │  • Operator assignment (assignEditorsOnActiveSheet)               │ │
│  │  • Analytics (hours summary, break plans, studio/set usage)       │ │
│  │  • Shared constants: STUDIO_ORDER, SET_ORDER, SET_TO_STUDIO       │ │
│  └───────────────────────────────────────────────────────────────────┘ │
│  ┌───────────────────────────────────────────────────────────────────┐ │
│  │  code.gs — Core Parsing Engine & Shared Helpers                  │ │
│  │  • detectDaySegments_() — Locate day columns in sheet             │ │
│  │  • _getTimelinePayloadForSheet_() — Extract all sessions          │ │
│  │  • _upsertDailyUsageRow_() — Write/update daily usage rows        │ │
│  │  • parseWeekHeader() — A1 text fallback for date detection        │ │
│  │  • _parseDateFromRow2Cell_() — Extract dates from Row 2 text      │ │
│  │  • _isSessionCrossedOut_() — Detect strikethrough sessions        │ │
│  │  • Timeline generators (getWeeklyTimelinePayload etc.)            │ │
│  └───────────────────────────────────────────────────────────────────┘ │
│  ┌───────────────────────────────────────────────────────────────────┐ │
│  │  Code Timelines - Daily Usage.gs — Tracking & Dashboard           │ │
│  │  • autoUpdateYesterdayData() — 2am daily trigger                  │ │
│  │  • backfillDailyStudioUsage/Set() — Populate historical data      │ │
│  │  • deduplicateDailyUsageSheets() — Fix duplicate rows             │ │
│  │  • showMonthlyStudioSummary/Set() — Usage popup dialogs           │ │
│  │  • setupDashboard() — One-time chart creation (script editor only) │ │
│  │  • updateDashboardData() — Monthly data refresh (menu)            │ │
│  │  • sendMonthlySummaryEmail() — Monthly email report               │ │
│  └───────────────────────────────────────────────────────────────────┘ │
│  ┌───────────────────────────────────────────────────────────────────┐ │
│  │  Data Integrity.gs — Session Validation                           │ │
│  │  • runSanityToday/Yesterday/WholeTab() — Data checks              │ │
│  │  • Camera, audio, timestamp drift validation                      │ │
│  │  • sendSanityTestEmail() — Email integrity report                 │ │
│  └───────────────────────────────────────────────────────────────────┘ │
│  ┌───────────────────────────────────────────────────────────────────┐ │
│  │  Photo Helpers.gs — ARW to JPG Conversion                         │ │
│  │  • createPhotoJpgMenu() — Called by onOpen()                      │ │
│  │  • scheduledAutoConverter() — Hourly trigger                      │ │
│  │  • createJpgFoldersFromBoard() — Manual run                       │ │
│  │  • installPhotoAutoConverterTrigger() — Install trigger           │ │
│  └───────────────────────────────────────────────────────────────────┘ │
└────────────────────────┬────────────────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────────────────┐
│                    Output & Visualization Layer                         │
│  ┌──────────────────────┐  ┌──────────────────────┐  ┌──────────────┐ │
│  │ HTML Timeline Dialogs│  │ Monthly Summary      │  │ Email Reports│ │
│  │ • timeline by room   │  │ Popups               │  │              │ │
│  │ • timeline by        │  │ • Styled cards       │  │ • Styled     │ │
│  │   operator           │  │ • Usage bars         │  │   summaries  │ │
│  │                      │  │ • Progress bars      │  │ • Sent 2nd   │ │
│  │                      │  │                      │  │   of month   │ │
│  └──────────────────────┘  └──────────────────────┘  └──────────────┘ │
│                                                                         │
│  ┌─────────────────────────────────────────────────────────────────┐   │
│  │ Daily Usage Sheets (Auto-maintained)                            │   │
│  │ • Studio Usage (Daily) — One row per date, upserted by date key  │   │
│  │ • Set Usage (Daily) — One row per date, upserted by date key     │   │
│  │ Format: Date | Name (h) | Name (%) | Name (#) | ...             │   │
│  └─────────────────────────────────────────────────────────────────┘   │
│                                                                         │
│  ┌─────────────────────────────────────────────────────────────────┐   │
│  │ Dashboard Sheet (data refreshed monthly, charts permanent)      │   │
│  │ • Overview: Studio % and Hours — full-width column charts       │   │
│  │ • Overview: Set % and Hours — full-width column charts          │   │
│  │ • Individual: one dual-axis chart per studio (excl. Other)      │   │
│  │ • Individual: one dual-axis chart per set   (excl. Other)       │   │
│  │ • Hidden data tables in cols 18-91 (pre-allocated 60 months)    │   │
│  └─────────────────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────────────────┐
│                    Time-Based Triggers                                  │
│  • Hourly: scheduledAutoConverter() — ARW to JPG processing             │
│  • Daily 2am: autoUpdateYesterdayData() — Update previous day           │
│  • Monthly 1st at 6am: updateDashboardData() — Refresh chart data        │
│  • Monthly 2nd at 6am: sendMonthlySummaryEmail() — Send report          │
└─────────────────────────────────────────────────────────────────────────┘
```

## Component Architecture

### 1. Menu System

**File**: `Code Scheduling.gs`

Single master `onOpen()` function registers all menus. Module-specific menus are built by calling dedicated menu builder functions.

```javascript
function onOpen(e)
  ├── Creates "🎨 Assign Operators" menu  [Code Scheduling.gs]
  ├── Creates "🕐 Timelines" menu         [code.gs]
  ├── Creates "🎬 Data Checker" menu      [Data Integrity.gs]
  └── Calls createPhotoJpgMenu()          [Photo Helpers.gs]
```

**Critical rule**: Only ONE `onOpen()` may exist across all files. New modules must provide a menu builder function (e.g. `createMyModuleMenu()`) called from this `onOpen()`.

### 2. Core Parsing Engine

#### 2.1 Day Detection System

**Function**: `detectDaySegments_(values, maxScanRows)` — `code.gs`

**Purpose**: Locates day headers in the spreadsheet and determines the column ranges for each day.

**Algorithm**:
```
1. Scan first maxScanRows (default 10) rows
2. For each cell, check if it starts with a day name from DAY_NAMES
3. When found, mark that row as headerRowIdx
4. Scan that row to build day segment array
5. Each segment: { day, startCol, endCol }
6. Return { headerRowIdx, dayCols }
```

**Output**:
```javascript
{
  headerRowIdx: 1,
  dayCols: [
    { day: "Monday", startCol: 2, endCol: 6 },
    { day: "Tuesday", startCol: 6, endCol: 10 },
  ]
}
```

#### 2.2 Date Detection (Three-Method System)

**Function**: `_getTimelinePayloadForSheet_(sheet)` — `code.gs`

Dates for each day column are determined using three methods in priority order:

```
Method 1 — Raw Date cell (post-Nov 2025 sheets)
  Row 2 cells contain actual Google Sheets Date objects
  → rawValues[headerRowIdx][col] instanceof Date → use directly

Method 2 — Text parsing of row 2 display value
  Row 2 cells display "Monday, 9 March 2026"
  → _parseDateFromRow2Cell_() regex match → use if valid

Method 3 — A1 week header text (pre-Nov 2025 sheets)
  Row 2 only has plain text "Monday", "Tuesday" etc.
  A1 contains e.g. "9th Mar - 15th Mar"
  → parseWeekHeader(a1) + day offset calculation

If all three fail → skip the day, log a warning
NEVER use fallbackWeekInfo() — it returns the current week
and would attribute historical sessions to today's dates
```

**Sheet format history**:
- Sheets before Nov 24 2025: Row 2 = plain text day names → Method 3
- Sheets from Nov 24 2025: Row 2 = actual date cells → Method 1

#### 2.3 Constants

**File**: `Code Scheduling.gs` (declared once, shared globally)

```javascript
const STUDIO_ORDER  = ['Studio 1', 'Studio 2', 'Studio 3', 'Studio 4', 'Other'];
const SET_ORDER     = ['Iris', 'Club', 'Nest', 'Exec', 'Nova', 'Soho', 'Other'];
const SET_TO_STUDIO = { 'iris': 'Studio 2', 'club': 'Studio 2', ... };
const UNKNOWN_STUDIO = 'Other';
```

**File**: `code.gs`

```javascript
const DAY_NAMES = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
```

**Critical Rule**: All constants declared **once** across all .gs files. Google Apps Script shares one global scope — duplicate `const` declarations cause a runtime error that silently prevents `onOpen()` from running and all menus disappearing.

### 3. Timeline Generation System

#### 3.1 Room Timeline Generator

**Function**: `getWeeklyTimelinePayload()`
**Lines**: 234-372

**Purpose**: Extracts session data organized by room/studio.

**Data Flow**:
```
Sheet Data
   ↓
detectDaySegments_() → Find day columns
   ↓
Three-method date detection (see §2.2) → Assign dates to day columns
   ↓
Scan for "Name/Time" rows
   ↓
Extract time ranges (HH:MM - HH:MM)
   ↓
Read client/host (2 cols left of time)
   ↓
Read room/seats from nearby row
   ↓
Create timeline items
   ↓
Return payload with items array
```

**Item Structure**:
```javascript
{
  group: "Monday — Iris",      // Day + Room
  label: "Client Name",         // Session client
  host: "Host Name",           // Session host
  room: "Iris",                // Studio/set name
  seats: "2",                  // Seat count
  start: Date,                 // Start datetime
  end: Date,                   // End datetime
  startMs: 1234567890,         // Start timestamp
  endMs: 1234567999            // End timestamp
}
```

#### 3.2 Operator Timeline Generator

**Function**: `getOperatorTimelinePayload()`
**Lines**: 403-556

**Purpose**: Extracts session data organized by operator assignment.

**Additional Features**:
- Reads operator name from first column of day segment
- Walks upward to handle merged cells
- Normalizes operator names (Naz, Syaz, Sufi)
- Groups sessions by "Day — Operator"

**Operator Normalization**:
```javascript
function normOp(x)
  • "naz" / "nazreen" → "Naz"
  • "syaz" / "syazwan" → "Syaz"
  • "fadz" / "fadzli" → "Sufi"
  • Others → Capitalized
```

#### 3.3 Safe Payload Wrappers

**Functions**:
- `getWeeklyTimelinePayloadSafe()` (lines 375-395)
- `getOperatorTimelinePayloadSafe()` (lines 559-580)

**Purpose**: Sanitize data for safe transmission to HTML dialogs.

**Sanitization**:
- Convert all fields to strings/numbers
- Filter invalid items (non-finite timestamps)
- Remove items where end <= start
- Ensure all required fields present

### 4. HTML Timeline Views

#### 4.1 Room Timeline View

**File**: `timeline by room.html`

**Features**:
- Visual timeline with time axis
- Color-coded sessions by room
- Hover tooltips with session details
- Groups organized by "Day — Room"
- Responsive layout

**Technology Stack**:
- HTML5
- CSS3 for styling
- JavaScript for rendering
- Google Apps Script client-side API

#### 4.2 Operator Timeline View

**File**: `timeline by operator.html`

**Features**:
- Visual timeline with time axis
- Color-coded sessions by operator
- Workload visualization
- Groups organized by "Day — Operator"
- Session overlap detection

**Rendering Flow**:
```
HTML Dialog Opened
   ↓
google.script.run.getOperatorTimelinePayloadSafe()
   ↓
Receive payload with items array
   ↓
Group items by operator and day
   ↓
Calculate time axis (earliest to latest)
   ↓
Render timeline bars positioned by time
   ↓
Add interactivity (tooltips, clicks)
```

## Data Model

### Sheet Structure

```
Row 1: [Week Header]                     (e.g., "6th - 12th October")
Row 2: [Day Headers with Dates]          (e.g., "Monday, 19 January 2026")
Row 3: [Name/Time] [Client] [Host] [Time] ...
Row 4: [Room/Seats] [Room] [Seats] ...
Row 5: [PP Status & Add-Ons] ...
Row 6: [Upload Status] ...
Row 7: [Notes] ...
Row 8: [Name/Time] [Client] [Host] [Time] ...  (Next session block)
...
```

### Session Block Pattern

Each session is represented by a cluster of rows:

```
Columns relative to time cell:
  [-2]      [-1]     [0]        [+1]    [+2]
[Client]  [Host]  [Time]   [More info] ...

Room row (1-5 rows below):
[Room Name] [Seat Count]
```

### Time Format

**Required Format**: `HH:MM - HH:MM`

**Examples**:
- ✅ "10:00 - 13:00"
- ✅ "14:30 - 15:30"
- ❌ "10am - 1pm"
- ❌ "10:00-13:00" (missing spaces)
- ❌ "10:00 to 13:00"

**Regex**: `/^\s*(\d{1,2}):(\d{2})\s*[-–—]\s*(\d{1,2}):(\d{2})\s*$/`

## Key Algorithms

### 1. Session Detection Algorithm

```
FOR each row after header row:
  IF row contains "Name/Time" label:
    Find corresponding "Room/Seats" row (scan next 6 rows)

    FOR each day segment:
      Extract segment values (startCol to endCol)

      FOR each cell in segment:
        IF cell matches time regex:
          Read client/host from 2 cells to the left
          Calculate absolute column position

          IF room row exists:
            Extract room/seats near this column position

          Create Date objects for start/end times
          Build timeline item
          Add to results array
```

### 2. Room/Seats Extraction Algorithm

```
pickRoomAndSeats(rowVals, seg, baseCol):
  Define search window: baseCol-2 to baseCol+4
  Limit window to segment bounds

  FOR each column in window:
    current = cell at column
    next = cell at column+1

    IF room not found AND current is text (not number):
      room = current

    IF seats not found AND next is number:
      seats = next

    IF both found:
      break

  RETURN { room, seats }
```

### 3. Operator Detection Algorithm

```
getOperatorFromSegment(rowIdx, seg):
  column = first column of segment

  FOR row from current upward (max 10 rows):
    value = cell at (row, column)

    IF value is not empty:
      RETURN normalized operator name

  RETURN empty string
```

**Rationale**: Handles merged cells where operator name spans multiple session rows.

### 5. Daily Usage Tracking System

**File**: `Code Timelines - Daily Usage.gs`

The daily usage tracking system provides automated data collection, monthly summaries, and email reporting for studio and set usage.

#### 5.1 Core Architecture

**Purpose**: Track daily studio/set usage hours and session counts, with automatic updates and monthly reporting.

**Data Flow**:
```
Weekly Timeline Sheets (Row 2 with dates)
   ↓
Parse dates from Row 2 ("Monday, 16 February 2026")
   ↓
Extract sessions for each date
   ↓
Check strikethrough (cancelled sessions)
   ↓
Deduplicate sessions
   ↓
Aggregate by studio/set
   ↓
Write to Daily Usage Sheets (upsert by date)
   ↓
Monthly summaries read from Daily sheets
   ↓
Email reports sent monthly
```

#### 5.2 Date Parsing System

**Function**: `_parseDateFromRow2Cell_(cellValue)`
**Critical Design Decision**: **ONLY** parse dates from Row 2 cells, never from tab names.

**Why Row 2 Only?**
- Tab names vary in format ("16th - 22nd Feb", "22nd - 28th Sep")
- Tab names may not include year (causes 2025/2026 confusion)
- Row 2 format is consistent: "Monday, 16 February 2026"
- Row 2 is the single source of truth for dates

**Algorithm**:
```
1. Extract cell value from Row 2
2. Match pattern: "DayName, DD MonthName YYYY"
3. Parse day number, month name, year
4. Convert month name to index (0-11)
5. Create Date object
6. Return Date or null if parsing fails
```

**Regex Pattern**: `/(\w+),\s+(\d{1,2})\s+(\w+)\s+(\d{4})/`

#### 5.3 Session Detection and Filtering

**Function**: `parseSheetForDates(sheetName, targetDateStr, type)`

**Multi-Stage Filtering**:

1. **Row Detection** (lines 852-856)
   ```javascript
   const firstCell = String(values[r][0] || '').trim().toLowerCase();
   const isNameTimeRow = /^name\s*\/\s*time/i.test(firstCell);
   ```
   - Only process rows where FIRST cell contains "Name/Time"
   - Prevents duplicate sessions from multiple "Name/Time" occurrences

2. **Strikethrough Detection** (lines 426-480)
   ```javascript
   function _isSessionCrossedOut_(sheet, item)
   ```
   - Stores row/column coordinates during parsing
   - Checks getFontLines() on specific cells where session was found
   - Checks name cell, time cell, and nearby cells
   - Returns true if ANY have 'line-through' formatting
   - **Purpose**: Exclude cancelled sessions from totals

3. **Session Deduplication** (lines 401-407)
   ```javascript
   const normalizedLabel = String(it.label || '').toLowerCase().trim();
   const normalizedRoom = String(it.room || '').toLowerCase().trim();
   const sessionId = normalizedLabel + '|' + normalizedRoom + '|' + startStr + '|' + endStr;
   ```
   - Creates unique ID from label + room + time
   - Normalizes to lowercase to catch "Esther Lussier" vs "esther lussier"
   - Prevents counting same session multiple times

4. **Date Matching**
   - Only includes sessions where Row 2 date matches target date
   - Format: YYYY-MM-DD for comparison

#### 5.4 Studio/Set Aggregation

**Studio Mapping** (lines 19-32):
```javascript
const SET_TO_STUDIO = {
  'iris': 'Studio 2',
  'club': 'Studio 2',
  'nova': 'Studio 3',
  'nest': 'Studio 1',
  'exec': 'Studio 1',
  'soho': 'Studio 4'
};
```

**Set Categorization** (lines 388-399):
```javascript
const knownSets = ['iris', 'club', 'nest', 'exec', 'nova', 'soho'];
const setName = knownSets.includes(roomName.toLowerCase()) ? roomName : 'Other';
```
- Known sets mapped to their names
- Unknown sets (Event, etc.) → 'Other'

**Aggregation Logic**:
```
FOR each valid session:
  Calculate hours = (endMs - startMs) / 3600000

  For Studio Usage:
    Map set → studio (via SET_TO_STUDIO)
    Accumulate hours and count by studio name

  For Set Usage:
    Use set name directly (or 'Other')
    Accumulate hours and count by set name
```

#### 5.5 Upsert Strategy

**Function**: `_writeDailyUsageRow_(sheetName, date, aggregated, labels, type)`
**Lines**: 265-343

**Purpose**: Insert or update daily usage data without duplicates.

**Algorithm**:
```
1. Get or create daily usage sheet
2. Ensure header row exists with columns for each studio/set
3. Read all existing data with getDisplayValues()
4. Search for existing row with matching date (compare as strings)
5. Build new row data: [Date, Hours1, Count1, Hours2, Count2, ..., Total]
6. IF row exists:
     Update existing row values
   ELSE:
     Append new row
7. Write back to sheet
```

**Why getDisplayValues()?**
- Dates in sheets might be formatted differently
- String comparison handles all date formats consistently
- Avoids date timezone issues

#### 5.6 Automatic Daily Updates

**Function**: `autoUpdateYesterdayData()`
**Lines**: 264-278
**Trigger**: Daily at 2am (installed via `installDailyTrigger()`)

**Process**:
```
1. Calculate yesterday's date
2. Format as YYYY-MM-DD
3. Call backfillDailyStudioUsage(yesterdayStr)
4. Call backfillDailySetUsage(yesterdayStr)
5. Log completion
```

**Why Yesterday?**
- Today's data is incomplete (day not finished)
- Yesterday is the last complete day
- Runs at 2am when all sheets are finalized

#### 5.7 Backfill System

**Functions**:
- `backfillDailyStudioUsage()` - User-triggered backfill
- `backfillDailySetUsage()` - User-triggered backfill
- `_backfillDailyUsage_(type, startDateStr, endDateStr)` - Core logic

**Features**:
- Prompts user to select month
- Processes entire month (1st to last day)
- Skips future dates
- Updates existing rows if re-run
- Creates "Studio Usage (Daily)" and "Set Usage (Daily)" sheets

**Performance**:
- Processes one date at a time
- Scans all timeline sheets for each date
- ~30 seconds for full month backfill

#### 5.8 Monthly Summary System

**Functions**:
- `showMonthlyStudioSummary()` - All studios
- `showMonthlySetSummary()` - All sets
- `showMonthlyStudioSummaryExclOther()` - Exclude "Other"
- `showMonthlySetSummaryExclOther()` - Exclude "Other"

**Data Source**: Reads from Daily Usage sheets (NOT re-parsing timelines)

**Performance Benefit**: ~30x faster than re-parsing

**Function**: `_buildMonthlySummaryHtml_(type, excludeOther)`
**Lines**: 487-760

**Process**:
```
1. Calculate last complete month
   now.getMonth() - 1 = previous month
2. Read from Daily Usage sheet
3. Filter rows by month
4. Aggregate hours and counts by studio/set
5. Calculate totals (with/without "Other")
6. Count working days in month
7. Calculate percentages
8. Generate styled HTML with:
   - Summary cards (total hours, avg/day, sessions)
   - Table with usage bars
   - Gradient styling and colors
9. Display in modal dialog (1100x750)
```

**Exclusion Mode** (`excludeOther=true`):
- Filters out "Other" category from table
- Calculates percentages based only on known studios/sets
- Shows "(Excl. Other)" in title
- Excludes "Other" sessions from total count

**Working Days Calculation**:
```javascript
function _countWorkingDays_(startDate, endDate)
  FOR each day in range:
    IF dayOfWeek is Monday-Friday:
      count++
  RETURN count
```

#### 5.9 Email Reporting System

**Functions**:
- `sendMonthlySummaryEmail()` - Main email function
- `sendTestMonthlySummaryEmail()` - Test email
- `installMonthlyEmailTrigger()` - Install trigger for 2nd of month at 6am

**Email Structure**:
```html
<html>
  <head>
    <style>
      /* All CSS from monthly summary popups */
      /* Gradient backgrounds, cards, tables, bars */
    </style>
  </head>
  <body>
    <h1>Monthly Usage Summary - January 2026</h1>

    <!-- Summary 1: Studio Usage (with Other) -->
    <div class="wrap">...</div>
    <hr>

    <!-- Summary 2: Studio Usage (Excl. Other) -->
    <div class="wrap">...</div>
    <hr>

    <!-- Summary 3: Set Usage (with Other) -->
    <div class="wrap">...</div>
    <hr>

    <!-- Summary 4: Set Usage (Excl. Other) -->
    <div class="wrap">...</div>
  </body>
</html>
```

**CSS Extraction** (`_extractCss_()`):
- Extracts `<style>` tags from first summary
- All 4 summaries share same CSS
- Preserves gradients, shadows, colors, bars

**Content Extraction** (`_extractBodyContent_()`):
- Extracts `.wrap` div content
- Removes Close button
- Preserves all styling and structure

**Recipient**: ben@poddster.com

**Trigger**: 2nd of each month at 6am

#### 5.10 Dashboard and Charts System

**File**: `Code Timelines - Daily Usage.gs`

The dashboard provides persistent visual analytics. Charts are created once and survive indefinitely — only the underlying data tables are rewritten on each monthly refresh.

##### Two-function design

| Function | When to run | What it does |
|---|---|---|
| `setupDashboard()` | Once (script editor only) | Full rebuild: deletes all charts, rewrites layout, creates new charts. Run when first setting up or after a major chart reset. |
| `updateDashboardData()` | Monthly (menu / trigger) | Rewrites hidden data tables only. Charts auto-refresh. All manual settings (data labels, etc.) are preserved. |

After running `setupDashboard()`, enable data labels on each individual chart once:
> Right-click chart → Edit chart → Customise → Series → tick **Data labels**

Labels persist through all future `updateDashboardData()` calls.

##### Hidden data table layout

All data lives in pre-allocated hidden columns (cols 18-91), keeping the visible area chart-only. Each table has a fixed header row plus `DASH_MAX_MONTHS = 60` pre-allocated data rows (5 years of headroom). Because the range never changes, chart ranges never need updating.

```
Cols 18-23   Overview Studio %      (Month + Studio 1/2/3/4/Other)
Cols 24-29   Overview Studio Hours  (same structure)
Cols 30-37   Overview Set %         (Month + Iris/Club/Nest/Exec/Nova/Soho/Other)
Cols 38-45   Overview Set Hours     (same structure)
Cols 46-65   Individual studio charts  (Studio 1, Studio 2, Studio 3, Studio 4 — 5 cols each)
Cols 66-91   Individual set charts     (Iris, Club, Nest, Exec, Nova, Soho — 5 cols each)
```

Each individual chart's 5-column block: `Month | Hours | HoursLabel | Pct | PctLabel`

##### Fixed visual layout

Chart positions are derived from named constants so the layout never shifts:

```
Row 4   Section header: Studio Overview
Row 5   Overview Studio % chart        (380px tall, full width)
Row 29  Overview Studio Hours chart
Row 53  Section header: Set Overview
Row 54  Overview Set % chart
Row 78  Overview Set Hours chart
Row 102 Section header: Individual Studios
Row 103 Studio 1 (col A) | Studio 2 (col I)
Row 135 Studio 3 (col A) | Studio 4 (col I)
Row 167 Section header: Individual Sets
Row 168 Iris (col A) | Club (col I)   (550px tall)
Row 200 Nest (col A) | Exec (col I)
Row 232 Nova (col A) | Soho (col I)
```

##### Data aggregation (`_getMonthlyAggregatedData_`)

```
1. Build month list: Oct 2025 → last completed month
2. Read Studio Usage (Daily) and Set Usage (Daily) sheets
3. Parse header to find "(h)" columns → studio/set labels
4. FOR each date row: accumulate hours by label by month key (YYYY-MM)
5. FOR each month: calculate % = (label_hours / total_hours) × 100
6. Return { studios, sets, months }
```

##### Color Consistency

Studio Colors (`_getStudioColors_()`):
```javascript
{ 'Studio 1': '#34a853', 'Studio 2': '#4285f4', 'Studio 3': '#fbbc04', 'Studio 4': '#ea4335', 'Other': '#9e9e9e' }
```

Set Colors (`_getSetColors_()`):
```javascript
{ 'Iris': '#9c27b0', 'Club': '#00bcd4', 'Nest': '#4caf50', 'Exec': '#ff9800', 'Nova': '#f44336', 'Soho': '#3f51b5', 'Other': '#9e9e9e' }
```

Individual charts use a lighter shade of each label's color for the % bar (`_lightenColor_(hex, 0.45)`).

**Trigger**: 1st of each month at 6am — calls `updateDashboardData()` only

## Design Patterns

### 1. Template Method Pattern

The timeline generation follows a template:
1. Detect structure (days, segments)
2. Parse temporal anchor (week header)
3. Scan for session markers
4. Extract session details
5. Map to timeline items
6. Return standardized payload

Both room and operator timelines follow this template with variations in grouping.

### 2. Strategy Pattern

Different parsing strategies for different contexts:
- Day detection: Flexible startsWith matching
- Week header: Multiple regex patterns with fallback
- Operator normalization: Prefix matching

### 3. Facade Pattern

Safe payload wrappers (`*Safe()` functions) provide a clean interface to the HTML layer, hiding data sanitization complexity.

## Performance Considerations

### Optimization Strategies

1. **Early Termination**
   - Day header scan stops after first match
   - Room/seats search stops when both found
   - Limited row scanning (maxScanRows = 10)

2. **Efficient Data Structures**
   - Pre-computed day segments reduce redundant scans
   - Single pass through rows for session detection
   - Column ranges prevent full-row processing

3. **Lazy Evaluation**
   - HTML dialogs only load when opened
   - Timeline payload generated on-demand
   - No background processing

### Scalability Limits

- **Max rows**: Effectively unlimited (scans entire sheet)
- **Max sessions per day**: Unlimited
- **Max days per week**: 7 (defined by DAY_NAMES)
- **Performance**: O(n*m) where n=rows, m=columns
- **Expected performance**: <2 seconds for typical weekly sheet

## Error Handling

### Error Scenarios

1. **No day headers found**
   - Error: "Could not find a row with day names in the first few rows."
   - Solution: Check row 2 contains day names

2. **Empty sheet**
   - Error: "Sheet looks empty."
   - Solution: Ensure sheet has data range

3. **Invalid time format**
   - Behavior: Session silently skipped
   - No error thrown (graceful degradation)

4. **Missing client/host**
   - Behavior: Session skipped if both empty
   - Fallback: "(TBC)" label if client empty

### Error Handling Strategy

- **Critical errors**: Throw with descriptive message
- **Data validation errors**: Skip item, continue processing
- **Missing data**: Use fallback values or empty strings
- **Safe wrappers**: Filter invalid items before HTML delivery

## Security Considerations

### Authorization Scopes

Required Google Apps Script permissions:
- `SpreadsheetApp`: Read sheet data
- `HtmlService`: Display dialogs
- `Ui`: Show menus and dialogs

### Data Access

- **Read-only**: Timeline functions only read data
- **Write operations**: Separate functions (not in timeline system)
- **User data**: Stays within Google Workspace
- **No external calls**: All processing server-side

### Input Validation

- Time strings validated via regex
- Date parsing with fallbacks
- Numeric conversions with bounds checking
- HTML output sanitized by Google Apps Script

## Extension Points

### Adding New Timeline Views

1. Create new HTML file (e.g., `timeline by client.html`)
2. Create payload function (e.g., `getClientTimelinePayload()`)
3. Implement grouping logic (group by client)
4. Add menu item in `onOpen()`
5. Add dialog function (e.g., `openClientTimeline()`)

### Adding New Operators

Modify `normOp()` function:
```javascript
function normOp(x) {
  const v = String(x || '').trim().toLowerCase().replace(/[^\p{L}\p{N}]+/gu, '');
  if (v.startsWith('naz'))  return 'Naz';
  if (v.startsWith('syaz')) return 'Syaz';
  if (v.startsWith('fadz')) return 'Sufi';
  if (v.startsWith('newop')) return 'NewOperator';  // Add here
  return v ? v.charAt(0).toUpperCase() + v.slice(1) : '';
}
```

### Adding New Day Detection Logic

Extend `detectDaySegments_()`:
```javascript
// Current: startsWith matching
if (lc.startsWith(dn)) { ... }

// Could add: Contains matching, regex, etc.
if (lc.includes(dn) || regexMatch(lc, dn)) { ... }
```

## Testing Strategy

### Manual Testing Checklist

- [ ] Open sheet with various day header formats
- [ ] Verify timelines display correctly
- [ ] Test with empty sessions
- [ ] Test with missing data (no room, no host)
- [ ] Test with merged cells
- [ ] Test with multiple session blocks
- [ ] Test week header parsing with various formats
- [ ] Test operator name variations
- [ ] Test with sessions at day boundaries (early/late)
- [ ] Test with overlapping sessions

### Test Data Requirements

Create test sheets with:
- Standard format (simple day names)
- Extended format (day names with dates)
- Edge cases (single session, many sessions)
- Invalid data (bad times, missing fields)
- Various operators (Naz, Syaz, Sufi, others)

## Dependencies

### Google Apps Script APIs

- `SpreadsheetApp`: Sheet access
- `HtmlService`: Dialog creation
- `Ui`: Menu and dialog display
- `Logger`: Debugging (optional)

### External Dependencies

None. Pure Google Apps Script implementation.

### Browser Requirements

HTML dialogs require modern browser with:
- JavaScript ES6+ support
- CSS3 support
- HTML5 support

## Future Architecture Considerations

### Potential Improvements

1. **Caching Layer**
   - Cache parsed sheet structure
   - Invalidate on sheet edit
   - Reduce re-parsing overhead

2. **Event-Driven Updates**
   - Use `onEdit()` trigger
   - Auto-update timelines on changes
   - Real-time collaboration support

3. **Modular Architecture**
   - Split `code.gs` into modules
   - Separate parsing, timeline, UI concerns
   - Easier testing and maintenance

4. **Data Layer Abstraction**
   - Abstract sheet format dependencies
   - Support multiple sheet layouts
   - Easier migration to new formats

5. **API Layer**
   - Expose timeline data via Web App
   - External system integration
   - Mobile app support

## Maintenance Notes

### Code Hygiene

- Keep functions under 100 lines
- Document complex algorithms
- Use descriptive variable names
- Add comments for non-obvious logic

### Version Control

- Track changes in CHANGELOG.md
- Use semantic versioning
- Tag releases in Git
- Document breaking changes

### Documentation

- Keep this ARCHITECTURE.md updated
- Update README.md for user-facing changes
- Add JSDoc comments to functions
- Include examples in documentation
