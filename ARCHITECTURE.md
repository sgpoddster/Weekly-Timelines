# Architecture Documentation

## System Overview

The Weekly Timelines system is a Google Apps Script application that integrates with Google Sheets to provide production scheduling, visualization, and analytics for podcast recording sessions.

## Architecture Diagram

```
┌─────────────────────────────────────────────────────────────┐
│                     Google Sheets UI                        │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐     │
│  │   Assign     │  │  Timelines   │  │ Data Checker │     │
│  │  Operators   │  │    Menu      │  │    Menu      │     │
│  └──────────────┘  └──────────────┘  └──────────────┘     │
└────────────────────────┬────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────┐
│              Google Apps Script (code.gs)                   │
│  ┌──────────────────────────────────────────────────────┐  │
│  │  Menu System (onOpen)                                │  │
│  └──────────────────────────────────────────────────────┘  │
│  ┌──────────────────────────────────────────────────────┐  │
│  │  Core Parsing Engine                                 │  │
│  │  • detectDaySegments_()                             │  │
│  │  • parseWeekHeader()                                │  │
│  │  • Time/date parsing utilities                      │  │
│  └──────────────────────────────────────────────────────┘  │
│  ┌──────────────────────────────────────────────────────┐  │
│  │  Timeline Generators                                 │  │
│  │  • getWeeklyTimelinePayload()                       │  │
│  │  • getOperatorTimelinePayload()                     │  │
│  └──────────────────────────────────────────────────────┘  │
└────────────────────────┬────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────┐
│              HTML Timeline Dialogs                          │
│  ┌──────────────────────┐  ┌──────────────────────┐        │
│  │ timeline by room.html│  │ timeline by          │        │
│  │ • Visual timeline    │  │ operator.html        │        │
│  │ • Session details    │  │ • Operator view      │        │
│  │ • Room grouping      │  │ • Workload tracking  │        │
│  └──────────────────────┘  └──────────────────────┘        │
└─────────────────────────────────────────────────────────────┘
```

## Component Architecture

### 1. Menu System

**File**: `code.gs` (lines 12-59)

The menu system provides user interface integration with Google Sheets through custom menus.

```javascript
function onOpen(e)
  ├── Creates "👩‍🎨 Assign Operators" menu
  ├── Creates "⏱️ Timelines" menu
  ├── Creates "🎬 Data Checker" menu
  └── Creates "📸 Photo JPG Helper" menu
```

**Responsibilities**:
- Register custom menus on sheet open
- Link menu items to corresponding functions
- Provide user-friendly access to all features

### 2. Core Parsing Engine

#### 2.1 Day Detection System

**Function**: `detectDaySegments_(values, maxScanRows)`
**Lines**: 95-157

**Purpose**: Locates day headers in the spreadsheet and determines the column ranges for each day.

**Algorithm**:
```
1. Scan first maxScanRows (default 10) rows
2. For each cell, check if it starts with a day name from DAY_NAMES
3. When found, mark that row as headerRowIdx
4. Scan that row to build day segment array
5. Each segment contains:
   - day: Day name (e.g., "Monday")
   - startCol: Starting column index
   - endCol: Ending column index (or sheet width)
6. Return { headerRowIdx, dayCols }
```

**Flexibility**:
- Accepts cells starting with day names (supports dates)
- Scans multiple rows to find headers
- Works with merged cells
- Handles any column layout

**Input**: 2D array of cell values from sheet
**Output**:
```javascript
{
  headerRowIdx: 1,  // Row index where day headers found
  dayCols: [
    { day: "Monday", startCol: 2, endCol: 6 },
    { day: "Tuesday", startCol: 6, endCol: 10 },
    // ...
  ]
}
```

#### 2.2 Week Header Parsing

**Function**: `parseWeekHeader(txt)`
**Lines**: 164-205

**Purpose**: Extracts the Monday date for the week from cell A1.

**Supported Formats**:
- "6th - 12th October"
- "6th - 12th October 2026"
- "13th January"
- "13th January 2026"

**Algorithm**:
```
1. Try to match date range pattern (e.g., "6th - 12th October")
2. If no match, try single date pattern (e.g., "13th January")
3. Extract day number, month name, and optional year
4. Convert month name to index (0-11)
5. Create Date object for the first day mentioned
6. Calculate Monday of that week
7. Return { mondayDate, year }
```

**Fallback**: If parsing fails, uses current week's Monday

#### 2.3 Constants

**Lines**: 5-9

```javascript
const DAY_NAMES = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
```

**Critical Dependency**: Used throughout the codebase for:
- Day header detection
- Day name standardization
- Day-to-date mapping
- Timeline generation

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
parseWeekHeader() → Get week anchor date
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
- Normalizes operator names (Naz, Syaz, Fadzli)
- Groups sessions by "Day — Operator"

**Operator Normalization**:
```javascript
function normOp(x)
  • "naz" / "nazreen" → "Naz"
  • "syaz" / "syazwan" → "Syaz"
  • "fadz" / "fadzli" → "Fadzli"
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
  if (v.startsWith('fadz')) return 'Fadzli';
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
- Various operators (Naz, Syaz, Fadzli, others)

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
