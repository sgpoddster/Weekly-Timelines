# Module Documentation

This project combines multiple feature modules into a single Google Apps Script project.

## Modules

### 1. Core Timelines (`code.gs`)
**Purpose:** Timeline parsing and display functions

**Key Functions:**
- `openTimeline()` - Show room timelines
- `openOperatorTimeline()` - Show operator timelines
- `detectDaySegments_()` - Parse weekly schedule structure
- `_getTimelinePayloadForSheet_()` - Extract timeline data

**Original Project:** SG Weekly Timelines (core)

---

### 2. Operator Assignment (`Code Scheduling.gs`)
**Purpose:** Automatic editor assignment to TBC slots

**Key Functions:**
- `onOpen()` - **Master menu builder** (calls all other menu builders)
- `assignEditorsOnActiveSheet()` - Assign editors to TBC
- `showBreakPlanDialogOnActiveSheet()` - Show break planning
- `showHoursSummaryDialogOnActiveSheet()` - Hours summary

**Constants Defined:**
- `STUDIO_ORDER`, `SET_ORDER`, `SET_TO_STUDIO`, `UNKNOWN_STUDIO`

**Original Project:** SG Weekly Timelines (core)

---

### 3. Daily Usage Tracking (`Code Timelines - Daily Usage.gs`)
**Purpose:** Track and visualize studio/set usage over time

**Key Functions:**
- `backfillDailyStudioUsage()` - Populate historical studio data
- `backfillDailySetUsage()` - Populate historical set data
- `showMonthlyStudioSummary()` - Monthly studio report
- `updateDashboard()` - Generate usage trend charts
- `installDashboardTrigger()` - Schedule monthly updates

**Charts:**
- Studio Usage Trends (last 6 months)
- Set Usage Trends (last 6 months)
- Current Month Comparison

**Original Project:** SG Weekly Timelines (dashboard feature added Feb 2025)

---

### 4. Data Integrity Checker (`Data Integrity.gs`)
**Purpose:** Validate session data completeness and file integrity

**Key Functions:**
- `runSanityToday()` - Check today's session data
- `runSanityYesterday()` - Check yesterday's data
- `runSanityWholeTab()` - Check entire sheet
- `sendSanityTestEmail()` - Email integrity report

**Checks:**
- Camera file counts
- Audio file counts
- Remote session handling
- Timestamp drift detection
- File size validation

**Original Project:** SG Weekly Board Sanity Check

---

### 5. Photo Management (`Photo Helpers.gs`)
**Purpose:** ARW to JPG conversion and photo folder management

**Key Functions:**
- `createPhotoJpgMenu()` - Build photo helper menu
- `convertArwFromActiveCell()` - Convert ARW files to JPG
- `createJpgFoldersFromBoard()` - Create JPG folders for sessions
- `sendTestEmailActiveRow()` - Test email delivery

**Integration:**
- Connects to GCP Cloud Run service for ARW conversion
- Creates organized JPG folder structure
- Generates HTML galleries

**Original Project:** Photo ARW Parser

---

## Module Dependencies

```
onOpen() [Code Scheduling.gs]
  ├── Calls: createPhotoJpgMenu() [Photo Helpers.gs]
  └── References: All menu functions from all modules

Timeline Functions [code.gs]
  └── Used by: Code Timelines - Daily Usage.gs

Constants [Code Scheduling.gs]
  ├── STUDIO_ORDER, SET_ORDER
  ├── SET_TO_STUDIO, UNKNOWN_STUDIO
  └── Used by: code.gs, Code Timelines - Daily Usage.gs

Helper Functions
  ├── pad2_() [Code Scheduling.gs]
  ├── detectDaySegments_() [code.gs]
  └── Various date/parsing helpers [code.gs]
```

## Adding New Modules

When adding a new module:

1. **DO NOT** add another `onOpen()` function
2. **DO** add a menu builder function (e.g., `createMyModuleMenu()`)
3. Call your menu builder from the master `onOpen()` in Code Scheduling.gs
4. Check for duplicate constants/functions before adding
5. Document the module in this file

## Changelog Organization

- **CHANGELOG.md** - Main project changelog (all modules)
- Each significant module update should note which module changed
- Format: `[Module Name] Description of change`

Example:
```
## [2025-02-15]
### Added
- [Dashboard] Studio and Set usage trend charts
- [Dashboard] Monthly comparison visualization

### Fixed
- [Data Integrity] Timestamp drift detection for remote sessions
```
