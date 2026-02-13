# Changelog

All notable changes to the Weekly Timelines project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.1.0] - 2026-02-13

### Added
- **Daily usage tracking** - Track studio and set usage day by day
  - `writeDailyStudioUsage()` - Write daily studio usage for active sheet
  - `writeDailySetUsage()` - Write daily set usage for active sheet
  - Automatically extracts dates from row 2 day headers (e.g., "Monday, 9 February 2026")
  - Creates "Studio Usage (Daily)" and "Set Usage (Daily)" sheets
  - Date format: YYYY-MM-DD for easy sorting and analysis
- **Backfill capability** - Populate historical data from beginning of year to today
  - `backfillDailyStudioUsage()` - Backfill all sheets from Jan 1 to today
  - `backfillDailySetUsage()` - Backfill all sheets from Jan 1 to today
  - Processes all sheets in workbook (skips usage sheets)
  - One-time operation to populate historical daily data
- Menu items for daily tracking under "👩‍🎨 Assign Operators"
  - "Write Daily Studio Usage (Active Sheet)"
  - "Write Daily Set Usage (Active Sheet)"
  - "🔄 Backfill Daily Studio Usage (Year to Date)"
  - "🔄 Backfill Daily Set Usage (Year to Date)"

### Changed
- Enhanced analytics capability from weekly-only to daily + weekly + monthly tracking

### Technical Details
- Added `_getDailyStatsForActiveSheet_()` - Get daily breakdown for active sheet
- Added `_getDailyStatsForSheet_(sheet)` - Get daily breakdown for any sheet
- Added `_parseDateFromDayHeader_(cellText)` - Parse dates from day headers
- Added `_formatDateLabel_(date)` - Format dates as YYYY-MM-DD
- Added `_extractDayFromGroup_(groupStr)` - Extract day name from group strings
- Added `_getTimelinePayloadForSheet_(sheet)` - Get timeline data for specific sheet
- Added `_upsertDailyUsageRow_(opts)` - Upsert daily usage row by date (updates existing or appends new)
- Daily usage sheets auto-format: hours (0.00), percentages (0.00%), counts (0)

## [1.0.0] - 2026-02-12

### Added
- Initial release of Weekly Timelines system
- Room timeline visualization showing studio/set usage across the week
- Operator timeline visualization tracking Naz, Syaz, and Sufi assignments
- Interactive HTML dialogs for both timeline views
- Operator assignment system for TBC sessions
- Hours summary reporting per operator
- Break plan management
- Studio and set usage analytics (weekly and monthly)
- Undo functionality for operator assignments (by sheet, by day, by active cell)
- Support for multiple day header formats including dates
- Week header parsing from cell A1 (e.g., "6th - 12th October")
- Robust day detection that scans first 10 rows for day headers
- Custom menu integration: "👩‍🎨 Assign Operators" and "⏱️ Timelines"
- Data integrity checker menu integration
- Photo JPG helper menu integration

### Fixed
- **CRITICAL**: Added missing `DAY_NAMES` constant definition that was causing "Could not find a row with day names" error
  - Added `const DAY_NAMES = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];`
  - This constant is used throughout the codebase for day detection and parsing
  - Fixed issue where operator timeline and room timeline functions would fail

### Technical Details
- Implemented `detectDaySegments_()` helper function for flexible day header detection
- Supports day headers in formats like "Monday", "Monday, 19 January 2026", etc.
- Uses `startsWith()` matching for day name detection to handle date suffixes
- Parses time ranges in format `HH:MM - HH:MM`
- Extracts client/host names from 2 columns left of time ranges
- Reads room/seats information from dedicated row near time entries
- Maps day names to actual dates based on Monday anchor from A1 week header
- Operator normalization function handles variations in operator names

### Known Issues
- None currently reported

---

## [Unreleased]

### Planned Features
- [ ] Export timeline to PDF
- [ ] Email notifications for operator assignments
- [ ] Capacity warning when operators are over-scheduled
- [ ] Multi-week view support
- [ ] Calendar integration for automated sync
- [ ] Mobile-responsive timeline views
- [ ] Filtering by studio/operator in timeline views
- [ ] Session conflict detection
- [ ] Automated break scheduling suggestions
- [ ] Historical usage analytics

### Under Consideration
- Integration with project management tools
- API for external system integration
- Real-time collaboration features
- Custom color coding for session types
- Drag-and-drop timeline editing

---

## Version History Summary

| Version | Date       | Major Changes                                    |
|---------|------------|--------------------------------------------------|
| 1.0.0   | 2026-02-12 | Initial release with timelines and operator mgmt |

---

## Migration Notes

### From Spreadsheet-Only Management
If you're migrating from manual spreadsheet management:
1. Ensure your day headers are in row 2
2. Format times as `HH:MM - HH:MM`
3. Place client/host names in the 2 columns before the time
4. Add "Name/Time" labels to identify session rows
5. Add "Room/Seats" rows below each Name/Time row

### Format Changes in 1.0.0
- **Day Header Format**: Now supports dates in day headers (e.g., "Monday, 19 January 2026")
  - Previously: Expected simple day names like "Monday"
  - Now: Accepts any cell that **starts with** a day name
  - This change is **backwards compatible** - old format still works

---

## Deprecation Notices

### None Currently

---

## Security Updates

### None Currently

---

## Breaking Changes

### None in 1.0.0
This is the initial release.

---

## Contributors

- Development Team at Poddster Singapore
- Bug fixes and improvements: Claude (AI Assistant)

---

## Notes

### Version Numbering
- **Major version** (X.0.0): Breaking changes, major feature additions
- **Minor version** (1.X.0): New features, non-breaking changes
- **Patch version** (1.0.X): Bug fixes, minor improvements

### Reporting Issues
When reporting issues, please include:
- Google Sheets structure/format
- Error messages from Apps Script logs
- Steps to reproduce
- Expected vs actual behavior
- Screenshots if applicable
