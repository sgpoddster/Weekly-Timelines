# Weekly Timelines - Poddster Singapore

A Google Apps Script project for managing and visualizing weekly podcast production timelines across studios, operators, and recording sessions.

## Overview

This project provides an integrated system for tracking podcast recording sessions in Google Sheets, with interactive timeline visualizations for both room scheduling and operator assignments. It's designed for production teams managing multiple recording studios with various operators and equipment configurations.

## Features

### 📊 Timeline Visualizations
- **Room Timeline**: Visual timeline showing studio/set usage across the week
- **Operator Timeline**: Track operator (Naz, Syaz, Sufi) assignments and availability
- Both timelines support interactive HTML dialogs with detailed session information

### 👥 Operator Management
- Assign operators to sessions marked as "TBC"
- Undo assignments by day or for the entire sheet
- Track hours summary per operator
- Break plan management
- Studio and set usage analytics

### 📈 Analytics & Reporting
- Hours summary by operator
- Studio and set usage tracking (weekly, daily, and monthly)
- **Daily usage tracking** - Track studio/set usage day by day with automatic date extraction
- **Automatic daily updates** - Runs at 2am daily to update yesterday's data automatically
- **Backfill capability** - Populate historical data for any month
- **Monthly summaries** - Beautiful popup summaries with charts and usage bars
- **Monthly summary emails** - Automatic email reports on the 2nd of each month
- **Dashboard with charts** - Visual trends showing studio/set usage over 6 months with consistent colors
- Detailed session breakdowns with strikethrough detection for cancelled sessions
- Usage statistics across multiple sheets
- Exclusion mode to view summaries without "Other/Event" category

### 🎬 Integration Features
- Session data integrity checking
- Photo JPG folder management
- HTML generation from existing JPGs
- Email notifications and calendar sync capabilities

## Project Structure

```
SG Weekly Timelines/
├── Code Scheduling.gs                  # Main menu and scheduling functions
├── Code Timelines - Daily Usage.gs     # Daily/monthly usage tracking
├── timeline by room.html               # Room timeline visualization UI
├── timeline by operator.html           # Operator timeline visualization UI
├── README.md                           # This file
├── CHANGELOG.md                        # Version history and changes
├── ARCHITECTURE.md                     # Technical architecture documentation
└── IMPLEMENTATION_STATUS.md            # Daily usage implementation guide
```

## Installation

### Prerequisites
- Google Workspace account
- Access to Google Sheets
- Basic understanding of Google Apps Script

### Setup Steps

1. **Create a Google Sheet** for your weekly production schedule
   - Set up your sheet with the expected format (see Sheet Format section)

2. **Open Script Editor**
   - In your Google Sheet, go to `Extensions > Apps Script`

3. **Copy the Code**
   - Copy the contents of `code.gs` into the script editor
   - Create HTML files for the timeline visualizations:
     - `File > New > HTML file` → Name it `timeline by room`
     - Copy contents from `timeline by room.html`
     - Repeat for `timeline by operator.html`

4. **Save and Authorize**
   - Save the project
   - Refresh your Google Sheet
   - You should see new menus appear: "👩‍🎨 Assign Operators", "⏱️ Timelines", etc.
   - Run any function and authorize the script when prompted

## Sheet Format

The script expects Google Sheets with the following structure:

### Header Structure
- **Row 1**: Week identifier (e.g., "6th - 12th October" or "V")
- **Row 2**: Day headers with dates (e.g., "Monday, 19 January 2026")

### Day Columns
Each day section should contain:
- **Name/Time**: Session name and time range (e.g., "10:00 - 13:00")
- **Room/Seats**: Studio/set name and seat count
- **PP Status & Add-Ons**: Post-production status
- **Upload Status**: Upload/Sent status
- **Notes**: Additional session notes

### Supported Day Formats
The script automatically detects day headers in these formats:
- Simple: "Monday", "Tuesday", etc.
- With date: "Monday, 19 January 2026"
- With full date: "Tuesday, 20 January 2026"

## Usage

### Viewing Timelines

1. **Room Timeline**
   - Go to `⏱️ Timelines > Show me the room timelines!`
   - View sessions organized by day and room
   - See host, client, and timing information

2. **Operator Timeline**
   - Go to `⏱️ Timelines > Show me the operator timelines!`
   - View operator assignments across the week
   - Track workload distribution

### Managing Operators

1. **Assign Operators**
   - Mark sessions with "TBC" where operators are needed
   - Go to `👩‍🎨 Assign Operators > Assign Editors to TBC on Active Sheet`
   - The script will automatically assign available operators

2. **Undo Assignments**
   - Undo all: `👩‍🎨 Assign Operators > Undo (replace names back to TBC on Active Sheet)`
   - Undo by day: `👩‍🎨 Assign Operators > Undo for Day…`
   - Undo for active cell's day: `👩‍🎨 Assign Operators > Undo for Day of Active Cell`

3. **View Analytics**
   - Hours summary: `👩‍🎨 Assign Operators > Show Hours Summary (Active Sheet)`
   - Break plan: `👩‍🎨 Assign Operators > Show Break Plan (Active Sheet)`
   - Studio usage: `👩‍🎨 Assign Operators > Show Studio & Set Usage (Active Sheet)`

### Tracking Daily Usage

1. **Automatic Daily Updates**
   - Install trigger: `🎨 Assign Operators > ⚙️ Install Daily Auto-Update (2am)`
   - Runs automatically at 2am every day
   - Updates yesterday's studio and set usage data
   - No manual intervention required

2. **Backfill Historical Data**
   - Studio backfill: `🎨 Assign Operators > 🔄 Backfill Daily Studio Usage`
   - Set backfill: `🎨 Assign Operators > 🔄 Backfill Daily Set Usage`
   - Select any month to backfill data
   - Creates/updates "Studio Usage (Daily)" and "Set Usage (Daily)" sheets
   - Date format: YYYY-MM-DD for easy sorting

3. **View Monthly Summaries**
   - `🎨 Assign Operators > 📊 Monthly Studio Summary` - All studios including Event/Other
   - `🎨 Assign Operators > 📊 Monthly Set Summary` - All sets including Event/Other
   - `🎨 Assign Operators > 📊 Monthly Studio Summary (Excl. Other)` - Known studios only
   - `🎨 Assign Operators > 📊 Monthly Set Summary (Excl. Other)` - Known sets only
   - Shows data for last complete month with beautiful charts and usage bars
   - Displays total hours, working days, average hours/day, session counts, and percentages

4. **Monthly Email Reports**
   - Install trigger: `🎨 Assign Operators > ⚙️ Install Monthly Email (2nd of month, 6am)`
   - Sends automatic email on 2nd of each month at 6am to ben@poddster.com
   - Contains all 4 monthly summaries (Studio/Set with and without "Other")
   - Test it: `🎨 Assign Operators > 📧 Send Test Monthly Summary Email`

5. **Dashboard with Charts**
   - Update dashboard: `🎨 Assign Operators > 📈 Update Dashboard`
   - Install trigger: `🎨 Assign Operators > ⚙️ Install Dashboard Update (1st of month, 6am)`
   - Creates "Dashboard" sheet with 3 charts:
     - **Studio Usage Trends** - Grouped bar chart showing % usage over last 6 months
     - **Set Usage Trends** - Grouped bar chart showing % usage over last 6 months
     - **Current Month Studio Comparison** - Bar chart comparing hours for last complete month
   - Each studio/set has consistent color across all charts for easy tracking
   - Automatically updates on 1st of each month

## Configuration

### Operator Names
The script recognizes these operators (case-insensitive):
- Naz (Reservist)
- Syaz
- Sufi
- Ben

To add more operators, modify the `normOp()` function in `code.gs`.

### Time Format
Sessions must use the format: `HH:MM - HH:MM` (e.g., "10:00 - 13:00")

### Studio/Room Names
Common rooms recognized:
- Nest
- Iris
- Exec
- Club
- Nova

## Troubleshooting

### "Could not find a row with day names"
- Ensure row 2 contains day names (Monday, Tuesday, etc.)
- Day names can include dates (e.g., "Monday, 19 January 2026")
- Check that the `DAY_NAMES` constant is defined in `code.gs`

### Timeline shows no sessions
- Verify that your sheet has "Name/Time" rows
- Check that time formats match `HH:MM - HH:MM`
- Ensure there are 2 columns to the left of times for client/host names

### Operator not recognized
- Check spelling in the operator column
- Verify the `normOp()` function includes your operator's name
- Operator names should be in the first column of each day segment

## Development

### Testing
The script includes various test functions accessible through the custom menus.

### Debugging
Use `Logger.log()` statements and view logs in Apps Script:
- `View > Logs` in the script editor

### Contributing
This is a custom internal tool. For modifications:
1. Test changes thoroughly on a copy of your sheet
2. Document any new features in CHANGELOG.md
3. Update this README with new usage instructions

## Related Projects

- **SG Weekly Board Sanity Check**: Session data integrity checker with remote/live handling

## License

Internal tool for Poddster Singapore. All rights reserved.

## Support

For issues or questions:
- Check ARCHITECTURE.md for technical details
- Review CHANGELOG.md for recent changes
- Contact the development team

---

**Version**: 1.0.0
**Last Updated**: February 2026
**Maintained by**: Poddster Singapore Team
