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
- **Backfill capability** - Populate historical data from beginning of year to today
- Detailed session breakdowns
- Usage statistics across multiple sheets

### 🎬 Integration Features
- Session data integrity checking
- Photo JPG folder management
- HTML generation from existing JPGs
- Email notifications and calendar sync capabilities

## Project Structure

```
SG Weekly Timelines/
├── code.gs                      # Main Google Apps Script file
├── timeline by room.html        # Room timeline visualization UI
├── timeline by operator.html    # Operator timeline visualization UI
├── README.md                    # This file
├── CHANGELOG.md                 # Version history and changes
└── ARCHITECTURE.md              # Technical architecture documentation
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

1. **Write Daily Usage for Active Sheet**
   - Studio usage: `👩‍🎨 Assign Operators > Write Daily Studio Usage (Active Sheet)`
   - Set usage: `👩‍🎨 Assign Operators > Write Daily Set Usage (Active Sheet)`
   - This extracts dates from row 2 and creates one row per day

2. **Backfill Historical Data** (One-time operation)
   - Studio backfill: `👩‍🎨 Assign Operators > 🔄 Backfill Daily Studio Usage (Year to Date)`
   - Set backfill: `👩‍🎨 Assign Operators > 🔄 Backfill Daily Set Usage (Year to Date)`
   - Processes all sheets from January 1st to today
   - Creates/updates "Studio Usage (Daily)" and "Set Usage (Daily)" sheets
   - Date format: YYYY-MM-DD for easy sorting

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
