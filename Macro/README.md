# Outlook Attachment Macro - Subject-Based Saving

This folder contains an advanced Outlook VBA macro that automatically saves email attachments based on configurable subject line matching rules.

## Overview

Unlike traditional attachment-saving macros that rely on filename patterns, this macro uses email subject lines to determine where and how to save attachments. It provides flexible matching options and saves attachments to both primary and backup locations.

## Features

- **Flexible Subject Matching**: Four different matching types (EXACT, CONTAINS, STARTS_WITH, ENDS_WITH)
- **Dual-Folder Saving**: Saves to both primary destination and backup folder
- **Configurable File Naming**: Support for timestamp placeholders in filenames
- **Comprehensive Logging**: Detailed logs with timestamps and recipient information
- **Auto-Folder Creation**: Automatically creates destination folders if they don't exist
- **CSV-Driven Configuration**: Easy-to-maintain configuration without code changes
- **Git-Safe**: Configuration files are ignored, only templates are committed

## Files Structure

```
Macro/
├── README.md                           # This documentation
├── SaveAttachmentsBySubject.bas        # Main VBA macro file
├── subject_config.csv.template         # Configuration template
├── subject_config.csv                  # Your actual configuration (Git-ignored)
├── .gitignore                          # Git ignore rules
└── Logs/                              # Auto-created log directory
    └── AttachmentLog.txt              # Detailed operation logs
```

## Installation

### 1. Set Up Configuration

1. Copy `subject_config.csv.template` to `subject_config.csv`
2. Edit `subject_config.csv` with your specific requirements
3. Remove comment lines (starting with #) and add your configuration rows

### 2. Install VBA Macro

1. Open Microsoft Outlook
2. Press `Alt + F11` to open the VBA Editor
3. In the Project Explorer, find your Outlook project
4. Right-click and select `Import File...`
5. Select the `SaveAttachmentsBySubject.bas` file
6. The macro will be imported and ready to use

### 3. Set Up Outlook Rule

1. In Outlook, go to `File > Manage Rules & Alerts`
2. Click `New Rule...`
3. Choose `Apply rule on messages I receive`
4. Set your conditions (e.g., specific senders, subjects, etc.)
5. In Actions, choose `run a script`
6. Select `SaveAttachmentsBySubject` from the list
7. Complete the rule setup

## Configuration Guide

### CSV Format

```csv
Subject,MatchType,DestinationFolder,BackupFolder,DestinationFileName
```

### Match Types

| Match Type | Description | Example |
|------------|-------------|---------|
| `EXACT` | Subject must match exactly (case-insensitive) | "Daily Report" matches only "Daily Report" |
| `CONTAINS` | Subject must contain the text | "Report" matches "Daily Report Summary" |
| `STARTS_WITH` | Subject must start with the text | "RE:" matches "RE: Important Update" |
| `ENDS_WITH` | Subject must end with the text | "Summary" matches "Monthly Financial Summary" |

### Filename Placeholders

| Placeholder | Description | Example Output |
|-------------|-------------|----------------|
| `{timestamp}` | Current date and time | `20250906_143022` |
| `{date}` | Current date | `20250906` |
| `{time}` | Current time | `143022` |

### Example Configuration

```csv
Subject,MatchType,DestinationFolder,BackupFolder,DestinationFileName
Daily Sales Report,EXACT,C:\Reports\Sales\,C:\Backup\Sales\,daily_sales_{date}.xlsx
Weekly Inventory,CONTAINS,C:\Reports\Inventory\,C:\Backup\Inventory\,inventory_{timestamp}.pdf
RE:,STARTS_WITH,C:\Reports\Replies\,C:\Backup\Replies\,reply_{timestamp}.msg
Summary,ENDS_WITH,C:\Reports\Summaries\,C:\Backup\Summaries\,summary_{date}.xlsx
```

## Usage Examples

### Scenario 1: Exact Match for Daily Reports
```csv
Daily Sales Report,EXACT,C:\Reports\Sales\,C:\Backup\Sales\,daily_sales.xlsx
```
- Matches: "Daily Sales Report"
- Doesn't match: "Daily Sales Report - January", "RE: Daily Sales Report"

### Scenario 2: Handle Reply Emails
```csv
RE:,STARTS_WITH,C:\Reports\Replies\,C:\Backup\Replies\,reply_{timestamp}.msg
```
- Matches: "RE: Important Update", "RE: Daily Report"
- Saves with timestamp: `reply_20250906_143022.msg`

### Scenario 3: Flexible Report Matching
```csv
Inventory,CONTAINS,C:\Reports\Inventory\,C:\Backup\Inventory\,inventory_{date}.pdf
```
- Matches: "Weekly Inventory Update", "Monthly Inventory Report", "Inventory Summary"

## Logging

The macro creates detailed logs in `Logs/AttachmentLog.txt` including:

- Email subject and received time
- Recipient email addresses
- Match type and pattern used
- File save locations (primary and backup)
- Success/failure status for each operation
- Error messages if any issues occur

### Sample Log Entry

```
=== Processing Email ===
Subject: Daily Sales Report
Received: 2025-09-06 14:30:22
Recipient: manager@company.com
Match found - Type: EXACT, Pattern: Daily Sales Report
2025-09-06 14:30:22 - Saved to primary: C:\Reports\Sales\daily_sales_20250906.xlsx
2025-09-06 14:30:22 - Saved to backup: C:\Backup\Sales\daily_sales_20250906.xlsx
```

## Troubleshooting

### Common Issues

1. **Configuration file not found**
   - Ensure `subject_config.csv` exists in the Macro folder
   - Check that the file path is correct

2. **No matches found**
   - Verify your subject patterns in the CSV
   - Check that match types are spelled correctly (EXACT, CONTAINS, etc.)
   - Review the log file for the actual email subject

3. **Permission errors when saving**
   - Ensure destination folders exist or can be created
   - Check Windows permissions for the target directories
   - Verify disk space availability

4. **Macro not running**
   - Confirm the Outlook rule is properly configured
   - Check that macros are enabled in Outlook security settings
   - Verify the macro was imported correctly

### Debug Tips

1. Check the log file for detailed error messages
2. Test with a simple EXACT match first
3. Verify folder paths use proper Windows format (backslashes)
4. Ensure CSV file doesn't have extra spaces or special characters

## Security Considerations

- Configuration files containing sensitive paths are Git-ignored
- Only template files are committed to version control
- Log files are also ignored to prevent sensitive information exposure
- Always review folder permissions before deployment

## Advanced Features

### Multiple Attachments
The macro processes all attachments in matching emails, saving each one to both primary and backup locations.

### Automatic Folder Creation
If destination folders don't exist, the macro will attempt to create them automatically.

### Error Handling
Comprehensive error handling ensures the macro continues processing even if individual operations fail, with detailed error logging.

### Case-Insensitive Matching
All subject matching is performed case-insensitively for better reliability.

## Support

For issues or questions:
1. Check the log file for detailed error information
2. Verify your configuration against the examples
3. Test with simple configurations first
4. Review Outlook's macro security settings

## Recent Updates (v1.1)

### Fixed Issues
- **Resolved ThisWorkbook.Path Error**: Fixed object error when macro runs from Outlook context by implementing robust path handling with environment variable fallback
- **Corrected File Path Construction**: Fixed missing backslashes in file paths that caused malformed save locations
- **Enhanced Filename Logic**: Improved filename replacement to properly use configured names instead of original attachment names
- **Added Comprehensive Logging**: Enhanced debugging with detailed logs showing configuration loading, attachment processing, and path construction

### Improvements
- **Robust Error Handling**: Added fallback mechanisms for path resolution
- **Better Debugging**: Detailed logging shows exactly what's happening during processing
- **Path Validation**: Automatic validation and correction of folder paths
- **Configuration Validation**: Better handling of CSV configuration loading

## Version History

- **v1.1** - Major bug fixes for path handling, filename logic, and enhanced logging
- **v1.0** - Initial release with flexible subject matching and dual-folder saving
