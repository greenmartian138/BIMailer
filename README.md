# BIMailer Automation System

A comprehensive automation system for converting Power BI report PNGs to PDFs and emailing them to specified recipients with full logging and archiving capabilities.

## Current Status: ✅ PRODUCTION READY - SMTP AUTOMATED

The BIMailer system is **fully functional and production-ready** with automated SMTP email sending via Gmail. No manual attachment process required.

## Overview

BIMailer automates the complete workflow of:
1. **Monitoring** input folders for PNG files from Power BI reports
2. **Converting** PNGs to PDFs while maintaining original dimensions
3. **Emailing** PDFs to configured recipients with enhanced attachment support
4. **Archiving** processed files with timestamps
5. **Comprehensive error handling** and admin notifications

## Project Structure

```
BIMailer/
├── Config/
│   ├── pdf_names.csv              # FolderName → PDF Name mapping
│   ├── mailing_list.csv           # PDF Name → Recipients, Subject, CC
│   └── settings.ini               # General settings
├── Input/
│   ├── ALL/                       # PNGs to include in every PDF
│   ├── GroupA/                    # Example folder for GroupA reports
│   ├── GroupB/                    # Example folder for GroupB reports
│   └── [Additional group folders as needed]
├── Output/
│   └── PDFs/                      # Generated PDFs (temporary)
├── Archive/
│   ├── PNGs/
│   │   └── [Year/Month structure]
│   └── PDFs/
│       └── [Year/Month structure]
├── Logs/
│   ├── file_processing_YYYY-MM.log
│   ├── email_log_YYYY-MM.log
│   ├── error_log_YYYY-MM.log
│   └── processing_summary_YYYY-MM-DD.log
├── Scripts/
│   ├── main.py                    # Main orchestration script
│   ├── pdf_generator.py           # PDF creation logic
│   ├── email_sender.py            # Email handling
│   ├── file_manager.py            # File operations and archiving
│   ├── config_manager.py          # Configuration handling
│   └── utils.py                   # Utility functions
├── Macro/                         # Outlook VBA Macro for attachment saving
│   ├── README.md                  # Macro documentation
│   ├── SaveAttachmentsBySubject.bas # VBA macro file
│   ├── subject_config.csv.template # Configuration template
│   └── .gitignore                 # Macro-specific ignore rules
├── requirements.txt               # Python dependencies
├── README.md                      # This file
└── bimailer_project_plan.md       # Detailed project specifications
```

## Installation

### Prerequisites
- Python 3.8 or higher
- Windows 10/11 (tested environment)

### Setup Steps

1. **Clone or download** the project to your desired location

2. **Install Python dependencies**:
   
   **Option A: Using the provided batch file (Windows)**
   ```cmd
   install_packages.bat
   ```
   
   **Option B: Manual installation**
   ```bash
   pip install Pillow>=9.0.0 reportlab>=3.6.0 pandas>=1.3.0
   ```
   
   **Option C: Using requirements.txt**
   ```bash
   pip install -r requirements.txt
   ```
   
   **Note for Virtual Environment Users:**
   If you're using a virtual environment, make sure it's activated first:
   ```cmd
   .venv\Scripts\activate
   pip install Pillow reportlab pandas
   ```

3. **Configure email settings**:
   ```bash
   # Copy template and configure with your credentials
   copy Config/settings.ini.template Config/settings.ini
   # Edit Config/settings.ini with your SMTP settings
   ```

4. **Set up recipient lists**:
   ```bash
   # Copy template and configure with your recipients
   copy Config/mailing_list.csv.template Config/mailing_list.csv
   # Edit Config/mailing_list.csv with your email recipients
   ```

5. **Configure folder mappings** in `Config/pdf_names.csv`

6. **Create required directories**:
   ```bash
   mkdir Input\ALL Input\GroupA Input\GroupB Input\GroupC
   mkdir Output\PDFs
   mkdir Archive\PNGs Archive\PDFs
   mkdir Logs
   ```

## Configuration

### Email Settings (`Config/settings.ini`)

```ini
[Email]
use_default_mailer = False          # CURRENT: Using automated SMTP sending
smtp_server = smtp.gmail.com        # Gmail SMTP server
smtp_port = 587                     # Port 587 with TLS encryption
smtp_username = your-email@gmail.com
smtp_password = your-gmail-app-password # Gmail App Password (16 characters)
use_tls = True                      # TLS encryption enabled
use_ssl = False                     # SSL disabled for port 587

[Admin]
admin_emails = admin@yourcompany.com
send_summary_email = True
send_error_notifications = True
```

**Configuration: Automated SMTP via Gmail**
- ✅ **Fully Automated**: No manual attachment process required
- ✅ **Gmail Integration**: Using Gmail SMTP with App Password authentication
- ✅ **Professional Sender**: Emails sent from your configured Gmail account
- ✅ **Secure Authentication**: Gmail App Password for enhanced security
- ✅ **TLS Encryption**: Secure email transmission via port 587

**Gmail Setup Requirements:**
1. **Gmail Account**: Configure your Gmail account in settings.ini
2. **2-Factor Authentication**: Required for App Passwords
3. **App Password**: 16-character password for SMTP authentication
4. **SMTP Settings**: Port 587 with TLS encryption

### PDF Names Mapping (`Config/pdf_names.csv`)

```csv
FolderName,PDFName
GroupA,Weekly Sales Report
GroupB,Marketing Dashboard
ALL,Global Headers
```

### Mailing List (`Config/mailing_list.csv`)

```csv
PDFName,Recipients,CC,Subject
Weekly Sales Report,john@company.com;mary@company.com,manager@company.com,Weekly Sales Report - [DATE]
Marketing Dashboard,marketing@company.com,cmo@company.com,Latest Marketing Dashboard - [DATE]
```

## Usage

### Basic Usage

Run the complete processing workflow:
```bash
python Scripts/main.py
```

**What happens when you run the system:**
1. **Automatic Processing**: System processes PNG files and creates PDFs
2. **Email Client Opens**: Your default email client opens for each recipient
3. **Pre-filled Content**: Emails have recipients, subject, and body already filled
4. **Attachment Instructions**: Clear instructions show which PDF file to attach
5. **File Location**: Console displays exact file path for easy attachment
6. **Manual Attachment**: Simply attach the specified PDF file and send

### Enhanced Default Mailer Workflow

When using `use_default_mailer = True` (recommended):

1. **System runs automatically** and processes all PNG files
2. **Email client opens** for each mailing entry with:
   - ✅ Pre-filled recipients and CC
   - ✅ Professional subject line with date
   - ✅ Complete email body with PDF metadata
   - ✅ Clear attachment instructions
3. **Console shows** formatted instructions:
   ```
   ============================================================
   EMAIL CLIENT OPENED - MANUAL ATTACHMENT REQUIRED
   ============================================================
   Subject: Weekly Sales Report - 2025-09-06
   Recipients: recipient@company.com
   CC: manager@company.com

   ATTACH THIS FILE: Weekly Sales Report_2025-09-06_12-08-54.pdf
   File Location: C:\...\Output\PDFs\Weekly Sales Report_2025-09-06_12-08-54.pdf
   ============================================================
   ```
4. **User simply attaches** the specified PDF file and sends the email
5. **System continues** processing remaining folders and emails

### Advanced Usage

**Test email system:**
```bash
python test_email.py
```

**Run system diagnostics:**
```bash
python test_basic_functionality.py
```

**Check package installation:**
```bash
python check_packages.py
```

### Automation Setup

**Windows Task Scheduler:**
1. Open Task Scheduler
2. Create Basic Task
3. Set trigger (e.g., daily at 8:00 AM)
4. Set action to run: `python C:\path\to\BIMailer\Scripts\main.py`
5. Configure additional settings as needed

## Features

### PDF Generation
- ✅ Combines all PNGs in a folder into a single PDF
- ✅ Includes ALL folder PNGs in every PDF
- ✅ Maintains original PNG dimensions and aspect ratios
- ✅ Processes PNGs in alphabetical order
- ✅ No compression applied to images

### Enhanced Email System
- ✅ **Enhanced Default Mailer**: Intelligent attachment handling with multiple methods
- ✅ **Outlook Automation**: Attempts direct sending via COM automation
- ✅ **Smart Fallback**: Email client with pre-filled content and attachment instructions
- ✅ **SMTP Support**: Full authentication with SSL/TLS support (ports 465, 587)
- ✅ **Multiple Recipients**: Separate emails per recipient with CC support
- ✅ **Attachment Validation**: Size validation (20MB limit) with clear error messages
- ✅ **Date Placeholders**: Automatic date replacement in subjects: `[DATE]` → today's date
- ✅ **File Location Guidance**: Console output and file explorer integration
- ✅ **Professional Templates**: Rich email bodies with PDF metadata

### File Management
- ✅ Processes all folders with unprocessed PNGs
- ✅ Never sends emails without attachments
- ✅ Archives PNGs after successful PDF creation
- ✅ Archives PDFs after successful email sending
- ✅ Renames archived PDFs with completion timestamp

### Error Handling
- ✅ Comprehensive logging of all operations
- ✅ Continues processing on individual failures
- ✅ Email admin summary with error details
- ✅ Processing lock to prevent concurrent runs
- ✅ Graceful handling of missing configuration

## Workflow

1. **Initialization**: Load and validate all configuration files
2. **Discovery**: Scan input folders for PNG files to process
3. **PDF Generation**: Create PDFs from PNGs (including ALL folder content)
4. **Email Sending**: Send PDFs to all configured recipients
5. **Archiving**: Move processed files to archive with timestamps
6. **Cleanup**: Remove old files based on retention policy
7. **Reporting**: Send admin summary and handle any errors

## Logging

The system maintains comprehensive logs:

- **File Processing**: `Logs/file_processing_YYYY-MM.log`
- **Email Operations**: `Logs/email_log_YYYY-MM.log`
- **Errors**: `Logs/error_log_YYYY-MM.log`
- **Daily Summary**: `Logs/processing_summary_YYYY-MM-DD.log`

## Troubleshooting

### Common Issues

**Configuration Errors:**
- Run `python main.py diagnostics` to check system status
- Verify all CSV files have correct headers
- Check email credentials and server settings

**Email Issues:**
- For Gmail: Use app-specific passwords, not regular password
- Check firewall settings for SMTP ports
- Verify recipient email addresses are valid

**File Permission Issues:**
- Ensure Python has read/write access to all directories
- Check Windows folder permissions
- Run as administrator if necessary

**PDF Generation Issues:**
- Verify PNG files are not corrupted
- Check available disk space
- Ensure Pillow and reportlab are properly installed

### Error Codes

- **Exit Code 0**: Success
- **Exit Code 1**: General error
- **Exit Code 130**: Interrupted by user (Ctrl+C)

## Security Considerations

- Store email passwords securely (use app-specific passwords)
- Limit file system permissions to necessary directories
- Regularly review and update recipient lists
- Monitor log files for suspicious activity

## Performance

- **Processing Speed**: ~1-2 seconds per PNG file
- **Memory Usage**: ~50-100MB during processing
- **Disk Space**: Archives grow over time (configure retention policy)
- **Concurrent Processing**: Prevented by lock mechanism

## Outlook Macro Integration

The BIMailer project includes an advanced Outlook VBA macro for automated attachment saving based on email subjects. This complements the main BIMailer system by providing intelligent email processing capabilities.

### Features

- **Subject-Based Matching**: Automatically save attachments based on email subject patterns
- **Flexible Matching Types**: EXACT, CONTAINS, STARTS_WITH, ENDS_WITH matching options
- **Dual-Folder Saving**: Save to both primary destination and backup folders
- **Configurable File Naming**: Support for timestamp placeholders in filenames
- **Comprehensive Logging**: Detailed operation logs with recipient tracking
- **CSV-Driven Configuration**: Easy maintenance without code changes
- **Git-Safe**: Configuration files ignored, only templates committed
- **Robust Error Handling**: Fixed path resolution and filename logic issues
- **Enhanced Debugging**: Detailed logging for troubleshooting attachment processing

### Quick Start

1. **Navigate to the Macro folder**: `cd Macro/`
2. **Copy configuration template**: `copy subject_config.csv.template subject_config.csv`
3. **Edit configuration** with your email subjects and destination folders
4. **Import VBA macro** into Outlook (see `Macro/README.md` for detailed instructions)
5. **Set up Outlook rule** to trigger the macro on incoming emails

### Example Configuration

```csv
Subject,MatchType,DestinationFolder,BackupFolder,DestinationFileName
Daily Sales Report,EXACT,C:\Reports\Sales\,C:\Backup\Sales\,daily_sales_{date}.xlsx
Weekly Inventory,CONTAINS,C:\Reports\Inventory\,C:\Backup\Inventory\,inventory_{timestamp}.pdf
RE:,STARTS_WITH,C:\Reports\Replies\,C:\Backup\Replies\,reply_{timestamp}.msg
```

For complete setup instructions, configuration options, and troubleshooting, see the dedicated documentation in `Macro/README.md`.

## Support

For issues or questions:
1. Check the logs in the `Logs/` directory
2. Run diagnostics: `python main.py diagnostics`
3. Review configuration files for errors
4. Check the project plan: `bimailer_project_plan.md`
5. For Macro issues: See `Macro/README.md` for detailed troubleshooting

## License

This project is provided as-is for internal use. Modify and distribute according to your organization's policies.

## Version History

- **v1.2.0**: Enhanced PDF Quality (CURRENT - PRODUCTION READY)
  - ✅ **Improved Text Clarity**: Fixed PDF text quality issues at all zoom levels
  - ✅ **Smart DPI Detection**: Automatic DPI reading from PNG metadata
  - ✅ **High-Resolution Support**: Optimized scaling for high-DPI images (>150 DPI)
  - ✅ **Quality Preservation**: Enhanced image processing with ImageReader
  - ✅ **Perfect Scaling**: 1:1 pixel mapping ensures PDFs look identical to PNGs
  - ✅ **Maintained Performance**: No impact on processing speed or functionality

- **v1.1.0**: Enhanced Default Mailer
  - ✅ **Enhanced Default Mailer**: Multi-method attachment handling
  - ✅ **Outlook Automation**: COM automation for direct sending
  - ✅ **Smart Fallback**: Email client with pre-filled content and instructions
  - ✅ **SSL/TLS Support**: Fixed connection handling for ports 465/587
  - ✅ **File Location Guidance**: Console output and file explorer integration
  - ✅ **Professional Templates**: Rich email bodies with PDF metadata
  - ✅ **Comprehensive Testing**: Full diagnostic and testing suite

- **v1.0.0**: Initial implementation with full feature set
  - PDF generation from PNGs
  - Email sending with SMTP and mailto support
  - File archiving and management
  - Comprehensive logging and error handling
  - Admin notifications and summaries
