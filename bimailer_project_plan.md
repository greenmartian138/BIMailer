# BIMailer Automation Project - Complete Implementation Plan

## Project Overview
Automate the process of converting Power BI report PNGs to PDFs and emailing them to specified recipients with comprehensive logging and archiving.

## Requirements Summary
1. Monitor folders for PNG files from Power BI reports
2. Create PDFs from PNGs (no compression, maintain original dimensions)
3. Email PDFs to configured recipients with detailed logs
4. Archive processed files with timestamps
5. Comprehensive error handling and admin notifications

## File Structure Implementation

```
C:\path\to\your\BIMailer\
├── Config\
│   ├── pdf_names.csv              # FolderName → PDF Name mapping
│   ├── mailing_list.csv           # PDF Name → Recipients, Subject, CC
│   └── settings.ini               # General settings
├── Input\
│   ├── ALL\                       # PNGs to include in every PDF
│   ├── GroupA\                    # Example folder for GroupA reports
│   ├── GroupB\                    # Example folder for GroupB reports
│   └── [Additional group folders as needed]
├── Output\
│   └── PDFs\                      # Generated PDFs (temporary)
├── Archive\
│   ├── PNGs\
│   │   ├── 2024\
│   │   │   ├── 09\               # Year/Month structure
│   │   │   └── [Monthly folders]
│   │   └── [Yearly folders]
│   └── PDFs\
│       ├── 2024\
│       │   ├── 09\               # Sent PDFs with timestamps
│       │   └── [Monthly folders]
│       └── [Yearly folders]
├── Logs\
│   ├── file_processing_YYYY-MM.log
│   ├── email_log_YYYY-MM.log
│   ├── error_log_YYYY-MM.log
│   └── processing_summary_YYYY-MM-DD.log
├── Scripts\
│   ├── main.py                    # Main orchestration script
│   ├── pdf_generator.py           # PDF creation logic
│   ├── email_sender.py            # Email handling
│   ├── file_manager.py            # File operations and archiving
│   ├── config_manager.py          # Configuration handling
│   └── utils.py                   # Utility functions
└── .lock                          # Processing lock file (created during execution)
```

## Configuration Files Specifications

### Config/pdf_names.csv
```csv
FolderName,PDFName
GroupA,Weekly Sales Report
GroupB,Marketing Dashboard
GroupC,Financial Summary
ALL,Global Headers
```

### Config/mailing_list.csv
```csv
PDFName,Recipients,CC,Subject
Weekly Sales Report,john@company.com;mary@company.com,manager@company.com,Weekly Sales Report - [DATE]
Weekly Sales Report,director@company.com,,Weekly Sales Report for Review - [DATE]
Marketing Dashboard,marketing@company.com,cmo@company.com,Latest Marketing Dashboard - [DATE]
Financial Summary,finance@company.com;cfo@company.com,audit@company.com,Financial Summary Report - [DATE]
```

### Config/settings.ini
```ini
[General]
log_retention_days = 60
max_attachment_size_mb = 20
processing_lock_timeout_minutes = 30

[Email]
use_default_mailer = False
smtp_server = smtp.gmail.com
smtp_port = 587
smtp_username = your_email@gmail.com
smtp_password = your_app_password
use_tls = True

[Admin]
admin_emails = admin1@company.com;admin2@company.com
send_summary_email = True
send_error_notifications = True

[Processing]
png_file_extensions = .png,.PNG
archive_after_processing = True
override_existing_lock = False
date_format = %Y-%m-%d
timestamp_format = %Y-%m-%d_%H-%M-%S
```

## Technical Implementation Details

### Technology Stack
- **Python 3.8+**
- **Libraries Required:**
  - `Pillow` - PNG image handling
  - `reportlab` - PDF creation
  - `smtplib` - SMTP email sending
  - `configparser` - Configuration file parsing
  - `pandas` - CSV file handling
  - `pathlib` - Modern path handling
  - `logging` - Comprehensive logging
  - `datetime` - Date/time operations
  - `os` and `shutil` - File operations

### Core Functionality Requirements

#### PDF Generation
- Combine all PNGs in a folder into single PDF
- Include ALL folder PNGs in every PDF
- Maintain original PNG dimensions and aspect ratios
- Process PNGs in alphabetical order
- No compression applied to images
- Each page matches source PNG dimensions

#### Email System
- Primary: SMTP with authentication
- Fallback: Default system mailer (mailto)
- Support multiple recipients per PDF (separate emails)
- CC support
- Attachment size validation (20MB limit)
- Date placeholder replacement in subjects: [DATE] → today's date

#### File Management
- Process all folders with unprocessed PNGs
- Never send emails without attachments
- Archive PNGs after successful PDF creation
- Archive PDFs after successful email sending
- Rename archived PDFs with completion timestamp

#### Error Handling
- Comprehensive logging of all operations
- Continue processing on individual failures
- Email admin summary with error details
- Processing lock to prevent concurrent runs
- Graceful handling of missing configuration

## Detailed Phase Implementation Plan

### Phase 1: Core Infrastructure (Week 1)
**Deliverables:**
1. Complete folder structure creation
2. Sample configuration files with validation
3. Logging system implementation
4. Processing lock mechanism
5. Configuration validation functions
6. Basic utility functions

**Key Files to Create:**
- `Scripts/utils.py` - Logging, lock management, date functions
- `Scripts/config_manager.py` - Configuration loading and validation
- All configuration template files
- Folder structure creation script

### Phase 2: PDF Generation (Week 2)
**Deliverables:**
1. PNG discovery and sorting functionality
2. PDF creation with original dimensions
3. ALL folder integration
4. File metadata extraction (creation dates)
5. PDF generation logging

**Key Files to Create:**
- `Scripts/pdf_generator.py` - Complete PDF creation logic
- PNG processing functions
- Metadata extraction utilities

### Phase 3: Email System (Week 3)
**Deliverables:**
1. SMTP email sending with authentication
2. Default mailer fallback implementation
3. Email template system with placeholders
4. Attachment size validation
5. Multiple recipient handling
6. Email delivery logging

**Key Files to Create:**
- `Scripts/email_sender.py` - Complete email functionality
- Email template generation
- Recipient list processing

### Phase 4: File Management & Archiving (Week 4)
**Deliverables:**
1. File archiving system with timestamp
2. Archive folder organization (Year/Month)
3. Processed file tracking
4. Archive cleanup based on retention policy
5. File operation error handling

**Key Files to Create:**
- `Scripts/file_manager.py` - Complete file operations
- Archive organization functions
- File tracking system

### Phase 5: Integration & Orchestration (Week 5)
**Deliverables:**
1. Main orchestration script
2. End-to-end workflow implementation
3. Admin notification system
4. Processing summary generation
5. Comprehensive error handling
6. Final testing and documentation

**Key Files to Create:**
- `Scripts/main.py` - Complete orchestration
- Admin notification system
- Processing summary templates

## Email Templates

### Standard Email Body Template
```
Dear Recipients,

Please find attached the {PDF_NAME} report.

Files included in this report:
{FILE_LIST_WITH_DATES}

Report generated on: {TODAY_DATE} at {GENERATION_TIME}

Best regards,
Automated Reporting System
```

### Admin Summary Email Template
```
Subject: BIMailer Processing Summary - {DATE}

Processing completed at: {TIMESTAMP}

Summary:
- Folders processed: {FOLDER_COUNT}
- PDFs created: {PDF_COUNT}
- Emails sent: {EMAIL_COUNT}
- Errors encountered: {ERROR_COUNT}

{DETAILED_RESULTS}

{ERROR_DETAILS}

Full logs available at: {LOG_PATH}
```

## Testing Strategy
1. **Unit Testing:** Individual functions and modules
2. **Integration Testing:** End-to-end workflow testing
3. **Error Testing:** Deliberate failure scenarios
4. **Configuration Testing:** Various config combinations
5. **Email Testing:** Both SMTP and fallback methods

## Deployment Instructions
1. Install Python 3.8+ and required libraries
2. Create folder structure
3. Configure email settings and recipient lists
4. Test with sample PNG files
5. Set up Windows Task Scheduler for automation
6. Configure error notification recipients

## Success Criteria
- ✅ Automatically process all PNG files in input folders
- ✅ Create properly formatted PDFs with correct dimensions
- ✅ Successfully send emails to all configured recipients
- ✅ Maintain comprehensive logs of all operations
- ✅ Archive processed files with proper organization
- ✅ Handle errors gracefully with admin notifications
- ✅ Prevent concurrent processing instances
- ✅ Support easy configuration changes via CSV/INI files

## Implementation Instructions for Claude
When implementing this project:

1. **Start with Phase 1** - Create the complete folder structure and configuration files
2. **Implement each phase sequentially** - Each phase builds on the previous
3. **Create working code** - All scripts should be functional and tested
4. **Include comprehensive error handling** - Every function should handle potential failures
5. **Add detailed logging** - Log all significant operations and errors
6. **Create sample configuration data** - Include realistic examples
7. **Test incrementally** - Verify each component works before integration

The project should be production-ready with proper documentation, error handling, and user-friendly configuration files.
