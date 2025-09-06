"""
Configuration management for BIMailer automation system.
Handles loading and validation of configuration files.
"""

import configparser
import pandas as pd
from pathlib import Path
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass
from utils import BIMailerLogger, validate_email, split_email_list


@dataclass
class EmailConfig:
    """Email configuration settings."""
    use_default_mailer: bool
    smtp_server: str
    smtp_port: int
    smtp_username: str
    smtp_password: str
    use_tls: bool
    use_ssl: bool


@dataclass
class GeneralConfig:
    """General configuration settings."""
    log_retention_days: int
    max_attachment_size_mb: int
    processing_lock_timeout_minutes: int


@dataclass
class AdminConfig:
    """Admin configuration settings."""
    admin_emails: List[str]
    send_summary_email: bool
    send_error_notifications: bool


@dataclass
class ProcessingConfig:
    """Processing configuration settings."""
    png_file_extensions: List[str]
    archive_after_processing: bool
    override_existing_lock: bool
    date_format: str
    timestamp_format: str


@dataclass
class MailingEntry:
    """Single mailing list entry."""
    pdf_name: str
    recipients: List[str]
    cc: List[str]
    subject: str


class ConfigManager:
    """Manages all configuration files and settings."""
    
    def __init__(self, config_dir: str = "Config"):
        # Handle relative paths when running from Scripts directory
        if not Path(config_dir).exists() and Path("..") / config_dir:
            config_dir = str(Path("..") / config_dir)
        
        self.config_dir = Path(config_dir)
        self.logger = BIMailerLogger()
        
        # Configuration file paths
        self.settings_file = self.config_dir / "settings.ini"
        self.pdf_names_file = self.config_dir / "pdf_names.csv"
        self.mailing_list_file = self.config_dir / "mailing_list.csv"
        
        # Configuration objects
        self.email_config: Optional[EmailConfig] = None
        self.general_config: Optional[GeneralConfig] = None
        self.admin_config: Optional[AdminConfig] = None
        self.processing_config: Optional[ProcessingConfig] = None
        
        # Configuration data
        self.pdf_names_mapping: Dict[str, str] = {}
        self.mailing_list: List[MailingEntry] = []
        
        self._load_all_configs()
    
    def _load_all_configs(self):
        """Load all configuration files."""
        try:
            self._load_settings()
            self._load_pdf_names()
            self._load_mailing_list()
            self.logger.log_file_operation("All configuration files loaded successfully")
        except Exception as e:
            self.logger.log_error("Failed to load configuration files", e)
            raise
    
    def _load_settings(self):
        """Load settings from INI file."""
        if not self.settings_file.exists():
            raise FileNotFoundError(f"Settings file not found: {self.settings_file}")
        
        config = configparser.ConfigParser()
        config.read(self.settings_file)
        
        # Load email configuration
        email_section = config['Email']
        self.email_config = EmailConfig(
            use_default_mailer=email_section.getboolean('use_default_mailer'),
            smtp_server=email_section.get('smtp_server'),
            smtp_port=email_section.getint('smtp_port'),
            smtp_username=email_section.get('smtp_username'),
            smtp_password=email_section.get('smtp_password'),
            use_tls=email_section.getboolean('use_tls'),
            use_ssl=email_section.getboolean('use_ssl', fallback=False)
        )
        
        # Load general configuration
        general_section = config['General']
        self.general_config = GeneralConfig(
            log_retention_days=general_section.getint('log_retention_days'),
            max_attachment_size_mb=general_section.getint('max_attachment_size_mb'),
            processing_lock_timeout_minutes=general_section.getint('processing_lock_timeout_minutes')
        )
        
        # Load admin configuration
        admin_section = config['Admin']
        admin_emails_str = admin_section.get('admin_emails', '')
        self.admin_config = AdminConfig(
            admin_emails=split_email_list(admin_emails_str),
            send_summary_email=admin_section.getboolean('send_summary_email'),
            send_error_notifications=admin_section.getboolean('send_error_notifications')
        )
        
        # Load processing configuration
        processing_section = config['Processing']
        extensions_str = processing_section.get('png_file_extensions', '.png,.PNG')
        extensions = [ext.strip() for ext in extensions_str.split(',')]
        
        self.processing_config = ProcessingConfig(
            png_file_extensions=extensions,
            archive_after_processing=processing_section.getboolean('archive_after_processing'),
            override_existing_lock=processing_section.getboolean('override_existing_lock'),
            date_format=processing_section.get('date_format'),
            timestamp_format=processing_section.get('timestamp_format')
        )
        
        self.logger.log_file_operation("Settings configuration loaded")
    
    def _load_pdf_names(self):
        """Load PDF names mapping from CSV file."""
        if not self.pdf_names_file.exists():
            raise FileNotFoundError(f"PDF names file not found: {self.pdf_names_file}")
        
        try:
            df = pd.read_csv(self.pdf_names_file)
            
            # Validate required columns
            required_columns = ['FolderName', 'PDFName']
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                raise ValueError(f"Missing required columns in PDF names file: {missing_columns}")
            
            # Create mapping
            self.pdf_names_mapping = dict(zip(df['FolderName'], df['PDFName']))
            
            self.logger.log_file_operation(f"PDF names mapping loaded: {len(self.pdf_names_mapping)} entries")
            
        except Exception as e:
            self.logger.log_error(f"Failed to load PDF names file: {self.pdf_names_file}", e)
            raise
    
    def _load_mailing_list(self):
        """Load mailing list from CSV file."""
        if not self.mailing_list_file.exists():
            raise FileNotFoundError(f"Mailing list file not found: {self.mailing_list_file}")
        
        try:
            df = pd.read_csv(self.mailing_list_file)
            
            # Validate required columns
            required_columns = ['PDFName', 'Recipients', 'Subject']
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                raise ValueError(f"Missing required columns in mailing list file: {missing_columns}")
            
            # Process each row
            self.mailing_list = []
            for _, row in df.iterrows():
                recipients = split_email_list(row['Recipients'])
                cc = split_email_list(row.get('CC', ''))
                
                if not recipients:
                    self.logger.log_error(f"No valid recipients for PDF: {row['PDFName']}")
                    continue
                
                entry = MailingEntry(
                    pdf_name=row['PDFName'],
                    recipients=recipients,
                    cc=cc,
                    subject=row['Subject']
                )
                self.mailing_list.append(entry)
            
            self.logger.log_file_operation(f"Mailing list loaded: {len(self.mailing_list)} entries")
            
        except Exception as e:
            self.logger.log_error(f"Failed to load mailing list file: {self.mailing_list_file}", e)
            raise
    
    def get_pdf_name_for_folder(self, folder_name: str) -> Optional[str]:
        """Get PDF name for a given folder name."""
        return self.pdf_names_mapping.get(folder_name)
    
    def get_mailing_entries_for_pdf(self, pdf_name: str) -> List[MailingEntry]:
        """Get all mailing entries for a given PDF name."""
        return [entry for entry in self.mailing_list if entry.pdf_name == pdf_name]
    
    def get_all_folder_names(self) -> List[str]:
        """Get all configured folder names."""
        return list(self.pdf_names_mapping.keys())
    
    def get_all_pdf_names(self) -> List[str]:
        """Get all configured PDF names."""
        return list(set(self.pdf_names_mapping.values()))
    
    def validate_configuration(self) -> Tuple[bool, List[str]]:
        """Validate all configuration settings."""
        errors = []
        
        # Validate email configuration
        if not self.email_config.use_default_mailer:
            if not self.email_config.smtp_server:
                errors.append("SMTP server not configured")
            if not self.email_config.smtp_username:
                errors.append("SMTP username not configured")
            if not self.email_config.smtp_password:
                errors.append("SMTP password not configured")
        
        # Validate admin emails
        if self.admin_config.send_summary_email or self.admin_config.send_error_notifications:
            if not self.admin_config.admin_emails:
                errors.append("Admin emails not configured but notifications are enabled")
        
        # Validate PDF names mapping
        if not self.pdf_names_mapping:
            errors.append("No PDF names mapping configured")
        
        # Validate mailing list
        if not self.mailing_list:
            errors.append("No mailing list entries configured")
        
        # Check for orphaned PDF names (in mailing list but not in PDF names mapping)
        configured_pdf_names = set(self.pdf_names_mapping.values())
        mailing_pdf_names = set(entry.pdf_name for entry in self.mailing_list)
        orphaned_pdfs = mailing_pdf_names - configured_pdf_names
        
        if orphaned_pdfs:
            errors.append(f"PDF names in mailing list but not in PDF names mapping: {orphaned_pdfs}")
        
        # Check for missing mailing entries (in PDF names mapping but not in mailing list)
        # Exclude "Global Headers" as it's meant to be included in all PDFs, not sent separately
        missing_mailing_entries = configured_pdf_names - mailing_pdf_names
        missing_mailing_entries.discard("Global Headers")  # Remove Global Headers from validation
        if missing_mailing_entries:
            errors.append(f"PDF names configured but no mailing entries found: {missing_mailing_entries}")
        
        is_valid = len(errors) == 0
        
        if is_valid:
            self.logger.log_file_operation("Configuration validation passed")
        else:
            for error in errors:
                self.logger.log_error(f"Configuration validation error: {error}")
        
        return is_valid, errors
    
    def reload_configuration(self):
        """Reload all configuration files."""
        self.logger.log_file_operation("Reloading configuration files")
        self._load_all_configs()
    
    def get_configuration_summary(self) -> Dict:
        """Get a summary of current configuration."""
        return {
            'email_config': {
                'use_default_mailer': self.email_config.use_default_mailer,
                'smtp_server': self.email_config.smtp_server,
                'smtp_port': self.email_config.smtp_port,
                'use_tls': self.email_config.use_tls
            },
            'general_config': {
                'log_retention_days': self.general_config.log_retention_days,
                'max_attachment_size_mb': self.general_config.max_attachment_size_mb,
                'processing_lock_timeout_minutes': self.general_config.processing_lock_timeout_minutes
            },
            'admin_config': {
                'admin_emails_count': len(self.admin_config.admin_emails),
                'send_summary_email': self.admin_config.send_summary_email,
                'send_error_notifications': self.admin_config.send_error_notifications
            },
            'processing_config': {
                'png_file_extensions': self.processing_config.png_file_extensions,
                'archive_after_processing': self.processing_config.archive_after_processing,
                'date_format': self.processing_config.date_format
            },
            'mappings': {
                'folder_count': len(self.pdf_names_mapping),
                'mailing_entries_count': len(self.mailing_list),
                'pdf_names_count': len(self.get_all_pdf_names())
            }
        }


if __name__ == "__main__":
    # Test the configuration manager
    try:
        config_manager = ConfigManager()
        is_valid, errors = config_manager.validate_configuration()
        
        print(f"Configuration valid: {is_valid}")
        if errors:
            print("Errors found:")
            for error in errors:
                print(f"  - {error}")
        
        print("\nConfiguration Summary:")
        summary = config_manager.get_configuration_summary()
        for section, details in summary.items():
            print(f"{section}: {details}")
            
    except Exception as e:
        print(f"Error testing configuration manager: {e}")
