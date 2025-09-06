"""
Utility functions for BIMailer automation system.
Handles logging, lock management, date functions, and common utilities.
"""

import os
import logging
import datetime
from pathlib import Path
from typing import Optional, List
import time
import json


class BIMailerLogger:
    """Centralized logging system for BIMailer."""
    
    def __init__(self, logs_dir: str = "Logs"):
        # Handle running from Scripts directory
        if not Path(logs_dir).exists() and Path("..") / logs_dir:
            logs_dir = str(Path("..") / logs_dir)
        
        self.logs_dir = Path(logs_dir)
        self.logs_dir.mkdir(exist_ok=True)
        self.current_date = datetime.datetime.now()
        self._setup_loggers()
    
    def _setup_loggers(self):
        """Set up different loggers for different types of operations."""
        # File processing logger
        self.file_logger = self._create_logger(
            'file_processing',
            f'file_processing_{self.current_date.strftime("%Y-%m")}.log'
        )
        
        # Email logger
        self.email_logger = self._create_logger(
            'email',
            f'email_log_{self.current_date.strftime("%Y-%m")}.log'
        )
        
        # Error logger
        self.error_logger = self._create_logger(
            'error',
            f'error_log_{self.current_date.strftime("%Y-%m")}.log'
        )
        
        # Daily summary logger
        self.summary_logger = self._create_logger(
            'summary',
            f'processing_summary_{self.current_date.strftime("%Y-%m-%d")}.log'
        )
    
    def _create_logger(self, name: str, filename: str) -> logging.Logger:
        """Create a logger with file and console handlers."""
        logger = logging.getLogger(name)
        logger.setLevel(logging.INFO)
        
        # Clear existing handlers
        logger.handlers.clear()
        
        # File handler
        file_handler = logging.FileHandler(self.logs_dir / filename, encoding='utf-8')
        file_handler.setLevel(logging.INFO)
        
        # Console handler
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        
        # Formatter
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)
        
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        
        return logger
    
    def log_file_operation(self, message: str, level: str = 'info'):
        """Log file processing operations."""
        getattr(self.file_logger, level.lower())(message)
    
    def log_email_operation(self, message: str, level: str = 'info'):
        """Log email operations."""
        getattr(self.email_logger, level.lower())(message)
    
    def log_error(self, message: str, exception: Optional[Exception] = None):
        """Log errors with optional exception details."""
        if exception:
            self.error_logger.error(f"{message}: {str(exception)}", exc_info=True)
        else:
            self.error_logger.error(message)
    
    def log_summary(self, message: str):
        """Log processing summary information."""
        self.summary_logger.info(message)


class ProcessingLock:
    """Manages processing lock to prevent concurrent runs."""
    
    def __init__(self, lock_file: str = ".lock", timeout_minutes: int = 30):
        self.lock_file = Path(lock_file)
        self.timeout_minutes = timeout_minutes
        self.logger = BIMailerLogger().file_logger
    
    def acquire_lock(self) -> bool:
        """Acquire processing lock. Returns True if successful."""
        try:
            if self.lock_file.exists():
                # Check if lock is stale
                lock_time = datetime.datetime.fromtimestamp(self.lock_file.stat().st_mtime)
                current_time = datetime.datetime.now()
                
                if (current_time - lock_time).total_seconds() > (self.timeout_minutes * 60):
                    self.logger.warning(f"Removing stale lock file (older than {self.timeout_minutes} minutes)")
                    self.lock_file.unlink()
                else:
                    self.logger.error("Processing already in progress (lock file exists)")
                    return False
            
            # Create lock file
            lock_data = {
                'timestamp': datetime.datetime.now().isoformat(),
                'pid': os.getpid()
            }
            
            with open(self.lock_file, 'w') as f:
                json.dump(lock_data, f, indent=2)
            
            self.logger.info("Processing lock acquired")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to acquire processing lock: {e}")
            return False
    
    def release_lock(self):
        """Release processing lock."""
        try:
            if self.lock_file.exists():
                self.lock_file.unlink()
                self.logger.info("Processing lock released")
        except Exception as e:
            self.logger.error(f"Failed to release processing lock: {e}")
    
    def __enter__(self):
        """Context manager entry."""
        if not self.acquire_lock():
            raise RuntimeError("Could not acquire processing lock")
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        self.release_lock()


def get_current_date(format_str: str = "%Y-%m-%d") -> str:
    """Get current date in specified format."""
    return datetime.datetime.now().strftime(format_str)


def get_current_timestamp(format_str: str = "%Y-%m-%d_%H-%M-%S") -> str:
    """Get current timestamp in specified format."""
    return datetime.datetime.now().strftime(format_str)


def get_file_creation_date(file_path: Path) -> datetime.datetime:
    """Get file creation date."""
    try:
        return datetime.datetime.fromtimestamp(file_path.stat().st_ctime)
    except Exception:
        return datetime.datetime.now()


def format_file_size(size_bytes: int) -> str:
    """Format file size in human readable format."""
    if size_bytes == 0:
        return "0 B"
    
    size_names = ["B", "KB", "MB", "GB"]
    i = 0
    while size_bytes >= 1024 and i < len(size_names) - 1:
        size_bytes /= 1024.0
        i += 1
    
    return f"{size_bytes:.1f} {size_names[i]}"


def validate_email(email: str) -> bool:
    """Basic email validation."""
    import re
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None


def split_email_list(email_string) -> List[str]:
    """Split semicolon-separated email list and validate."""
    # Handle NaN values from pandas
    if email_string is None or (hasattr(email_string, '__class__') and email_string.__class__.__name__ == 'float'):
        return []
    
    # Convert to string and check if empty
    email_string = str(email_string)
    if not email_string or email_string.strip() == '' or email_string.lower() == 'nan':
        return []
    
    emails = [email.strip() for email in email_string.split(';')]
    valid_emails = [email for email in emails if validate_email(email)]
    
    if len(valid_emails) != len(emails):
        logger = BIMailerLogger().error_logger
        invalid_emails = [email for email in emails if not validate_email(email)]
        logger.warning(f"Invalid email addresses found: {invalid_emails}")
    
    return valid_emails


def ensure_directory_exists(directory: Path):
    """Ensure directory exists, create if it doesn't."""
    directory.mkdir(parents=True, exist_ok=True)


def clean_filename(filename: str) -> str:
    """Clean filename by removing invalid characters."""
    import re
    # Remove invalid characters for Windows filenames
    invalid_chars = r'[<>:"/\\|?*]'
    cleaned = re.sub(invalid_chars, '_', filename)
    return cleaned.strip()


def get_png_files(directory: Path, extensions: List[str] = ['.png', '.PNG']) -> List[Path]:
    """Get all PNG files in directory sorted alphabetically."""
    png_files = []
    
    if not directory.exists():
        return png_files
    
    for ext in extensions:
        png_files.extend(directory.glob(f'*{ext}'))
    
    return sorted(png_files)


def replace_date_placeholders(text: str, date_format: str = "%Y-%m-%d") -> str:
    """Replace [DATE] placeholders with current date."""
    current_date = get_current_date(date_format)
    return text.replace('[DATE]', current_date)


if __name__ == "__main__":
    # Test the utility functions
    logger = BIMailerLogger()
    logger.log_file_operation("Testing file operation logging")
    logger.log_email_operation("Testing email operation logging")
    logger.log_summary("Testing summary logging")
    
    print(f"Current date: {get_current_date()}")
    print(f"Current timestamp: {get_current_timestamp()}")
    print(f"Email validation test: {validate_email('test@example.com')}")
    print(f"Split emails test: {split_email_list('a@test.com;b@test.com;invalid-email')}")
