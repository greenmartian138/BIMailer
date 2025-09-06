"""
File management and archiving functionality for BIMailer automation system.
Handles file operations, archiving, and cleanup with comprehensive logging.
"""

import shutil
from pathlib import Path
from typing import List, Dict, Optional, Tuple
import datetime
import os
from utils import BIMailerLogger, get_current_timestamp, ensure_directory_exists
from config_manager import ConfigManager


class FileManager:
    """Handles file operations and archiving."""
    
    def __init__(self, config_manager: ConfigManager):
        self.config = config_manager
        self.logger = BIMailerLogger()
        
        # Directory paths - handle running from Scripts directory
        base_path = Path("..") if not Path("Input").exists() and Path("../Input").exists() else Path(".")
        
        self.input_dir = base_path / "Input"
        self.output_dir = base_path / "Output/PDFs"
        self.archive_dir = base_path / "Archive"
        self.png_archive_dir = self.archive_dir / "PNGs"
        self.pdf_archive_dir = self.archive_dir / "PDFs"
        
        # Ensure archive directories exist
        ensure_directory_exists(self.png_archive_dir)
        ensure_directory_exists(self.pdf_archive_dir)
    
    def archive_processed_pngs(self, folder_name: str) -> bool:
        """Archive PNG files after successful PDF creation."""
        try:
            if not self.config.processing_config.archive_after_processing:
                self.logger.log_file_operation("Archiving disabled in configuration")
                return True
            
            folder_path = self.input_dir / folder_name
            if not folder_path.exists():
                self.logger.log_error(f"Source folder does not exist: {folder_path}")
                return False
            
            # Get PNG files to archive
            png_extensions = self.config.processing_config.png_file_extensions
            png_files = []
            for ext in png_extensions:
                png_files.extend(folder_path.glob(f'*{ext}'))
            
            if not png_files:
                self.logger.log_file_operation(f"No PNG files to archive in folder: {folder_name}")
                return True
            
            # Create archive directory structure (Year/Month)
            current_date = datetime.datetime.now()
            archive_folder = self.png_archive_dir / str(current_date.year) / f"{current_date.month:02d}"
            ensure_directory_exists(archive_folder)
            
            # Create subfolder for this processing batch
            timestamp = get_current_timestamp(self.config.processing_config.timestamp_format)
            batch_folder = archive_folder / f"{folder_name}_{timestamp}"
            ensure_directory_exists(batch_folder)
            
            # Archive each PNG file
            archived_count = 0
            for png_file in png_files:
                try:
                    destination = batch_folder / png_file.name
                    shutil.move(str(png_file), str(destination))
                    archived_count += 1
                    self.logger.log_file_operation(f"Archived PNG: {png_file.name} -> {destination}")
                    
                except Exception as e:
                    self.logger.log_error(f"Failed to archive PNG file: {png_file}", e)
                    continue
            
            self.logger.log_file_operation(
                f"PNG archiving completed for {folder_name}: {archived_count}/{len(png_files)} files archived"
            )
            
            return archived_count == len(png_files)
            
        except Exception as e:
            self.logger.log_error(f"Failed to archive PNGs for folder: {folder_name}", e)
            return False
    
    def archive_sent_pdf(self, pdf_path: Path) -> bool:
        """Archive PDF file after successful email sending."""
        try:
            if not self.config.processing_config.archive_after_processing:
                return True
            
            if not pdf_path.exists():
                self.logger.log_error(f"PDF file does not exist: {pdf_path}")
                return False
            
            # Create archive directory structure (Year/Month)
            current_date = datetime.datetime.now()
            archive_folder = self.pdf_archive_dir / str(current_date.year) / f"{current_date.month:02d}"
            ensure_directory_exists(archive_folder)
            
            # Create archived filename with completion timestamp
            timestamp = get_current_timestamp(self.config.processing_config.timestamp_format)
            archived_name = f"{pdf_path.stem}_sent_{timestamp}{pdf_path.suffix}"
            destination = archive_folder / archived_name
            
            # Move PDF to archive
            shutil.move(str(pdf_path), str(destination))
            
            # Also move metadata file if it exists
            metadata_path = pdf_path.with_suffix('.json')
            if metadata_path.exists():
                metadata_destination = destination.with_suffix('.json')
                shutil.move(str(metadata_path), str(metadata_destination))
                self.logger.log_file_operation(f"Archived PDF metadata: {metadata_destination}")
            
            self.logger.log_file_operation(f"Archived sent PDF: {pdf_path.name} -> {destination}")
            return True
            
        except Exception as e:
            self.logger.log_error(f"Failed to archive PDF: {pdf_path}", e)
            return False
    
    def cleanup_old_archives(self) -> Dict[str, int]:
        """Clean up old archived files based on retention policy."""
        try:
            retention_days = self.config.general_config.log_retention_days
            cutoff_date = datetime.datetime.now() - datetime.timedelta(days=retention_days)
            
            cleanup_results = {
                'png_files_cleaned': 0,
                'pdf_files_cleaned': 0,
                'directories_cleaned': 0
            }
            
            # Clean PNG archives
            png_cleaned = self._cleanup_directory(self.png_archive_dir, cutoff_date)
            cleanup_results['png_files_cleaned'] = png_cleaned
            
            # Clean PDF archives
            pdf_cleaned = self._cleanup_directory(self.pdf_archive_dir, cutoff_date)
            cleanup_results['pdf_files_cleaned'] = pdf_cleaned
            
            # Clean empty directories
            empty_dirs_cleaned = self._cleanup_empty_directories(self.archive_dir)
            cleanup_results['directories_cleaned'] = empty_dirs_cleaned
            
            total_cleaned = sum(cleanup_results.values())
            if total_cleaned > 0:
                self.logger.log_file_operation(
                    f"Archive cleanup completed: {cleanup_results['png_files_cleaned']} PNG files, "
                    f"{cleanup_results['pdf_files_cleaned']} PDF files, "
                    f"{cleanup_results['directories_cleaned']} empty directories removed"
                )
            
            return cleanup_results
            
        except Exception as e:
            self.logger.log_error("Failed to cleanup old archives", e)
            return {'png_files_cleaned': 0, 'pdf_files_cleaned': 0, 'directories_cleaned': 0}
    
    def _cleanup_directory(self, directory: Path, cutoff_date: datetime.datetime) -> int:
        """Clean up files in a directory older than cutoff date."""
        cleaned_count = 0
        
        if not directory.exists():
            return cleaned_count
        
        try:
            for file_path in directory.rglob('*'):
                if file_path.is_file():
                    file_date = datetime.datetime.fromtimestamp(file_path.stat().st_mtime)
                    if file_date < cutoff_date:
                        file_path.unlink()
                        cleaned_count += 1
                        self.logger.log_file_operation(f"Cleaned up old file: {file_path}")
            
        except Exception as e:
            self.logger.log_error(f"Error during directory cleanup: {directory}", e)
        
        return cleaned_count
    
    def _cleanup_empty_directories(self, root_directory: Path) -> int:
        """Remove empty directories recursively."""
        cleaned_count = 0
        
        if not root_directory.exists():
            return cleaned_count
        
        try:
            # Walk directories from deepest to shallowest
            for directory in sorted(root_directory.rglob('*'), key=lambda p: len(p.parts), reverse=True):
                if directory.is_dir() and directory != root_directory:
                    try:
                        # Try to remove if empty
                        directory.rmdir()
                        cleaned_count += 1
                        self.logger.log_file_operation(f"Removed empty directory: {directory}")
                    except OSError:
                        # Directory not empty, skip
                        pass
            
        except Exception as e:
            self.logger.log_error(f"Error during empty directory cleanup: {root_directory}", e)
        
        return cleaned_count
    
    def get_folder_processing_status(self) -> Dict[str, Dict]:
        """Get processing status for all configured folders."""
        status = {}
        
        try:
            folder_names = self.config.get_all_folder_names()
            folder_names = [name for name in folder_names if name != 'ALL']
            
            for folder_name in folder_names:
                folder_path = self.input_dir / folder_name
                
                # Count PNG files
                png_count = 0
                if folder_path.exists():
                    png_extensions = self.config.processing_config.png_file_extensions
                    for ext in png_extensions:
                        png_count += len(list(folder_path.glob(f'*{ext}')))
                
                # Check for recent PDFs in output
                recent_pdfs = self._get_recent_pdfs_for_folder(folder_name)
                
                status[folder_name] = {
                    'folder_exists': folder_path.exists(),
                    'png_files_count': png_count,
                    'has_files_to_process': png_count > 0,
                    'recent_pdfs': len(recent_pdfs),
                    'last_processed': self._get_last_processing_time(folder_name)
                }
            
        except Exception as e:
            self.logger.log_error("Failed to get folder processing status", e)
        
        return status
    
    def _get_recent_pdfs_for_folder(self, folder_name: str, hours: int = 24) -> List[Path]:
        """Get recent PDFs generated for a specific folder."""
        recent_pdfs = []
        
        try:
            pdf_name = self.config.get_pdf_name_for_folder(folder_name)
            if not pdf_name:
                return recent_pdfs
            
            cutoff_time = datetime.datetime.now() - datetime.timedelta(hours=hours)
            
            # Check output directory
            for pdf_file in self.output_dir.glob('*.pdf'):
                if pdf_name.lower() in pdf_file.name.lower():
                    file_time = datetime.datetime.fromtimestamp(pdf_file.stat().st_mtime)
                    if file_time > cutoff_time:
                        recent_pdfs.append(pdf_file)
            
        except Exception as e:
            self.logger.log_error(f"Failed to get recent PDFs for folder: {folder_name}", e)
        
        return recent_pdfs
    
    def _get_last_processing_time(self, folder_name: str) -> Optional[str]:
        """Get the last processing time for a folder from archives."""
        try:
            # Check PNG archives for most recent processing
            latest_time = None
            
            for year_dir in self.png_archive_dir.iterdir():
                if not year_dir.is_dir():
                    continue
                
                for month_dir in year_dir.iterdir():
                    if not month_dir.is_dir():
                        continue
                    
                    for batch_dir in month_dir.iterdir():
                        if batch_dir.is_dir() and batch_dir.name.startswith(folder_name):
                            dir_time = datetime.datetime.fromtimestamp(batch_dir.stat().st_mtime)
                            if latest_time is None or dir_time > latest_time:
                                latest_time = dir_time
            
            return latest_time.strftime('%Y-%m-%d %H:%M:%S') if latest_time else None
            
        except Exception as e:
            self.logger.log_error(f"Failed to get last processing time for folder: {folder_name}", e)
            return None
    
    def create_processing_backup(self, folder_name: str) -> Optional[Path]:
        """Create a backup of folder contents before processing."""
        try:
            folder_path = self.input_dir / folder_name
            if not folder_path.exists():
                return None
            
            # Create backup directory
            timestamp = get_current_timestamp(self.config.processing_config.timestamp_format)
            backup_dir = Path("Backup") / f"{folder_name}_{timestamp}"
            ensure_directory_exists(backup_dir)
            
            # Copy all files
            for file_path in folder_path.iterdir():
                if file_path.is_file():
                    shutil.copy2(str(file_path), str(backup_dir / file_path.name))
            
            self.logger.log_file_operation(f"Created processing backup: {backup_dir}")
            return backup_dir
            
        except Exception as e:
            self.logger.log_error(f"Failed to create backup for folder: {folder_name}", e)
            return None
    
    def validate_file_permissions(self) -> Dict[str, bool]:
        """Validate file permissions for all required directories."""
        permissions = {}
        
        directories_to_check = [
            self.input_dir,
            self.output_dir,
            self.archive_dir,
            self.png_archive_dir,
            self.pdf_archive_dir
        ]
        
        for directory in directories_to_check:
            try:
                # Test read permission
                can_read = os.access(directory, os.R_OK)
                
                # Test write permission
                can_write = os.access(directory, os.W_OK)
                
                permissions[str(directory)] = {
                    'exists': directory.exists(),
                    'readable': can_read,
                    'writable': can_write,
                    'valid': directory.exists() and can_read and can_write
                }
                
            except Exception as e:
                self.logger.log_error(f"Failed to check permissions for: {directory}", e)
                permissions[str(directory)] = {
                    'exists': False,
                    'readable': False,
                    'writable': False,
                    'valid': False
                }
        
        return permissions
    
    def get_disk_usage_info(self) -> Dict[str, Dict]:
        """Get disk usage information for key directories."""
        usage_info = {}
        
        directories = [
            ('Input', self.input_dir),
            ('Output', self.output_dir),
            ('Archive', self.archive_dir)
        ]
        
        for name, directory in directories:
            try:
                if directory.exists():
                    total_size = sum(f.stat().st_size for f in directory.rglob('*') if f.is_file())
                    file_count = len([f for f in directory.rglob('*') if f.is_file()])
                    
                    usage_info[name] = {
                        'total_size_mb': total_size / 1024 / 1024,
                        'file_count': file_count,
                        'exists': True
                    }
                else:
                    usage_info[name] = {
                        'total_size_mb': 0,
                        'file_count': 0,
                        'exists': False
                    }
                    
            except Exception as e:
                self.logger.log_error(f"Failed to get disk usage for: {name}", e)
                usage_info[name] = {
                    'total_size_mb': 0,
                    'file_count': 0,
                    'exists': False,
                    'error': str(e)
                }
        
        return usage_info


if __name__ == "__main__":
    # Test the file manager
    try:
        config_manager = ConfigManager()
        file_manager = FileManager(config_manager)
        
        # Test folder status
        status = file_manager.get_folder_processing_status()
        print("Folder Processing Status:")
        for folder, info in status.items():
            print(f"  {folder}: {info}")
        
        # Test permissions
        permissions = file_manager.validate_file_permissions()
        print("\nFile Permissions:")
        for directory, perms in permissions.items():
            print(f"  {directory}: {perms}")
        
        # Test disk usage
        usage = file_manager.get_disk_usage_info()
        print("\nDisk Usage:")
        for directory, info in usage.items():
            print(f"  {directory}: {info}")
            
    except Exception as e:
        print(f"Error testing file manager: {e}")
