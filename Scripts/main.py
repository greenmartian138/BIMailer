"""
Main orchestration script for BIMailer automation system.
Coordinates PDF generation, email sending, and file archiving with comprehensive error handling.
"""

import sys
import traceback
from pathlib import Path
from typing import Dict, List, Optional
import datetime

# Import all BIMailer modules
from utils import BIMailerLogger, ProcessingLock
from config_manager import ConfigManager
from pdf_generator import PDFGenerator
from email_sender import EmailSender
from file_manager import FileManager


class BIMailerOrchestrator:
    """Main orchestrator for the BIMailer automation system."""
    
    def __init__(self):
        """Initialize the orchestrator with all required components."""
        try:
            # Initialize logger first
            self.logger = BIMailerLogger()
            self.logger.log_summary("BIMailer system starting up")
            
            # Load configuration
            self.config = ConfigManager()
            
            # Validate configuration
            is_valid, errors = self.config.validate_configuration()
            if not is_valid:
                error_msg = f"Configuration validation failed: {'; '.join(errors)}"
                self.logger.log_error(error_msg)
                raise RuntimeError(error_msg)
            
            # Initialize components
            self.pdf_generator = PDFGenerator(self.config)
            self.email_sender = EmailSender(self.config)
            self.file_manager = FileManager(self.config)
            
            self.logger.log_summary("BIMailer system initialized successfully")
            
        except Exception as e:
            if hasattr(self, 'logger'):
                self.logger.log_error("Failed to initialize BIMailer system", e)
            else:
                print(f"Critical error during initialization: {e}")
            raise
    
    def run_full_processing(self) -> Dict[str, Dict]:
        """Run the complete processing workflow for all folders."""
        processing_results = {}
        start_time = datetime.datetime.now()
        
        try:
            self.logger.log_summary("Starting full processing workflow")
            
            # Get folders that have PNG files to process
            folders_to_process = self.pdf_generator.get_folders_with_new_pngs()
            
            if not folders_to_process:
                self.logger.log_summary("No folders with PNG files found for processing")
                return {}
            
            self.logger.log_summary(f"Found {len(folders_to_process)} folders to process: {folders_to_process}")
            
            # Process each folder
            for folder_name in folders_to_process:
                folder_result = self._process_single_folder(folder_name)
                processing_results[folder_name] = folder_result
            
            # Generate and send admin summary
            self._send_admin_summary(processing_results, start_time)
            
            # Cleanup old files
            self._perform_cleanup()
            
            end_time = datetime.datetime.now()
            duration = (end_time - start_time).total_seconds()
            
            self.logger.log_summary(f"Full processing completed in {duration:.2f} seconds")
            
            return processing_results
            
        except Exception as e:
            self.logger.log_error("Critical error during full processing", e)
            self._send_error_notification("Critical Processing Error", str(e))
            raise
    
    def _process_single_folder(self, folder_name: str) -> Dict:
        """Process a single folder through the complete workflow."""
        result = {
            'folder_name': folder_name,
            'pdf_created': False,
            'emails_sent': False,
            'files_archived': False,
            'error': None,
            'pdf_path': None,
            'start_time': datetime.datetime.now().isoformat()
        }
        
        try:
            self.logger.log_summary(f"Processing folder: {folder_name}")
            
            # Step 1: Generate PDF
            pdf_path = self.pdf_generator.generate_pdf_for_folder(folder_name)
            
            if not pdf_path:
                result['error'] = "PDF generation failed"
                return result
            
            result['pdf_created'] = True
            result['pdf_path'] = str(pdf_path)
            
            # Step 2: Validate PDF size
            if not self.pdf_generator.validate_pdf_size(pdf_path):
                result['error'] = "PDF size validation failed"
                return result
            
            # Step 3: Get PDF metadata for email
            pdf_metadata = self.pdf_generator.get_pdf_metadata(pdf_path)
            if not pdf_metadata:
                self.logger.log_error(f"Could not load PDF metadata for: {pdf_path}")
                pdf_metadata = {'pdf_name': folder_name, 'files': []}
            
            # Step 4: Send emails
            emails_success = self.email_sender.send_pdf_emails(pdf_path, pdf_metadata)
            result['emails_sent'] = emails_success
            
            if not emails_success:
                result['error'] = "Email sending failed"
                # Don't return here - still try to archive
            
            # Step 5: Archive files (only if emails were sent successfully)
            if emails_success:
                # Archive the PDF
                pdf_archived = self.file_manager.archive_sent_pdf(pdf_path)
                
                # Archive the PNG files
                pngs_archived = self.file_manager.archive_processed_pngs(folder_name)
                
                result['files_archived'] = pdf_archived and pngs_archived
                
                if not result['files_archived']:
                    result['error'] = "File archiving failed"
            else:
                self.logger.log_file_operation(f"Skipping archiving for {folder_name} due to email failures")
            
            # Log completion
            if result['emails_sent'] and result['files_archived']:
                self.logger.log_summary(f"Successfully completed processing for folder: {folder_name}")
            else:
                self.logger.log_error(f"Partial failure for folder {folder_name}: {result['error']}")
            
            result['end_time'] = datetime.datetime.now().isoformat()
            return result
            
        except Exception as e:
            result['error'] = f"Processing exception: {str(e)}"
            result['end_time'] = datetime.datetime.now().isoformat()
            self.logger.log_error(f"Failed to process folder: {folder_name}", e)
            return result
    
    def _send_admin_summary(self, results: Dict[str, Dict], start_time: datetime.datetime):
        """Send processing summary to administrators."""
        try:
            if not self.config.admin_config.send_summary_email:
                return
            
            summary_body = self.email_sender.generate_processing_summary(results)
            subject = f"BIMailer Processing Summary - {datetime.datetime.now().strftime('%Y-%m-%d')}"
            
            success = self.email_sender.send_admin_notification(subject, summary_body, is_error=False)
            
            if success:
                self.logger.log_email_operation("Admin summary email sent successfully")
            else:
                self.logger.log_error("Failed to send admin summary email")
                
        except Exception as e:
            self.logger.log_error("Failed to send admin summary", e)
    
    def _send_error_notification(self, subject: str, error_details: str):
        """Send error notification to administrators."""
        try:
            if not self.config.admin_config.send_error_notifications:
                return
            
            timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            error_body = f"""BIMailer Error Notification

Error occurred at: {timestamp}

Error Details:
{error_details}

Please check the system logs for more information.

System: BIMailer Automation
"""
            
            success = self.email_sender.send_admin_notification(subject, error_body, is_error=True)
            
            if success:
                self.logger.log_email_operation("Error notification sent to administrators")
            else:
                self.logger.log_error("Failed to send error notification")
                
        except Exception as e:
            self.logger.log_error("Failed to send error notification", e)
    
    def _perform_cleanup(self):
        """Perform system cleanup tasks."""
        try:
            self.logger.log_file_operation("Starting cleanup tasks")
            
            # Clean up old archived files
            cleanup_results = self.file_manager.cleanup_old_archives()
            
            # Clean up old PDFs in output directory
            self.pdf_generator.cleanup_old_pdfs(keep_days=7)
            
            self.logger.log_file_operation("Cleanup tasks completed")
            
        except Exception as e:
            self.logger.log_error("Failed to perform cleanup tasks", e)
    
    def run_diagnostics(self) -> Dict:
        """Run system diagnostics and return status information."""
        try:
            self.logger.log_summary("Running system diagnostics")
            
            diagnostics = {
                'timestamp': datetime.datetime.now().isoformat(),
                'configuration': {
                    'valid': True,
                    'summary': self.config.get_configuration_summary()
                },
                'folders': self.file_manager.get_folder_processing_status(),
                'permissions': self.file_manager.validate_file_permissions(),
                'disk_usage': self.file_manager.get_disk_usage_info(),
                'system_status': 'healthy'
            }
            
            # Check for any critical issues
            critical_issues = []
            
            # Check folder permissions
            for path, perms in diagnostics['permissions'].items():
                if not perms.get('valid', False):
                    critical_issues.append(f"Invalid permissions for: {path}")
            
            # Check for folders with files but no processing capability
            for folder, status in diagnostics['folders'].items():
                if status.get('has_files_to_process') and not status.get('folder_exists'):
                    critical_issues.append(f"Folder {folder} has files but doesn't exist")
            
            if critical_issues:
                diagnostics['system_status'] = 'issues_found'
                diagnostics['critical_issues'] = critical_issues
                self.logger.log_error(f"System diagnostics found issues: {critical_issues}")
            else:
                self.logger.log_summary("System diagnostics completed - no issues found")
            
            return diagnostics
            
        except Exception as e:
            self.logger.log_error("Failed to run system diagnostics", e)
            return {
                'timestamp': datetime.datetime.now().isoformat(),
                'system_status': 'error',
                'error': str(e)
            }
    
    def process_specific_folder(self, folder_name: str) -> Dict:
        """Process a specific folder only."""
        try:
            self.logger.log_summary(f"Processing specific folder: {folder_name}")
            
            # Validate folder exists in configuration
            if folder_name not in self.config.get_all_folder_names():
                error_msg = f"Folder '{folder_name}' not found in configuration"
                self.logger.log_error(error_msg)
                return {'error': error_msg}
            
            # Process the folder
            result = self._process_single_folder(folder_name)
            
            # Send notification if there were errors
            if result.get('error'):
                self._send_error_notification(
                    f"BIMailer Processing Error - {folder_name}",
                    f"Error processing folder {folder_name}: {result['error']}"
                )
            
            return result
            
        except Exception as e:
            self.logger.log_error(f"Failed to process specific folder: {folder_name}", e)
            return {'error': str(e)}


def main():
    """Main entry point for the BIMailer system."""
    try:
        # Check command line arguments
        if len(sys.argv) > 1:
            command = sys.argv[1].lower()
            
            if command == 'diagnostics':
                # Run diagnostics only
                orchestrator = BIMailerOrchestrator()
                diagnostics = orchestrator.run_diagnostics()
                
                print("=== BIMailer System Diagnostics ===")
                print(f"Timestamp: {diagnostics['timestamp']}")
                print(f"System Status: {diagnostics['system_status']}")
                
                if diagnostics.get('critical_issues'):
                    print("\nCritical Issues Found:")
                    for issue in diagnostics['critical_issues']:
                        print(f"  - {issue}")
                
                return 0 if diagnostics['system_status'] == 'healthy' else 1
            
            elif command.startswith('folder:'):
                # Process specific folder
                folder_name = command.split(':', 1)[1]
                orchestrator = BIMailerOrchestrator()
                
                with ProcessingLock(timeout_minutes=orchestrator.config.general_config.processing_lock_timeout_minutes):
                    result = orchestrator.process_specific_folder(folder_name)
                
                if result.get('error'):
                    print(f"Error processing folder {folder_name}: {result['error']}")
                    return 1
                else:
                    print(f"Successfully processed folder: {folder_name}")
                    return 0
            
            else:
                print("Usage: python main.py [diagnostics|folder:FOLDER_NAME]")
                return 1
        
        # Default: Run full processing
        orchestrator = BIMailerOrchestrator()
        
        # Use processing lock to prevent concurrent runs
        with ProcessingLock(timeout_minutes=orchestrator.config.general_config.processing_lock_timeout_minutes):
            results = orchestrator.run_full_processing()
        
        # Print summary
        successful_folders = len([r for r in results.values() if r.get('emails_sent')])
        total_folders = len(results)
        
        print(f"BIMailer processing completed: {successful_folders}/{total_folders} folders processed successfully")
        
        return 0 if successful_folders == total_folders else 1
        
    except KeyboardInterrupt:
        print("\nBIMailer processing interrupted by user")
        return 130
    
    except Exception as e:
        print(f"Critical error in BIMailer: {e}")
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
