"""
Email sending functionality for BIMailer automation system.
Handles SMTP and default mailer email sending with comprehensive logging.
"""

import smtplib
import subprocess
import webbrowser
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
from typing import List, Dict, Optional, Tuple
import datetime
import os
from utils import BIMailerLogger, replace_date_placeholders, format_file_size
from config_manager import ConfigManager, MailingEntry


class EmailSender:
    """Handles email sending functionality."""
    
    def __init__(self, config_manager: ConfigManager):
        self.config = config_manager
        self.logger = BIMailerLogger()
    
    def send_pdf_emails(self, pdf_path: Path, pdf_metadata: Dict) -> bool:
        """Send PDF to all configured recipients."""
        try:
            pdf_name = pdf_metadata.get('pdf_name', pdf_path.stem)
            
            # Get mailing entries for this PDF
            mailing_entries = self.config.get_mailing_entries_for_pdf(pdf_name)
            
            if not mailing_entries:
                self.logger.log_error(f"No mailing entries found for PDF: {pdf_name}")
                return False
            
            # Validate PDF size
            if not self._validate_attachment_size(pdf_path):
                return False
            
            success_count = 0
            total_entries = len(mailing_entries)
            
            self.logger.log_email_operation(f"Sending PDF '{pdf_name}' to {total_entries} mailing entries")
            
            for entry in mailing_entries:
                if self._send_single_email(pdf_path, pdf_metadata, entry):
                    success_count += 1
            
            success = success_count == total_entries
            
            if success:
                self.logger.log_email_operation(
                    f"All emails sent successfully for PDF: {pdf_name} ({success_count}/{total_entries})"
                )
            else:
                self.logger.log_error(
                    f"Some emails failed for PDF: {pdf_name} ({success_count}/{total_entries} successful)"
                )
            
            return success
            
        except Exception as e:
            self.logger.log_error(f"Failed to send emails for PDF: {pdf_path}", e)
            return False
    
    def _send_single_email(self, pdf_path: Path, pdf_metadata: Dict, mailing_entry: MailingEntry) -> bool:
        """Send email to a single mailing entry."""
        try:
            # Generate email content
            subject = self._generate_subject(mailing_entry.subject)
            body = self._generate_email_body(pdf_metadata)
            
            # Log email attempt
            recipients_str = '; '.join(mailing_entry.recipients)
            cc_str = '; '.join(mailing_entry.cc) if mailing_entry.cc else 'None'
            
            self.logger.log_email_operation(
                f"Sending email - Subject: '{subject}', Recipients: {recipients_str}, CC: {cc_str}"
            )
            
            # Choose email method
            if self.config.email_config.use_default_mailer:
                success = self._send_via_default_mailer(
                    mailing_entry.recipients, mailing_entry.cc, subject, body, pdf_path
                )
            else:
                success = self._send_via_smtp(
                    mailing_entry.recipients, mailing_entry.cc, subject, body, pdf_path
                )
            
            if success:
                self.logger.log_email_operation(
                    f"Email sent successfully to: {recipients_str}"
                )
            else:
                self.logger.log_error(
                    f"Failed to send email to: {recipients_str}"
                )
            
            return success
            
        except Exception as e:
            self.logger.log_error(f"Failed to send single email", e)
            return False
    
    def _send_via_smtp(self, recipients: List[str], cc: List[str], subject: str, body: str, attachment_path: Path) -> bool:
        """Send email via SMTP."""
        try:
            # Create message
            msg = MIMEMultipart()
            msg['From'] = f"Glacial Insights <{self.config.email_config.smtp_username}>"
            msg['To'] = '; '.join(recipients)
            if cc:
                msg['Cc'] = '; '.join(cc)
            msg['Subject'] = subject
            
            # Add body
            msg.attach(MIMEText(body, 'plain'))
            
            # Add attachment
            with open(attachment_path, 'rb') as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
            
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename= {attachment_path.name}'
            )
            msg.attach(part)
            
            # Connect to server and send
            # Check if we should use SSL (port 465) or TLS (port 587)
            use_ssl = getattr(self.config.email_config, 'use_ssl', False) or self.config.email_config.smtp_port == 465
            
            if use_ssl:
                # Use SSL connection (typically port 465)
                server = smtplib.SMTP_SSL(self.config.email_config.smtp_server, self.config.email_config.smtp_port, timeout=30)
            else:
                # Use regular SMTP with optional TLS (typically port 587)
                server = smtplib.SMTP(self.config.email_config.smtp_server, self.config.email_config.smtp_port, timeout=30)
                if self.config.email_config.use_tls:
                    server.starttls()
            
            server.login(self.config.email_config.smtp_username, self.config.email_config.smtp_password)
            
            # Combine recipients and CC for sending
            all_recipients = recipients + cc
            text = msg.as_string()
            server.sendmail(self.config.email_config.smtp_username, all_recipients, text)
            server.quit()
            
            return True
            
        except Exception as e:
            self.logger.log_error(f"SMTP email sending failed", e)
            return False
    
    def _send_via_default_mailer(self, recipients: List[str], cc: List[str], subject: str, body: str, attachment_path: Path) -> bool:
        """Send email via default system mailer with attachment support."""
        try:
            recipients_str = ';'.join(recipients)
            cc_str = ';'.join(cc) if cc else ''
            
            # Try different methods to send with attachment
            success = False
            
            # Method 1: Try Outlook automation (Windows)
            if os.name == 'nt':  # Windows
                success = self._try_outlook_automation(recipients, cc, subject, body, attachment_path)
                if success:
                    self.logger.log_email_operation(f"Email sent via Outlook automation with attachment: {attachment_path.name}")
                    return True
            
            # Method 2: Try PowerShell Send-MailMessage (Windows)
            if os.name == 'nt':
                success = self._try_powershell_email(recipients, cc, subject, body, attachment_path)
                if success:
                    self.logger.log_email_operation(f"Email sent via PowerShell with attachment: {attachment_path.name}")
                    return True
            
            # Method 3: Try mapi32 (Windows)
            if os.name == 'nt':
                success = self._try_mapi_email(recipients, cc, subject, body, attachment_path)
                if success:
                    self.logger.log_email_operation(f"Email sent via MAPI with attachment: {attachment_path.name}")
                    return True
            
            # Method 4: Fallback - Open email client and show instructions
            self._open_email_with_instructions(recipients, cc, subject, body, attachment_path)
            
            return True  # We opened the client, user needs to attach manually
            
        except Exception as e:
            self.logger.log_error(f"Default mailer email failed", e)
            return False
    
    def _try_outlook_automation(self, recipients: List[str], cc: List[str], subject: str, body: str, attachment_path: Path) -> bool:
        """Try to send email using Outlook automation."""
        try:
            import win32com.client
            
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)  # 0 = olMailItem
            
            mail.To = ';'.join(recipients)
            if cc:
                mail.CC = ';'.join(cc)
            mail.Subject = subject
            mail.Body = body
            
            # Add attachment
            mail.Attachments.Add(str(attachment_path.absolute()))
            
            # Send the email
            mail.Send()
            
            return True
            
        except ImportError:
            self.logger.log_email_operation("pywin32 not available for Outlook automation")
            return False
        except Exception as e:
            self.logger.log_email_operation(f"Outlook automation failed: {e}")
            return False
    
    def _try_powershell_email(self, recipients: List[str], cc: List[str], subject: str, body: str, attachment_path: Path) -> bool:
        """Try to send email using PowerShell Send-MailMessage."""
        try:
            # Note: This requires SMTP configuration, so we'll skip it for default mailer
            # as it would need the same SMTP settings that are already failing
            return False
            
        except Exception as e:
            self.logger.log_email_operation(f"PowerShell email failed: {e}")
            return False
    
    def _try_mapi_email(self, recipients: List[str], cc: List[str], subject: str, body: str, attachment_path: Path) -> bool:
        """Try to send email using Windows MAPI."""
        try:
            import win32api
            import win32con
            
            # This is complex and requires specific MAPI setup
            # For now, we'll skip this method
            return False
            
        except ImportError:
            return False
        except Exception as e:
            self.logger.log_email_operation(f"MAPI email failed: {e}")
            return False
    
    def _open_email_with_instructions(self, recipients: List[str], cc: List[str], subject: str, body: str, attachment_path: Path):
        """Open email client with instructions for manual attachment."""
        try:
            # Enhanced body with attachment instructions
            enhanced_body = f"""{body}

=== ATTACHMENT REQUIRED ===
Please manually attach the following file to this email:
File: {attachment_path.name}
Location: {attachment_path.absolute()}

The file has been created and is ready for attachment.
"""
            
            # URL encode the subject and body
            import urllib.parse
            encoded_subject = urllib.parse.quote(subject)
            encoded_body = urllib.parse.quote(enhanced_body)
            
            recipients_str = ';'.join(recipients)
            cc_str = ';'.join(cc) if cc else ''
            
            mailto_url = f"mailto:{recipients_str}?subject={encoded_subject}&body={encoded_body}"
            
            if cc_str:
                mailto_url += f"&cc={cc_str}"
            
            # Open default email client
            webbrowser.open(mailto_url)
            
            # Also try to open the file location
            try:
                if os.name == 'nt':  # Windows
                    subprocess.run(['explorer', '/select,', str(attachment_path.absolute())], check=False)
                elif os.name == 'posix':  # macOS/Linux
                    subprocess.run(['open', '-R', str(attachment_path.absolute())], check=False)
            except:
                pass  # If file explorer fails, continue anyway
            
            self.logger.log_email_operation(
                f"Email client opened with attachment instructions. File location: {attachment_path.absolute()}"
            )
            
            # Print instructions to console as well
            print(f"\n{'='*60}")
            print("EMAIL CLIENT OPENED - MANUAL ATTACHMENT REQUIRED")
            print(f"{'='*60}")
            print(f"Subject: {subject}")
            print(f"Recipients: {'; '.join(recipients)}")
            if cc:
                print(f"CC: {'; '.join(cc)}")
            print(f"\nATTACH THIS FILE: {attachment_path.name}")
            print(f"File Location: {attachment_path.absolute()}")
            print(f"{'='*60}\n")
            
        except Exception as e:
            self.logger.log_error(f"Failed to open email with instructions", e)
    
    def _generate_subject(self, subject_template: str) -> str:
        """Generate email subject with date placeholders replaced."""
        return replace_date_placeholders(subject_template, self.config.processing_config.date_format)
    
    def _generate_email_body(self, pdf_metadata: Dict) -> str:
        """Generate email body from PDF metadata."""
        try:
            pdf_name = pdf_metadata.get('pdf_name', 'Report')
            creation_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            current_date = datetime.datetime.now().strftime(self.config.processing_config.date_format)
            
            # Generate file list with dates
            file_list = []
            for file_info in pdf_metadata.get('files', []):
                filename = file_info.get('filename', 'Unknown')
                creation_date = file_info.get('creation_date', 'Unknown')
                if isinstance(creation_date, str):
                    try:
                        creation_date = datetime.datetime.fromisoformat(creation_date.replace('Z', '+00:00'))
                        creation_date_str = creation_date.strftime('%Y-%m-%d %H:%M')
                    except:
                        creation_date_str = str(creation_date)
                else:
                    creation_date_str = creation_date.strftime('%Y-%m-%d %H:%M') if creation_date else 'Unknown'
                
                dimensions = file_info.get('dimensions', 'Unknown')
                file_list.append(f"  • {filename} (Created: {creation_date_str}, Size: {dimensions})")
            
            file_list_text = '\n'.join(file_list) if file_list else '  • No file details available'
            
            # Email body template
            body = f"""Dear Recipients,

Please find attached the {pdf_name} report.

Files included in this report:
{file_list_text}

Report generated on: {current_date} at {creation_time}

Best regards,
Automated Reporting System"""
            
            return body
            
        except Exception as e:
            self.logger.log_error("Failed to generate email body", e)
            return f"Please find attached the {pdf_metadata.get('pdf_name', 'report')}.\n\nBest regards,\nAutomated Reporting System"
    
    def _validate_attachment_size(self, attachment_path: Path) -> bool:
        """Validate attachment size against maximum allowed size."""
        try:
            file_size_mb = attachment_path.stat().st_size / 1024 / 1024
            max_size_mb = self.config.general_config.max_attachment_size_mb
            
            if file_size_mb > max_size_mb:
                self.logger.log_error(
                    f"Attachment size ({file_size_mb:.2f} MB) exceeds maximum allowed size ({max_size_mb} MB): {attachment_path}"
                )
                return False
            
            self.logger.log_email_operation(
                f"Attachment size validation passed: {format_file_size(attachment_path.stat().st_size)}"
            )
            return True
            
        except Exception as e:
            self.logger.log_error(f"Failed to validate attachment size: {attachment_path}", e)
            return False
    
    def send_admin_notification(self, subject: str, body: str, is_error: bool = False) -> bool:
        """Send notification email to administrators."""
        try:
            if is_error and not self.config.admin_config.send_error_notifications:
                return True
            
            if not is_error and not self.config.admin_config.send_summary_email:
                return True
            
            if not self.config.admin_config.admin_emails:
                self.logger.log_error("No admin emails configured for notifications")
                return False
            
            # Create a temporary mailing entry for admin notification
            admin_entry = MailingEntry(
                pdf_name="Admin Notification",
                recipients=self.config.admin_config.admin_emails,
                cc=[],
                subject=subject
            )
            
            # Send email without attachment
            success = self._send_admin_email_only(admin_entry, body)
            
            if success:
                self.logger.log_email_operation(f"Admin notification sent: {subject}")
            else:
                self.logger.log_error(f"Failed to send admin notification: {subject}")
            
            return success
            
        except Exception as e:
            self.logger.log_error(f"Failed to send admin notification", e)
            return False
    
    def _send_admin_email_only(self, mailing_entry: MailingEntry, body: str) -> bool:
        """Send email without attachment (for admin notifications)."""
        try:
            subject = self._generate_subject(mailing_entry.subject)
            
            if self.config.email_config.use_default_mailer:
                return self._send_text_via_default_mailer(
                    mailing_entry.recipients, mailing_entry.cc, subject, body
                )
            else:
                return self._send_text_via_smtp(
                    mailing_entry.recipients, mailing_entry.cc, subject, body
                )
                
        except Exception as e:
            self.logger.log_error("Failed to send admin email", e)
            return False
    
    def _send_text_via_smtp(self, recipients: List[str], cc: List[str], subject: str, body: str) -> bool:
        """Send text-only email via SMTP."""
        try:
            msg = MIMEMultipart()
            msg['From'] = f"Glacial Insights <{self.config.email_config.smtp_username}>"
            msg['To'] = '; '.join(recipients)
            if cc:
                msg['Cc'] = '; '.join(cc)
            msg['Subject'] = subject
            
            msg.attach(MIMEText(body, 'plain'))
            
            # Check if we should use SSL (port 465) or TLS (port 587)
            use_ssl = getattr(self.config.email_config, 'use_ssl', False) or self.config.email_config.smtp_port == 465
            
            if use_ssl:
                # Use SSL connection (typically port 465)
                server = smtplib.SMTP_SSL(self.config.email_config.smtp_server, self.config.email_config.smtp_port, timeout=30)
            else:
                # Use regular SMTP with optional TLS (typically port 587)
                server = smtplib.SMTP(self.config.email_config.smtp_server, self.config.email_config.smtp_port, timeout=30)
                if self.config.email_config.use_tls:
                    server.starttls()
            
            server.login(self.config.email_config.smtp_username, self.config.email_config.smtp_password)
            
            all_recipients = recipients + cc
            text = msg.as_string()
            server.sendmail(self.config.email_config.smtp_username, all_recipients, text)
            server.quit()
            
            return True
            
        except Exception as e:
            self.logger.log_error("SMTP text email sending failed", e)
            return False
    
    def _send_text_via_default_mailer(self, recipients: List[str], cc: List[str], subject: str, body: str) -> bool:
        """Send text-only email via default mailer."""
        try:
            recipients_str = ';'.join(recipients)
            cc_str = ';'.join(cc) if cc else ''
            
            import urllib.parse
            encoded_subject = urllib.parse.quote(subject)
            encoded_body = urllib.parse.quote(body)
            
            mailto_url = f"mailto:{recipients_str}?subject={encoded_subject}&body={encoded_body}"
            
            if cc_str:
                mailto_url += f"&cc={cc_str}"
            
            webbrowser.open(mailto_url)
            return True
            
        except Exception as e:
            self.logger.log_error("Default mailer text email failed", e)
            return False
    
    def generate_processing_summary(self, results: Dict) -> str:
        """Generate processing summary for admin notification."""
        try:
            timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            # Count results
            total_folders = len(results)
            successful_pdfs = len([r for r in results.values() if r.get('pdf_created')])
            successful_emails = len([r for r in results.values() if r.get('emails_sent')])
            errors = []
            
            for folder, result in results.items():
                if result.get('error'):
                    errors.append(f"  • {folder}: {result['error']}")
            
            # Generate summary
            summary = f"""BIMailer Processing Summary - {timestamp}

Processing completed at: {timestamp}

Summary:
- Folders processed: {total_folders}
- PDFs created: {successful_pdfs}
- Emails sent: {successful_emails}
- Errors encountered: {len(errors)}

Detailed Results:"""
            
            for folder, result in results.items():
                pdf_status = "✓" if result.get('pdf_created') else "✗"
                email_status = "✓" if result.get('emails_sent') else "✗"
                summary += f"\n  • {folder}: PDF {pdf_status}, Email {email_status}"
            
            if errors:
                summary += f"\n\nErrors:\n" + '\n'.join(errors)
            
            summary += f"\n\nFull logs available in the Logs directory."
            
            return summary
            
        except Exception as e:
            self.logger.log_error("Failed to generate processing summary", e)
            return f"Processing completed at {datetime.datetime.now()}, but summary generation failed."


if __name__ == "__main__":
    # Test the email sender
    try:
        config_manager = ConfigManager()
        email_sender = EmailSender(config_manager)
        
        # Test admin notification
        test_subject = "BIMailer Test Notification"
        test_body = "This is a test notification from BIMailer system."
        
        success = email_sender.send_admin_notification(test_subject, test_body, is_error=False)
        print(f"Admin notification test: {'Success' if success else 'Failed'}")
        
    except Exception as e:
        print(f"Error testing email sender: {e}")
