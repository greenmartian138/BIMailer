"""
Email configuration test script for BIMailer.
Tests email settings and sends a test email to verify configuration.
"""

import sys
from pathlib import Path

# Add Scripts directory to path
scripts_dir = Path("Scripts")
sys.path.insert(0, str(scripts_dir))

from config_manager import ConfigManager
from email_sender import EmailSender
from utils import BIMailerLogger

def test_email_configuration():
    """Test email configuration and send a test email."""
    try:
        print("=== BIMailer Email Configuration Test ===\n")
        
        # Load configuration
        print("Loading configuration...")
        config = ConfigManager()
        
        # Display email configuration
        email_config = config.email_config
        print(f"Email Configuration:")
        print(f"  Use Default Mailer: {email_config.use_default_mailer}")
        print(f"  SMTP Server: {email_config.smtp_server}")
        print(f"  SMTP Port: {email_config.smtp_port}")
        print(f"  SMTP Username: {email_config.smtp_username}")
        print(f"  SMTP Password: {'*' * len(email_config.smtp_password) if email_config.smtp_password else 'Not set'}")
        print(f"  Use TLS: {email_config.use_tls}")
        print()
        
        # Display admin configuration
        admin_config = config.admin_config
        print(f"Admin Configuration:")
        print(f"  Admin Emails: {admin_config.admin_emails}")
        print(f"  Send Summary Email: {admin_config.send_summary_email}")
        print(f"  Send Error Notifications: {admin_config.send_error_notifications}")
        print()
        
        # Check if we have valid email configuration
        if not email_config.use_default_mailer:
            missing_config = []
            if not email_config.smtp_server:
                missing_config.append("SMTP Server")
            if not email_config.smtp_username:
                missing_config.append("SMTP Username")
            if not email_config.smtp_password:
                missing_config.append("SMTP Password")
            
            if missing_config:
                print(f"❌ Missing SMTP configuration: {', '.join(missing_config)}")
                print("Please update Config/settings.ini with your email settings.")
                return False
        
        # Initialize email sender
        print("Initializing email sender...")
        email_sender = EmailSender(config)
        
        # Test admin notification (this is simpler than PDF email)
        print("Testing admin notification email...")
        
        test_subject = "BIMailer Email Test"
        test_body = f"""This is a test email from the BIMailer system.

Email Configuration Test Results:
- SMTP Server: {email_config.smtp_server}
- SMTP Port: {email_config.smtp_port}
- Use TLS: {email_config.use_tls}
- Use Default Mailer: {email_config.use_default_mailer}

If you receive this email, your email configuration is working correctly!

Test sent at: {config.processing_config.date_format}

Best regards,
BIMailer Test System
"""
        
        print(f"Sending test email to: {admin_config.admin_emails}")
        print("This may take a moment...")
        
        success = email_sender.send_admin_notification(test_subject, test_body, is_error=False)
        
        if success:
            print("✅ Test email sent successfully!")
            print("Check your inbox to confirm receipt.")
        else:
            print("❌ Test email failed to send.")
            print("Check the logs for detailed error information.")
        
        return success
        
    except Exception as e:
        print(f"❌ Error during email test: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_smtp_connection():
    """Test SMTP connection without sending email."""
    try:
        print("\n=== SMTP Connection Test ===")
        
        config = ConfigManager()
        email_config = config.email_config
        
        if email_config.use_default_mailer:
            print("Using default mailer (mailto) - no SMTP connection to test.")
            return True
        
        print(f"Testing connection to {email_config.smtp_server}:{email_config.smtp_port}")
        print(f"SSL: {getattr(email_config, 'use_ssl', False)}, TLS: {email_config.use_tls}")
        
        import smtplib
        import socket
        
        # Set connection timeout
        socket.setdefaulttimeout(30)
        
        server = None
        try:
            # Check if we should use SSL (port 465) or regular SMTP with TLS (port 587)
            use_ssl = getattr(email_config, 'use_ssl', False)
            
            if use_ssl and email_config.smtp_port == 465:
                print("Using SSL connection (SMTP_SSL)...")
                server = smtplib.SMTP_SSL(email_config.smtp_server, email_config.smtp_port, timeout=30)
            else:
                print("Using regular SMTP connection...")
                server = smtplib.SMTP(email_config.smtp_server, email_config.smtp_port, timeout=30)
                
                if email_config.use_tls:
                    print("Starting TLS...")
                    server.starttls()
            
            print("Connection established. Attempting login...")
            server.login(email_config.smtp_username, email_config.smtp_password)
            
            print("✅ SMTP connection and authentication successful!")
            server.quit()
            return True
            
        except socket.timeout:
            print("❌ Connection timed out. Check server address and port.")
            if server:
                server.quit()
            return False
        except smtplib.SMTPAuthenticationError as e:
            print(f"❌ Authentication failed: {e}")
            if server:
                server.quit()
            return False
        except smtplib.SMTPConnectError as e:
            print(f"❌ Connection failed: {e}")
            return False
        except Exception as e:
            print(f"❌ Unexpected error: {e}")
            if server:
                server.quit()
            return False
        
    except Exception as e:
        print(f"❌ SMTP connection test failed: {e}")
        return False

def main():
    """Main test function."""
    print("BIMailer Email Configuration Test\n")
    
    # Test SMTP connection first
    smtp_success = test_smtp_connection()
    
    if not smtp_success:
        print("\n⚠️  SMTP connection failed. Check your email settings in Config/settings.ini")
        print("\nCommon issues:")
        print("- Incorrect server address or port")
        print("- Wrong username or password")
        print("- Firewall blocking connection")
        print("- Server requires app-specific password (Gmail, Outlook)")
        return 1
    
    # Test full email sending
    email_success = test_email_configuration()
    
    if email_success:
        print("\n🎉 Email configuration test completed successfully!")
        print("Your BIMailer system is ready to send emails.")
        return 0
    else:
        print("\n❌ Email test failed. Please check your configuration and try again.")
        return 1

if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
