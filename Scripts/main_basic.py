"""
Basic version of BIMailer main script that works without external dependencies.
Use this for testing configuration and basic functionality before installing packages.
"""

import sys
import traceback
from pathlib import Path
from typing import Dict, List, Optional
import datetime

# Import BIMailer modules (only those that don't require external packages)
from utils import BIMailerLogger, ProcessingLock
from config_manager import ConfigManager
from file_manager import FileManager


class BIMailerBasicOrchestrator:
    """Basic orchestrator for testing without external dependencies."""
    
    def __init__(self):
        """Initialize the orchestrator with basic components."""
        try:
            # Initialize logger first
            self.logger = BIMailerLogger()
            self.logger.log_summary("BIMailer basic system starting up")
            
            # Load configuration
            self.config = ConfigManager()
            
            # Validate configuration
            is_valid, errors = self.config.validate_configuration()
            if not is_valid:
                self.logger.log_error(f"Configuration validation failed: {'; '.join(errors)}")
                print("⚠️  Configuration issues found (this is normal for initial setup):")
                for error in errors:
                    print(f"   - {error}")
            
            # Initialize file manager
            self.file_manager = FileManager(self.config)
            
            self.logger.log_summary("BIMailer basic system initialized successfully")
            
        except Exception as e:
            if hasattr(self, 'logger'):
                self.logger.log_error("Failed to initialize BIMailer basic system", e)
            else:
                print(f"Critical error during initialization: {e}")
            raise
    
    def run_diagnostics(self) -> Dict:
        """Run system diagnostics without external dependencies."""
        try:
            self.logger.log_summary("Running basic system diagnostics")
            
            diagnostics = {
                'timestamp': datetime.datetime.now().isoformat(),
                'configuration': {
                    'valid': True,
                    'summary': self.config.get_configuration_summary()
                },
                'folders': self.file_manager.get_folder_processing_status(),
                'permissions': self.file_manager.validate_file_permissions(),
                'disk_usage': self.file_manager.get_disk_usage_info(),
                'system_status': 'healthy',
                'external_packages': self._check_external_packages()
            }
            
            # Check for any critical issues
            critical_issues = []
            
            # Check folder permissions
            for path, perms in diagnostics['permissions'].items():
                if not perms.get('valid', False):
                    critical_issues.append(f"Invalid permissions for: {path}")
            
            # Check external packages
            if not diagnostics['external_packages']['all_installed']:
                critical_issues.append("External packages not installed (run install_packages.bat)")
            
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
    
    def _check_external_packages(self) -> Dict:
        """Check if external packages are installed."""
        packages = {
            'Pillow': False,
            'reportlab': False,
            'pandas': False
        }
        
        # Check each package with more detailed error handling
        for package in packages:
            try:
                if package == 'Pillow':
                    import PIL
                    packages[package] = True
                elif package == 'reportlab':
                    import reportlab
                    packages[package] = True
                elif package == 'pandas':
                    import pandas
                    packages[package] = True
            except ImportError as e:
                packages[package] = False
            except Exception as e:
                # Other errors might indicate partial installation
                packages[package] = False
        
        return {
            'packages': packages,
            'all_installed': all(packages.values()),
            'missing': [pkg for pkg, installed in packages.items() if not installed]
        }


def main():
    """Main entry point for the basic BIMailer system."""
    try:
        print("=== BIMailer Basic System ===")
        print("This version runs without external packages for initial testing.\n")
        
        # Check command line arguments
        if len(sys.argv) > 1:
            command = sys.argv[1].lower()
            
            if command == 'diagnostics':
                # Run diagnostics only
                orchestrator = BIMailerBasicOrchestrator()
                diagnostics = orchestrator.run_diagnostics()
                
                print("=== BIMailer System Diagnostics ===")
                print(f"Timestamp: {diagnostics['timestamp']}")
                print(f"System Status: {diagnostics['system_status']}")
                
                # Show package status
                pkg_info = diagnostics.get('external_packages', {})
                print(f"\nExternal Packages:")
                for pkg, installed in pkg_info.get('packages', {}).items():
                    status = "✅ Installed" if installed else "❌ Missing"
                    print(f"  {pkg}: {status}")
                
                if not pkg_info.get('all_installed', False):
                    print(f"\n⚠️  Missing packages: {', '.join(pkg_info.get('missing', []))}")
                    print("   Run 'install_packages.bat' to install them.")
                
                if diagnostics.get('critical_issues'):
                    print("\nCritical Issues Found:")
                    for issue in diagnostics['critical_issues']:
                        print(f"  - {issue}")
                
                # Show configuration summary
                config_summary = diagnostics.get('configuration', {}).get('summary', {})
                print(f"\nConfiguration Summary:")
                for section, details in config_summary.items():
                    print(f"  {section}: {details}")
                
                return 0 if diagnostics['system_status'] == 'healthy' else 1
            
            else:
                print("Available commands:")
                print("  diagnostics  - Run system diagnostics")
                print("\nFor full functionality, install packages first:")
                print("  install_packages.bat")
                return 1
        
        # Default: Show status and instructions
        orchestrator = BIMailerBasicOrchestrator()
        diagnostics = orchestrator.run_diagnostics()
        
        pkg_info = diagnostics.get('external_packages', {})
        
        if pkg_info.get('all_installed', False):
            print("✅ All packages installed! You can now use the full system:")
            print("   python main.py")
        else:
            print("⚠️  External packages not installed.")
            print(f"   Missing: {', '.join(pkg_info.get('missing', []))}")
            print("\n📦 To install packages, run:")
            print("   install_packages.bat")
            print("\n🔧 Or install manually:")
            print("   pip install Pillow reportlab pandas")
            print("\n📋 Then test with:")
            print("   python main.py diagnostics")
        
        return 0
        
    except KeyboardInterrupt:
        print("\nBIMailer basic system interrupted by user")
        return 130
    
    except Exception as e:
        print(f"Critical error in BIMailer basic system: {e}")
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
