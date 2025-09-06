"""
Basic functionality test for BIMailer system without external dependencies.
Tests configuration loading and basic utilities.
"""

import sys
import os
from pathlib import Path

# Add Scripts directory to path
scripts_dir = Path("Scripts")
sys.path.insert(0, str(scripts_dir))

def test_configuration():
    """Test configuration loading."""
    try:
        from config_manager import ConfigManager
        
        print("✓ Configuration manager imported successfully")
        
        config = ConfigManager()
        print("✓ Configuration loaded successfully")
        
        # Test validation
        is_valid, errors = config.validate_configuration()
        print(f"✓ Configuration validation: {'Valid' if is_valid else 'Invalid'}")
        
        if errors:
            print("  Validation errors:")
            for error in errors:
                print(f"    - {error}")
        
        # Test configuration summary
        summary = config.get_configuration_summary()
        print("✓ Configuration summary generated")
        
        return True
        
    except Exception as e:
        print(f"✗ Configuration test failed: {e}")
        return False

def test_utilities():
    """Test utility functions."""
    try:
        from utils import BIMailerLogger, get_current_date, validate_email, split_email_list
        
        print("✓ Utilities imported successfully")
        
        # Test logger
        logger = BIMailerLogger()
        logger.log_summary("Test log message")
        print("✓ Logger working")
        
        # Test date functions
        current_date = get_current_date()
        print(f"✓ Current date: {current_date}")
        
        # Test email validation
        valid_email = validate_email("test@example.com")
        invalid_email = validate_email("invalid-email")
        print(f"✓ Email validation: valid={valid_email}, invalid={invalid_email}")
        
        # Test email list splitting
        emails = split_email_list("a@test.com;b@test.com")
        print(f"✓ Email list splitting: {emails}")
        
        return True
        
    except Exception as e:
        print(f"✗ Utilities test failed: {e}")
        return False

def test_file_structure():
    """Test file structure exists."""
    try:
        required_dirs = [
            "Config", "Input", "Output", "Archive", "Logs", "Scripts",
            "Input/ALL", "Input/GroupA", "Input/GroupB",
            "Output/PDFs", "Archive/PNGs", "Archive/PDFs"
        ]
        
        missing_dirs = []
        for dir_path in required_dirs:
            if not Path(dir_path).exists():
                missing_dirs.append(dir_path)
        
        if missing_dirs:
            print(f"✗ Missing directories: {missing_dirs}")
            return False
        else:
            print("✓ All required directories exist")
            return True
            
    except Exception as e:
        print(f"✗ File structure test failed: {e}")
        return False

def test_config_files():
    """Test configuration files exist and are readable."""
    try:
        config_files = [
            "Config/settings.ini",
            "Config/pdf_names.csv", 
            "Config/mailing_list.csv"
        ]
        
        missing_files = []
        for file_path in config_files:
            if not Path(file_path).exists():
                missing_files.append(file_path)
        
        if missing_files:
            print(f"✗ Missing configuration files: {missing_files}")
            return False
        else:
            print("✓ All configuration files exist")
            return True
            
    except Exception as e:
        print(f"✗ Configuration files test failed: {e}")
        return False

def main():
    """Run all basic tests."""
    print("=== BIMailer Basic Functionality Test ===\n")
    
    tests = [
        ("File Structure", test_file_structure),
        ("Configuration Files", test_config_files),
        ("Utilities", test_utilities),
        ("Configuration Loading", test_configuration),
    ]
    
    passed = 0
    total = len(tests)
    
    for test_name, test_func in tests:
        print(f"\n--- Testing {test_name} ---")
        try:
            if test_func():
                passed += 1
                print(f"✓ {test_name} PASSED")
            else:
                print(f"✗ {test_name} FAILED")
        except Exception as e:
            print(f"✗ {test_name} FAILED with exception: {e}")
    
    print(f"\n=== Test Results ===")
    print(f"Passed: {passed}/{total}")
    print(f"Success Rate: {(passed/total)*100:.1f}%")
    
    if passed == total:
        print("\n🎉 All basic tests passed! The core system is working correctly.")
        print("\nNote: PDF generation and email sending require additional packages:")
        print("  pip install Pillow reportlab pandas")
    else:
        print(f"\n⚠️  {total-passed} test(s) failed. Please check the issues above.")
    
    return passed == total

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
