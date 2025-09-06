"""
Simple package checker to verify what's actually installed.
"""

import sys
print(f"Python executable: {sys.executable}")
print(f"Python version: {sys.version}")
print()

packages_to_check = ['PIL', 'reportlab', 'pandas']

print("Package Status:")
for package in packages_to_check:
    try:
        if package == 'PIL':
            import PIL
            print(f"✅ Pillow (PIL): Available - Version {PIL.__version__}")
        elif package == 'reportlab':
            import reportlab
            print(f"✅ reportlab: Available - Version {reportlab.Version}")
        elif package == 'pandas':
            import pandas
            print(f"✅ pandas: Available - Version {pandas.__version__}")
    except ImportError:
        print(f"❌ {package}: Not available")
    except Exception as e:
        print(f"⚠️  {package}: Error - {e}")

print()
print("If packages show as 'Not available' but you installed them:")
print("1. Make sure you're using the same Python environment")
print("2. If using a virtual environment, activate it first:")
print("   .venv\\Scripts\\activate")
print("3. Then install packages in that environment:")
print("   pip install Pillow reportlab pandas")
