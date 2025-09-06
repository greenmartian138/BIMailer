# Gmail SMTP Setup Instructions

## Issue Identified
Gmail is rejecting the regular password for SMTP authentication. Gmail requires **App Passwords** for third-party applications like our BIMailer system.

## Solution: Create Gmail App Password

### Step 1: Enable 2-Factor Authentication
1. Go to [Google Account Settings](https://myaccount.google.com/)
2. Click on "Security" in the left sidebar
3. Under "Signing in to Google", click "2-Step Verification"
4. Follow the setup process to enable 2FA (required for App Passwords)

### Step 2: Generate App Password
1. After 2FA is enabled, go back to "Security"
2. Under "Signing in to Google", click "App passwords"
3. Select "Mail" as the app
4. Select "Other (Custom name)" as the device
5. Enter "BIMailer" as the custom name
6. Click "Generate"
7. **Copy the 16-character password** (it will look like: `abcd efgh ijkl mnop`)

### Step 3: Update BIMailer Configuration
Replace the current password in `Config/settings.ini` with the App Password:

```ini
[Email]
smtp_username = your-email@gmail.com
smtp_password = [YOUR_16_CHARACTER_APP_PASSWORD]
```

## Current Status
- ✅ Gmail account created: `your-email@gmail.com`
- ✅ SMTP connection successful (both port 587 and 465)
- ❌ Authentication failing due to regular password usage
- 🔄 **Next step**: Generate App Password and update configuration

## Test Command
After updating the password, test with:
```bash
python test_smtp_simple.py
```

## Expected Result
Once the App Password is configured, both port 587 (TLS) and port 465 (SSL) should work successfully.
