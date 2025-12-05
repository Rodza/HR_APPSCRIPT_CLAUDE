# Google Apps Script Reauthorization Guide

## Issue
You're encountering authorization errors because the deployed web app needs to be reauthorized with the OAuth scopes that were recently added:
- `https://www.googleapis.com/auth/drive` (for DriveApp.createFile)
- `https://www.googleapis.com/auth/script.send_mail` (for MailApp.sendEmail)

## Current Status
✅ OAuth scopes are correctly configured in `apps-script/appsscript.json`
❌ Deployed web app needs to be updated and reauthorized

## Solution: Redeploy and Reauthorize

### Option 1: Using Clasp (Recommended for developers)

If you have clasp installed on your local machine:

```bash
# 1. Push the updated manifest to Apps Script
clasp push

# 2. Create a new deployment (or update existing)
clasp deploy --description "Fixed OAuth permissions for Drive and Mail"

# 3. Open the Apps Script project in browser
clasp open
```

Then follow the steps in Option 2 starting from step 2.

### Option 2: Using Apps Script Web Interface

1. **Open your Apps Script project:**
   - Go to https://script.google.com/home/projects/1AnBSu3JL1YkqNfhWaNuh34hKjkSbPoBAh886NMGFZP9c6FE7kkYI2f3d/edit
   - OR open your Google Sheet → Extensions → Apps Script

2. **Verify the manifest file:**
   - Click on "Project Settings" (gear icon) in the left sidebar
   - Check "Show 'appsscript.json' manifest file in editor" if not already checked
   - Click on `appsscript.json` in the files list
   - Verify it contains both scopes:
     ```json
     "oauthScopes": [
       "https://www.googleapis.com/auth/spreadsheets",
       "https://www.googleapis.com/auth/script.external_request",
       "https://www.googleapis.com/auth/documents",
       "https://www.googleapis.com/auth/drive",
       "https://www.googleapis.com/auth/drive.file",
       "https://www.googleapis.com/auth/drive.readonly",
       "https://www.googleapis.com/auth/userinfo.email",
       "https://www.googleapis.com/auth/script.send_mail"
     ]
     ```

3. **Create a new deployment:**
   - Click "Deploy" → "New deployment" (top right)
   - Click the gear icon next to "Select type"
   - Select "Web app"
   - Configure settings:
     - **Description:** "Fixed OAuth permissions" (or similar)
     - **Execute as:** Me
     - **Who has access:** Anyone with Google account (or your preference)
   - Click "Deploy"

4. **Authorize the application:**
   - A popup will appear asking you to authorize
   - Click "Authorize access"
   - Select your Google account
   - **Important:** You may see a warning "Google hasn't verified this app"
     - Click "Advanced"
     - Click "Go to [Your App Name] (unsafe)"
     - This is normal for personal Apps Script projects
   - Review the permissions requested:
     - ✅ View and manage your spreadsheets
     - ✅ View and manage your Google Drive files
     - ✅ Send email on your behalf
     - ✅ Connect to external services
   - Click "Allow"

5. **Copy the new Web App URL:**
   - After deployment, copy the new Web App URL
   - Update this URL in your documentation/bookmarks if needed

6. **Test the application:**
   - Open the Web App URL in a browser
   - Try importing an Excel file (tests DriveApp.createFile)
   - Try the forgot password feature (tests MailApp.sendEmail)

### Option 3: Quick Reauthorization (If deployment exists)

If you already have a deployment and just need to reauthorize:

1. Open Apps Script project
2. Click on any function in `Code.gs` (e.g., `doGet`)
3. Click "Run" at the top
4. Click "Review permissions" when prompted
5. Follow the authorization flow (steps 4 above)
6. After authorization, test your web app

## Alternative: Test Deployment Locally

Before creating a production deployment, you can test:

1. In Apps Script editor, click "Deploy" → "Test deployments"
2. This creates a temporary deployment for testing
3. Click the Web App URL to test
4. If it works, create a new production deployment

## Verification Checklist

After reauthorization, verify:
- [ ] Import Excel functionality works (no DriveApp.createFile error)
- [ ] Password reset emails send successfully (no MailApp.sendEmail error)
- [ ] All existing features still work
- [ ] No authorization errors in logs

## Troubleshooting

### Error: "You do not have permission to call..."
- Solution: Follow the reauthorization steps above
- The deployment needs to be updated with the new scopes

### Error: "Google hasn't verified this app"
- This is expected for personal Apps Script projects
- Click "Advanced" → "Go to [App Name] (unsafe)"
- This is safe because it's your own code

### Error: "Access blocked: This app's request is invalid"
- Check that appsscript.json has all required scopes
- Ensure no typos in scope URLs
- Try creating a completely new deployment

### Authorization popup doesn't appear
- Disable popup blockers for script.google.com
- Try in an incognito window
- Clear browser cache and cookies

## Prevention

To avoid this issue in the future:
1. Always update appsscript.json when adding new Google services
2. Test in a development deployment before production
3. Document required OAuth scopes in your README
4. Create new deployments when scopes change (don't just update existing)

## Questions?

If you continue to have issues:
1. Check the Apps Script execution logs (View → Logs)
2. Verify all sheets exist and are accessible
3. Ensure the deploying user has edit access to the spreadsheet
4. Try with a different Google account to isolate account-specific issues
