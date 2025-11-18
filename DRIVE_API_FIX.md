# CRITICAL: Drive API Authorization Required for Excel Import

## The Problem

Your diagnostic tests show everything passes EXCEPT when you try to import Excel files. This is because:

**Excel import requires the Drive Advanced Service API**, which needs special authorization that basic functions don't need.

### What Happens During Import:
1. User selects Excel file
2. JavaScript reads file as base64
3. `importClockDataFromBase64()` decodes it
4. `parseClockDataExcel()` tries to use **Drive.Files.copy()** ← **THIS FAILS**
5. Drive API not authorized → Returns null → "Authorization required" error

### Why Other Tests Pass:
- ✅ `getTimeConfig()` - Only uses PropertiesService
- ✅ `updateTimeConfig()` - Only uses PropertiesService
- ✅ `listPendingTimesheets()` - Only reads spreadsheet
- ❌ `importClockDataFromBase64()` - **Needs Drive API v3**

## The Solution

You need to enable and authorize the Drive Advanced Service in Apps Script.

### Step 1: Enable Drive API in Apps Script Editor

1. **Open Apps Script Editor:**
   - Extensions → Apps Script

2. **Enable Drive API:**
   - Click on **Services** (+ icon) in the left sidebar
   - Find **Google Drive API**
   - Set version to **v3**
   - Click **Add**

   OR

   - Go to **Project Settings** (gear icon)
   - Scroll to **Advanced Google Services**
   - Find **Google Drive API**
   - Toggle it **ON**
   - Set version to **v3**

### Step 2: Add OAuth Scope to appsscript.json

1. **Show appsscript.json:**
   - In Apps Script editor, go to **Project Settings**
   - Check "Show appsscript.json manifest file in editor"

2. **Edit appsscript.json:**
   - Click on `appsscript.json` in the file list
   - Add this scope if not present:
   ```json
   {
     "timeZone": "Africa/Johannesburg",
     "dependencies": {
       "enabledAdvancedServices": [
         {
           "userSymbol": "Drive",
           "version": "v3",
           "serviceId": "drive"
         }
       ]
     },
     "oauthScopes": [
       "https://www.googleapis.com/auth/spreadsheets",
       "https://www.googleapis.com/auth/script.scriptapp",
       "https://www.googleapis.com/auth/drive.file"
     ]
   }
   ```

### Step 3: Force Re-Authorization

1. **In Apps Script Editor:**
   - Select function: `forceAuthorization`
   - Click **Run**

2. **Grant Permissions:**
   - Click **Review Permissions**
   - Choose your Google account
   - Click **Advanced** → **Go to [Your Project]**
   - Click **Allow**
   - Grant access to Drive

### Step 4: Test Again

1. **Run the diagnostics again** - it should now show Drive API test
2. **Try uploading an Excel file** - should work now!

## Quick Fix Commands

### For Apps Script Editor Console:

```javascript
// Test Drive API access
function testDriveAPI() {
  try {
    // Create test file
    var testFile = DriveApp.createFile('test', 'test');
    var fileId = testFile.getId();

    // Test Drive API v3
    var file = Drive.Files.get(fileId);
    Logger.log('✅ Drive API works! File: ' + file.name);

    // Cleanup
    DriveApp.getFileById(fileId).setTrashed(true);
    Logger.log('✅ Drive API test passed!');

  } catch (error) {
    Logger.log('❌ Drive API FAILED: ' + error.message);
    Logger.log('');
    Logger.log('FIX:');
    Logger.log('1. Services → Add Google Drive API v3');
    Logger.log('2. Run forceAuthorization()');
  }
}
```

Run this in the Apps Script editor to test Drive API access.

## Expected Behavior After Fix

### Before Fix:
```
✅ All diagnostics pass
❌ Excel import: "Authorization required or server returned no response"
```

### After Fix:
```
✅ All diagnostics pass
✅ Drive API test: PASSED
✅ Excel import: Works correctly
```

## Verification

Run the updated diagnostics:
1. Open Debugger in web app
2. Click "Run All Tests"
3. Look for authorization test - it should now say:
   ```
   ✅ Drive API v3 access: OK
   ```

If it says:
```
❌ CRITICAL: Drive API access FAILED
```

Then Drive API is not enabled - follow steps above.

## Why This Wasn't Caught Earlier

The initial diagnostic tests only checked:
- Spreadsheet access ✅
- PropertiesService access ✅
- User session ✅

But didn't test Drive API because it requires special configuration. The Excel import feature is the ONLY part of the timesheet system that needs Drive API (to convert Excel → Google Sheets).

## Additional Notes

**Drive API is Required For:**
- Importing clock-in Excel files
- Converting Excel to Google Sheets format
- Processing uploaded timesheets

**Drive API is NOT Required For:**
- Viewing timesheet settings
- Approving timesheets
- Listing pending timesheets
- Updating configurations

That's why everything else works but import doesn't!
