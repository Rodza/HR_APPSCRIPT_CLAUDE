# Timesheet Module Debugging Guide

## Overview

This guide provides comprehensive debugging tools and procedures to diagnose and fix "Authorization required or server returned no response" errors and other timesheet module issues.

## Debugging Tools

### 1. Timesheet Debugger (Client-Side)
**File:** `apps-script/TimesheetDebugger.html`

Interactive HTML debugging tool that tests all client-server communication.

**How to Access:**
1. Open your Google Sheet
2. Go to **HR System → Timesheets → 🔍 Debugger**
3. A dialog window will open with the debugging interface

**Features:**
- ✅ **Ping Test** - Tests basic server connectivity
- ✅ **getTimeConfig Test** - Tests configuration loading
- ✅ **updateTimeConfig Test** - Tests configuration updates
- ✅ **listPendingTimesheets Test** - Tests timesheet retrieval
- ✅ **Server Diagnostics** - Runs comprehensive server-side tests
- ✅ **Response Format Test** - Tests various response formats
- ✅ **Master Debug Log** - Consolidated log of all tests
- ✅ **Export Log** - Download debug log as text file

### 2. Server-Side Diagnostics
**File:** `apps-script/TimesheetDebug.gs`

Comprehensive server-side diagnostic functions.

**How to Run:**
1. Open **Extensions → Apps Script**
2. Select `runTimesheetDiagnostics` from the function dropdown
3. Click **Run**
4. Check **Execution log** (View → Logs or Ctrl+Enter)

**Functions Available:**
- `runTimesheetDiagnostics()` - Main diagnostic function
- `testGetTimeConfig()` - Test config loading
- `testUpdateTimeConfig()` - Test config updates
- `testResetTimeConfig()` - Test config reset
- `testExportTimeConfig()` - Test config export
- `testListPendingTimesheets()` - Test timesheet listing
- `testSheetAccess()` - Test spreadsheet access
- `testPropertiesAccess()` - Test PropertiesService
- `testAuthorizationStatus()` - Test OAuth status
- `pingTest()` - Simple connectivity test

## Step-by-Step Debugging Procedure

### Step 1: Run Client-Side Debugger

1. **Open the Debugger:**
   - HR System → Timesheets → 🔍 Debugger

2. **Click "Run All Tests":**
   - This will run all tests sequentially
   - Watch the logs in real-time

3. **Check for Errors:**
   - ✅ Green = Passed
   - ❌ Red = Failed
   - ⚠️ Yellow = Warning

4. **Export the Log:**
   - Click "Export Debug Log"
   - Save the file for reference

### Step 2: Analyze the Results

#### If Ping Test Fails:
```
❌ ERROR: Server returned null/undefined
```
**Cause:** Server-side function not responding
**Solutions:**
1. Check if code is deployed correctly
2. Verify function names match exactly
3. Check browser console (F12) for errors
4. Try refreshing the page

#### If getTimeConfig Fails with Null:
```
❌ ERROR: Server returned null/undefined - THIS IS THE PROBLEM!
```
**Cause:** `getTimeConfig()` returning null or wrong format
**Solutions:**
1. Verify TimesheetConfig.gs is deployed
2. Check function returns `{success: true, data: config}`
3. Run server-side diagnostics (Step 3)

#### If Success Handler Not Called:
```
Failure handler called
Error: Authorization required
```
**Cause:** Real OAuth authorization issue
**Solutions:**
1. Check OAuth scopes in appsscript.json
2. Run `forceAuthorization()` in Apps Script editor
3. Clear browser cache and cookies
4. Try incognito mode

### Step 3: Run Server-Side Diagnostics

1. **Open Apps Script Editor:**
   - Extensions → Apps Script

2. **Select Function:**
   - Choose `runTimesheetDiagnostics` from dropdown

3. **Run and Check Logs:**
   ```
   Execution log
   ========== TIMESHEET MODULE DIAGNOSTICS ==========

   --- Testing: getTimeConfig() ---
   ✅ PASSED: Function returned correct format

   --- Testing: updateTimeConfig() ---
   ✅ PASSED: Function returned result with success property

   ...

   ========== DIAGNOSTIC SUMMARY ==========
   Total Tests: 8
   ✅ Passed: 8
   ❌ Failed: 0
   ⚠️  Warnings: 0
   ```

4. **Review Failed Tests:**
   - Each failed test shows the error message
   - Check stack traces for detailed info

### Step 4: Common Issues and Fixes

#### Issue 1: Return Format Mismatch

**Symptom:**
```javascript
if (!result) {  // Passes when it shouldn't
    showError('Authorization required...');
}
```

**Fix:**
Ensure all server functions return:
```javascript
return {
    success: true,
    data: yourData
};

// OR on error:
return {
    success: false,
    error: errorMessage
};
```

#### Issue 2: Real Authorization Error

**Symptom:**
```
Exception: You do not have permission to call...
```

**Fix:**
1. Open Apps Script editor
2. Run `forceAuthorization()`
3. Grant all requested permissions
4. Redeploy if necessary

#### Issue 3: Sheet Not Found

**Symptom:**
```
PENDING_TIMESHEETS sheet not found
```

**Fix:**
1. Run `setupPendingTimesheetsSheet()`
2. Or create sheet manually with required columns

#### Issue 4: PropertiesService Error

**Symptom:**
```
Cannot access PropertiesService
```

**Fix:**
1. Check OAuth scopes include script properties
2. Verify script permissions
3. Try clearing properties and resetting

## Manual Testing Scripts

### Test 1: Basic Connectivity

Run in Apps Script editor:
```javascript
function testBasicConnectivity() {
  Logger.log('Testing basic connectivity...');

  // Test 1: Spreadsheet access
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('✅ Spreadsheet: ' + ss.getName());

  // Test 2: Properties access
  var props = PropertiesService.getScriptProperties();
  props.setProperty('TEST', 'value');
  Logger.log('✅ Properties: ' + props.getProperty('TEST'));
  props.deleteProperty('TEST');

  // Test 3: User session
  var user = Session.getActiveUser().getEmail();
  Logger.log('✅ User: ' + user);

  Logger.log('All connectivity tests passed!');
}
```

### Test 2: Config Function Format

Run in Apps Script editor:
```javascript
function testConfigFormat() {
  var result = getTimeConfig();

  Logger.log('Result type: ' + typeof result);
  Logger.log('Is null: ' + (result === null));
  Logger.log('Is undefined: ' + (result === undefined));
  Logger.log('Has success: ' + result.hasOwnProperty('success'));
  Logger.log('Has data: ' + result.hasOwnProperty('data'));
  Logger.log('Success value: ' + result.success);
  Logger.log('Data type: ' + typeof result.data);
  Logger.log('Data keys: ' + Object.keys(result.data).length);

  if (result.success && result.data) {
    Logger.log('✅ FORMAT CORRECT');
  } else {
    Logger.log('❌ FORMAT INCORRECT');
  }
}
```

### Test 3: HTML Success Handler

Add to HTML file temporarily:
```javascript
function debugSuccessHandler(functionName) {
    google.script.run
        .withSuccessHandler(function(result) {
            console.log('===== SUCCESS HANDLER DEBUG =====');
            console.log('Function:', functionName);
            console.log('Result:', result);
            console.log('Result type:', typeof result);
            console.log('Is null:', result === null);
            console.log('Is undefined:', result === undefined);
            console.log('Is object:', typeof result === 'object');

            if (result) {
                console.log('Has success:', result.hasOwnProperty('success'));
                console.log('Has data:', result.hasOwnProperty('data'));
                console.log('Keys:', Object.keys(result));
            }

            console.log('================================');
        })
        .withFailureHandler(function(error) {
            console.error('===== FAILURE HANDLER =====');
            console.error('Function:', functionName);
            console.error('Error:', error);
            console.error('===========================');
        })
        .getTimeConfig();  // Replace with your function
}

// Call it
debugSuccessHandler('getTimeConfig');
```

## Browser Console Debugging

### Enable Console Logging

1. **Open Browser Console:**
   - Press F12
   - Go to "Console" tab

2. **Add Logging to HTML:**
   ```javascript
   console.log('Before calling server function');

   google.script.run
       .withSuccessHandler(function(result) {
           console.log('Success handler result:', result);
       })
       .withFailureHandler(function(error) {
           console.error('Failure handler error:', error);
       })
       .yourFunction();
   ```

3. **Watch for Errors:**
   - Look for red errors
   - Check network tab for failed requests
   - Look for CORS errors or blocked requests

## Troubleshooting Checklist

- [ ] Code deployed to Apps Script editor
- [ ] All files saved (Ctrl+S)
- [ ] Browser cache cleared
- [ ] Tried incognito/private mode
- [ ] OAuth permissions granted
- [ ] No typos in function names
- [ ] Return format is `{success:  true, data: ...}`
- [ ] Success handler checks `if (!result)`
- [ ] PENDING_TIMESHEETS sheet exists
- [ ] PropertiesService accessible
- [ ] No console errors (F12)
- [ ] Tried on different browser

## Getting Help

### Information to Provide

When asking for help, provide:

1. **Debug Log Export:**
   - Run all tests in Debugger
   - Export debug log
   - Share the log file

2. **Server Diagnostic Results:**
   - Run `runTimesheetDiagnostics()`
   - Copy the execution log
   - Include in your help request

3. **Browser Console Log:**
   - Open F12 console
   - Reproduce the error
   - Screenshot or copy console output

4. **Specific Error Message:**
   - Exact text of the error
   - When it occurs (which action)
   - Which module (Settings, Approval, Import)

5. **Environment Info:**
   - Browser and version
   - Operating system
   - Whether issue occurs in incognito mode

## Advanced Debugging

### Network Tab Analysis

1. **Open Network Tab:**
   - F12 → Network tab
   - Reload page

2. **Filter for Script Calls:**
   - Look for requests to `script.google.com`
   - Check for failed requests (red)

3. **Inspect Response:**
   - Click on failed request
   - Check "Response" tab
   - Look for error messages

### Apps Script Execution Log

1. **View Detailed Logs:**
   - Apps Script editor → View → Logs
   - Or press Ctrl+Enter

2. **Look for:**
   - `⚠️` Warnings
   - `❌` Errors
   - Exception messages
   - Stack traces

### OAuth Scope Verification

1. **Check appsscript.json:**
   ```json
   {
     "oauthScopes": [
       "https://www.googleapis.com/auth/spreadsheets",
       "https://www.googleapis.com/auth/script.scriptapp",
       "https://www.googleapis.com/auth/script.external_request"
     ]
   }
   ```

2. **Verify in Project Settings:**
   - Settings (gear icon)
   - Show "appsscript.json" in editor
   - Check OAuth scopes section

## Summary

The debugging tools provide comprehensive diagnostics for timesheet module issues. Follow the step-by-step procedure to identify and fix problems systematically. Most issues can be resolved by ensuring proper return formats and OAuth authorization.

For persistent issues, export debug logs and server diagnostics before seeking help.
