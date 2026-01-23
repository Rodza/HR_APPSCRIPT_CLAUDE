/**
 * TIMESHEETS.GS - Timesheet Approval Workflow Module
 * HR Payroll System for SA Grinding Wheels & Scorpio Abrasives
 *
 * Handles time approval workflow:
 * 1. Import Excel from HTML time clock analyzer
 * 2. Display in PendingTimesheets with editable data
 * 3. User reviews and approves/rejects
 * 4. Approved records create payslips in MASTERSALARY
 *

/**
 * DIAGNOSTIC FUNCTION - Run this directly from Apps Script editor
 * This will analyze the current RAW_CLOCK_DATA to find the timezone issue
 *
 * HOW TO RUN:
 * 1. Open Apps Script editor
 * 2. Select "diagnosePunchTimeIssue" from function dropdown at top
 * 3. Click Run ▶ button
 * 4. View logs in Execution log below
 */
function diagnosePunchTimeIssue() {
  Logger.log('\n========== PUNCH TIME DIAGNOSTIC ==========\n');

  // Get timezone info
  const scriptTz = Session.getScriptTimeZone();
  Logger.log('🌍 TIMEZONE INFORMATION:');
  Logger.log('   Script timezone: ' + scriptTz);

  const now = new Date();
  Logger.log('   Current time: ' + now.toString());
  Logger.log('   UTC time: ' + now.toUTCString());
  Logger.log('   Timezone offset: ' + now.getTimezoneOffset() + ' minutes (' + (now.getTimezoneOffset() / -60) + ' hours from UTC)');

  // Get spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('\n📊 SPREADSHEET:');
  Logger.log('   Name: ' + ss.getName());
  Logger.log('   Timezone: ' + ss.getSpreadsheetTimeZone());

  // Get RAW_CLOCK_DATA sheet
  const sheet = ss.getSheetByName('RAW_CLOCK_DATA');
  if (!sheet) {
    Logger.log('\n❌ RAW_CLOCK_DATA sheet not found!');
    return;
  }

  Logger.log('\n📋 RAW_CLOCK_DATA SHEET:');
  Logger.log('   Last row: ' + sheet.getLastRow());

  // Get PUNCH_TIME column (column 9)
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('   No data rows found');
    return;
  }

  // Read first 5 data rows
  const numRows = Math.min(5, lastRow - 1);
  const punchTimeCol = 9; // PUNCH_TIME column

  const rawValues = sheet.getRange(2, punchTimeCol, numRows, 1).getValues();
  const displayValues = sheet.getRange(2, punchTimeCol, numRows, 1).getDisplayValues();

  Logger.log('\n🔍 ANALYZING FIRST ' + numRows + ' PUNCH TIME VALUES:\n');

  for (let i = 0; i < numRows; i++) {
    const rawVal = rawValues[i][0];
    const dispVal = displayValues[i][0];

    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('ROW ' + (i + 2) + ':');
    Logger.log('   Display value: "' + dispVal + '"');
    Logger.log('   Raw value type: ' + typeof rawVal);

    if (rawVal instanceof Date) {
      Logger.log('   Is Date object: YES');
      Logger.log('   .toString(): ' + rawVal.toString());
      Logger.log('   .toISOString(): ' + rawVal.toISOString());
      Logger.log('   .toUTCString(): ' + rawVal.toUTCString());
      Logger.log('   .toLocaleString(): ' + rawVal.toLocaleString());
      Logger.log('');
      Logger.log('   LOCAL components:');
      Logger.log('      Date: ' + rawVal.getFullYear() + '-' + (rawVal.getMonth()+1) + '-' + rawVal.getDate());
      Logger.log('      Time: ' + rawVal.getHours() + ':' + rawVal.getMinutes() + ':' + rawVal.getSeconds());
      Logger.log('');
      Logger.log('   UTC components:');
      Logger.log('      Date: ' + rawVal.getUTCFullYear() + '-' + (rawVal.getUTCMonth()+1) + '-' + rawVal.getUTCDate());
      Logger.log('      Time: ' + rawVal.getUTCHours() + ':' + rawVal.getUTCMinutes() + ':' + rawVal.getUTCSeconds());
      Logger.log('');
      Logger.log('   🎯 PROBLEM CHECK:');
      Logger.log('      Excel should show: 13:04:57 (expected)');
      Logger.log('      Display shows: ' + dispVal + ' ❌ if it shows 23:04:57');
      Logger.log('      getHours() returns: ' + rawVal.getHours() + ' ❌ if this is 23, should be 13');
      Logger.log('      getUTCHours() returns: ' + rawVal.getUTCHours() + ' ✅ if this is 13');
    } else {
      Logger.log('   Is Date object: NO');
      Logger.log('   Value: ' + rawVal);
    }
    Logger.log('');
  }

  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('\n📝 INTERPRETATION GUIDE:');
  Logger.log('   If getHours() = 23 and getUTCHours() = 13:');
  Logger.log('      → The Date object contains the WRONG time (23 instead of 13)');
  Logger.log('      → This means the issue happened BEFORE storing to sheet');
  Logger.log('      → We need to fix the parseClockDataExcel function');
  Logger.log('');
  Logger.log('   If getHours() = 13 but display shows 23:04:57:');
  Logger.log('      → The Date object has the CORRECT time');
  Logger.log('      → The display formatting is wrong');
  Logger.log('      → We need to fix the number format or timezone setting');
  Logger.log('');
  Logger.log('========== DIAGNOSTIC COMPLETE ==========\n');
}

/**
 * Sheet: PendingTimesheets
 * Columns: ID, EMPLOYEE NAME, WEEKENDING, HOURS, MINUTES, OVERTIMEHOURS,
 *          OVERTIMEMINUTES, NOTES, STATUS, IMPORTED_BY, IMPORTED_DATE,
 *          REVIEWED_BY, REVIEWED_DATE
 *
 * Excel Import Format (from HTML analyzer):
 * Employee Name, Week Ending, Standard Hours, Standard Minutes, Overtime Hours, Overtime Minutes, Notes
 */

// ==================== SETUP AND INITIALIZATION ====================

/**
 * Setup PENDING_TIMESHEETS sheet if it doesn't exist
 * Run this manually from Script Editor if the sheet is missing
 */
function setupPendingTimesheetsSheet() {
  try {
    Logger.log('========== SETUP PENDING_TIMESHEETS SHEET ==========');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = null;

    // Check if sheet already exists (try both names for compatibility)
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      const sheetName = sheets[i].getName();
      if (sheetName === 'PENDING_TIMESHEETS' || sheetName === 'PendingTimesheets') {
        sheet = sheets[i];
        Logger.log('✓ Found existing sheet: ' + sheetName);
        break;
      }
    }

    // Create if doesn't exist (use PendingTimesheets for compatibility)
    if (!sheet) {
      Logger.log('Creating PendingTimesheets sheet...');
      sheet = ss.insertSheet('PendingTimesheets');
      Logger.log('✓ Sheet created');
    }

    // Set up headers from Config.gs
    const headers = PENDING_TIMESHEETS_COLUMNS;
    Logger.log('Setting up ' + headers.length + ' column headers...');

    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f3f3f3');

    // Freeze header row
    sheet.setFrozenRows(1);

    // Auto-resize columns
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }

    Logger.log('✓ PendingTimesheets sheet setup complete!');
    Logger.log('Sheet now has ' + headers.length + ' columns');
    Logger.log('Headers: ' + headers.join(', '));
    Logger.log('========== SETUP COMPLETE ==========');

    return { success: true, message: 'PendingTimesheets sheet created successfully' };

  } catch (error) {
    Logger.log('❌ ERROR: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    return { success: false, error: error.message };
  }
}

/**
 * Setup RAW_CLOCK_DATA sheet if it doesn't exist
 * Run this manually from Script Editor if the sheet is missing
 */
function setupRawClockDataSheet() {
  try {
    Logger.log('========== SETUP RAW_CLOCK_DATA SHEET ==========');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = null;

    // Check if sheet already exists
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getName() === 'RAW_CLOCK_DATA') {
        sheet = sheets[i];
        Logger.log('✓ RAW_CLOCK_DATA sheet already exists');
        break;
      }
    }

    // Create if doesn't exist
    if (!sheet) {
      Logger.log('Creating RAW_CLOCK_DATA sheet...');
      sheet = ss.insertSheet('RAW_CLOCK_DATA');
      Logger.log('✓ Sheet created');
    }

    // Set up headers from Config.gs
    const headers = RAW_CLOCK_DATA_COLUMNS;
    Logger.log('Setting up ' + headers.length + ' column headers...');

    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f3f3f3');

    // Freeze header row
    sheet.setFrozenRows(1);

    // Format PUNCH_TIME column (column 9) to show datetime format
    // This ensures time is visible, not just date
    const punchTimeCol = 9; // PUNCH_TIME in RAW_CLOCK_DATA_COLUMNS
    const lastRow = Math.max(2, sheet.getLastRow()); // At least row 2
    const punchTimeRange = sheet.getRange(2, punchTimeCol, lastRow - 1, 1);
    punchTimeRange.setNumberFormat('yyyy-mm-dd hh:mm:ss');

    // Auto-resize columns
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }

    Logger.log('✓ RAW_CLOCK_DATA sheet setup complete!');
    Logger.log('Sheet now has ' + headers.length + ' columns');
    Logger.log('Headers: ' + headers.join(', '));
    Logger.log('========== SETUP COMPLETE ==========');

    return { success: true, message: 'RAW_CLOCK_DATA sheet created successfully' };

  } catch (error) {
    Logger.log('❌ ERROR: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    return { success: false, error: error.message };
  }
}

/**
 * Setup CLOCK_IN_IMPORTS sheet if it doesn't exist
 * Run this manually from Script Editor if the sheet is missing
 */
function setupClockInImportsSheet() {
  try {
    Logger.log('========== SETUP CLOCK_IN_IMPORTS SHEET ==========');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = null;

    // Check if sheet already exists
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getName() === 'CLOCK_IN_IMPORTS') {
        sheet = sheets[i];
        Logger.log('✓ CLOCK_IN_IMPORTS sheet already exists');
        break;
      }
    }

    // Create if doesn't exist
    if (!sheet) {
      Logger.log('Creating CLOCK_IN_IMPORTS sheet...');
      sheet = ss.insertSheet('CLOCK_IN_IMPORTS');
      Logger.log('✓ Sheet created');
    }

    // Set up headers from Config.gs
    const headers = CLOCK_IN_IMPORTS_COLUMNS;
    Logger.log('Setting up ' + headers.length + ' column headers...');

    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f3f3f3');

    // Freeze header row
    sheet.setFrozenRows(1);

    // Auto-resize columns
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }

    Logger.log('✓ CLOCK_IN_IMPORTS sheet setup complete!');
    Logger.log('Sheet now has ' + headers.length + ' columns');
    Logger.log('Headers: ' + headers.join(', '));
    Logger.log('========== SETUP COMPLETE ==========');

    return { success: true, message: 'CLOCK_IN_IMPORTS sheet created successfully' };

  } catch (error) {
    Logger.log('❌ ERROR: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    return { success: false, error: error.message };
  }
}

/**
 * Fix RAW_CLOCK_DATA headers if they're incomplete
 * This function adds missing headers to existing sheet
 */
function fixRawClockDataHeaders() {
  try {
    Logger.log('========== FIX RAW_CLOCK_DATA HEADERS ==========');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = null;

    // Find the sheet
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getName() === 'RAW_CLOCK_DATA') {
        sheet = sheets[i];
        break;
      }
    }

    if (!sheet) {
      throw new Error('RAW_CLOCK_DATA sheet not found');
    }

    // Get current headers
    const currentHeaderRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    const currentHeaders = currentHeaderRange.getValues()[0];
    Logger.log('Current headers (' + currentHeaders.length + '): ' + currentHeaders.join(', '));

    // Expected headers
    const expectedHeaders = RAW_CLOCK_DATA_COLUMNS;
    Logger.log('Expected headers (' + expectedHeaders.length + '): ' + expectedHeaders.join(', '));

    if (currentHeaders.length === expectedHeaders.length) {
      Logger.log('✓ Headers are already complete (' + currentHeaders.length + ' columns)');
      return { success: true, message: 'Headers already correct' };
    }

    if (currentHeaders.length > expectedHeaders.length) {
      throw new Error('Sheet has more columns than expected: ' + currentHeaders.length + ' > ' + expectedHeaders.length);
    }

    // Set all headers to correct values
    Logger.log('🔧 Updating headers to correct values...');
    const headerRange = sheet.getRange(1, 1, 1, expectedHeaders.length);
    headerRange.setValues([expectedHeaders]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f3f3f3');

    // Auto-resize all columns
    for (let i = 1; i <= expectedHeaders.length; i++) {
      sheet.autoResizeColumn(i);
    }

    SpreadsheetApp.flush();

    Logger.log('✅ Headers fixed!');
    Logger.log('   Before: ' + currentHeaders.length + ' columns');
    Logger.log('   After: ' + expectedHeaders.length + ' columns');
    Logger.log('   Missing headers added: ' + expectedHeaders.slice(currentHeaders.length).join(', '));
    Logger.log('========== FIX COMPLETE ==========');

    return {
      success: true,
      message: 'Headers fixed successfully',
      before: currentHeaders.length,
      after: expectedHeaders.length,
      added: expectedHeaders.slice(currentHeaders.length)
    };

  } catch (error) {
    Logger.log('❌ ERROR: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    return { success: false, error: error.message };
  }
}

/**
 * Fix PendingTimesheets headers to match Config.gs
 * This function corrects wrong column header names
 */
function fixPendingTimesheetsHeaders() {
  try {
    Logger.log('========== FIX PENDINGTIMESHEETS HEADERS ==========');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = null;

    // Find the sheet (try both names)
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      const sheetName = sheets[i].getName();
      if (sheetName === 'PENDING_TIMESHEETS' || sheetName === 'PendingTimesheets') {
        sheet = sheets[i];
        Logger.log('✓ Found sheet: ' + sheetName);
        break;
      }
    }

    if (!sheet) {
      throw new Error('PendingTimesheets sheet not found');
    }

    // Get current headers
    const currentHeaderRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    const currentHeaders = currentHeaderRange.getValues()[0];
    Logger.log('Current headers (' + currentHeaders.length + '): ' + currentHeaders.join(', '));

    // Expected headers
    const expectedHeaders = PENDING_TIMESHEETS_COLUMNS;
    Logger.log('Expected headers (' + expectedHeaders.length + '): ' + expectedHeaders.join(', '));

    // Check if headers match
    let mismatchCount = 0;
    const mismatches = [];
    for (let i = 0; i < Math.min(currentHeaders.length, expectedHeaders.length); i++) {
      if (currentHeaders[i] !== expectedHeaders[i]) {
        mismatchCount++;
        mismatches.push({
          position: i + 1,
          current: currentHeaders[i],
          expected: expectedHeaders[i]
        });
      }
    }

    if (mismatchCount === 0 && currentHeaders.length === expectedHeaders.length) {
      Logger.log('✓ Headers are already correct');
      return { success: true, message: 'Headers already correct' };
    }

    Logger.log('🔧 Found ' + mismatchCount + ' header mismatches');
    mismatches.forEach(function(m) {
      Logger.log('   Column ' + m.position + ': "' + m.current + '" → "' + m.expected + '"');
    });

    // Set all headers to correct values
    Logger.log('🔧 Updating headers to correct values...');
    const headerRange = sheet.getRange(1, 1, 1, expectedHeaders.length);
    headerRange.setValues([expectedHeaders]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f3f3f3');

    // Auto-resize all columns
    for (let i = 1; i <= expectedHeaders.length; i++) {
      sheet.autoResizeColumn(i);
    }

    SpreadsheetApp.flush();

    Logger.log('✅ Headers fixed!');
    Logger.log('   Corrected ' + mismatchCount + ' header(s)');
    Logger.log('========== FIX COMPLETE ==========');

    return {
      success: true,
      message: 'Headers fixed successfully',
      corrected: mismatchCount,
      mismatches: mismatches
    };

  } catch (error) {
    Logger.log('❌ ERROR: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    return { success: false, error: error.message };
  }
}

/**
 * Setup ALL timesheet-related sheets at once
 * Convenience function to set up all required sheets
 */
function setupAllTimesheetSheets() {
  try {
    Logger.log('\n========== SETUP ALL TIMESHEET SHEETS ==========\n');

    const results = [];

    // Setup RAW_CLOCK_DATA
    Logger.log('1/3: Setting up RAW_CLOCK_DATA...');
    const rawClockResult = setupRawClockDataSheet();
    results.push({ sheet: 'RAW_CLOCK_DATA', result: rawClockResult });

    // Setup CLOCK_IN_IMPORTS
    Logger.log('\n2/3: Setting up CLOCK_IN_IMPORTS...');
    const clockImportsResult = setupClockInImportsSheet();
    results.push({ sheet: 'CLOCK_IN_IMPORTS', result: clockImportsResult });

    // Setup PENDING_TIMESHEETS
    Logger.log('\n3/3: Setting up PENDING_TIMESHEETS...');
    const pendingResult = setupPendingTimesheetsSheet();
    results.push({ sheet: 'PENDING_TIMESHEETS', result: pendingResult });

    Logger.log('\n========== SETUP SUMMARY ==========');
    results.forEach(function(item) {
      const status = item.result.success ? '✅ SUCCESS' : '❌ FAILED';
      Logger.log(status + ': ' + item.sheet);
      if (!item.result.success) {
        Logger.log('   Error: ' + item.result.error);
      }
    });
    Logger.log('========== ALL SETUP COMPLETE ==========\n');

    return {
      success: true,
      message: 'All sheets setup completed',
      results: results
    };

  } catch (error) {
    Logger.log('❌ ERROR in setupAllTimesheetSheets: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    return { success: false, error: error.message };
  }
}

/**
 * Fix PUNCH_TIME column format in RAW_CLOCK_DATA to show datetime
 * Run this to fix existing data that only shows dates without times
 */
function fixRawClockDataPunchTimeFormat() {
  try {
    Logger.log('========== FIX PUNCH_TIME COLUMN FORMAT ==========');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = null;

    // Find the sheet
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getName() === 'RAW_CLOCK_DATA') {
        sheet = sheets[i];
        break;
      }
    }

    if (!sheet) {
      throw new Error('RAW_CLOCK_DATA sheet not found');
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('✓ No data rows to format');
      return { success: true, message: 'No data to format' };
    }

    // Format PUNCH_TIME column (column 9) to show both date and time
    const punchTimeCol = 9;
    const punchTimeRange = sheet.getRange(2, punchTimeCol, lastRow - 1, 1);
    punchTimeRange.setNumberFormat('yyyy-mm-dd hh:mm:ss');

    SpreadsheetApp.flush();

    Logger.log('✅ Formatted ' + (lastRow - 1) + ' PUNCH_TIME cells');
    Logger.log('   Column 9 (PUNCH_TIME) now shows: YYYY-MM-DD HH:MM:SS');
    Logger.log('========== FIX COMPLETE ==========');

    return {
      success: true,
      message: 'PUNCH_TIME column formatted successfully',
      rowsFormatted: lastRow - 1
    };

  } catch (error) {
    Logger.log('❌ ERROR: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    return { success: false, error: error.message };
  }
}

/**
 * DIAGNOSTIC: Display current PendingTimesheets column headers
 * Run this from Apps Script editor to see what headers are actually in the sheet
 */
function checkPendingTimesheetsHeaders() {
  try {
    Logger.log('\n========== CHECK PENDINGTIMESHEETS HEADERS ==========\n');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = null;

    // Find the sheet (try both names)
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      const sheetName = sheets[i].getName();
      if (sheetName === 'PENDING_TIMESHEETS' || sheetName === 'PendingTimesheets') {
        sheet = sheets[i];
        Logger.log('✓ Found sheet: ' + sheetName);
        break;
      }
    }

    if (!sheet) {
      throw new Error('PendingTimesheets sheet not found');
    }

    // Get current headers
    const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Expected headers from Config.gs (inlined to avoid dependency issues)
    const expectedHeaders = [
      'ID',
      'RAW_DATA_IMPORT_ID',
      'EMPLOYEE_ID',
      'EMPLOYEE NAME',
      'EMPLOYEE_CLOCK_REF',
      'WEEKENDING',
      'CALCULATED_TOTAL_HOURS',
      'CALCULATED_TOTAL_MINUTES',
      'HOURS',
      'MINUTES',
      'OVERTIMEHOURS',
      'OVERTIMEMINUTES',
      'LUNCH_DEDUCTION_MIN',
      'BATHROOM_TIME_MIN',
      'RECON_DETAILS',
      'WARNINGS',
      'NOTES',
      'STATUS',
      'IMPORTED_BY',
      'IMPORTED_DATE',
      'REVIEWED_BY',
      'REVIEWED_DATE',
      'PAYSLIP_ID',
      'IS_LOCKED',
      'LOCKED_DATE'
    ];

    Logger.log('Total columns: ' + currentHeaders.length);
    Logger.log('Expected columns: ' + expectedHeaders.length);
    Logger.log('\n========== CURRENT HEADERS ==========');

    for (let i = 0; i < currentHeaders.length; i++) {
      const colNum = i + 1;
      const current = currentHeaders[i];
      const expected = expectedHeaders[i] || 'N/A';
      const match = current === expected ? '✓' : '✗';

      Logger.log(match + ' Column ' + colNum + ': "' + current + '"' +
                 (current !== expected ? ' (Expected: "' + expected + '")' : ''));
    }

    Logger.log('\n========== EXPECTED HEADERS (Config.gs) ==========');
    for (let i = 0; i < expectedHeaders.length; i++) {
      Logger.log('Column ' + (i + 1) + ': ' + expectedHeaders[i]);
    }

    Logger.log('\n========== CHECK COMPLETE ==========\n');

    return { success: true, currentHeaders: currentHeaders, expectedHeaders: expectedHeaders };

  } catch (error) {
    Logger.log('❌ ERROR: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    return { success: false, error: error.message };
  }
}

// ==================== IMPORT EXCEL TIMESHEET ====================

/**
 * Imports timesheet data from Excel file
 * Parses the file and creates pending timesheet records
 *
 * @param {Blob} fileBlob - Excel file blob
 * @returns {Object} Result with success flag and import stats
 */
function importTimesheetExcel(fileBlob) {
  try {
    Logger.log('\n========== IMPORT TIMESHEET EXCEL ==========');

    // Parse Excel data
    const parseResult = parseExcelData(fileBlob);
    if (!parseResult.success) {
      throw new Error('Failed to parse Excel: ' + parseResult.error);
    }

    const records = parseResult.data;
    Logger.log('📊 Found ' + records.length + ' records in Excel');

    // Import each record
    let successCount = 0;
    let errorCount = 0;
    const errors = [];

    for (let i = 0; i < records.length; i++) {
      const record = records[i];

      const result = addPendingTimesheet(record);

      if (result.success) {
        successCount++;
      } else {
        errorCount++;
        errors.push('Row ' + (i + 2) + ': ' + result.error);
      }
    }

    Logger.log('✅ Import complete: ' + successCount + ' success, ' + errorCount + ' errors');

    if (errors.length > 0) {
      Logger.log('⚠️ Errors:');
      errors.forEach(err => Logger.log('  - ' + err));
    }

    Logger.log('========== IMPORT TIMESHEET EXCEL COMPLETE ==========\n');

    return {
      success: true,
      data: {
        totalRecords: records.length,
        successCount: successCount,
        errorCount: errorCount,
        errors: errors
      }
    };

  } catch (error) {
    Logger.log('❌ ERROR in importTimesheetExcel: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== IMPORT TIMESHEET EXCEL FAILED ==========\n');
    return { success: false, error: error.message };
  }
}

/**
 * Parses Excel file and extracts timesheet data
 *
 * @param {Blob} fileBlob - Excel file blob
 * @returns {Object} Result with parsed data
 */
function parseExcelData(fileBlob) {
  try {
    Logger.log('📖 Parsing Excel file...');

    // Convert Excel to CSV using Drive API
    const file = DriveApp.createFile(fileBlob);
    const fileId = file.getId();

    // Import as Google Sheets to parse
    const resource = {
      title: fileBlob.getName(),
      mimeType: MimeType.GOOGLE_SHEETS
    };

    const importedFile = Drive.Files.copy(resource, fileId);
    const spreadsheet = SpreadsheetApp.openById(importedFile.id);
    const sheet = spreadsheet.getSheets()[0];
    const data = sheet.getDataRange().getValues();

    // Delete temporary files
    DriveApp.getFileById(fileId).setTrashed(true);
    DriveApp.getFileById(importedFile.id).setTrashed(true);

    // Expected format:
    // Employee Name, Week Ending, Standard Hours, Standard Minutes, Overtime Hours, Overtime Minutes, Notes
    const headers = data[0];
    Logger.log('📋 Headers: ' + headers.join(', '));

    const records = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Skip empty rows
      if (!row[0]) continue;

      const record = {
        employeeName: row[0],
        weekEnding: row[1],
        hours: row[2] || 0,
        minutes: row[3] || 0,
        overtimeHours: row[4] || 0,
        overtimeMinutes: row[5] || 0,
        notes: row[6] || ''
      };

      records.push(record);
    }

    Logger.log('✅ Parsed ' + records.length + ' records');

    return { success: true, data: records };

  } catch (error) {
    Logger.log('❌ ERROR in parseExcelData: ' + error.message);
    return { success: false, error: error.message };
  }
}

// ==================== ADD PENDING TIMESHEET ====================

/**
 * Adds a pending timesheet record for review
 *
 * @param {Object} data - Timesheet data
 * @param {string} data.employeeName - Employee REFNAME (required)
 * @param {Date|string} data.weekEnding - Week ending date (required)
 * @param {number} data.hours - Standard hours worked (required)
 * @param {number} data.minutes - Standard minutes worked (required)
 * @param {number} data.overtimeHours - Overtime hours (required)
 * @param {number} data.overtimeMinutes - Overtime minutes (required)
 * @param {string} [data.notes] - Notes (optional)
 *
 * @returns {Object} Result with success flag and timesheet record
 */
function addPendingTimesheet(data) {
  try {
    Logger.log('➕ Adding pending timesheet for: ' + data.employeeName);

    // Validate input
    validateTimesheet(data);

    // Get sheets
    const sheets = getSheets();
    const pendingSheet = sheets.pending;

    if (!pendingSheet) {
      throw new Error('PendingTimesheets sheet not found');
    }

    // Check for duplicate (same employee + same week)
    const duplicate = checkDuplicatePendingTimesheet(data.employeeName, data.weekEnding);
    if (duplicate) {
      throw new Error('Pending timesheet already exists for ' + data.employeeName + ' week ending ' + formatDate(data.weekEnding));
    }

    // Generate ID
    const id = generateFullUUID();
    const importedBy = getCurrentUser();
    const importedDate = normalizeToDateOnly(new Date());
    const weekEnding = parseDate(data.weekEnding);

    // Prepare row data
    const rowData = [
      id,                                 // ID
      data.employeeName,                  // EMPLOYEE NAME
      weekEnding,                         // WEEKENDING
      data.hours || 0,                    // HOURS
      data.minutes || 0,                  // MINUTES
      data.overtimeHours || 0,            // OVERTIMEHOURS
      data.overtimeMinutes || 0,          // OVERTIMEMINUTES
      data.notes || '',                   // NOTES
      'Pending',                          // STATUS
      importedBy,                         // IMPORTED_BY
      importedDate,                       // IMPORTED_DATE
      '',                                 // REVIEWED_BY
      ''                                  // REVIEWED_DATE
    ];

    // Append to sheet
    pendingSheet.appendRow(rowData);

    Logger.log('✅ Pending timesheet added');

    return { success: true, data: { id: id } };

  } catch (error) {
    Logger.log('❌ ERROR in addPendingTimesheet: ' + error.message);
    return { success: false, error: error.message };
  }
}

// ==================== UPDATE PENDING TIMESHEET ====================

/**
 * Updates a pending timesheet record (for user edits before approval)
 *
 * @param {string} id - Timesheet ID
 * @param {Object} data - Updated timesheet data
 * @returns {Object} Result with success flag
 */
function updatePendingTimesheet(id, data) {
  try {
    Logger.log('\n========== UPDATE PENDING TIMESHEET ==========');
    Logger.log('ID: ' + id);

    const sheets = getSheets();
    const pendingSheet = sheets.pending;

    if (!pendingSheet) {
      throw new Error('PendingTimesheets sheet not found');
    }

    const sheetData = pendingSheet.getDataRange().getValues();
    const headers = sheetData[0];

    // Find column indexes
    const idCol = findColumnIndex(headers, 'ID');
    const hoursCol = findColumnIndex(headers, 'HOURS');
    const minutesCol = findColumnIndex(headers, 'MINUTES');
    const overtimeHoursCol = findColumnIndex(headers, 'OVERTIMEHOURS');
    const overtimeMinutesCol = findColumnIndex(headers, 'OVERTIMEMINUTES');
    const notesCol = findColumnIndex(headers, 'NOTES');

    // Find the record
    let rowIndex = -1;
    for (let i = 1; i < sheetData.length; i++) {
      if (sheetData[i][idCol] === id) {
        rowIndex = i + 1;  // Sheet row number (1-based)
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error('Pending timesheet not found: ' + id);
    }

    // Update fields
    if (data.hours !== undefined) {
      pendingSheet.getRange(rowIndex, hoursCol + 1).setValue(data.hours);
    }
    if (data.minutes !== undefined) {
      pendingSheet.getRange(rowIndex, minutesCol + 1).setValue(data.minutes);
    }
    if (data.overtimeHours !== undefined) {
      pendingSheet.getRange(rowIndex, overtimeHoursCol + 1).setValue(data.overtimeHours);
    }
    if (data.overtimeMinutes !== undefined) {
      pendingSheet.getRange(rowIndex, overtimeMinutesCol + 1).setValue(data.overtimeMinutes);
    }
    if (data.notes !== undefined) {
      pendingSheet.getRange(rowIndex, notesCol + 1).setValue(data.notes);
    }

    SpreadsheetApp.flush();

    Logger.log('✅ Pending timesheet updated');
    Logger.log('========== UPDATE PENDING TIMESHEET COMPLETE ==========\n');

    return { success: true, data: 'Updated successfully' };

  } catch (error) {
    Logger.log('❌ ERROR in updatePendingTimesheet: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== UPDATE PENDING TIMESHEET FAILED ==========\n');
    return { success: false, error: error.message };
  }
}

// ==================== APPROVE TIMESHEET ====================

/**
 * Approves a pending timesheet and creates payslip
 * CRITICAL: Creates MASTERSALARY record and triggers loan sync
 *
 * @param {string} id - Timesheet ID
 * @returns {Object} Result with success flag and payslip record number
 */
function approveTimesheet(id) {
  try {
    Logger.log('\n========== APPROVE TIMESHEET ==========');
    Logger.log('ID: ' + id);

    const sheets = getSheets();
    const pendingSheet = sheets.pending;

    if (!pendingSheet) {
      throw new Error('PendingTimesheets sheet not found');
    }

    // Get the pending timesheet
    const sheetData = pendingSheet.getDataRange().getValues();
    const headers = sheetData[0];

    const idCol = findColumnIndex(headers, 'ID');

    let timesheetRecord = null;
    let rowIndex = -1;

    for (let i = 1; i < sheetData.length; i++) {
      if (sheetData[i][idCol] === id) {
        timesheetRecord = buildObjectFromRow(sheetData[i], headers);
        rowIndex = i + 1;
        break;
      }
    }

    if (!timesheetRecord) {
      throw new Error('Pending timesheet not found: ' + id);
    }

    // Check if already approved
    if (timesheetRecord['STATUS'] === 'Approved') {
      throw new Error('Timesheet already approved');
    }

    // Verify employee exists
    const empResult = getEmployeeByName(timesheetRecord['EMPLOYEE NAME']);
    if (!empResult.success) {
      throw new Error('Employee not found: ' + timesheetRecord['EMPLOYEE NAME']);
    }

    // Check if payslip already exists for this employee/week
    // Calculate week ending as the Friday of the import week
    const importDate = timesheetRecord['IMPORTED_DATE'] ? new Date(timesheetRecord['IMPORTED_DATE']) : new Date();
    const dayOfWeek = importDate.getDay();
    // Calculate days to get to Friday (day 5)
    // If today is Friday (5), use today. Otherwise find the Friday of this week.
    let daysToFriday = 5 - dayOfWeek;
    if (daysToFriday < 0) {
      daysToFriday += 7; // Go to next Friday if we're past Friday
    }
    const weekEndingDate = new Date(importDate);
    weekEndingDate.setDate(importDate.getDate() + daysToFriday);

    // Normalize to date-only format (strip time components)
    const weekEnding = normalizeToDateOnly(weekEndingDate);
    const duplicatePayslip = checkDuplicatePayslip(timesheetRecord['EMPLOYEE NAME'], weekEnding);

    if (duplicatePayslip) {
      throw new Error('Payslip already exists for this employee and week');
    }

    // Create payslip data
    const payslipData = {
      employeeName: timesheetRecord['EMPLOYEE NAME'],
      weekEnding: weekEnding,
      hours: timesheetRecord['HOURS'] || 0,
      minutes: timesheetRecord['MINUTES'] || 0,
      overtimeHours: timesheetRecord['OVERTIMEHOURS'] || 0,
      overtimeMinutes: timesheetRecord['OVERTIMEMINUTES'] || 0,
      notes: timesheetRecord['NOTES'] || '',
      // Default values for other fields
      leavePay: 0,
      bonusPay: 0,
      otherIncome: 0,
      otherDeductions: 0,
      otherDeductionsText: '',
      loanDeductionThisWeek: 0,
      newLoanThisWeek: 0,
      loanDisbursementType: 'Separate'
    };

    // Create payslip
    Logger.log('📝 Creating payslip from timesheet...');
    const payslipResult = createPayslip(payslipData);

    if (!payslipResult.success) {
      throw new Error('Failed to create payslip: ' + payslipResult.error);
    }

    Logger.log('✅ Payslip created: #' + payslipResult.data.recordNumber);

    // Update pending timesheet status
    const statusCol = findColumnIndex(headers, 'STATUS');
    const reviewedByCol = findColumnIndex(headers, 'REVIEWED_BY');
    const reviewedDateCol = findColumnIndex(headers, 'REVIEWED_DATE');

    pendingSheet.getRange(rowIndex, statusCol + 1).setValue('Approved');
    pendingSheet.getRange(rowIndex, reviewedByCol + 1).setValue(getCurrentUser());
    pendingSheet.getRange(rowIndex, reviewedDateCol + 1).setValue(normalizeToDateOnly(new Date()));

    SpreadsheetApp.flush();

    Logger.log('✅ Timesheet approved');
    Logger.log('========== APPROVE TIMESHEET COMPLETE ==========\n');

    return {
      success: true,
      data: {
        timesheetId: id,
        payslipRecordNumber: payslipResult.data.recordNumber
      }
    };

  } catch (error) {
    Logger.log('❌ ERROR in approveTimesheet: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== APPROVE TIMESHEET FAILED ==========\n');
    return { success: false, error: error.message };
  }
}

// ==================== REJECT TIMESHEET ====================

/**
 * Rejects a pending timesheet
 *
 * @param {string} id - Timesheet ID
 * @param {string} [reason] - Rejection reason (optional)
 * @returns {Object} Result with success flag
 */
function rejectTimesheet(id, reason) {
  try {
    Logger.log('\n========== REJECT TIMESHEET ==========');
    Logger.log('ID: ' + id);
    Logger.log('Reason: ' + (reason || 'Not specified'));

    const sheets = getSheets();
    const pendingSheet = sheets.pending;

    if (!pendingSheet) {
      throw new Error('PendingTimesheets sheet not found');
    }

    const sheetData = pendingSheet.getDataRange().getValues();
    const headers = sheetData[0];

    const idCol = findColumnIndex(headers, 'ID');
    const statusCol = findColumnIndex(headers, 'STATUS');
    const reviewedByCol = findColumnIndex(headers, 'REVIEWED_BY');
    const reviewedDateCol = findColumnIndex(headers, 'REVIEWED_DATE');
    const notesCol = findColumnIndex(headers, 'NOTES');

    // Find the record
    let rowIndex = -1;
    for (let i = 1; i < sheetData.length; i++) {
      if (sheetData[i][idCol] === id) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error('Pending timesheet not found: ' + id);
    }

    // Update status
    pendingSheet.getRange(rowIndex, statusCol + 1).setValue('Rejected');
    pendingSheet.getRange(rowIndex, reviewedByCol + 1).setValue(getCurrentUser());
    pendingSheet.getRange(rowIndex, reviewedDateCol + 1).setValue(normalizeToDateOnly(new Date()));

    if (reason) {
      const currentNotes = sheetData[rowIndex - 1][notesCol] || '';
      const updatedNotes = currentNotes + '\n[REJECTED: ' + reason + ']';
      pendingSheet.getRange(rowIndex, notesCol + 1).setValue(updatedNotes);
    }

    SpreadsheetApp.flush();

    Logger.log('✅ Timesheet rejected');
    Logger.log('========== REJECT TIMESHEET COMPLETE ==========\n');

    return { success: true, data: 'Timesheet rejected' };

  } catch (error) {
    Logger.log('❌ ERROR in rejectTimesheet: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== REJECT TIMESHEET FAILED ==========\n');
    return { success: false, error: error.message };
  }
}

// ==================== LIST PENDING TIMESHEETS ====================

/**
 * Lists pending timesheets with optional filters
 *
 * @param {Object} [filters] - Filter options
 * @param {string} [filters.status] - Filter by status (Pending/Approved/Rejected)
 * @param {string} [filters.employeeName] - Filter by employee name
 * @param {Date|string} [filters.weekEnding] - Filter by week ending
 *
 * @returns {Object} Result with success flag and timesheet records
 */
function listPendingTimesheets(filters) {
  try {
    Logger.log('\n========== LIST PENDING TIMESHEETS ==========');
    Logger.log('Filters: ' + JSON.stringify(filters || {}));

    const sheets = getSheets();
    const pendingSheet = sheets.pending;

    if (!pendingSheet) {
      throw new Error('PendingTimesheets sheet not found');
    }

    Logger.log('📋 Using sheet: ' + pendingSheet.getName());

    const data = pendingSheet.getDataRange().getValues();
    Logger.log('📊 Total rows in sheet: ' + data.length);

    if (data.length === 0) {
      Logger.log('⚠️ Sheet is completely empty!');
      return { success: true, data: [] };
    }

    const headers = data[0];
    Logger.log('📋 Headers (' + headers.length + '): ' + headers.join(', '));

    // Convert all rows to objects
    let records = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // Skip completely empty rows
      if (row.every(cell => !cell || cell === '')) {
        Logger.log('  Row ' + (i + 1) + ': Empty, skipping');
        continue;
      }
      records.push(buildObjectFromRow(row, headers));
    }

    Logger.log('📊 Total records before filtering: ' + records.length);

    if (records.length > 0) {
      Logger.log('📋 Sample record (first): ' + JSON.stringify(records[0]));
    }

    // Apply filters if provided
    if (filters) {
      if (filters.status) {
        Logger.log('🔍 Filtering by STATUS: ' + filters.status);
        const beforeCount = records.length;
        records = records.filter(r => {
          const match = r['STATUS'] === filters.status;
          if (!match && beforeCount < 5) {
            Logger.log('  Record STATUS "' + r['STATUS'] + '" does not match "' + filters.status + '"');
          }
          return match;
        });
        Logger.log('  After STATUS filter: ' + records.length + ' records (was ' + beforeCount + ')');
      }

      if (filters.employeeName) {
        records = records.filter(r => r['EMPLOYEE NAME'] === filters.employeeName);  // ✅ Fixed: Use bracket notation
      }

      if (filters.weekEnding) {
        const filterWeekEnding = parseDate(filters.weekEnding);
        records = records.filter(r => {
          // Skip records with empty WEEKENDING
          if (!r.WEEKENDING) return false;
          const recordWeekEnding = parseDate(r.WEEKENDING);  // ✅ Fixed: WEEKENDING not WEEK_ENDING
          return recordWeekEnding.getTime() === filterWeekEnding.getTime();
        });
      }
    }

    // Sort by week ending (descending), then employee name
    records.sort((a, b) => {
      // Handle empty WEEKENDING values - treat as oldest date
      const dateA = a.WEEKENDING ? parseDate(a.WEEKENDING) : new Date(0);
      const dateB = b.WEEKENDING ? parseDate(b.WEEKENDING) : new Date(0);

      if (dateA.getTime() !== dateB.getTime()) {
        return dateB - dateA;  // Descending
      }

      return a['EMPLOYEE NAME'].localeCompare(b['EMPLOYEE NAME']);  // ✅ Fixed: Use bracket notation for space
    });

    Logger.log('✅ Found ' + records.length + ' pending timesheets');
    Logger.log('========== LIST PENDING TIMESHEETS COMPLETE ==========\n');

    // Sanitize records for web serialization (dates to DD-MM-YYYY format)
    const sanitizedRecords = records.map(function(record) {
      return sanitizeForWeb(record);
    });

    return { success: true, data: sanitizedRecords };

  } catch (error) {
    Logger.log('❌ ERROR in listPendingTimesheets: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== LIST PENDING TIMESHEETS FAILED ==========\n');
    return { success: false, error: error.message };
  }
}

// ==================== VALIDATION ====================

/**
 * Validates timesheet data
 *
 * @param {Object} data - Timesheet data to validate
 * @throws {Error} If validation fails
 */
function validateTimesheet(data) {
  const errors = [];

  // Required fields
  if (!data.employeeName) {
    errors.push('Employee Name is required');
  }

  if (!data.weekEnding) {
    errors.push('Week Ending is required');
  }

  if (data.hours === undefined || data.hours === null) {
    errors.push('Hours is required');
  }

  if (data.minutes === undefined || data.minutes === null) {
    errors.push('Minutes is required');
  }

  if (data.overtimeHours === undefined || data.overtimeHours === null) {
    errors.push('Overtime Hours is required');
  }

  if (data.overtimeMinutes === undefined || data.overtimeMinutes === null) {
    errors.push('Overtime Minutes is required');
  }

  // Validate time values
  if (data.hours < 0 || data.minutes < 0 || data.overtimeHours < 0 || data.overtimeMinutes < 0) {
    errors.push('Time values cannot be negative');
  }

  if (data.minutes >= 60) {
    errors.push('Minutes must be less than 60');
  }

  if (data.overtimeMinutes >= 60) {
    errors.push('Overtime Minutes must be less than 60');
  }

  // Verify employee exists
  if (data.employeeName) {
    const empResult = getEmployeeByName(data.employeeName);
    if (!empResult.success) {
      errors.push('Employee not found: ' + data.employeeName);
    }
  }

  if (errors.length > 0) {
    throw new Error(errors.join('; '));
  }
}

/**
 * Checks if pending timesheet already exists for employee/week
 *
 * @param {string} employeeName - Employee REFNAME
 * @param {Date|string} weekEnding - Week ending date
 * @returns {boolean} True if duplicate exists
 */
function checkDuplicatePendingTimesheet(employeeName, weekEnding) {
  try {
    const sheets = getSheets();
    const pendingSheet = sheets.pending;

    if (!pendingSheet) return false;

    const data = pendingSheet.getDataRange().getValues();
    const headers = data[0];

    const empNameCol = findColumnIndex(headers, 'EMPLOYEE NAME');
    const weekEndingCol = findColumnIndex(headers, 'WEEKENDING');
    const statusCol = findColumnIndex(headers, 'STATUS');

    const searchDate = parseDate(weekEnding);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowEmpName = row[empNameCol];
      const rowWeekEnding = parseDate(row[weekEndingCol]);
      const rowStatus = row[statusCol];

      // Only check pending records
      if (rowStatus === 'Pending' &&
          rowEmpName === employeeName &&
          rowWeekEnding.getTime() === searchDate.getTime()) {
        return true;
      }
    }

    return false;

  } catch (error) {
    Logger.log('⚠️ Error checking duplicate: ' + error.message);
    return false;
  }
}

// ==================== CLOCK-IN IMPORT FUNCTIONS ====================

/**
 * Import clock-in data from base64 encoded file (client-side wrapper)
 * Converts base64 data to blob and calls importClockData
 *
 * @param {string} base64Data - Base64 encoded file data
 * @param {string} filename - Original filename
 * @param {string} mimeType - File MIME type
 * @param {boolean} override - Override warnings and import anyway
 * @returns {Object} Result with success flag and import details
 */
function importClockDataFromBase64(base64Data, filename, mimeType, override) {
  try {
    Logger.log('\n========== IMPORT CLOCK DATA FROM BASE64 ==========');
    Logger.log('📥 Function called with:');
    Logger.log('   Filename: ' + filename);
    Logger.log('   MIME Type: ' + mimeType);
    Logger.log('   Override: ' + override);
    Logger.log('   Base64 length: ' + (base64Data ? base64Data.length : 'NULL'));

    // Validate inputs
    if (!base64Data) {
      Logger.log('❌ ERROR: base64Data is null or undefined');
      return { success: false, error: 'No file data provided' };
    }

    if (!filename) {
      Logger.log('❌ ERROR: filename is null or undefined');
      return { success: false, error: 'No filename provided' };
    }

    Logger.log('✅ Input validation passed');
    Logger.log('🔄 Decoding base64...');

    // Decode base64 and create blob (server-side Utilities API)
    const bytes = Utilities.base64Decode(base64Data);
    Logger.log('✅ Base64 decoded, bytes length: ' + bytes.length);

    Logger.log('🔄 Creating blob...');
    const fileBlob = Utilities.newBlob(bytes, mimeType, filename);
    Logger.log('✅ Blob created: ' + fileBlob.getName());

    // Call the main import function
    Logger.log('🔄 Calling importClockData...');
    const result = importClockData(fileBlob, filename, override);

    Logger.log('✅ importClockData returned');
    Logger.log('   Result type: ' + typeof result);
    Logger.log('   Result is null: ' + (result === null));
    Logger.log('   Result is undefined: ' + (result === undefined));

    if (result) {
      Logger.log('   Result.success: ' + result.success);
      Logger.log('   Result has error: ' + (!!result.error));
      Logger.log('   Result has data: ' + (!!result.data));
    }

    Logger.log('========== IMPORT CLOCK DATA FROM BASE64 COMPLETE ==========\n');
    return result;

  } catch (error) {
    // Extract detailed error information
    const errorMessage = error.message || error.toString() || 'Unknown error occurred during import';
    Logger.log('❌ ERROR in importClockDataFromBase64: ' + errorMessage);
    if (error.stack) Logger.log('   Stack: ' + error.stack);
    Logger.log('========== IMPORT CLOCK DATA FROM BASE64 FAILED ==========\n');
    return { success: false, error: errorMessage };
  }
}

/**
 * Import clock-in data from Excel file
 * This function processes raw clock-in data and validates against employee records
 *
 * @param {Blob} fileBlob - Excel file blob
 * @param {string} filename - Original filename
 * @param {boolean} override - Override warnings and import anyway
 * @returns {Object} Result with success flag and import details
 */
function importClockData(fileBlob, filename, override) {
  try {
    Logger.log('\n========== IMPORT CLOCK DATA ==========');
    Logger.log('Filename: ' + filename);
    Logger.log('Override: ' + (override ? 'YES' : 'NO'));

    // Parse Excel file (v2 - cache busting rename)
    const parseResult = parseClockDataExcel_v2(fileBlob);
    if (!parseResult.success) {
      throw new Error('Failed to parse Excel: ' + parseResult.error);
    }

    const clockRecords = parseResult.data;
    Logger.log('📊 Parsed ' + clockRecords.length + ' clock records');

    // Generate file hash for duplicate detection
    const fileHash = generateImportHash(JSON.stringify(clockRecords));
    Logger.log('🔑 File hash: ' + fileHash);

    // Determine week ending
    const weekEnding = determineWeekEnding(clockRecords);
    Logger.log('📅 Week ending: ' + formatDate(weekEnding));

    // Check for duplicate import
    const dupCheck = checkDuplicateImport(weekEnding, fileHash);
    if (dupCheck.isDuplicate && !override) {
      Logger.log('⚠️ Duplicate import detected');
      return {
        success: false,
        error: 'DUPLICATE_IMPORT',
        data: {
          existingImportId: dupCheck.existingImportId,
          existingImportDate: dupCheck.existingImportDate,
          canReplace: true
        }
      };
    }

    // Extract unique clock references
    const clockRefs = getUniqueValues(clockRecords.map(function(r) {
      return String(r.ClockInRef).trim();
    }));

    // Validate against employee table
    const validation = validateEmployeeClockRefs(clockRefs);
    if (!validation.success) {
      throw new Error('Validation failed: ' + validation.error);
    }

    const unmatchedRefs = validation.data.unmatched;

    // If there are unmatched refs and no override, return warning
    if (unmatchedRefs.length > 0 && !override) {
      Logger.log('⚠️ Found ' + unmatchedRefs.length + ' unmatched clock references');
      return {
        success: false,
        error: 'UNMATCHED_EMPLOYEES',
        data: {
          unmatchedRefs: unmatchedRefs,
          unmatchedDetails: buildUnmatchedDetails(clockRecords, unmatchedRefs),
          totalRecords: clockRecords.length,
          matchedCount: validation.data.matched.length,
          canOverride: true
        }
      };
    }

    // Create import record
    const importId = generateFullUUID();
    const importResult = createImportRecord({
      importId: importId,
      filename: filename,
      fileHash: fileHash,
      weekEnding: weekEnding,
      totalRecords: clockRecords.length,
      matchedEmployees: validation.data.matched.length,
      unmatchedRefs: unmatchedRefs,
      overrideApplied: override
    });

    if (!importResult.success) {
      throw new Error('Failed to create import record: ' + importResult.error);
    }

    // If duplicate, mark old import as replaced
    if (dupCheck.isDuplicate && override) {
      replaceImport(dupCheck.existingImportId, importId);
    }

    // Store raw clock data
    const storeResult = storeRawClockData(clockRecords, importId);
    if (!storeResult.success) {
      throw new Error('Failed to store clock data: ' + storeResult.error);
    }

    // Process raw data and create pending timesheets
    const processResult = processRawClockData(importId);
    if (!processResult.success) {
      Logger.log('⚠️ Warning: Processing completed with errors: ' + processResult.error);
    }

    Logger.log('✅ Import complete: ' + clockRecords.length + ' records imported');
    Logger.log('========== IMPORT CLOCK DATA COMPLETE ==========\n');

    return {
      success: true,
      data: {
        importId: importId,
        totalRecords: clockRecords.length,
        matchedEmployees: validation.data.matched.length,
        unmatchedRefs: unmatchedRefs,
        timesheetsCreated: processResult.data ? processResult.data.timesheetsCreated : 0
      }
    };

  } catch (error) {
    // Extract detailed error information
    const errorMessage = error.message || error.toString() || 'Unknown error occurred during import';
    Logger.log('❌ ERROR in importClockData: ' + errorMessage);
    if (error.stack) Logger.log('Stack: ' + error.stack);
    Logger.log('========== IMPORT CLOCK DATA FAILED ==========\n');
    return {
      success: false,
      error: errorMessage
    };
  }
}

/**
 * Retry a function with exponential backoff for rate limit errors
 *
 * @param {Function} fn - Function to retry
 * @param {number} maxRetries - Maximum number of retries (default: 3)
 * @param {number} initialDelay - Initial delay in milliseconds (default: 1000)
 * @returns {*} Result of the function
 */
function retryWithBackoff(fn, maxRetries, initialDelay) {
  maxRetries = maxRetries || 3;
  initialDelay = initialDelay || 1000;

  let lastError;
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      return fn();
    } catch (error) {
      lastError = error;

      // Check if it's a rate limit error (429) or quota error
      const errorMsg = error.message || '';
      const isRateLimit = errorMsg.indexOf('429') >= 0 ||
                         errorMsg.toLowerCase().indexOf('rate limit') >= 0 ||
                         errorMsg.toLowerCase().indexOf('quota') >= 0 ||
                         errorMsg.toLowerCase().indexOf('too many requests') >= 0;

      // If it's the last attempt or not a rate limit error, throw
      if (attempt === maxRetries || !isRateLimit) {
        throw error;
      }

      // Calculate delay with exponential backoff
      const delay = initialDelay * Math.pow(2, attempt);
      Logger.log('⚠️ Rate limit error (attempt ' + (attempt + 1) + '/' + maxRetries + '): ' + errorMsg);
      Logger.log('⏳ Waiting ' + delay + 'ms before retry...');

      Utilities.sleep(delay);
    }
  }

  throw lastError;
}

/**
 * Parse clock data from Excel file - VERSION 2 (Cache-busting rename)
 *
 * @param {Blob} fileBlob - Excel file blob
 * @returns {Object} Result with parsed data
 */
function parseClockDataExcel_v2(fileBlob) {
  try {
    Logger.log('🔴 CODE VERSION: 2025-11-19-v6 WITH 10-HOUR CORRECTION (CACHE BUSTED) 🔴');
    Logger.log('📖 Parsing clock data Excel file...');
    Logger.log('📄 File name: ' + fileBlob.getName());
    Logger.log('📦 File type: ' + fileBlob.getContentType());

    // Convert Excel to Google Sheets using Drive Advanced Service
    const file = DriveApp.createFile(fileBlob);
    const fileId = file.getId();
    Logger.log('📁 Created temp file with ID: ' + fileId);

    let data;
    let convertedFileId;

    try {
      // Use Drive Advanced Service (v3) - now enabled in appsscript.json
      Logger.log('🔄 Converting Excel to Google Sheets using Drive API v3...');

      const resource = {
        name: fileBlob.getName() + '_converted',
        mimeType: MimeType.GOOGLE_SHEETS
      };

      // Wrap Drive API call with retry logic for rate limit errors
      const importedFile = retryWithBackoff(function() {
        return Drive.Files.copy(resource, fileId);
      }, 3, 2000); // 3 retries, starting with 2 second delay

      convertedFileId = importedFile.id;

      Logger.log('✅ Conversion successful! Converted file ID: ' + convertedFileId);

      // Open the converted spreadsheet
      const spreadsheet = SpreadsheetApp.openById(convertedFileId);
      const sheet = spreadsheet.getSheets()[0];

      // DEBUG: Get timezone information
      const scriptTimezone = Session.getScriptTimeZone();
      const spreadsheetTimezone = spreadsheet.getSpreadsheetTimeZone();
      Logger.log('\n========== TIMEZONE DEBUG INFO ==========');
      Logger.log('🌍 Script timezone: ' + scriptTimezone);
      Logger.log('🌍 Spreadsheet timezone: ' + spreadsheetTimezone);
      Logger.log('🌍 Current server time: ' + new Date().toString());
      Logger.log('🌍 Server UTC time: ' + new Date().toUTCString());

      // Get timezone offset in hours
      const tzOffset = new Date().getTimezoneOffset();
      Logger.log('🌍 Timezone offset: ' + tzOffset + ' minutes (' + (tzOffset / -60) + ' hours from UTC)');

      // Get both raw values and display values for comparison
      const range = sheet.getDataRange();
      data = range.getValues(); // Get Date objects
      const displayData = range.getDisplayValues(); // Get formatted strings

      Logger.log('📊 Retrieved ' + data.length + ' rows from converted sheet');
      Logger.log('========================================\n');

      // Clean up temporary files
      Logger.log('🗑️ Cleaning up temporary files...');
      DriveApp.getFileById(fileId).setTrashed(true);
      DriveApp.getFileById(convertedFileId).setTrashed(true);
      Logger.log('✅ Cleanup complete');

    } catch (conversionError) {
      // Extract detailed error information
      const errorMessage = conversionError.message || conversionError.toString() || 'Unknown error';
      const errorDetails = conversionError.details || '';
      const errorStack = conversionError.stack || '';

      Logger.log('❌ ERROR during Excel conversion: ' + errorMessage);
      if (errorDetails) Logger.log('Error details: ' + errorDetails);
      if (errorStack) Logger.log('Stack: ' + errorStack);

      // Clean up on error
      try {
        Logger.log('🗑️ Cleaning up temp file on error...');
        DriveApp.getFileById(fileId).setTrashed(true);
      } catch (e) {
        Logger.log('⚠️ Could not clean up temp file: ' + e.message);
      }

      // Provide user-friendly error messages
      let userMessage = 'Excel conversion failed: ' + errorMessage;
      if (errorMessage.indexOf('429') >= 0 || errorMessage.toLowerCase().indexOf('rate limit') >= 0) {
        userMessage = 'Google API rate limit exceeded. Please wait a few moments and try again.';
      } else if (errorMessage.toLowerCase().indexOf('quota') >= 0) {
        userMessage = 'Google API quota exceeded. Please try again later.';
      }

      throw new Error(userMessage);
    }

    // Actual clock-in system export format:
    // Row 1: "Transaction Reports" (title header - skip)
    // Row 2: Column headers - Person ID, Person Name, Department, Type, Source, Punch Time, Time Zone, Verification Mode, Mobile Punch, Device SN, Device Name, Upload Time
    // Row 3+: Data rows

    // Determine starting row (skip title row if present)
    let headerRow = 0;
    let dataStartRow = 1;

    // Check if first row is a title (single cell or "Transaction Reports")
    if (data.length > 0 && (String(data[0][0]).toLowerCase().indexOf('transaction') >= 0 || data[0].length === 1)) {
      headerRow = 1;
      dataStartRow = 2;
      Logger.log('📋 Detected title row, using row 2 as headers');
    }

    const headers = data[headerRow];
    Logger.log('📋 Headers: ' + headers.join(', '));

    // Find column indices (case-insensitive)
    const colMap = {};
    for (let i = 0; i < headers.length; i++) {
      const header = String(headers[i]).trim().toLowerCase();
      if (header.indexOf('person id') >= 0) colMap.personId = i;
      else if (header.indexOf('person name') >= 0) colMap.personName = i;
      else if (header.indexOf('department') >= 0) colMap.department = i;
      else if (header.indexOf('punch time') >= 0) colMap.punchTime = i;
      else if (header.indexOf('device name') >= 0) colMap.deviceName = i;
      else if (header.indexOf('device sn') >= 0) colMap.deviceSn = i;
      else if (header.indexOf('type') >= 0) colMap.type = i;
      else if (header.indexOf('source') >= 0) colMap.source = i;
    }

    Logger.log('📋 Column mapping: ' + JSON.stringify(colMap));

    const records = [];

    for (let i = dataStartRow; i < data.length; i++) {
      const row = data[i];

      // Skip empty rows (check Person ID column)
      if (!row[colMap.personId] || String(row[colMap.personId]).trim() === '') continue;

      // Parse Punch Time - CRITICAL: Use local components to preserve Excel time
      // Google Sheets stores Excel datetimes as Date objects
      // We extract local (not UTC) components to preserve the exact time shown in Excel
      const punchTimeValue = row[colMap.punchTime];

      if (!punchTimeValue) {
        Logger.log('⚠️ Row ' + (i + 1) + ': Empty punch time, skipping');
        continue;
      }

      let punchDateTime;

      try {
        if (punchTimeValue instanceof Date && !isNaN(punchTimeValue.getTime())) {
          // DEBUG: Detailed logging for first 5 records
          if (i < dataStartRow + 5) {
            Logger.log('\n========== ROW ' + (i + 1) + ' PUNCH TIME DEBUG ==========');
            Logger.log('📥 RAW VALUE FROM GOOGLE SHEETS:');
            Logger.log('   Type: ' + typeof punchTimeValue);
            Logger.log('   Is Date: ' + (punchTimeValue instanceof Date));
            Logger.log('   Display value: ' + displayData[i][colMap.punchTime]);
            Logger.log('   .toString(): ' + punchTimeValue.toString());
            Logger.log('   .toISOString(): ' + punchTimeValue.toISOString());
            Logger.log('   .toUTCString(): ' + punchTimeValue.toUTCString());
            Logger.log('   .toLocaleString(): ' + punchTimeValue.toLocaleString());
            Logger.log('   .getTime(): ' + punchTimeValue.getTime());

            Logger.log('\n🔍 LOCAL COMPONENTS (what we use):');
            Logger.log('   getFullYear(): ' + punchTimeValue.getFullYear());
            Logger.log('   getMonth(): ' + punchTimeValue.getMonth() + ' (0-indexed, so +1 = ' + (punchTimeValue.getMonth() + 1) + ')');
            Logger.log('   getDate(): ' + punchTimeValue.getDate());
            Logger.log('   getHours(): ' + punchTimeValue.getHours());
            Logger.log('   getMinutes(): ' + punchTimeValue.getMinutes());
            Logger.log('   getSeconds(): ' + punchTimeValue.getSeconds());
            Logger.log('   => Extracted time: ' + punchTimeValue.getHours() + ':' +
                      punchTimeValue.getMinutes() + ':' + punchTimeValue.getSeconds());

            Logger.log('\n🌍 UTC COMPONENTS (for comparison):');
            Logger.log('   getUTCFullYear(): ' + punchTimeValue.getUTCFullYear());
            Logger.log('   getUTCMonth(): ' + punchTimeValue.getUTCMonth() + ' (0-indexed, so +1 = ' + (punchTimeValue.getUTCMonth() + 1) + ')');
            Logger.log('   getUTCDate(): ' + punchTimeValue.getUTCDate());
            Logger.log('   getUTCHours(): ' + punchTimeValue.getUTCHours());
            Logger.log('   getUTCMinutes(): ' + punchTimeValue.getUTCMinutes());
            Logger.log('   getUTCSeconds(): ' + punchTimeValue.getUTCSeconds());
            Logger.log('   => UTC time: ' + punchTimeValue.getUTCHours() + ':' +
                      punchTimeValue.getUTCMinutes() + ':' + punchTimeValue.getUTCSeconds());
          }

          // CRITICAL FIX: Google Sheets Excel conversion applies incorrect timezone
          // The diagnostic showed: getHours() = 23 but should be 13 (10-hour offset)
          // We need to subtract 10 hours to get the correct South Africa time

          const TIMEZONE_CORRECTION_HOURS = 10; // Adjust if needed based on your timezone

          // Create Date with current (wrong) time
          let tempDateTime = new Date(
            punchTimeValue.getFullYear(),
            punchTimeValue.getMonth(),
            punchTimeValue.getDate(),
            punchTimeValue.getHours(),
            punchTimeValue.getMinutes(),
            punchTimeValue.getSeconds()
          );

          // Subtract 10 hours to get correct time
          punchDateTime = new Date(tempDateTime.getTime() - (TIMEZONE_CORRECTION_HOURS * 60 * 60 * 1000));

          // DEBUG: Log correction
          if (i < dataStartRow + 5) {
            Logger.log('\n🔧 TIMEZONE CORRECTION APPLIED:');
            Logger.log('   Before correction: ' + tempDateTime.getHours() + ':' + tempDateTime.getMinutes() + ':' + tempDateTime.getSeconds());
            Logger.log('   After correction:  ' + punchDateTime.getHours() + ':' + punchDateTime.getMinutes() + ':' + punchDateTime.getSeconds());
            Logger.log('   Corrected -' + TIMEZONE_CORRECTION_HOURS + ' hours');
          }

          // DEBUG: Log created date
          if (i < dataStartRow + 5) {
            Logger.log('\n✅ CREATED DATE OBJECT:');
            Logger.log('   .toString(): ' + punchDateTime.toString());
            Logger.log('   .toISOString(): ' + punchDateTime.toISOString());
            Logger.log('   .toLocaleString(): ' + punchDateTime.toLocaleString());
            Logger.log('   Time components: ' + punchDateTime.getHours() + ':' +
                      punchDateTime.getMinutes() + ':' + punchDateTime.getSeconds());
            Logger.log('========================================\n');
          }
        } else if (typeof punchTimeValue === 'string' && punchTimeValue.trim()) {
          // Fallback for string values: parse manually
          const punchTimeStr = punchTimeValue.trim();
          const parts = punchTimeStr.split(' ');

          if (parts.length >= 2) {
            const datePart = parts[0];
            const timePart = parts[1];

            let year, month, day;
            if (datePart.indexOf('-') >= 0) {
              const dateParts = datePart.split('-');
              year = parseInt(dateParts[0]);
              month = parseInt(dateParts[1]) - 1;
              day = parseInt(dateParts[2]);
            } else if (datePart.indexOf('/') >= 0) {
              const dateParts = datePart.split('/');
              if (parseInt(dateParts[0]) > 31) {
                year = parseInt(dateParts[0]);
                month = parseInt(dateParts[1]) - 1;
                day = parseInt(dateParts[2]);
              } else {
                day = parseInt(dateParts[0]);
                month = parseInt(dateParts[1]) - 1;
                year = parseInt(dateParts[2]);
              }
            } else {
              Logger.log('⚠️ Row ' + (i + 1) + ': Unrecognized date format: ' + datePart);
              continue;
            }

            const timeParts = timePart.split(':');
            const hours = parseInt(timeParts[0]);
            const minutes = parseInt(timeParts[1]);
            const seconds = parseInt(timeParts[2] || 0);

            punchDateTime = new Date(year, month, day, hours, minutes, seconds);
          } else {
            Logger.log('⚠️ Row ' + (i + 1) + ': Invalid punch time format: ' + punchTimeStr);
            continue;
          }
        } else {
          Logger.log('⚠️ Row ' + (i + 1) + ': Invalid punch time value, skipping');
          continue;
        }

        // Validate the parsed date
        if (!punchDateTime || isNaN(punchDateTime.getTime())) {
          Logger.log('⚠️ Row ' + (i + 1) + ': Failed to parse punch time');
          continue;
        }

      } catch (parseError) {
        Logger.log('⚠️ Row ' + (i + 1) + ': Error parsing punch time: ' + parseError.message);
        continue;
      }

      // Extract date and time components
      const punchDate = new Date(punchDateTime.getFullYear(), punchDateTime.getMonth(), punchDateTime.getDate());

      const record = {
        ClockInRef: String(row[colMap.personId] || '').trim(),
        EmployeeName: String(row[colMap.personName] || '').trim(),
        Department: String(row[colMap.department] || '').trim(),
        PunchDate: punchDate,
        PunchTime: punchDateTime,
        DeviceName: String(row[colMap.deviceName] || '').trim(),
        DeviceSN: colMap.deviceSn !== undefined ? String(row[colMap.deviceSn] || '').trim() : '',
        Type: colMap.type !== undefined ? String(row[colMap.type] || '').trim() : '',
        Source: colMap.source !== undefined ? String(row[colMap.source] || '').trim() : ''
      };

      records.push(record);
    }

    Logger.log('✅ Parsed ' + records.length + ' clock records');
    return { success: true, data: records };

  } catch (error) {
    // Extract detailed error information
    const errorMessage = error.message || error.toString() || 'Unknown error occurred';
    Logger.log('❌ ERROR in parseClockDataExcel: ' + errorMessage);
    if (error.stack) Logger.log('Stack: ' + error.stack);
    return { success: false, error: errorMessage };
  }
}

/**
 * Determine week ending from clock records
 *
 * @param {Array} records - Clock records
 * @returns {Date} Week ending date
 */
function determineWeekEnding(records) {
  if (records.length === 0) {
    throw new Error('No records to determine week ending');
  }

  // Use WeekEnding from first record if available
  if (records[0].WeekEnding) {
    return records[0].WeekEnding;
  }

  // Otherwise, calculate from punch dates
  const dates = records.map(function(r) {
    return r.PunchDate;
  });

  // Find latest date
  const latestDate = new Date(Math.max.apply(null, dates));

  // Get week ending (Saturday)
  return getWeekEnding(latestDate);
}

/**
 * Validate employee clock references against employee table
 *
 * @param {Array} clockRefs - Array of clock reference numbers
 * @returns {Object} Validation result
 */
function validateEmployeeClockRefs(clockRefs) {
  return validateClockRefs(clockRefs);
}

/**
 * Build unmatched details for reporting
 *
 * @param {Array} clockRecords - All clock records
 * @param {Array} unmatchedRefs - Unmatched clock references
 * @returns {Array} Unmatched details with transaction counts
 */
function buildUnmatchedDetails(clockRecords, unmatchedRefs) {
  const details = [];

  unmatchedRefs.forEach(function(ref) {
    const count = clockRecords.filter(function(r) {
      return String(r.ClockInRef).trim() === ref;
    }).length;

    const employeeName = clockRecords.find(function(r) {
      return String(r.ClockInRef).trim() === ref;
    }).EmployeeName;

    details.push({
      clockRef: ref,
      employeeName: employeeName || 'Unknown-' + ref,
      transactionCount: count
    });
  });

  return details;
}

/**
 * Create import record in CLOCK_IN_IMPORTS sheet
 *
 * @param {Object} data - Import data
 * @returns {Object} Result
 */
function createImportRecord(data) {
  try {
    const sheets = getSheets();
    const importSheet = sheets.clockImports;

    if (!importSheet) {
      throw new Error('CLOCK_IN_IMPORTS sheet not found');
    }

    const rowData = [
      data.importId,
      new Date(),
      data.weekEnding,
      data.filename,
      data.fileHash,
      data.totalRecords,
      data.matchedEmployees,
      JSON.stringify(data.unmatchedRefs),
      'Active',
      getCurrentUser(),
      data.overrideApplied || false,
      '',
      '',
      data.notes || ''
    ];

    importSheet.appendRow(rowData);
    SpreadsheetApp.flush();

    Logger.log('✅ Import record created: ' + data.importId);
    return { success: true, data: { importId: data.importId } };

  } catch (error) {
    Logger.log('❌ ERROR in createImportRecord: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Store raw clock data in RAW_CLOCK_DATA sheet
 *
 * @param {Array} clockRecords - Clock records
 * @param {string} importId - Import ID
 * @returns {Object} Result
 */
function storeRawClockData(clockRecords, importId) {
  try {
    const sheets = getSheets();
    const rawDataSheet = sheets.rawClockData;

    if (!rawDataSheet) {
      throw new Error('RAW_CLOCK_DATA sheet not found');
    }

    // Resolve employee IDs from clock references
    const employeeMap = buildEmployeeClockRefMap();

    const rows = clockRecords.map(function(record, idx) {
      const empInfo = employeeMap[String(record.ClockInRef).trim()] || {};

      // DEBUG: Log first 3 records being stored
      if (idx < 3) {
        Logger.log('\n========== STORING RECORD ' + (idx + 1) + ' TO SHEET ==========');
        Logger.log('📋 Record.PunchTime:');
        Logger.log('   Type: ' + typeof record.PunchTime);
        Logger.log('   Is Date: ' + (record.PunchTime instanceof Date));
        if (record.PunchTime instanceof Date) {
          Logger.log('   .toString(): ' + record.PunchTime.toString());
          Logger.log('   .toISOString(): ' + record.PunchTime.toISOString());
          Logger.log('   .toLocaleString(): ' + record.PunchTime.toLocaleString());
          Logger.log('   Hours: ' + record.PunchTime.getHours());
          Logger.log('   Minutes: ' + record.PunchTime.getMinutes());
          Logger.log('   Seconds: ' + record.PunchTime.getSeconds());
        } else {
          Logger.log('   Value: ' + record.PunchTime);
        }
        Logger.log('========================================\n');
      }

      return [
        generateFullUUID(),                    // ID
        importId,                              // IMPORT_ID
        String(record.ClockInRef).trim(),      // EMPLOYEE_CLOCK_REF
        empInfo.name || record.EmployeeName,   // EMPLOYEE_NAME
        empInfo.id || '',                      // EMPLOYEE_ID
        record.WeekEnding,                     // WEEK_ENDING
        record.PunchDate,                      // PUNCH_DATE
        record.DeviceName,                     // DEVICE_NAME
        record.PunchTime,                      // PUNCH_TIME
        record.Department,                     // DEPARTMENT
        'Draft',                               // STATUS
        new Date(),                            // CREATED_DATE
        '',                                    // LOCKED_DATE
        ''                                     // LOCKED_BY
      ];
    });

    // Batch append for performance
    if (rows.length > 0) {
      const startRow = rawDataSheet.getLastRow() + 1;
      const range = rawDataSheet.getRange(
        startRow,
        1,
        rows.length,
        rows[0].length
      );
      range.setValues(rows);

      // Format PUNCH_TIME column (column 9) to show date AND time
      // Without this, Google Sheets may only show the date part
      const punchTimeCol = 9; // PUNCH_TIME is the 9th column in RAW_CLOCK_DATA_COLUMNS
      const punchTimeRange = rawDataSheet.getRange(startRow, punchTimeCol, rows.length, 1);
      punchTimeRange.setNumberFormat('yyyy-mm-dd hh:mm:ss');

      SpreadsheetApp.flush();

      // DEBUG: Read back the first 3 records to see what was actually stored
      Logger.log('\n========== VERIFICATION: READING BACK FROM SHEET ==========');
      const verifyRows = Math.min(3, rows.length);
      const storedData = rawDataSheet.getRange(startRow, punchTimeCol, verifyRows, 1).getValues();
      const storedDisplay = rawDataSheet.getRange(startRow, punchTimeCol, verifyRows, 1).getDisplayValues();

      for (let i = 0; i < verifyRows; i++) {
        Logger.log('\n📖 Stored Record ' + (i + 1) + ':');
        Logger.log('   Raw value (getValues): ' + storedData[i][0]);
        Logger.log('   Display value (getDisplayValues): ' + storedDisplay[i][0]);
        if (storedData[i][0] instanceof Date) {
          Logger.log('   Stored as Date object: ' + storedData[i][0].toString());
          Logger.log('   Hours: ' + storedData[i][0].getHours());
          Logger.log('   Minutes: ' + storedData[i][0].getMinutes());
        }
      }
      Logger.log('========================================\n');
    }

    Logger.log('✅ Stored ' + rows.length + ' raw clock records');
    return { success: true, data: { recordsStored: rows.length } };

  } catch (error) {
    Logger.log('❌ ERROR in storeRawClockData: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Build map of clock references to employee details
 *
 * @returns {Object} Map of clockRef => {id, name}
 */
function buildEmployeeClockRefMap() {
  const sheets = getSheets();
  const empSheet = sheets.empdetails;

  if (!empSheet) {
    return {};
  }

  const data = empSheet.getDataRange().getValues();
  const headers = data[0];

  const idCol = findColumnIndex(headers, 'id');
  const nameCol = findColumnIndex(headers, 'REFNAME');
  const clockRefCol = findColumnIndex(headers, 'ClockInRef');

  const map = {};

  for (let i = 1; i < data.length; i++) {
    const clockRef = String(data[i][clockRefCol] || '').trim();
    if (clockRef) {
      map[clockRef] = {
        id: data[i][idCol],
        name: data[i][nameCol]
      };
    }
  }

  return map;
}

/**
 * Process raw clock data and create pending timesheets
 *
 * @param {string} importId - Import ID to process
 * @returns {Object} Result
 */
function processRawClockData(importId) {
  try {
    Logger.log('\n========== PROCESS RAW CLOCK DATA ==========');
    Logger.log('Import ID: ' + importId);

    const sheets = getSheets();
    const rawDataSheet = sheets.rawClockData;

    if (!rawDataSheet) {
      throw new Error('RAW_CLOCK_DATA sheet not found');
    }

    // Get all raw data for this import
    const data = rawDataSheet.getDataRange().getValues();
    const headers = data[0];
    const importIdCol = findColumnIndex(headers, 'IMPORT_ID');

    const importRecords = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][importIdCol] === importId) {
        importRecords.push(buildObjectFromRow(data[i], headers));
      }
    }

    Logger.log('📊 Found ' + importRecords.length + ' records to process');

    // Group by employee
    const byEmployee = {};

    importRecords.forEach(function(record) {
      const empId = record.EMPLOYEE_ID;
      if (!empId) return;

      if (!byEmployee[empId]) {
        byEmployee[empId] = [];
      }
      byEmployee[empId].push(record);
    });

    // Process each employee
    const configResult = getTimeConfig();
    const config = configResult.success ? configResult.data : null;
    if (!config) {
      throw new Error('Failed to load time configuration');
    }
    let timesheetsCreated = 0;

    for (const empId in byEmployee) {
      const empRecords = byEmployee[empId];

      // Process clock data
      const processResult = processClockData(empRecords, config);
      if (!processResult.success) {
        Logger.log('⚠️ Failed to process for employee ' + empId + ': ' + processResult.error);
        continue;
      }

      const processed = processResult.data;

      // Create pending timesheet
      const timesheetData = {
        rawDataImportId: importId,
        employeeId: empId,
        employeeName: empRecords[0].EMPLOYEE_NAME,
        employeeClockRef: empRecords[0].EMPLOYEE_CLOCK_REF,
        weekEnding: empRecords[0].WEEK_ENDING,
        calculatedTotalHours: processed.calculatedTotalHours,
        calculatedTotalMinutes: processed.calculatedTotalMinutes,
        hours: processed.calculatedTotalHours,  // Default to calculated
        minutes: processed.calculatedTotalMinutes,
        overtimeHours: 0,
        overtimeMinutes: 0,
        lunchDeductionMin: processed.lunchDeductionMinutes,
        bathroomTimeMin: processed.bathroomTimeMinutes,
        reconDetails: JSON.stringify(processed),
        warnings: JSON.stringify(processed.warnings)
      };

      const createResult = createEnhancedPendingTimesheet(timesheetData);
      if (createResult.success) {
        timesheetsCreated++;
      }
    }

    Logger.log('✅ Created ' + timesheetsCreated + ' pending timesheets');
    Logger.log('========== PROCESS RAW CLOCK DATA COMPLETE ==========\n');

    return {
      success: true,
      data: { timesheetsCreated: timesheetsCreated }
    };

  } catch (error) {
    Logger.log('❌ ERROR in processRawClockData: ' + error.message);
    Logger.log('========== PROCESS RAW CLOCK DATA FAILED ==========\n');
    return { success: false, error: error.message };
  }
}

/**
 * Create enhanced pending timesheet with all new fields
 *
 * @param {Object} data - Timesheet data
 * @returns {Object} Result
 */
function createEnhancedPendingTimesheet(data) {
  try {
    const sheets = getSheets();
    const pendingSheet = sheets.pending;

    if (!pendingSheet) {
      throw new Error('PendingTimesheets sheet not found');
    }

    // Check for duplicate
    const duplicate = checkDuplicatePendingTimesheet(data.employeeName, data.weekEnding);
    if (duplicate) {
      Logger.log('⚠️ Duplicate timesheet for ' + data.employeeName);
      return { success: false, error: 'Duplicate timesheet' };
    }

    const id = generateFullUUID();

    // Match column order in PendingTimesheets sheet header (Config.gs:326-351)
    const rowData = [
      id,                              // 1. ID
      data.rawDataImportId || '',      // 2. RAW_DATA_IMPORT_ID
      data.employeeId || '',           // 3. EMPLOYEE_ID
      data.employeeName,               // 4. EMPLOYEE NAME
      data.employeeClockRef || '',     // 5. EMPLOYEE_CLOCK_REF
      data.weekEnding,                 // 6. WEEKENDING
      data.calculatedTotalHours || 0,  // 7. CALCULATED_TOTAL_HOURS
      data.calculatedTotalMinutes || 0,// 8. CALCULATED_TOTAL_MINUTES
      data.hours || 0,                 // 9. HOURS
      data.minutes || 0,               // 10. MINUTES
      data.overtimeHours || 0,         // 11. OVERTIMEHOURS
      data.overtimeMinutes || 0,       // 12. OVERTIMEMINUTES
      data.lunchDeductionMin || 0,     // 13. LUNCH_DEDUCTION_MIN
      data.bathroomTimeMin || 0,       // 14. BATHROOM_TIME_MIN
      data.reconDetails || '',         // 15. RECON_DETAILS
      data.warnings || '[]',           // 16. WARNINGS
      data.notes || '',                // 17. NOTES ✅ Fixed position
      'Pending',                       // 18. STATUS ✅ Fixed position
      getCurrentUser(),                // 19. IMPORTED_BY ✅ Fixed position
      normalizeToDateOnly(new Date()), // 20. IMPORTED_DATE ✅ Fixed position
      '',                              // 21. REVIEWED_BY ✅ Fixed position
      '',                              // 22. REVIEWED_DATE ✅ Fixed position
      '',                              // 23. PAYSLIP_ID ✅ Fixed position
      false,                           // 24. IS_LOCKED
      ''                               // 25. LOCKED_DATE
    ];

    pendingSheet.appendRow(rowData);
    SpreadsheetApp.flush();

    return { success: true, data: { id: id } };

  } catch (error) {
    Logger.log('❌ ERROR in createEnhancedPendingTimesheet: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Check for duplicate import based on week ending and file hash
 *
 * @param {Date} weekEnding - Week ending date
 * @param {string} fileHash - File content hash
 * @returns {Object} Duplicate check result
 */
function checkDuplicateImport(weekEnding, fileHash) {
  try {
    const sheets = getSheets();
    const importSheet = sheets.clockImports;

    if (!importSheet) {
      return { isDuplicate: false };
    }

    const data = importSheet.getDataRange().getValues();
    const headers = data[0];

    const weekEndingCol = findColumnIndex(headers, 'WEEK_ENDING');
    const fileHashCol = findColumnIndex(headers, 'FILE_HASH');
    const statusCol = findColumnIndex(headers, 'STATUS');
    const importIdCol = findColumnIndex(headers, 'IMPORT_ID');
    const importDateCol = findColumnIndex(headers, 'IMPORT_DATE');

    const searchDate = parseDate(weekEnding);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowWeekEnding = parseDate(row[weekEndingCol]);
      const rowHash = row[fileHashCol];
      const rowStatus = row[statusCol];

      // Check if same week and same hash, and still active
      if (rowStatus === 'Active' &&
          rowWeekEnding.getTime() === searchDate.getTime() &&
          rowHash === fileHash) {
        return {
          isDuplicate: true,
          existingImportId: row[importIdCol],
          existingImportDate: formatDate(row[importDateCol])  // ✅ Convert Date to string for serialization
        };
      }
    }

    return { isDuplicate: false };

  } catch (error) {
    Logger.log('⚠️ Error checking duplicate: ' + error.message);
    return { isDuplicate: false };
  }
}

/**
 * Replace existing import with new one
 *
 * @param {string} oldImportId - Old import ID to replace
 * @param {string} newImportId - New import ID
 * @returns {Object} Result
 */
function replaceImport(oldImportId, newImportId) {
  try {
    const sheets = getSheets();
    const importSheet = sheets.clockImports;

    if (!importSheet) {
      throw new Error('CLOCK_IN_IMPORTS sheet not found');
    }

    const data = importSheet.getDataRange().getValues();
    const headers = data[0];

    const importIdCol = findColumnIndex(headers, 'IMPORT_ID');
    const statusCol = findColumnIndex(headers, 'STATUS');
    const replacedByCol = findColumnIndex(headers, 'REPLACED_BY_IMPORT_ID');
    const replacedDateCol = findColumnIndex(headers, 'REPLACED_DATE');

    // Find old import row
    for (let i = 1; i < data.length; i++) {
      if (data[i][importIdCol] === oldImportId) {
        const rowIndex = i + 1;

        importSheet.getRange(rowIndex, statusCol + 1).setValue('Replaced');
        importSheet.getRange(rowIndex, replacedByCol + 1).setValue(newImportId);
        importSheet.getRange(rowIndex, replacedDateCol + 1).setValue(new Date());

        SpreadsheetApp.flush();
        Logger.log('✅ Marked import ' + oldImportId + ' as replaced');
        break;
      }
    }

    return { success: true };

  } catch (error) {
    Logger.log('❌ ERROR in replaceImport: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Get reconciliation details for a timesheet
 *
 * @param {string} id - Timesheet ID
 * @returns {Object} Reconciliation details
 */
function getTimesheetReconData(id) {
  try {
    const sheets = getSheets();
    const pendingSheet = sheets.pending;

    if (!pendingSheet) {
      throw new Error('PendingTimesheets sheet not found');
    }

    const data = pendingSheet.getDataRange().getValues();
    const headers = data[0];

    const idCol = findColumnIndex(headers, 'ID');
    const reconCol = findColumnIndex(headers, 'RECON_DETAILS');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === id) {
        const reconJson = data[i][reconCol];

        if (reconJson) {
          return {
            success: true,
            data: JSON.parse(reconJson)
          };
        } else {
          return {
            success: false,
            error: 'No reconciliation data available'
          };
        }
      }
    }

    return {
      success: false,
      error: 'Timesheet not found'
    };

  } catch (error) {
    Logger.log('❌ ERROR in getTimesheetReconData: ' + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

// ==================== TIMESHEET BREAKDOWN DASHBOARD ====================

/**
 * Get timesheet breakdown data for the dashboard
 * Returns all pending timesheets with their daily breakdown details
 *
 * @returns {Object} Result with array of employee breakdown data
 */
function getTimesheetBreakdownData() {
  try {
    Logger.log('\n========== GET TIMESHEET BREAKDOWN DATA ==========');

    const sheets = getSheets_v2();
    const rawClockSheet = sheets.rawClockData;

    if (!rawClockSheet) {
      throw new Error('RAW_CLOCK_DATA sheet not found');
    }

    // Get raw clock data
    const rawData = rawClockSheet.getDataRange().getValues();
    const rawHeaders = rawData[0];

    Logger.log('📊 Raw data rows: ' + rawData.length);
    Logger.log('📊 Headers: ' + JSON.stringify(rawHeaders));

    // Find column indices for raw clock data
    const rawEmpNameCol = findColumnIndex(rawHeaders, 'EMPLOYEE_NAME');
    const rawPunchDateCol = findColumnIndex(rawHeaders, 'PUNCH_DATE');
    const rawPunchTimeCol = findColumnIndex(rawHeaders, 'PUNCH_TIME');
    const rawDeviceCol = findColumnIndex(rawHeaders, 'DEVICE_NAME');

    // Log column indices for debugging
    Logger.log('📊 Column indices - EMPLOYEE_NAME: ' + rawEmpNameCol +
               ', PUNCH_DATE: ' + rawPunchDateCol +
               ', PUNCH_TIME: ' + rawPunchTimeCol +
               ', DEVICE_NAME: ' + rawDeviceCol);

    // Check if critical columns were found
    if (rawEmpNameCol === -1 || rawPunchTimeCol === -1) {
      throw new Error('Required columns not found. EMPLOYEE_NAME: ' + rawEmpNameCol + ', PUNCH_TIME: ' + rawPunchTimeCol);
    }

    // Get time config
    const configResult = getTimeConfig();
    const config = configResult.success ? configResult.data : DEFAULT_TIME_CONFIG;

    // Group raw clock data by employee
    const employeeData = {};
    const deviceTypes = {}; // Track device types for diagnostics

    for (let j = 1; j < rawData.length; j++) {
      const rawRow = rawData[j];
      const empName = rawRow[rawEmpNameCol];

      if (!empName) continue;

      if (!employeeData[empName]) {
        employeeData[empName] = [];
      }

      const deviceName = rawRow[rawDeviceCol] || 'Main Unit';

      // Track device types for diagnostics
      if (!deviceTypes[deviceName]) {
        deviceTypes[deviceName] = 0;
      }
      deviceTypes[deviceName]++;

      employeeData[empName].push({
        EMPLOYEE_NAME: empName,
        PUNCH_DATE: rawRow[rawPunchDateCol],
        PUNCH_TIME: rawRow[rawPunchTimeCol],
        DEVICE_NAME: deviceName
      });
    }

    // Log device type distribution
    const employeeCount = Object.keys(employeeData).length;
    Logger.log('📊 Found ' + employeeCount + ' employees with data');
    Logger.log('📊 Device type distribution:');
    for (const device in deviceTypes) {
      Logger.log('   - ' + device + ': ' + deviceTypes[device] + ' punches');
    }

    const results = [];
    let processedCount = 0;
    let failedCount = 0;

    // Process each employee's data
    for (const employeeName in employeeData) {
      const clockData = employeeData[employeeName];

      if (clockData.length === 0) continue;

      // Process the clock data to get daily breakdown
      const processResult = processClockData(clockData, config);

      if (processResult.success) {
        processedCount++;
        const data = processResult.data;

        // Determine week ending from the data (use last date + days to Friday)
        let weekEnding = '';
        if (data.dailyBreakdown && data.dailyBreakdown.length > 0) {
          const dates = data.dailyBreakdown.map(d => d.date).sort();
          const lastDate = new Date(dates[dates.length - 1]);
          // Find next Friday
          const dayOfWeek = lastDate.getDay();
          const daysToFriday = dayOfWeek <= 5 ? (5 - dayOfWeek) : (5 + 7 - dayOfWeek);
          const friday = new Date(lastDate);
          friday.setDate(friday.getDate() + daysToFriday);
          weekEnding = Utilities.formatDate(friday, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }

        results.push({
          id: employeeName, // Use employee name as ID since we're not using pending timesheets
          employeeName: employeeName,
          weekEnding: weekEnding,
          totalMinutes: data.rawTotalMinutes || 0,
          lunchDeductionMinutes: data.lunchDeductionMinutes || 0,
          status: 'Raw Data',
          dailyBreakdown: data.dailyBreakdown || []
        });
      } else {
        // Log when processing fails for an employee
        failedCount++;
        Logger.log('⚠️ Processing failed for ' + employeeName + ': ' + (processResult.error || 'Unknown error'));
      }
    }

    Logger.log('✅ Retrieved breakdown for ' + results.length + ' employees from RAW_CLOCK_DATA');
    if (failedCount > 0) {
      Logger.log('⚠️ Failed to process ' + failedCount + ' employees');
    }
    Logger.log('========== GET TIMESHEET BREAKDOWN DATA COMPLETE ==========\n');

    const returnValue = {
      success: true,
      data: results
    };

    Logger.log('🔄 Returning: success=' + returnValue.success + ', employees=' + returnValue.data.length);

    return returnValue;

  } catch (error) {
    Logger.log('❌ ERROR in getTimesheetBreakdownData: ' + error.message);
    Logger.log('Stack: ' + (error.stack || 'N/A'));
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Show the timesheet breakdown dashboard
 */
function showTimesheetBreakdown() {
  const html = HtmlService.createHtmlOutputFromFile('TimesheetBreakdown')
    .setWidth(1200)
    .setHeight(800)
    .setTitle('Timesheet Breakdown Dashboard');

  SpreadsheetApp.getUi().showModalDialog(html, 'Timesheet Breakdown Dashboard');
}

/**
 * Lock timesheet record after payslip generation
 *
 * @param {string} id - Timesheet ID
 * @returns {Object} Result
 */
function lockTimesheetRecord(id) {
  try {
    const sheets = getSheets();
    const pendingSheet = sheets.pending;

    if (!pendingSheet) {
      throw new Error('PendingTimesheets sheet not found');
    }

    const data = pendingSheet.getDataRange().getValues();
    const headers = data[0];

    const idCol = findColumnIndex(headers, 'ID');
    const isLockedCol = findColumnIndex(headers, 'IS_LOCKED');
    const lockedDateCol = findColumnIndex(headers, 'LOCKED_DATE');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === id) {
        const rowIndex = i + 1;

        pendingSheet.getRange(rowIndex, isLockedCol + 1).setValue(true);
        pendingSheet.getRange(rowIndex, lockedDateCol + 1).setValue(new Date());

        SpreadsheetApp.flush();
        Logger.log('✅ Locked timesheet: ' + id);

        return { success: true };
      }
    }

    return {
      success: false,
      error: 'Timesheet not found'
    };

  } catch (error) {
    Logger.log('❌ ERROR in lockTimesheetRecord: ' + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

// ==================== TEST FUNCTIONS ====================

/**
 * Test function for timesheet workflow
 */
function test_timesheetWorkflow() {
  Logger.log('\n========== TEST: TIMESHEET WORKFLOW ==========');

  // Test data (update with actual employee name)
  const testData = {
    employeeName: 'Test Employee',
    weekEnding: new Date(),
    hours: 40,
    minutes: 0,
    overtimeHours: 2,
    overtimeMinutes: 30,
    notes: 'Test timesheet'
  };

  // Step 1: Add pending timesheet
  const addResult = addPendingTimesheet(testData);
  Logger.log('Add result: ' + (addResult.success ? 'SUCCESS' : 'FAILED - ' + addResult.error));

  if (addResult.success) {
    const timesheetId = addResult.data.id;

    // Step 2: List pending
    const listResult = listPendingTimesheets({ status: 'Pending' });
    Logger.log('List result: ' + (listResult.success ? listResult.data.length + ' records' : 'FAILED'));

    // Step 3: Approve (commented out to prevent accidental payslip creation)
    // const approveResult = approveTimesheet(timesheetId);
    // Logger.log('Approve result: ' + (approveResult.success ? 'SUCCESS' : 'FAILED'));
  }

  Logger.log('========== TEST COMPLETE ==========\n');
}
