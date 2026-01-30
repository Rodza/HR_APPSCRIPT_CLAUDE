/**
 * FIX DATE LOCALE ISSUES
 *
 * This script fixes date parsing issues where dates like "09/01/2026" are being
 * interpreted as US format (September 1st) instead of South African format (9th January).
 *
 * The issue occurs when:
 * 1. Spreadsheet locale is set to United States
 * 2. Dates are stored as text strings instead of Date objects
 * 3. Dates are entered in DD/MM/YYYY format but parsed as MM/DD/YYYY
 */

/**
 * Check current spreadsheet locale settings
 * @returns {Object} Locale information
 */
function checkSpreadsheetLocale() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const info = {
    locale: ss.getSpreadsheetLocale(),
    timeZone: ss.getSpreadsheetTimeZone(),
    recommendedLocale: 'en_ZA', // South Africa
    recommendedTimeZone: 'Africa/Johannesburg'
  };

  Logger.log('\n========================================');
  Logger.log('📍 SPREADSHEET LOCALE SETTINGS');
  Logger.log('========================================');
  Logger.log('Current Locale: ' + info.locale);
  Logger.log('Current TimeZone: ' + info.timeZone);
  Logger.log('');
  Logger.log('Recommended Locale: ' + info.recommendedLocale);
  Logger.log('Recommended TimeZone: ' + info.recommendedTimeZone);
  Logger.log('');

  if (info.locale !== info.recommendedLocale) {
    Logger.log('⚠️  WARNING: Locale is not set to South Africa!');
    Logger.log('   This causes dates like "09/01/2026" to be parsed as MM/DD/YYYY (US format)');
    Logger.log('   instead of DD/MM/YYYY (South African format).');
    Logger.log('');
    Logger.log('✅ To fix: Run setSpreadsheetLocale() or manually:');
    Logger.log('   File → Settings → Locale → South Africa');
  } else {
    Logger.log('✅ Locale is correctly set to South Africa');
  }
  Logger.log('========================================\n');

  return info;
}

/**
 * Set spreadsheet locale to South Africa
 * This fixes date parsing to use DD/MM/YYYY format
 */
function setSpreadsheetLocale() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log('\n========================================');
  Logger.log('🔧 SETTING SPREADSHEET LOCALE');
  Logger.log('========================================');

  const oldLocale = ss.getSpreadsheetLocale();
  const oldTimeZone = ss.getSpreadsheetTimeZone();

  Logger.log('Old Locale: ' + oldLocale);
  Logger.log('Old TimeZone: ' + oldTimeZone);
  Logger.log('');

  // Set to South Africa
  ss.setSpreadsheetLocale('en_ZA');
  ss.setSpreadsheetTimeZone('Africa/Johannesburg');

  Logger.log('New Locale: ' + ss.getSpreadsheetLocale());
  Logger.log('New TimeZone: ' + ss.getSpreadsheetTimeZone());
  Logger.log('');
  Logger.log('✅ Spreadsheet locale updated to South Africa!');
  Logger.log('   Dates will now be parsed as DD/MM/YYYY format.');
  Logger.log('========================================\n');

  return {
    success: true,
    oldLocale: oldLocale,
    newLocale: ss.getSpreadsheetLocale(),
    oldTimeZone: oldTimeZone,
    newTimeZone: ss.getSpreadsheetTimeZone()
  };
}

/**
 * Parse date string explicitly as DD/MM/YYYY format
 * Use this when you need to parse South African format dates
 *
 * @param {string} dateString - Date in DD/MM/YYYY format (e.g., "09/01/2026")
 * @returns {Date} Parsed date object
 */
function parseDateDDMMYYYY(dateString) {
  if (!dateString || typeof dateString !== 'string') {
    throw new Error('Invalid date string');
  }

  // Try to parse DD/MM/YYYY or DD-MM-YYYY format
  const parts = dateString.split(/[\/\-]/);

  if (parts.length !== 3) {
    throw new Error('Invalid date format. Expected DD/MM/YYYY or DD-MM-YYYY');
  }

  const day = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10) - 1; // JavaScript months are 0-indexed
  const year = parseInt(parts[2], 10);

  if (isNaN(day) || isNaN(month) || isNaN(year)) {
    throw new Error('Invalid date components');
  }

  if (day < 1 || day > 31) {
    throw new Error('Invalid day: ' + day);
  }

  if (month < 0 || month > 11) {
    throw new Error('Invalid month: ' + (month + 1));
  }

  if (year < 1900 || year > 2100) {
    throw new Error('Invalid year: ' + year);
  }

  const date = new Date(year, month, day, 0, 0, 0, 0);

  // Verify the date is valid (handles things like Feb 31)
  if (date.getDate() !== day || date.getMonth() !== month || date.getFullYear() !== year) {
    throw new Error('Invalid date: ' + dateString);
  }

  return date;
}

/**
 * Fix text dates in a specific column
 * Converts text strings like "09/01/2026" to proper Date objects
 * using South African format (DD/MM/YYYY)
 *
 * @param {string} sheetName - Name of the sheet
 * @param {string} columnName - Name of the date column
 * @returns {Object} Fix results
 */
function fixTextDatesInColumn(sheetName, columnName) {
  Logger.log('\n========================================');
  Logger.log('🔧 FIXING TEXT DATES IN COLUMN');
  Logger.log('========================================');
  Logger.log('Sheet: ' + sheetName);
  Logger.log('Column: ' + columnName);
  Logger.log('');

  const result = {
    sheetName: sheetName,
    columnName: columnName,
    recordsChecked: 0,
    textDatesFixed: 0,
    errors: [],
    fixedRows: []
  };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      result.errors.push('Sheet not found: ' + sheetName);
      return result;
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log('⚠️  No data rows to process');
      return result;
    }

    const headers = data[0];
    const colIndex = findColumnIndex(headers, columnName);

    if (colIndex === -1) {
      result.errors.push('Column not found: ' + columnName);
      return result;
    }

    Logger.log('Found column at index: ' + colIndex);
    Logger.log('');

    // Check each row
    for (let i = 1; i < data.length; i++) {
      result.recordsChecked++;
      const row = data[i];
      const cellValue = row[colIndex];

      // Skip empty cells
      if (!cellValue) continue;

      // Check if it's a text string (not a Date object)
      if (typeof cellValue === 'string') {
        try {
          // Try to parse as DD/MM/YYYY
          const parsedDate = parseDateDDMMYYYY(cellValue);

          // Update the cell with proper Date object
          row[colIndex] = parsedDate;
          result.textDatesFixed++;
          result.fixedRows.push({
            row: i + 1,
            oldValue: cellValue,
            newValue: parsedDate
          });

          Logger.log('Row ' + (i + 1) + ': "' + cellValue + '" → ' + formatDate(parsedDate));
        } catch (error) {
          result.errors.push({
            row: i + 1,
            value: cellValue,
            error: error.message
          });
          Logger.log('⚠️  Row ' + (i + 1) + ': Error parsing "' + cellValue + '" - ' + error.message);
        }
      }
      // Check if it's already a Date but might have been parsed wrong
      else if (cellValue instanceof Date && !isNaN(cellValue.getTime())) {
        // Date is already a valid Date object, leave it
        continue;
      }
    }

    // Write changes back to sheet
    if (result.textDatesFixed > 0) {
      Logger.log('');
      Logger.log('📝 Writing ' + result.textDatesFixed + ' fixed dates back to sheet...');
      sheet.getRange(2, 1, data.length - 1, headers.length).setValues(data.slice(1));
      SpreadsheetApp.flush();
      Logger.log('✅ Successfully fixed ' + result.textDatesFixed + ' text dates');
    } else {
      Logger.log('✅ No text dates found - all dates are already Date objects');
    }

    if (result.errors.length > 0) {
      Logger.log('');
      Logger.log('❌ Errors: ' + result.errors.length);
      result.errors.forEach(err => {
        if (typeof err === 'string') {
          Logger.log('  - ' + err);
        } else {
          Logger.log('  - Row ' + err.row + ': ' + err.error);
        }
      });
    }

  } catch (error) {
    result.errors.push('Fatal error: ' + error.message);
    Logger.log('❌ Fatal error: ' + error.message);
  }

  Logger.log('========================================\n');

  return result;
}

/**
 * Fix all text dates across all sheets
 * @returns {Object} Results for all sheets
 */
function fixAllTextDates() {
  Logger.log('\n========================================');
  Logger.log('🔧 FIX ALL TEXT DATES');
  Logger.log('========================================\n');

  const results = {
    timestamp: new Date(),
    totalFixed: 0,
    sheets: {}
  };

  // Fix each sheet
  results.sheets.mastersalary = fixTextDatesInColumn('MASTERSALARY', 'WEEKENDING');
  results.sheets.pendingTimesheets = [
    fixTextDatesInColumn('PendingTimesheets', 'WEEKENDING'),
    fixTextDatesInColumn('PendingTimesheets', 'IMPORTED_DATE'),
    fixTextDatesInColumn('PendingTimesheets', 'REVIEWED_DATE')
  ];
  results.sheets.employeeDetails = [
    fixTextDatesInColumn('Employee Details', 'DATE OF BIRTH'),
    fixTextDatesInColumn('Employee Details', 'EMPLOYMENT DATE'),
    fixTextDatesInColumn('Employee Details', 'TERMINATION DATE')
  ];
  results.sheets.employeeLoans = [
    fixTextDatesInColumn('EmployeeLoans', 'Timestamp'),
    fixTextDatesInColumn('EmployeeLoans', 'TransactionDate')
  ];
  results.sheets.leave = [
    fixTextDatesInColumn('Leave', 'TIMESTAMP'),
    fixTextDatesInColumn('Leave', 'STARTDATE.LEAVE'),
    fixTextDatesInColumn('Leave', 'RETURNDATE.LEAVE')
  ];

  // Calculate totals
  for (const sheetKey in results.sheets) {
    const sheetResults = results.sheets[sheetKey];
    if (Array.isArray(sheetResults)) {
      sheetResults.forEach(r => results.totalFixed += r.textDatesFixed);
    } else {
      results.totalFixed += sheetResults.textDatesFixed;
    }
  }

  Logger.log('\n========================================');
  Logger.log('✅ COMPLETE');
  Logger.log('========================================');
  Logger.log('Total text dates fixed: ' + results.totalFixed);
  Logger.log('========================================\n');

  return results;
}

/**
 * Complete fix workflow
 * 1. Check locale
 * 2. Set to South Africa if needed
 * 3. Fix all text dates
 * 4. Validate
 */
function completeLocaleFix() {
  Logger.log('\n========================================');
  Logger.log('🔧 COMPLETE LOCALE FIX WORKFLOW');
  Logger.log('========================================\n');

  // Step 1: Check locale
  Logger.log('Step 1: Checking current locale...');
  const localeInfo = checkSpreadsheetLocale();

  // Step 2: Set locale if needed
  if (localeInfo.locale !== 'en_ZA') {
    Logger.log('\nStep 2: Setting locale to South Africa...');
    setSpreadsheetLocale();
  } else {
    Logger.log('\nStep 2: Locale already correct, skipping...');
  }

  // Step 3: Fix text dates
  Logger.log('\nStep 3: Fixing text dates in all sheets...');
  const fixResults = fixAllTextDates();

  // Step 4: Validate
  Logger.log('\nStep 4: Validating all dates...');
  const validation = validateAllDates();

  Logger.log('\n========================================');
  Logger.log('✅ WORKFLOW COMPLETE');
  Logger.log('========================================');
  Logger.log('Text dates fixed: ' + fixResults.totalFixed);
  Logger.log('Remaining issues: ' + validation.totalIssuesFound);
  Logger.log('========================================\n');

  return {
    locale: localeInfo,
    fixed: fixResults,
    validation: validation
  };
}
