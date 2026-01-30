/**
 * Migration Script: Normalize All Dates Across All Sheets
 *
 * This script normalizes all date fields across the entire system to ensure:
 * - All dates are stored at midnight (00:00:00)
 * - No time components in date-only fields
 * - Consistent date format across all tables
 *
 * WARNING: This script modifies data directly in sheets.
 * Make sure to backup your spreadsheet before running!
 *
 * To run: Call normalizeDatesAcrossAllSheets() from the Apps Script editor
 */

/**
 * Main migration function - normalizes dates across all sheets
 * @returns {Object} Migration results with counts and errors
 */
function normalizeDatesAcrossAllSheets() {
  try {
    Logger.log('\n========================================');
    Logger.log('🔧 STARTING DATE MIGRATION');
    Logger.log('========================================\n');

    const results = {
      success: true,
      timestamp: new Date(),
      sheets: {},
      totalRecordsProcessed: 0,
      totalDatesNormalized: 0,
      errors: []
    };

    // Get all sheets
    const sheets = getSheets();

    // 1. Normalize MASTERSALARY (Payslips)
    Logger.log('📊 Processing MASTERSALARY sheet...');
    results.sheets.mastersalary = normalizeSalarySheetDates(sheets.salary);

    // 2. Normalize PendingTimesheets
    Logger.log('\n📊 Processing PendingTimesheets sheet...');
    results.sheets.pendingTimesheets = normalizePendingTimesheetsSheetDates(sheets.pending);

    // 3. Normalize Employee Details
    Logger.log('\n📊 Processing Employee Details sheet...');
    results.sheets.employeeDetails = normalizeEmployeeDetailsSheetDates(sheets.empdetails);

    // 4. Normalize EmployeeLoans
    Logger.log('\n📊 Processing EmployeeLoans sheet...');
    results.sheets.employeeLoans = normalizeEmployeeLoansSheetDates(sheets.loans);

    // 5. Normalize Leave
    Logger.log('\n📊 Processing Leave sheet...');
    results.sheets.leave = normalizeLeaveSheetDates(sheets.leave);

    // Calculate totals
    for (const sheetName in results.sheets) {
      const sheetResult = results.sheets[sheetName];
      results.totalRecordsProcessed += sheetResult.recordsProcessed;
      results.totalDatesNormalized += sheetResult.datesNormalized;
      if (sheetResult.errors && sheetResult.errors.length > 0) {
        results.errors = results.errors.concat(sheetResult.errors);
      }
    }

    Logger.log('\n========================================');
    Logger.log('✅ DATE MIGRATION COMPLETE');
    Logger.log('========================================');
    Logger.log('📊 Total records processed: ' + results.totalRecordsProcessed);
    Logger.log('📅 Total dates normalized: ' + results.totalDatesNormalized);
    Logger.log('❌ Total errors: ' + results.errors.length);
    Logger.log('========================================\n');

    return results;

  } catch (error) {
    Logger.log('❌ FATAL ERROR in normalizeDatesAcrossAllSheets: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    return {
      success: false,
      error: error.message,
      stack: error.stack
    };
  }
}

/**
 * Normalize dates in MASTERSALARY sheet (Payslips)
 * Date columns: WEEKENDING
 */
function normalizeSalarySheetDates(sheet) {
  const result = {
    sheetName: 'MASTERSALARY',
    recordsProcessed: 0,
    datesNormalized: 0,
    errors: []
  };

  try {
    if (!sheet) {
      result.errors.push('Sheet not found');
      return result;
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log('  ⚠️ No data rows to process');
      return result;
    }

    const headers = data[0];
    const weekEndingCol = findColumnIndex(headers, 'WEEKENDING');

    if (weekEndingCol === -1) {
      result.errors.push('WEEKENDING column not found');
      return result;
    }

    let changedRows = [];

    for (let i = 1; i < data.length; i++) {
      result.recordsProcessed++;
      const row = data[i];

      // Normalize WEEKENDING
      if (row[weekEndingCol] && row[weekEndingCol] instanceof Date) {
        const originalDate = new Date(row[weekEndingCol]);
        const normalizedDate = normalizeToDateOnly(originalDate);

        // Check if normalization changed the date
        if (originalDate.getTime() !== normalizedDate.getTime()) {
          row[weekEndingCol] = normalizedDate;
          changedRows.push(i + 1); // Store 1-based row number
          result.datesNormalized++;
        }
      }
    }

    // Write changes back to sheet if any
    if (changedRows.length > 0) {
      Logger.log('  📝 Writing ' + changedRows.length + ' normalized dates back to sheet...');
      sheet.getRange(2, 1, data.length - 1, headers.length).setValues(data.slice(1));
      SpreadsheetApp.flush();
      Logger.log('  ✅ MASTERSALARY: Normalized ' + result.datesNormalized + ' dates in ' + result.recordsProcessed + ' records');
    } else {
      Logger.log('  ✅ MASTERSALARY: All dates already normalized (' + result.recordsProcessed + ' records checked)');
    }

  } catch (error) {
    result.errors.push('Error in MASTERSALARY: ' + error.message);
    Logger.log('  ❌ Error: ' + error.message);
  }

  return result;
}

/**
 * Normalize dates in PendingTimesheets sheet
 * Date columns: WEEKENDING, IMPORTED_DATE, REVIEWED_DATE
 */
function normalizePendingTimesheetsSheetDates(sheet) {
  const result = {
    sheetName: 'PendingTimesheets',
    recordsProcessed: 0,
    datesNormalized: 0,
    errors: []
  };

  try {
    if (!sheet) {
      result.errors.push('Sheet not found');
      return result;
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log('  ⚠️ No data rows to process');
      return result;
    }

    const headers = data[0];
    const weekEndingCol = findColumnIndex(headers, 'WEEKENDING');
    const importedDateCol = findColumnIndex(headers, 'IMPORTED_DATE');
    const reviewedDateCol = findColumnIndex(headers, 'REVIEWED_DATE');

    let changedRows = [];

    for (let i = 1; i < data.length; i++) {
      result.recordsProcessed++;
      const row = data[i];
      let rowChanged = false;

      // Normalize WEEKENDING
      if (weekEndingCol !== -1 && row[weekEndingCol] && row[weekEndingCol] instanceof Date) {
        const originalDate = new Date(row[weekEndingCol]);
        const normalizedDate = normalizeToDateOnly(originalDate);
        if (originalDate.getTime() !== normalizedDate.getTime()) {
          row[weekEndingCol] = normalizedDate;
          result.datesNormalized++;
          rowChanged = true;
        }
      }

      // Normalize IMPORTED_DATE
      if (importedDateCol !== -1 && row[importedDateCol] && row[importedDateCol] instanceof Date) {
        const originalDate = new Date(row[importedDateCol]);
        const normalizedDate = normalizeToDateOnly(originalDate);
        if (originalDate.getTime() !== normalizedDate.getTime()) {
          row[importedDateCol] = normalizedDate;
          result.datesNormalized++;
          rowChanged = true;
        }
      }

      // Normalize REVIEWED_DATE
      if (reviewedDateCol !== -1 && row[reviewedDateCol] && row[reviewedDateCol] instanceof Date) {
        const originalDate = new Date(row[reviewedDateCol]);
        const normalizedDate = normalizeToDateOnly(originalDate);
        if (originalDate.getTime() !== normalizedDate.getTime()) {
          row[reviewedDateCol] = normalizedDate;
          result.datesNormalized++;
          rowChanged = true;
        }
      }

      if (rowChanged) {
        changedRows.push(i + 1);
      }
    }

    // Write changes back to sheet if any
    if (changedRows.length > 0) {
      Logger.log('  📝 Writing ' + result.datesNormalized + ' normalized dates back to sheet...');
      sheet.getRange(2, 1, data.length - 1, headers.length).setValues(data.slice(1));
      SpreadsheetApp.flush();
      Logger.log('  ✅ PendingTimesheets: Normalized ' + result.datesNormalized + ' dates in ' + result.recordsProcessed + ' records');
    } else {
      Logger.log('  ✅ PendingTimesheets: All dates already normalized (' + result.recordsProcessed + ' records checked)');
    }

  } catch (error) {
    result.errors.push('Error in PendingTimesheets: ' + error.message);
    Logger.log('  ❌ Error: ' + error.message);
  }

  return result;
}

/**
 * Normalize dates in Employee Details sheet
 * Date columns: DATE OF BIRTH, EMPLOYMENT DATE, TERMINATION DATE
 */
function normalizeEmployeeDetailsSheetDates(sheet) {
  const result = {
    sheetName: 'Employee Details',
    recordsProcessed: 0,
    datesNormalized: 0,
    errors: []
  };

  try {
    if (!sheet) {
      result.errors.push('Sheet not found');
      return result;
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log('  ⚠️ No data rows to process');
      return result;
    }

    const headers = data[0];
    const dobCol = findColumnIndex(headers, 'DATE OF BIRTH');
    const employmentDateCol = findColumnIndex(headers, 'EMPLOYMENT DATE');
    const terminationDateCol = findColumnIndex(headers, 'TERMINATION DATE');

    let changedRows = [];

    for (let i = 1; i < data.length; i++) {
      result.recordsProcessed++;
      const row = data[i];
      let rowChanged = false;

      // Normalize DATE OF BIRTH
      if (dobCol !== -1 && row[dobCol] && row[dobCol] instanceof Date) {
        const originalDate = new Date(row[dobCol]);
        const normalizedDate = normalizeToDateOnly(originalDate);
        if (originalDate.getTime() !== normalizedDate.getTime()) {
          row[dobCol] = normalizedDate;
          result.datesNormalized++;
          rowChanged = true;
        }
      }

      // Normalize EMPLOYMENT DATE
      if (employmentDateCol !== -1 && row[employmentDateCol] && row[employmentDateCol] instanceof Date) {
        const originalDate = new Date(row[employmentDateCol]);
        const normalizedDate = normalizeToDateOnly(originalDate);
        if (originalDate.getTime() !== normalizedDate.getTime()) {
          row[employmentDateCol] = normalizedDate;
          result.datesNormalized++;
          rowChanged = true;
        }
      }

      // Normalize TERMINATION DATE
      if (terminationDateCol !== -1 && row[terminationDateCol] && row[terminationDateCol] instanceof Date) {
        const originalDate = new Date(row[terminationDateCol]);
        const normalizedDate = normalizeToDateOnly(originalDate);
        if (originalDate.getTime() !== normalizedDate.getTime()) {
          row[terminationDateCol] = normalizedDate;
          result.datesNormalized++;
          rowChanged = true;
        }
      }

      if (rowChanged) {
        changedRows.push(i + 1);
      }
    }

    // Write changes back to sheet if any
    if (changedRows.length > 0) {
      Logger.log('  📝 Writing ' + result.datesNormalized + ' normalized dates back to sheet...');
      sheet.getRange(2, 1, data.length - 1, headers.length).setValues(data.slice(1));
      SpreadsheetApp.flush();
      Logger.log('  ✅ Employee Details: Normalized ' + result.datesNormalized + ' dates in ' + result.recordsProcessed + ' records');
    } else {
      Logger.log('  ✅ Employee Details: All dates already normalized (' + result.recordsProcessed + ' records checked)');
    }

  } catch (error) {
    result.errors.push('Error in Employee Details: ' + error.message);
    Logger.log('  ❌ Error: ' + error.message);
  }

  return result;
}

/**
 * Normalize dates in EmployeeLoans sheet
 * Date columns: Timestamp, TransactionDate
 */
function normalizeEmployeeLoansSheetDates(sheet) {
  const result = {
    sheetName: 'EmployeeLoans',
    recordsProcessed: 0,
    datesNormalized: 0,
    errors: []
  };

  try {
    if (!sheet) {
      result.errors.push('Sheet not found');
      return result;
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log('  ⚠️ No data rows to process');
      return result;
    }

    const headers = data[0];
    const timestampCol = findColumnIndex(headers, 'Timestamp');
    const transactionDateCol = findColumnIndex(headers, 'TransactionDate');

    let changedRows = [];

    for (let i = 1; i < data.length; i++) {
      result.recordsProcessed++;
      const row = data[i];
      let rowChanged = false;

      // Normalize Timestamp
      if (timestampCol !== -1 && row[timestampCol] && row[timestampCol] instanceof Date) {
        const originalDate = new Date(row[timestampCol]);
        const normalizedDate = normalizeToDateOnly(originalDate);
        if (originalDate.getTime() !== normalizedDate.getTime()) {
          row[timestampCol] = normalizedDate;
          result.datesNormalized++;
          rowChanged = true;
        }
      }

      // Normalize TransactionDate
      if (transactionDateCol !== -1 && row[transactionDateCol] && row[transactionDateCol] instanceof Date) {
        const originalDate = new Date(row[transactionDateCol]);
        const normalizedDate = normalizeToDateOnly(originalDate);
        if (originalDate.getTime() !== normalizedDate.getTime()) {
          row[transactionDateCol] = normalizedDate;
          result.datesNormalized++;
          rowChanged = true;
        }
      }

      if (rowChanged) {
        changedRows.push(i + 1);
      }
    }

    // Write changes back to sheet if any
    if (changedRows.length > 0) {
      Logger.log('  📝 Writing ' + result.datesNormalized + ' normalized dates back to sheet...');
      sheet.getRange(2, 1, data.length - 1, headers.length).setValues(data.slice(1));
      SpreadsheetApp.flush();
      Logger.log('  ✅ EmployeeLoans: Normalized ' + result.datesNormalized + ' dates in ' + result.recordsProcessed + ' records');
    } else {
      Logger.log('  ✅ EmployeeLoans: All dates already normalized (' + result.recordsProcessed + ' records checked)');
    }

  } catch (error) {
    result.errors.push('Error in EmployeeLoans: ' + error.message);
    Logger.log('  ❌ Error: ' + error.message);
  }

  return result;
}

/**
 * Normalize dates in Leave sheet
 * Date columns: TIMESTAMP, STARTDATE.LEAVE, RETURNDATE.LEAVE
 */
function normalizeLeaveSheetDates(sheet) {
  const result = {
    sheetName: 'Leave',
    recordsProcessed: 0,
    datesNormalized: 0,
    errors: []
  };

  try {
    if (!sheet) {
      result.errors.push('Sheet not found');
      return result;
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log('  ⚠️ No data rows to process');
      return result;
    }

    const headers = data[0];
    const timestampCol = findColumnIndex(headers, 'TIMESTAMP');
    const startDateCol = findColumnIndex(headers, 'STARTDATE.LEAVE');
    const returnDateCol = findColumnIndex(headers, 'RETURNDATE.LEAVE');

    let changedRows = [];

    for (let i = 1; i < data.length; i++) {
      result.recordsProcessed++;
      const row = data[i];
      let rowChanged = false;

      // Normalize TIMESTAMP
      if (timestampCol !== -1 && row[timestampCol] && row[timestampCol] instanceof Date) {
        const originalDate = new Date(row[timestampCol]);
        const normalizedDate = normalizeToDateOnly(originalDate);
        if (originalDate.getTime() !== normalizedDate.getTime()) {
          row[timestampCol] = normalizedDate;
          result.datesNormalized++;
          rowChanged = true;
        }
      }

      // Normalize STARTDATE.LEAVE
      if (startDateCol !== -1 && row[startDateCol] && row[startDateCol] instanceof Date) {
        const originalDate = new Date(row[startDateCol]);
        const normalizedDate = normalizeToDateOnly(originalDate);
        if (originalDate.getTime() !== normalizedDate.getTime()) {
          row[startDateCol] = normalizedDate;
          result.datesNormalized++;
          rowChanged = true;
        }
      }

      // Normalize RETURNDATE.LEAVE
      if (returnDateCol !== -1 && row[returnDateCol] && row[returnDateCol] instanceof Date) {
        const originalDate = new Date(row[returnDateCol]);
        const normalizedDate = normalizeToDateOnly(originalDate);
        if (originalDate.getTime() !== normalizedDate.getTime()) {
          row[returnDateCol] = normalizedDate;
          result.datesNormalized++;
          rowChanged = true;
        }
      }

      if (rowChanged) {
        changedRows.push(i + 1);
      }
    }

    // Write changes back to sheet if any
    if (changedRows.length > 0) {
      Logger.log('  📝 Writing ' + result.datesNormalized + ' normalized dates back to sheet...');
      sheet.getRange(2, 1, data.length - 1, headers.length).setValues(data.slice(1));
      SpreadsheetApp.flush();
      Logger.log('  ✅ Leave: Normalized ' + result.datesNormalized + ' dates in ' + result.recordsProcessed + ' records');
    } else {
      Logger.log('  ✅ Leave: All dates already normalized (' + result.recordsProcessed + ' records checked)');
    }

  } catch (error) {
    result.errors.push('Error in Leave: ' + error.message);
    Logger.log('  ❌ Error: ' + error.message);
  }

  return result;
}

/**
 * Preview migration - shows what would be changed without making changes
 * @returns {Object} Preview results
 */
function previewDateMigration() {
  Logger.log('\n========================================');
  Logger.log('🔍 PREVIEWING DATE MIGRATION (DRY RUN)');
  Logger.log('========================================\n');

  const preview = {
    sheets: {},
    totalRecordsToProcess: 0,
    totalDatesNeedingNormalization: 0
  };

  const sheets = getSheets();

  // Preview each sheet
  preview.sheets.mastersalary = previewSheetDates(sheets.salary, 'MASTERSALARY', ['WEEKENDING']);
  preview.sheets.pendingTimesheets = previewSheetDates(sheets.pending, 'PendingTimesheets', ['WEEKENDING', 'IMPORTED_DATE', 'REVIEWED_DATE']);
  preview.sheets.employeeDetails = previewSheetDates(sheets.empdetails, 'Employee Details', ['DATE OF BIRTH', 'EMPLOYMENT DATE', 'TERMINATION DATE']);
  preview.sheets.employeeLoans = previewSheetDates(sheets.loans, 'EmployeeLoans', ['Timestamp', 'TransactionDate']);
  preview.sheets.leave = previewSheetDates(sheets.leave, 'Leave', ['TIMESTAMP', 'STARTDATE.LEAVE', 'RETURNDATE.LEAVE']);

  // Calculate totals
  for (const sheetName in preview.sheets) {
    const sheetPreview = preview.sheets[sheetName];
    preview.totalRecordsToProcess += sheetPreview.recordCount;
    preview.totalDatesNeedingNormalization += sheetPreview.datesNeedingNormalization;
  }

  Logger.log('\n========================================');
  Logger.log('📊 PREVIEW SUMMARY');
  Logger.log('========================================');
  Logger.log('📊 Total records to process: ' + preview.totalRecordsToProcess);
  Logger.log('📅 Total dates needing normalization: ' + preview.totalDatesNeedingNormalization);
  Logger.log('========================================\n');

  return preview;
}

/**
 * Preview dates in a specific sheet
 */
function previewSheetDates(sheet, sheetName, dateColumns) {
  const preview = {
    sheetName: sheetName,
    recordCount: 0,
    datesNeedingNormalization: 0,
    columns: {}
  };

  try {
    if (!sheet) {
      Logger.log('  ⚠️ ' + sheetName + ': Sheet not found');
      return preview;
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log('  ⚠️ ' + sheetName + ': No data rows');
      return preview;
    }

    const headers = data[0];
    preview.recordCount = data.length - 1;

    // Check each date column
    for (const columnName of dateColumns) {
      const colIndex = findColumnIndex(headers, columnName);
      if (colIndex === -1) continue;

      let needsNormalization = 0;
      for (let i = 1; i < data.length; i++) {
        const value = data[i][colIndex];
        if (value && value instanceof Date) {
          const normalized = normalizeToDateOnly(new Date(value));
          if (value.getTime() !== normalized.getTime()) {
            needsNormalization++;
          }
        }
      }

      preview.columns[columnName] = needsNormalization;
      preview.datesNeedingNormalization += needsNormalization;
    }

    Logger.log('  📊 ' + sheetName + ': ' + preview.recordCount + ' records, ' + preview.datesNeedingNormalization + ' dates need normalization');

  } catch (error) {
    Logger.log('  ❌ Error previewing ' + sheetName + ': ' + error.message);
  }

  return preview;
}
