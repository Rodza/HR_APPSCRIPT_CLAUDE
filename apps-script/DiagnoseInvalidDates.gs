/**
 * DIAGNOSE SPECIFIC INVALID DATE RECORDS
 *
 * This script identifies exactly which payslip records have date issues
 * that cause "Invalid Date" to display in the web UI.
 */

/**
 * Deep inspection of all payslip dates
 * Checks dates at a deeper level than validation
 */
function diagnosePayslipDates() {
  Logger.log('\n========================================');
  Logger.log('🔍 DEEP PAYSLIP DATE DIAGNOSIS');
  Logger.log('========================================\n');

  const sheets = getSheets();
  const salarySheet = sheets.salary;

  if (!salarySheet) {
    Logger.log('❌ MASTERSALARY sheet not found');
    return;
  }

  const data = salarySheet.getDataRange().getValues();
  const headers = data[0];

  const weekEndingCol = findColumnIndex(headers, 'WEEKENDING');
  const recordNumberCol = findColumnIndex(headers, 'RECORDNUMBER');
  const employeeNameCol = findColumnIndex(headers, 'EMPLOYEE NAME');

  Logger.log('📊 Checking ' + (data.length - 1) + ' payslip records...\n');

  const issues = [];
  let validCount = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const recordNumber = row[recordNumberCol];
    const employeeName = row[employeeNameCol];
    const weekEnding = row[weekEndingCol];

    const issue = {
      row: i + 1,
      recordNumber: recordNumber,
      employeeName: employeeName,
      problems: []
    };

    // Check if WEEKENDING exists
    if (!weekEnding) {
      issue.problems.push('WEEKENDING is null/undefined/empty');
    } else {
      // Check type
      if (!(weekEnding instanceof Date)) {
        issue.problems.push('WEEKENDING is not a Date object (type: ' + typeof weekEnding + ')');
      } else {
        // Check if valid
        if (isNaN(weekEnding.getTime())) {
          issue.problems.push('WEEKENDING is Invalid Date object');
        } else {
          // Check if can be formatted
          try {
            const formatted = formatDateDDMMYYYY(weekEnding);
            if (!formatted || formatted === '') {
              issue.problems.push('WEEKENDING formats to empty string');
            } else {
              validCount++;
            }
          } catch (e) {
            issue.problems.push('WEEKENDING formatting error: ' + e.message);
          }
        }
      }
    }

    if (issue.problems.length > 0) {
      issues.push(issue);
    }
  }

  Logger.log('========================================');
  Logger.log('📊 RESULTS');
  Logger.log('========================================');
  Logger.log('✅ Valid records: ' + validCount);
  Logger.log('❌ Invalid records: ' + issues.length);
  Logger.log('');

  if (issues.length > 0) {
    Logger.log('========================================');
    Logger.log('🔍 DETAILED ISSUES');
    Logger.log('========================================\n');

    issues.forEach(issue => {
      Logger.log('Record #' + issue.recordNumber + ' (Row ' + issue.row + ')');
      Logger.log('  Employee: ' + issue.employeeName);
      Logger.log('  Problems:');
      issue.problems.forEach(p => Logger.log('    - ' + p));
      Logger.log('');
    });

    Logger.log('========================================');
    Logger.log('🔧 SUGGESTED FIXES');
    Logger.log('========================================');
    Logger.log('');
    Logger.log('For each record above, you can:');
    Logger.log('1. Open MASTERSALARY sheet');
    Logger.log('2. Go to the row number listed');
    Logger.log('3. Find the WEEKENDING column (column G)');
    Logger.log('4. Delete the cell content');
    Logger.log('5. Re-enter the date in format: YYYY-MM-DD');
    Logger.log('');
    Logger.log('Or run: fixSpecificPayslipDate(recordNumber, "YYYY-MM-DD")');
    Logger.log('');
  }

  Logger.log('========================================\n');

  return {
    totalRecords: data.length - 1,
    validRecords: validCount,
    invalidRecords: issues.length,
    issues: issues
  };
}

/**
 * Fix a specific payslip's date
 * @param {string|number} recordNumber - Record number to fix
 * @param {string} newDateStr - New date in YYYY-MM-DD format
 */
function fixSpecificPayslipDate(recordNumber, newDateStr) {
  Logger.log('\n🔧 Fixing payslip record #' + recordNumber);

  try {
    // Parse the new date
    const newDate = parseDate(newDateStr);
    Logger.log('New date: ' + formatDate(newDate));

    const sheets = getSheets();
    const salarySheet = sheets.salary;

    if (!salarySheet) {
      throw new Error('MASTERSALARY sheet not found');
    }

    const data = salarySheet.getDataRange().getValues();
    const headers = data[0];

    const weekEndingCol = findColumnIndex(headers, 'WEEKENDING');
    const recordNumberCol = findColumnIndex(headers, 'RECORDNUMBER');

    // Find the row
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][recordNumberCol] == recordNumber) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error('Record not found: ' + recordNumber);
    }

    Logger.log('Found at row: ' + (rowIndex + 1));

    // Update the date
    salarySheet.getRange(rowIndex + 1, weekEndingCol + 1).setValue(newDate);
    SpreadsheetApp.flush();

    Logger.log('✅ Successfully updated!');
    Logger.log('   Record #' + recordNumber + ' now has WEEKENDING: ' + formatDate(newDate));

    return { success: true, recordNumber: recordNumber, newDate: newDate };

  } catch (error) {
    Logger.log('❌ Error: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Get detailed info about a specific payslip record
 * @param {string|number} recordNumber - Record number to inspect
 */
function inspectPayslipRecord(recordNumber) {
  Logger.log('\n========================================');
  Logger.log('🔍 INSPECTING PAYSLIP RECORD #' + recordNumber);
  Logger.log('========================================\n');

  const sheets = getSheets();
  const salarySheet = sheets.salary;

  if (!salarySheet) {
    Logger.log('❌ MASTERSALARY sheet not found');
    return;
  }

  const data = salarySheet.getDataRange().getValues();
  const headers = data[0];

  const recordNumberCol = findColumnIndex(headers, 'RECORDNUMBER');
  const weekEndingCol = findColumnIndex(headers, 'WEEKENDING');
  const employeeNameCol = findColumnIndex(headers, 'EMPLOYEE NAME');

  // Find the row
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][recordNumberCol] == recordNumber) {
      rowIndex = i;
      break;
    }
  }

  if (rowIndex === -1) {
    Logger.log('❌ Record not found: ' + recordNumber);
    return;
  }

  const row = data[rowIndex];
  const weekEnding = row[weekEndingCol];

  Logger.log('Row number: ' + (rowIndex + 1));
  Logger.log('Record #: ' + row[recordNumberCol]);
  Logger.log('Employee: ' + row[employeeNameCol]);
  Logger.log('');
  Logger.log('WEEKENDING Analysis:');
  Logger.log('  Value: ' + weekEnding);
  Logger.log('  Type: ' + typeof weekEnding);
  Logger.log('  Is Date?: ' + (weekEnding instanceof Date));

  if (weekEnding instanceof Date) {
    Logger.log('  getTime(): ' + weekEnding.getTime());
    Logger.log('  isNaN?: ' + isNaN(weekEnding.getTime()));
    Logger.log('  toString(): ' + weekEnding.toString());

    try {
      const formatted = formatDate(weekEnding);
      Logger.log('  formatDate(): ' + formatted);
    } catch (e) {
      Logger.log('  formatDate(): ERROR - ' + e.message);
    }

    try {
      const formatted = formatDateDDMMYYYY(weekEnding);
      Logger.log('  formatDateDDMMYYYY(): ' + formatted);
    } catch (e) {
      Logger.log('  formatDateDDMMYYYY(): ERROR - ' + e.message);
    }
  }

  Logger.log('\n========================================\n');

  return {
    row: rowIndex + 1,
    recordNumber: row[recordNumberCol],
    employeeName: row[employeeNameCol],
    weekEnding: weekEnding,
    weekEndingType: typeof weekEnding,
    isDate: weekEnding instanceof Date,
    isValid: weekEnding instanceof Date ? !isNaN(weekEnding.getTime()) : false
  };
}

/**
 * Bulk fix all invalid payslip dates
 * Sets them all to a default date (Saturday of current week)
 * USE WITH CAUTION - review diagnosePayslipDates() output first
 */
function bulkFixInvalidPayslipDates() {
  Logger.log('\n⚠️  BULK FIX INVALID PAYSLIP DATES\n');
  Logger.log('This will set ALL invalid WEEKENDING dates to the most recent Saturday.');
  Logger.log('');

  // Get most recent Saturday
  const today = new Date();
  const dayOfWeek = today.getDay();
  const daysUntilSaturday = (6 - dayOfWeek + 7) % 7;
  const lastSaturday = new Date(today);
  lastSaturday.setDate(today.getDate() - (dayOfWeek === 6 ? 0 : (dayOfWeek + 1)));
  lastSaturday.setHours(0, 0, 0, 0);

  Logger.log('Default date will be: ' + formatDate(lastSaturday) + ' (last Saturday)');
  Logger.log('');

  // First, diagnose to find issues
  const diagnosis = diagnosePayslipDates();

  if (diagnosis.invalidRecords === 0) {
    Logger.log('✅ No invalid records found!');
    return;
  }

  Logger.log('\n🔧 Fixing ' + diagnosis.invalidRecords + ' records...\n');

  const sheets = getSheets();
  const salarySheet = sheets.salary;
  const data = salarySheet.getDataRange().getValues();
  const headers = data[0];
  const weekEndingCol = findColumnIndex(headers, 'WEEKENDING');

  let fixed = 0;

  diagnosis.issues.forEach(issue => {
    const rowIndex = issue.row - 1; // Convert to 0-based
    data[rowIndex][weekEndingCol] = lastSaturday;
    fixed++;
    Logger.log('Fixed Record #' + issue.recordNumber + ' (Row ' + issue.row + ')');
  });

  // Write back
  Logger.log('\n📝 Writing changes to sheet...');
  salarySheet.getRange(2, 1, data.length - 1, headers.length).setValues(data.slice(1));
  SpreadsheetApp.flush();

  Logger.log('✅ Fixed ' + fixed + ' records!');
  Logger.log('\n⚠️  IMPORTANT: Review the payslips and update WEEKENDING dates to correct values!\n');

  return {
    fixed: fixed,
    defaultDate: lastSaturday
  };
}
