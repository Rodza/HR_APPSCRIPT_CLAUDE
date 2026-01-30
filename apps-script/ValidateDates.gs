/**
 * DATE VALIDATION SCRIPT
 *
 * This script scans all date columns across all sheets and identifies issues:
 * - Invalid Date objects
 * - Text values instead of Date objects
 * - Dates with time components (not normalized to midnight)
 * - Empty cells in required date fields
 *
 * Run: validateAllDates() from Apps Script editor
 * Output: Detailed report with specific row/column references
 */

/**
 * Main validation function - checks all dates across all sheets
 * @returns {Object} Validation report with all issues found
 */
function validateAllDates() {
  Logger.log('\n========================================');
  Logger.log('🔍 DATE VALIDATION REPORT');
  Logger.log('========================================\n');
  Logger.log('Started: ' + new Date().toLocaleString());
  Logger.log('');

  const report = {
    timestamp: new Date(),
    totalSheets: 0,
    totalRecordsChecked: 0,
    totalIssuesFound: 0,
    sheets: {},
    summary: {
      invalidDates: 0,
      textDates: 0,
      datesWithTime: 0,
      emptyRequired: 0,
      nullDates: 0
    }
  };

  try {
    const sheets = getSheets();

    // Validate each sheet
    report.sheets.employeeDetails = validateSheetDates(
      sheets.empdetails,
      'Employee Details',
      [
        { name: 'DATE OF BIRTH', required: false },
        { name: 'EMPLOYMENT DATE', required: false },
        { name: 'TERMINATION DATE', required: false }
      ]
    );

    report.sheets.leave = validateSheetDates(
      sheets.leave,
      'Leave',
      [
        { name: 'TIMESTAMP', required: true },
        { name: 'STARTDATE.LEAVE', required: true },
        { name: 'RETURNDATE.LEAVE', required: true }
      ]
    );

    report.sheets.employeeLoans = validateSheetDates(
      sheets.loans,
      'EmployeeLoans',
      [
        { name: 'Timestamp', required: true },
        { name: 'TransactionDate', required: true }
      ]
    );

    report.sheets.mastersalary = validateSheetDates(
      sheets.salary,
      'MASTERSALARY',
      [
        { name: 'WEEKENDING', required: true }
      ]
    );

    report.sheets.pendingTimesheets = validateSheetDates(
      sheets.pending,
      'PendingTimesheets',
      [
        { name: 'WEEKENDING', required: true },
        { name: 'IMPORTED_DATE', required: false },
        { name: 'REVIEWED_DATE', required: false }
      ]
    );

    // Calculate totals
    for (const sheetKey in report.sheets) {
      const sheetReport = report.sheets[sheetKey];
      report.totalSheets++;
      report.totalRecordsChecked += sheetReport.recordsChecked;
      report.totalIssuesFound += sheetReport.issues.length;

      // Count issue types
      sheetReport.issues.forEach(issue => {
        switch (issue.type) {
          case 'INVALID_DATE':
            report.summary.invalidDates++;
            break;
          case 'TEXT_DATE':
            report.summary.textDates++;
            break;
          case 'DATE_WITH_TIME':
            report.summary.datesWithTime++;
            break;
          case 'EMPTY_REQUIRED':
            report.summary.emptyRequired++;
            break;
          case 'NULL_DATE':
            report.summary.nullDates++;
            break;
        }
      });
    }

    // Print summary
    Logger.log('\n========================================');
    Logger.log('📊 SUMMARY');
    Logger.log('========================================');
    Logger.log('Sheets checked: ' + report.totalSheets);
    Logger.log('Records checked: ' + report.totalRecordsChecked);
    Logger.log('Total issues found: ' + report.totalIssuesFound);
    Logger.log('');
    Logger.log('Issue breakdown:');
    Logger.log('  Invalid Date objects: ' + report.summary.invalidDates);
    Logger.log('  Text instead of Date: ' + report.summary.textDates);
    Logger.log('  Dates with time: ' + report.summary.datesWithTime);
    Logger.log('  Empty required fields: ' + report.summary.emptyRequired);
    Logger.log('  Null dates: ' + report.summary.nullDates);
    Logger.log('');

    // Print detailed issues by sheet
    if (report.totalIssuesFound > 0) {
      Logger.log('========================================');
      Logger.log('🔍 DETAILED ISSUES BY SHEET');
      Logger.log('========================================\n');

      for (const sheetKey in report.sheets) {
        const sheetReport = report.sheets[sheetKey];
        if (sheetReport.issues.length > 0) {
          Logger.log('\n📋 ' + sheetReport.sheetName + ' (' + sheetReport.issues.length + ' issues)');
          Logger.log('─'.repeat(60));

          sheetReport.issues.forEach(issue => {
            Logger.log('');
            Logger.log('  Row: ' + issue.row + ' | Column: ' + issue.column);
            Logger.log('  Issue: ' + issue.description);
            Logger.log('  Current value: ' + (issue.currentValue || '(empty)'));
            Logger.log('  Type: ' + issue.valueType);
            if (issue.suggestion) {
              Logger.log('  ✅ Suggestion: ' + issue.suggestion);
            }
          });
        }
      }
    } else {
      Logger.log('✅ No issues found! All dates are valid and properly formatted.');
    }

    Logger.log('\n========================================');
    Logger.log('Completed: ' + new Date().toLocaleString());
    Logger.log('========================================\n');

    return report;

  } catch (error) {
    Logger.log('❌ ERROR in validateAllDates: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    return {
      success: false,
      error: error.message,
      stack: error.stack
    };
  }
}

/**
 * Validate dates in a single sheet
 * @param {Sheet} sheet - Google Sheets sheet object
 * @param {string} sheetName - Name of the sheet for reporting
 * @param {Array} dateColumns - Array of {name, required} objects
 * @returns {Object} Validation report for this sheet
 */
function validateSheetDates(sheet, sheetName, dateColumns) {
  const sheetReport = {
    sheetName: sheetName,
    recordsChecked: 0,
    issues: []
  };

  try {
    if (!sheet) {
      sheetReport.issues.push({
        row: 'N/A',
        column: 'N/A',
        type: 'SHEET_NOT_FOUND',
        description: 'Sheet not found in spreadsheet',
        suggestion: 'Verify sheet name and existence'
      });
      return sheetReport;
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      // Only header row or empty sheet
      return sheetReport;
    }

    const headers = data[0];

    // Find column indexes for all date columns
    const dateColumnIndexes = {};
    dateColumns.forEach(col => {
      const index = findColumnIndex(headers, col.name);
      if (index !== -1) {
        dateColumnIndexes[col.name] = {
          index: index,
          required: col.required
        };
      } else {
        sheetReport.issues.push({
          row: 'Header',
          column: col.name,
          type: 'COLUMN_NOT_FOUND',
          description: 'Date column "' + col.name + '" not found in sheet',
          suggestion: 'Verify column name spelling and existence'
        });
      }
    });

    // Check each data row
    for (let i = 1; i < data.length; i++) {
      sheetReport.recordsChecked++;
      const row = data[i];
      const rowNumber = i + 1; // 1-based row number for display

      // Check each date column in this row
      for (const columnName in dateColumnIndexes) {
        const colInfo = dateColumnIndexes[columnName];
        const cellValue = row[colInfo.index];

        const issue = validateDateCell(
          cellValue,
          rowNumber,
          columnName,
          colInfo.required
        );

        if (issue) {
          sheetReport.issues.push(issue);
        }
      }
    }

  } catch (error) {
    sheetReport.issues.push({
      row: 'Error',
      column: 'N/A',
      type: 'VALIDATION_ERROR',
      description: 'Error validating sheet: ' + error.message,
      suggestion: 'Check sheet structure and permissions'
    });
  }

  return sheetReport;
}

/**
 * Validate a single date cell
 * @param {*} cellValue - The cell value to validate
 * @param {number} row - Row number (1-based)
 * @param {string} column - Column name
 * @param {boolean} required - Whether this field is required
 * @returns {Object|null} Issue object if invalid, null if valid
 */
function validateDateCell(cellValue, row, column, required) {
  // Check for empty/null values
  if (cellValue === null || cellValue === undefined || cellValue === '') {
    if (required) {
      return {
        row: row,
        column: column,
        type: 'EMPTY_REQUIRED',
        description: 'Required date field is empty',
        currentValue: '(empty)',
        valueType: 'empty',
        suggestion: 'Enter a valid date in format YYYY-MM-DD'
      };
    }
    // Optional field that's empty is OK
    return null;
  }

  // Check if it's a Date object
  if (!(cellValue instanceof Date)) {
    return {
      row: row,
      column: column,
      type: 'TEXT_DATE',
      description: 'Cell contains ' + typeof cellValue + ' instead of Date object',
      currentValue: String(cellValue),
      valueType: typeof cellValue,
      suggestion: 'Convert to Date: Select cell → Format → Number → Date'
    };
  }

  // Check if the Date object is valid
  if (isNaN(cellValue.getTime())) {
    return {
      row: row,
      column: column,
      type: 'INVALID_DATE',
      description: 'Cell contains invalid Date object',
      currentValue: String(cellValue),
      valueType: 'Invalid Date',
      suggestion: 'Delete cell content and re-enter date in format YYYY-MM-DD'
    };
  }

  // Check if date is normalized to midnight (no time component)
  const hours = cellValue.getHours();
  const minutes = cellValue.getMinutes();
  const seconds = cellValue.getSeconds();
  const milliseconds = cellValue.getMilliseconds();

  if (hours !== 0 || minutes !== 0 || seconds !== 0 || milliseconds !== 0) {
    return {
      row: row,
      column: column,
      type: 'DATE_WITH_TIME',
      description: 'Date has time component (should be midnight)',
      currentValue: cellValue.toLocaleString(),
      valueType: 'Date with time',
      suggestion: 'Run normalizeDatesAcrossAllSheets() to fix automatically'
    };
  }

  // Date is valid!
  return null;
}

/**
 * Generate a CSV report of all date issues
 * Outputs to Logger in CSV format for easy copy-paste to spreadsheet
 */
function generateDateIssuesCSV() {
  Logger.log('Generating CSV report...\n');

  const report = validateAllDates();

  Logger.log('\n========================================');
  Logger.log('📄 CSV EXPORT (Copy-paste to spreadsheet)');
  Logger.log('========================================\n');

  // CSV Header
  Logger.log('Sheet,Row,Column,Issue Type,Description,Current Value,Value Type,Suggestion');

  // CSV Data
  for (const sheetKey in report.sheets) {
    const sheetReport = report.sheets[sheetKey];
    sheetReport.issues.forEach(issue => {
      const csvRow = [
        sheetReport.sheetName,
        issue.row,
        issue.column,
        issue.type,
        '"' + issue.description + '"',
        '"' + (issue.currentValue || '') + '"',
        issue.valueType || '',
        '"' + (issue.suggestion || '') + '"'
      ].join(',');

      Logger.log(csvRow);
    });
  }

  Logger.log('\n========================================\n');

  return report;
}

/**
 * Quick check - just count issues without details
 * @returns {Object} Summary counts
 */
function quickDateCheck() {
  Logger.log('Running quick date check...\n');

  const report = validateAllDates();

  const summary = {
    totalIssues: report.totalIssuesFound,
    bySheet: {}
  };

  for (const sheetKey in report.sheets) {
    const sheetReport = report.sheets[sheetKey];
    summary.bySheet[sheetReport.sheetName] = sheetReport.issues.length;
  }

  Logger.log('\n📊 QUICK SUMMARY:');
  Logger.log('  Total issues: ' + summary.totalIssues);
  Logger.log('  By sheet:');
  for (const sheetName in summary.bySheet) {
    Logger.log('    ' + sheetName + ': ' + summary.bySheet[sheetName]);
  }
  Logger.log('');

  return summary;
}

/**
 * Auto-fix issues that can be fixed automatically
 * Fixes:
 * - Dates with time components (normalizes to midnight)
 *
 * Does NOT fix:
 * - Invalid Date objects (require manual entry)
 * - Text dates (require format change)
 * - Empty required fields (require data entry)
 */
function autoFixDateIssues() {
  Logger.log('\n========================================');
  Logger.log('🔧 AUTO-FIX DATE ISSUES');
  Logger.log('========================================\n');
  Logger.log('This will automatically fix dates with time components.');
  Logger.log('Other issues require manual correction.\n');

  // Run the existing migration script which fixes time components
  const result = normalizeDatesAcrossAllSheets();

  Logger.log('\n✅ Auto-fix complete!');
  Logger.log('Run validateAllDates() again to check remaining issues.\n');

  return result;
}
