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
 * Sheet: PendingTimesheets
 * Columns: ID, EMPLOYEE NAME, WEEKENDING, HOURS, MINUTES, OVERTIMEHOURS,
 *          OVERTIMEMINUTES, NOTES, STATUS, IMPORTED_BY, IMPORTED_DATE,
 *          REVIEWED_BY, REVIEWED_DATE
 *
 * Excel Import Format (from HTML analyzer):
 * Employee Name, Week Ending, Standard Hours, Standard Minutes, Overtime Hours, Overtime Minutes, Notes
 */

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
    const importedDate = new Date();
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
        timesheetRecord = rowToObject(sheetData[i], headers);
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
    const weekEnding = parseDate(timesheetRecord['WEEKENDING']);
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
    pendingSheet.getRange(rowIndex, reviewedDateCol + 1).setValue(new Date());

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
    pendingSheet.getRange(rowIndex, reviewedDateCol + 1).setValue(new Date());

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

    const data = pendingSheet.getDataRange().getValues();
    const headers = data[0];

    // Convert all rows to objects
    let records = [];
    for (let i = 1; i < data.length; i++) {
      records.push(rowToObject(data[i], headers));
    }

    // Apply filters if provided
    if (filters) {
      if (filters.status) {
        records = records.filter(r => r['STATUS'] === filters.status);
      }

      if (filters.employeeName) {
        records = records.filter(r => r['EMPLOYEE NAME'] === filters.employeeName);
      }

      if (filters.weekEnding) {
        const filterWeekEnding = parseDate(filters.weekEnding);
        records = records.filter(r => {
          const recordWeekEnding = parseDate(r['WEEKENDING']);
          return recordWeekEnding.getTime() === filterWeekEnding.getTime();
        });
      }
    }

    // Sort by week ending (descending), then employee name
    records.sort((a, b) => {
      const dateA = parseDate(a['WEEKENDING']);
      const dateB = parseDate(b['WEEKENDING']);

      if (dateA.getTime() !== dateB.getTime()) {
        return dateB - dateA;  // Descending
      }

      return a['EMPLOYEE NAME'].localeCompare(b['EMPLOYEE NAME']);
    });

    Logger.log('✅ Found ' + records.length + ' pending timesheets');
    Logger.log('========== LIST PENDING TIMESHEETS COMPLETE ==========\n');

    return { success: true, data: records };

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
