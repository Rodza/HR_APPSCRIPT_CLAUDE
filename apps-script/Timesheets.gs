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
      records.push(buildObjectFromRow(data[i], headers));
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
    Logger.log('📥 Converting base64 to blob: ' + filename);

    // Decode base64 and create blob (server-side Utilities API)
    const bytes = Utilities.base64Decode(base64Data);
    const fileBlob = Utilities.newBlob(bytes, mimeType, filename);

    // Call the main import function
    return importClockData(fileBlob, filename, override);

  } catch (error) {
    Logger.log('❌ ERROR in importClockDataFromBase64: ' + error.message);
    return { success: false, error: error.message };
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

    // Parse Excel file
    const parseResult = parseClockDataExcel(fileBlob);
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
    Logger.log('❌ ERROR in importClockData: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== IMPORT CLOCK DATA FAILED ==========\n');
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Parse clock data from Excel file
 *
 * @param {Blob} fileBlob - Excel file blob
 * @returns {Object} Result with parsed data
 */
function parseClockDataExcel(fileBlob) {
  try {
    Logger.log('📖 Parsing clock data Excel file...');

    // Convert Excel to Google Sheets using Drive API REST endpoint
    const file = DriveApp.createFile(fileBlob);
    const fileId = file.getId();

    try {
      // Use Drive API v3 via UrlFetchApp (no advanced service needed)
      const accessToken = ScriptApp.getOAuthToken();
      const copyUrl = 'https://www.googleapis.com/drive/v3/files/' + fileId + '/copy';

      const copyPayload = {
        name: fileBlob.getName() + '_converted',
        mimeType: 'application/vnd.google-apps.spreadsheet'
      };

      const copyOptions = {
        method: 'post',
        contentType: 'application/json',
        headers: {
          'Authorization': 'Bearer ' + accessToken
        },
        payload: JSON.stringify(copyPayload),
        muteHttpExceptions: true
      };

      const copyResponse = UrlFetchApp.fetch(copyUrl, copyOptions);
      const copyResult = JSON.parse(copyResponse.getContentText());

      if (!copyResult.id) {
        throw new Error('Failed to convert Excel to Sheets: ' + copyResponse.getContentText());
      }

      const convertedFileId = copyResult.id;

      // Open the converted spreadsheet
      const spreadsheet = SpreadsheetApp.openById(convertedFileId);
      const sheet = spreadsheet.getSheets()[0];
      const data = sheet.getDataRange().getValues();

      // Clean up temporary files
      DriveApp.getFileById(fileId).setTrashed(true);
      DriveApp.getFileById(convertedFileId).setTrashed(true);

    } catch (conversionError) {
      // Clean up on error
      try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {}
      throw new Error('Excel conversion failed: ' + conversionError.message);
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

      // Parse Punch Time which contains both date and time: "2025-11-14 07:33:34"
      const punchTimeValue = row[colMap.punchTime];
      let punchDateTime;

      if (punchTimeValue instanceof Date) {
        punchDateTime = punchTimeValue;
      } else if (typeof punchTimeValue === 'string') {
        // Parse "YYYY-MM-DD HH:MM:SS" format
        punchDateTime = new Date(punchTimeValue);
      } else if (typeof punchTimeValue === 'number') {
        // Excel serial date
        punchDateTime = parseExcelDate(punchTimeValue);
      } else {
        Logger.log('⚠️ Row ' + (i + 1) + ': Invalid punch time format: ' + punchTimeValue);
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
    Logger.log('❌ ERROR in parseClockDataExcel: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    return { success: false, error: error.message };
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

    const rows = clockRecords.map(function(record) {
      const empInfo = employeeMap[String(record.ClockInRef).trim()] || {};

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
      const range = rawDataSheet.getRange(
        rawDataSheet.getLastRow() + 1,
        1,
        rows.length,
        rows[0].length
      );
      range.setValues(rows);
      SpreadsheetApp.flush();
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
    const config = getTimeConfig();
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

    const rowData = [
      id,
      data.rawDataImportId || '',
      data.employeeId || '',
      data.employeeName,
      data.employeeClockRef || '',
      data.weekEnding,
      data.calculatedTotalHours || 0,
      data.calculatedTotalMinutes || 0,
      data.hours || 0,
      data.minutes || 0,
      data.overtimeHours || 0,
      data.overtimeMinutes || 0,
      data.lunchDeductionMin || 0,
      data.bathroomTimeMin || 0,
      data.reconDetails || '',
      data.warnings || '[]',
      data.notes || '',
      'Pending',
      getCurrentUser(),
      new Date(),
      '',
      '',
      '',
      false,
      ''
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
          existingImportDate: row[importDateCol]
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
