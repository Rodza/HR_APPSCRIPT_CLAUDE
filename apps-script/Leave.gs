/**
 * LEAVE.GS - Leave Management Module
 * HR Payroll System for SA Grinding Wheels & Scorpio Abrasives
 *
 * Handles leave tracking (record keeping only - NOT linked to salary calculations)
 * Leave Pay is entered manually in MASTERSALARY if applicable.
 *
 * Sheet: LEAVE
 * Columns: TIMESTAMP, EMPLOYEE NAME, STARTDATE.LEAVE, RETURNDATE.LEAVE,
 *          TOTALDAYS.LEAVE, REASON, NOTES, USER, IMAGE
 */

// ==================== ADD LEAVE RECORD ====================

/**
 * Records a new leave entry
 *
 * @param {Object} data - Leave record data
 * @param {string} data.employeeName - Employee REFNAME (required)
 * @param {Date|string} data.startDate - Leave start date (required)
 * @param {Date|string} data.returnDate - Expected return date (required)
 * @param {string} data.reason - Leave reason from LEAVE_REASONS (required)
 * @param {string} [data.notes] - Additional notes (optional)
 * @param {string} [data.imageUrl] - Supporting document URL (optional)
 *
 * @returns {Object} Result with success flag and data/error
 *
 * @example
 * const result = addLeave({
 *   employeeName: 'Archie Patrick',
 *   startDate: '2025-10-20',
 *   returnDate: '2025-10-22',
 *   reason: 'Sick Leave',
 *   notes: 'Doctor appointment'
 * });
 */
function addLeave(data) {
  try {
    Logger.log('\n========== ADD LEAVE ==========');
    Logger.log('Input: ' + JSON.stringify(data));

    // Validate input
    validateLeave(data);

    // Get sheets
    const sheets = getSheets();
    const leaveSheet = sheets.leave;

    if (!leaveSheet) {
      throw new Error('Leave sheet not found');
    }

    // Parse dates
    const startDate = parseDate(data.startDate);
    const returnDate = parseDate(data.returnDate);

    // Calculate total days (inclusive)
    const totalDays = calculateDaysBetween(startDate, returnDate) + 1;

    Logger.log('📅 Leave period: ' + formatDate(startDate) + ' to ' + formatDate(returnDate));
    Logger.log('📊 Total days: ' + totalDays);

    // Prepare row data
    const timestamp = new Date();
    const user = getCurrentUser();

    const rowData = [
      timestamp,                          // TIMESTAMP
      data.employeeName,                  // EMPLOYEE NAME
      startDate,                          // STARTDATE.LEAVE
      returnDate,                         // RETURNDATE.LEAVE
      totalDays,                          // TOTALDAYS.LEAVE
      data.reason,                        // REASON
      data.notes || '',                   // NOTES
      user,                               // USER
      data.imageUrl || ''                 // IMAGE
    ];

    // Append to sheet
    leaveSheet.appendRow(rowData);

    const result = {
      timestamp: timestamp,
      employeeName: data.employeeName,
      startDate: startDate,
      returnDate: returnDate,
      totalDays: totalDays,
      reason: data.reason,
      notes: data.notes || '',
      user: user
    };

    // Sanitize result for web serialization (convert Date objects to ISO strings)
    const sanitizedResult = sanitizeForWeb(result);

    Logger.log('✅ Leave record added successfully');
    Logger.log('========== ADD LEAVE COMPLETE ==========\n');

    return { success: true, data: sanitizedResult };

  } catch (error) {
    Logger.log('❌ ERROR in addLeave: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== ADD LEAVE FAILED ==========\n');
    return { success: false, error: error.message };
  }
}

// ==================== GET LEAVE HISTORY ====================

/**
 * Gets leave history for a specific employee
 *
 * @param {string} employeeName - Employee REFNAME
 * @returns {Object} Result with success flag and leave records
 */
function getLeaveHistory(employeeName) {
  try {
    Logger.log('\n========== GET LEAVE HISTORY ==========');
    Logger.log('Employee: ' + employeeName);

    const sheets = getSheets();
    const leaveSheet = sheets.leave;

    if (!leaveSheet) {
      throw new Error('Leave sheet not found');
    }

    const data = leaveSheet.getDataRange().getValues();
    const headers = data[0];

    // Find column indexes
    const empNameCol = findColumnIndex(headers, 'EMPLOYEE NAME');

    if (empNameCol === -1) {
      throw new Error('EMPLOYEE NAME column not found in Leave sheet');
    }

    // Filter records for this employee
    const records = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowEmpName = row[empNameCol];

      if (rowEmpName === employeeName) {
        records.push(buildObjectFromRow(row, headers));
      }
    }

    // Sanitize records for web serialization (convert Date objects to ISO strings)
    Logger.log('Sanitizing ' + records.length + ' records for web...');
    const sanitizedRecords = [];
    for (let i = 0; i < records.length; i++) {
      try {
        const sanitized = sanitizeForWeb(records[i]);
        sanitizedRecords.push(sanitized);
      } catch (sanitizeError) {
        Logger.log('❌ Error sanitizing record ' + (i + 1) + ': ' + sanitizeError.message);
      }
    }

    Logger.log('✅ Found ' + sanitizedRecords.length + ' leave records (after sanitization)');
    Logger.log('========== GET LEAVE HISTORY COMPLETE ==========\n');

    return { success: true, data: sanitizedRecords };

  } catch (error) {
    Logger.log('❌ ERROR in getLeaveHistory: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== GET LEAVE HISTORY FAILED ==========\n');
    return { success: false, error: error.message };
  }
}

// ==================== LIST ALL LEAVE ====================

/**
 * Lists all leave records with optional filters
 *
 * @param {Object} [filters] - Filter options
 * @param {string} [filters.employeeName] - Filter by employee name
 * @param {string} [filters.reason] - Filter by leave reason
 * @param {Date|string} [filters.startDate] - Filter by start date (after this date)
 * @param {Date|string} [filters.endDate] - Filter by end date (before this date)
 *
 * @returns {Object} Result with success flag and leave records
 *
 * @example
 * const result = listLeave({ reason: 'Sick Leave' });
 */
function listLeave(filters) {
  try {
    Logger.log('\n========== LIST LEAVE ==========');
    Logger.log('Filters: ' + JSON.stringify(filters || {}));

    const sheets = getSheets();
    const leaveSheet = sheets.leave;

    if (!leaveSheet) {
      throw new Error('Leave sheet not found');
    }

    const data = leaveSheet.getDataRange().getValues();
    const headers = data[0];

    // Convert all rows to objects
    let records = [];
    for (let i = 1; i < data.length; i++) {
      records.push(buildObjectFromRow(data[i], headers));
    }

    // Apply filters if provided
    if (filters) {
      if (filters.employeeName) {
        records = records.filter(r => r['EMPLOYEE NAME'] === filters.employeeName);
      }

      if (filters.reason) {
        records = records.filter(r => r['REASON'] === filters.reason);
      }

      if (filters.startDate) {
        const filterStartDate = parseDate(filters.startDate);
        records = records.filter(r => {
          const leaveStart = parseDate(r['STARTDATE.LEAVE']);
          return leaveStart >= filterStartDate;
        });
      }

      if (filters.endDate) {
        const filterEndDate = parseDate(filters.endDate);
        records = records.filter(r => {
          const leaveEnd = parseDate(r['RETURNDATE.LEAVE']);
          return leaveEnd <= filterEndDate;
        });
      }
    }

    // Sort by start date (descending - most recent first)
    records.sort((a, b) => {
      const dateA = parseDate(a['STARTDATE.LEAVE']);
      const dateB = parseDate(b['STARTDATE.LEAVE']);
      return dateB - dateA;
    });

    // Sanitize records for web serialization (convert Date objects to ISO strings)
    Logger.log('Sanitizing ' + records.length + ' records for web...');
    const sanitizedRecords = [];
    for (let i = 0; i < records.length; i++) {
      try {
        const sanitized = sanitizeForWeb(records[i]);
        sanitizedRecords.push(sanitized);
      } catch (sanitizeError) {
        Logger.log('❌ Error sanitizing record ' + (i + 1) + ': ' + sanitizeError.message);
      }
    }

    Logger.log('✅ Found ' + sanitizedRecords.length + ' leave records (after filtering and sanitization)');
    Logger.log('========== LIST LEAVE COMPLETE ==========\n');

    return { success: true, data: sanitizedRecords };

  } catch (error) {
    Logger.log('❌ ERROR in listLeave: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== LIST LEAVE FAILED ==========\n');
    return { success: false, error: error.message };
  }
}

// ==================== VALIDATION ====================

/**
 * Validates leave data
 *
 * @param {Object} data - Leave data to validate
 * @throws {Error} If validation fails
 */
function validateLeave(data) {
  const errors = [];

  // Required fields
  if (!data.employeeName) {
    errors.push('Employee Name is required');
  }

  if (!data.startDate) {
    errors.push('Start Date is required');
  }

  if (!data.returnDate) {
    errors.push('Return Date is required');
  }

  if (!data.reason) {
    errors.push('Reason is required');
  }

  // Validate dates if both provided
  if (data.startDate && data.returnDate) {
    try {
      const startDate = parseDate(data.startDate);
      const returnDate = parseDate(data.returnDate);

      if (returnDate < startDate) {
        errors.push('Return date cannot be before start date');
      }
    } catch (e) {
      errors.push('Invalid date format: ' + e.message);
    }
  }

  // Validate reason
  if (data.reason && !LEAVE_REASONS.includes(data.reason)) {
    errors.push('Invalid leave reason. Must be one of: ' + LEAVE_REASONS.join(', '));
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

// ==================== TEST FUNCTIONS ====================

/**
 * Test function for adding leave
 */
function test_addLeave() {
  Logger.log('\n========== TEST: ADD LEAVE ==========');

  const testData = {
    employeeName: 'Test Employee',  // Change to actual employee name
    startDate: new Date('2025-10-20'),
    returnDate: new Date('2025-10-22'),
    reason: 'Sick Leave',
    notes: 'Doctor appointment - test entry'
  };

  const result = addLeave(testData);

  if (result.success) {
    Logger.log('✅ TEST PASSED: Leave added successfully');
    Logger.log('Total days: ' + result.data.totalDays);
  } else {
    Logger.log('❌ TEST FAILED: ' + result.error);
  }

  Logger.log('========== TEST COMPLETE ==========\n');
}

/**
 * Test function for listing leave
 */
function test_listLeave() {
  Logger.log('\n========== TEST: LIST LEAVE ==========');

  // Test 1: List all leave
  const result1 = listLeave();
  Logger.log('Total records: ' + (result1.success ? result1.data.length : 0));

  // Test 2: Filter by reason
  const result2 = listLeave({ reason: 'Sick Leave' });
  Logger.log('Sick leave records: ' + (result2.success ? result2.data.length : 0));

  if (result1.success && result2.success) {
    Logger.log('✅ TEST PASSED: List functions working');
  } else {
    Logger.log('❌ TEST FAILED');
  }

  Logger.log('========== TEST COMPLETE ==========\n');
}
