/**
 * LEAVE.GS - Leave Management Module
 * HR Payroll System for SA Grinding Wheels & Scorpio Abrasives
 *
 * Handles leave tracking (record keeping only - NOT linked to salary calculations)
 * Leave Pay is entered manually in MASTERSALARY if applicable.
 *
 * Sheet: LEAVE
 * Columns: TIMESTAMP, EMPLOYEE NAME, STARTDATE.LEAVE, RETURNDATE.LEAVE,
 *          TOTALDAYS.LEAVE, REASON, NOTES, USER, IMAGE, WEEK.DAY
 *
 * WEEK.DAY - Automatically calculated from STARTDATE.LEAVE (e.g., "Monday", "Tuesday")
 */

// ==================== HELPER FUNCTIONS ====================

/**
 * Get day of week name from a date
 *
 * @param {Date} date - Date object
 * @returns {string} Day name (e.g., "Monday", "Tuesday")
 */
function getDayOfWeek(date) {
  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  return days[date.getDay()];
}

/**
 * Get or create the Leave Documents folder in Google Drive
 *
 * @returns {Folder} Google Drive folder for leave documents
 */
function getLeaveDocumentsFolder() {
  const folderName = 'Leave Documents';
  const folders = DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  } else {
    // Create folder if it doesn't exist
    const folder = DriveApp.createFolder(folderName);
    Logger.log('Created Leave Documents folder: ' + folder.getId());
    return folder;
  }
}

/**
 * Upload a leave document to Google Drive
 *
 * @param {Object} fileData - File data object
 * @param {string} fileData.fileName - Original file name
 * @param {string} fileData.mimeType - File MIME type
 * @param {string} fileData.content - Base64 encoded file content
 * @param {string} fileData.employeeName - Employee name for file naming
 * @param {string} fileData.startDate - Leave start date for file naming
 *
 * @returns {Object} Result with success flag and file URL/error
 *
 * @example
 * const result = uploadLeaveDocument({
 *   fileName: 'doctors_note.pdf',
 *   mimeType: 'application/pdf',
 *   content: 'base64string...',
 *   employeeName: 'John Doe',
 *   startDate: '2025-11-20'
 * });
 */
function uploadLeaveDocument(fileData) {
  try {
    Logger.log('\n========== UPLOAD LEAVE DOCUMENT ==========');
    Logger.log('File name: ' + fileData.fileName);
    Logger.log('MIME type: ' + fileData.mimeType);
    Logger.log('Employee: ' + fileData.employeeName);

    // Validate input
    if (!fileData.fileName || !fileData.mimeType || !fileData.content) {
      throw new Error('File name, MIME type, and content are required');
    }

    // Get or create Leave Documents folder
    const folder = getLeaveDocumentsFolder();

    // Create a unique file name
    const timestamp = new Date().getTime();
    const sanitizedEmployeeName = fileData.employeeName.replace(/[^a-zA-Z0-9]/g, '_');
    const sanitizedStartDate = fileData.startDate ? fileData.startDate.replace(/[^0-9-]/g, '') : 'unknown';
    const fileExtension = fileData.fileName.split('.').pop();
    const uniqueFileName = `Leave_${sanitizedEmployeeName}_${sanitizedStartDate}_${timestamp}.${fileExtension}`;

    // Decode base64 content
    const blob = Utilities.newBlob(
      Utilities.base64Decode(fileData.content),
      fileData.mimeType,
      uniqueFileName
    );

    // Create file in Drive
    const file = folder.createFile(blob);

    // Make file accessible to anyone with the link (for easy access)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileUrl = file.getUrl();
    const fileId = file.getId();

    Logger.log('✅ File uploaded successfully');
    Logger.log('File ID: ' + fileId);
    Logger.log('File URL: ' + fileUrl);
    Logger.log('========== UPLOAD LEAVE DOCUMENT COMPLETE ==========\n');

    return {
      success: true,
      data: {
        fileUrl: fileUrl,
        fileId: fileId,
        fileName: uniqueFileName
      }
    };

  } catch (error) {
    Logger.log('❌ ERROR in uploadLeaveDocument: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== UPLOAD LEAVE DOCUMENT FAILED ==========\n');
    return { success: false, error: error.message };
  }
}

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

    // Calculate total working days (Monday to Friday only, inclusive)
    const totalDays = calculateWorkingDaysBetween(startDate, returnDate);

    // Calculate day of week from start date
    const weekDay = getDayOfWeek(startDate);

    Logger.log('📅 Leave period: ' + formatDate(startDate) + ' to ' + formatDate(returnDate));
    Logger.log('📊 Total days: ' + totalDays);
    Logger.log('📅 Start day: ' + weekDay);

    // Prepare row data
    const timestamp = normalizeToDateOnly(new Date());
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
      data.imageUrl || '',                // IMAGE
      weekDay                             // WEEK.DAY (automatically calculated)
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
        const record = buildObjectFromRow(row, headers);
        // Only include records with valid required date fields
        if (record['STARTDATE.LEAVE'] && record['RETURNDATE.LEAVE']) {
          records.push(record);
        } else {
          Logger.log('⚠️ Skipping record with missing dates for employee: ' + employeeName);
        }
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

    // Convert all rows to objects and filter out records with missing required date fields
    let records = [];
    let skippedOnLoad = 0;
    for (let i = 1; i < data.length; i++) {
      const record = buildObjectFromRow(data[i], headers);
      // Only include records with valid required date fields
      if (record['STARTDATE.LEAVE'] && record['RETURNDATE.LEAVE']) {
        records.push(record);
      } else {
        skippedOnLoad++;
      }
    }

    if (skippedOnLoad > 0) {
      Logger.log('⚠️ Skipped ' + skippedOnLoad + ' records with missing dates on initial load');
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
          if (!r['STARTDATE.LEAVE']) return false;
          try {
            const leaveStart = parseDate(r['STARTDATE.LEAVE']);
            return leaveStart >= filterStartDate;
          } catch (e) {
            return false;
          }
        });
      }

      if (filters.endDate) {
        const filterEndDate = parseDate(filters.endDate);
        records = records.filter(r => {
          if (!r['RETURNDATE.LEAVE']) return false;
          try {
            const leaveEnd = parseDate(r['RETURNDATE.LEAVE']);
            return leaveEnd <= filterEndDate;
          } catch (e) {
            return false;
          }
        });
      }
    }

    // Sort records by start date (descending - most recent first)
    // Handle null/invalid dates by sorting them to the end
    records.sort((a, b) => {
      try {
        const dateA = a['STARTDATE.LEAVE'] ? parseDate(a['STARTDATE.LEAVE']) : new Date(0);
        const dateB = b['STARTDATE.LEAVE'] ? parseDate(b['STARTDATE.LEAVE']) : new Date(0);
        return dateB - dateA;
      } catch (e) {
        return 0;
      }
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

// ==================== EDIT LEAVE RECORD ====================

/**
 * Edit an existing leave record
 *
 * @param {number} rowNumber - Row number of the leave record (1-based, excluding header)
 * @param {Object} data - Updated leave data (only include fields to update)
 * @param {string} [data.employeeName] - Employee REFNAME
 * @param {Date|string} [data.startDate] - Leave start date
 * @param {Date|string} [data.returnDate] - Expected return date
 * @param {string} [data.reason] - Leave reason from LEAVE_REASONS
 * @param {string} [data.notes] - Additional notes
 * @param {string} [data.imageUrl] - Supporting document URL
 *
 * @returns {Object} Result with success flag and data/error
 *
 * @example
 * const result = editLeave(5, {
 *   startDate: '2025-11-05',
 *   returnDate: '2025-11-07',
 *   notes: 'Extended sick leave'
 * });
 */
function editLeave(rowNumber, data) {
  try {
    Logger.log('\n========== EDIT LEAVE ==========');
    Logger.log('Row Number: ' + rowNumber);
    Logger.log('Update Data: ' + JSON.stringify(data));

    if (!rowNumber || rowNumber < 1) {
      throw new Error('Valid row number is required (must be >= 1)');
    }

    // Get sheets
    const sheets = getSheets();
    const leaveSheet = sheets.leave;

    if (!leaveSheet) {
      throw new Error('Leave sheet not found');
    }

    // Get all data
    const allData = leaveSheet.getDataRange().getValues();
    const headers = allData[0];

    // Calculate actual sheet row (add 1 for header row)
    const sheetRowIndex = rowNumber + 1;

    if (sheetRowIndex > allData.length) {
      throw new Error('Row number out of range: ' + rowNumber);
    }

    // Get current record
    const currentRow = allData[sheetRowIndex - 1];
    const currentRecord = buildObjectFromRow(currentRow, headers);

    Logger.log('📋 Current record: ' + JSON.stringify(currentRecord));

    // Merge updates with current data
    const updatedRecord = {
      employeeName: data.employeeName || currentRecord['EMPLOYEE NAME'],
      startDate: data.startDate || currentRecord['STARTDATE.LEAVE'],
      returnDate: data.returnDate || currentRecord['RETURNDATE.LEAVE'],
      reason: data.reason || currentRecord['REASON'],
      notes: data.hasOwnProperty('notes') ? data.notes : currentRecord['NOTES'],
      imageUrl: data.hasOwnProperty('imageUrl') ? data.imageUrl : currentRecord['IMAGE']
    };

    // Validate updated record
    validateLeave(updatedRecord);

    // Parse dates
    const startDate = parseDate(updatedRecord.startDate);
    const returnDate = parseDate(updatedRecord.returnDate);

    // Recalculate total working days (Monday to Friday only, inclusive) and week day
    const totalDays = calculateWorkingDaysBetween(startDate, returnDate);
    const weekDay = getDayOfWeek(startDate);

    Logger.log('📅 Updated leave period: ' + formatDate(startDate) + ' to ' + formatDate(returnDate));
    Logger.log('📊 Updated total days: ' + totalDays);
    Logger.log('📅 Updated start day: ' + weekDay);

    // Prepare updated row data (keep original timestamp and user)
    const updatedRowData = [
      currentRecord['TIMESTAMP'],                // TIMESTAMP (keep original)
      updatedRecord.employeeName,                // EMPLOYEE NAME
      startDate,                                 // STARTDATE.LEAVE
      returnDate,                                // RETURNDATE.LEAVE
      totalDays,                                 // TOTALDAYS.LEAVE
      updatedRecord.reason,                      // REASON
      updatedRecord.notes || '',                 // NOTES
      currentRecord['USER'],                     // USER (keep original)
      updatedRecord.imageUrl || '',              // IMAGE
      weekDay                                    // WEEK.DAY (recalculated)
    ];

    // Update the row in the sheet
    const range = leaveSheet.getRange(sheetRowIndex, 1, 1, headers.length);
    range.setValues([updatedRowData]);

    const result = {
      rowNumber: rowNumber,
      employeeName: updatedRecord.employeeName,
      startDate: startDate,
      returnDate: returnDate,
      totalDays: totalDays,
      weekDay: weekDay,
      reason: updatedRecord.reason,
      notes: updatedRecord.notes || '',
      imageUrl: updatedRecord.imageUrl || ''
    };

    // Sanitize result for web serialization
    const sanitizedResult = sanitizeForWeb(result);

    Logger.log('✅ Leave record updated successfully');
    Logger.log('========== EDIT LEAVE COMPLETE ==========\n');

    return { success: true, data: sanitizedResult };

  } catch (error) {
    Logger.log('❌ ERROR in editLeave: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== EDIT LEAVE FAILED ==========\n');
    return { success: false, error: error.message };
  }
}

// ==================== DELETE LEAVE RECORD ====================

/**
 * Delete a leave record
 *
 * @param {number} rowNumber - Row number of the leave record to delete (1-based, excluding header)
 *
 * @returns {Object} Result with success flag and data/error
 *
 * @example
 * const result = deleteLeave(5);
 */
function deleteLeave(rowNumber) {
  try {
    Logger.log('\n========== DELETE LEAVE ==========');
    Logger.log('Row Number: ' + rowNumber);

    if (!rowNumber || rowNumber < 1) {
      throw new Error('Valid row number is required (must be >= 1)');
    }

    // Get sheets
    const sheets = getSheets();
    const leaveSheet = sheets.leave;

    if (!leaveSheet) {
      throw new Error('Leave sheet not found');
    }

    // Get all data to verify row exists
    const allData = leaveSheet.getDataRange().getValues();
    const headers = allData[0];

    // Calculate actual sheet row (add 1 for header row)
    const sheetRowIndex = rowNumber + 1;

    if (sheetRowIndex > allData.length) {
      throw new Error('Row number out of range: ' + rowNumber);
    }

    // Get the record before deleting (for logging and return)
    const recordToDelete = allData[sheetRowIndex - 1];
    const deletedRecord = buildObjectFromRow(recordToDelete, headers);

    Logger.log('📋 Deleting record: ' + JSON.stringify(deletedRecord));

    // Delete the row
    leaveSheet.deleteRow(sheetRowIndex);

    const result = {
      rowNumber: rowNumber,
      deletedRecord: deletedRecord
    };

    // Sanitize result for web serialization
    const sanitizedResult = sanitizeForWeb(result);

    Logger.log('✅ Leave record deleted successfully');
    Logger.log('========== DELETE LEAVE COMPLETE ==========\n');

    return { success: true, data: sanitizedResult };

  } catch (error) {
    Logger.log('❌ ERROR in deleteLeave: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== DELETE LEAVE FAILED ==========\n');
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
      errors.push('Invalid date format. Please enter valid dates for Start Date and Return Date.');
    }
  } else if (data.startDate && !data.returnDate) {
    // Only start date provided - validate it
    try {
      parseDate(data.startDate);
    } catch (e) {
      errors.push('Invalid Start Date format. Please enter a valid date.');
    }
  } else if (!data.startDate && data.returnDate) {
    // Only return date provided - validate it
    try {
      parseDate(data.returnDate);
    } catch (e) {
      errors.push('Invalid Return Date format. Please enter a valid date.');
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

/**
 * Test function for editing leave
 */
function test_editLeave() {
  Logger.log('\n========== TEST: EDIT LEAVE ==========');

  // WARNING: Change row number to an actual existing row in your sheet
  const testRowNumber = 2; // Change this to test an actual row

  const testData = {
    startDate: '2025-11-10',
    returnDate: '2025-11-12',
    notes: 'Updated via test - extended leave period'
  };

  const result = editLeave(testRowNumber, testData);

  if (result.success) {
    Logger.log('✅ TEST PASSED: Leave record edited successfully');
    Logger.log('Updated record: ' + JSON.stringify(result.data));
  } else {
    Logger.log('❌ TEST FAILED: ' + result.error);
  }

  Logger.log('========== TEST COMPLETE ==========\n');
}

/**
 * Test function for deleting leave
 * WARNING: This will actually delete a row from your sheet!
 * Only run this on test data.
 */
function test_deleteLeave() {
  Logger.log('\n========== TEST: DELETE LEAVE ==========');
  Logger.log('⚠️ WARNING: This will delete an actual row from your sheet!');

  // WARNING: Change row number to an actual row you want to delete
  const testRowNumber = 999; // Set to a high number to avoid accidental deletion

  Logger.log('Attempting to delete row: ' + testRowNumber);

  const result = deleteLeave(testRowNumber);

  if (result.success) {
    Logger.log('✅ TEST PASSED: Leave record deleted successfully');
    Logger.log('Deleted record: ' + JSON.stringify(result.data));
  } else {
    Logger.log('❌ TEST FAILED: ' + result.error);
  }

  Logger.log('========== TEST COMPLETE ==========\n');
}
