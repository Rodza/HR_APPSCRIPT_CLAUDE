/**
 * IncompleteWeekDetection.gs
 *
 * Detects incomplete work weeks and prompts users to add missing days to the leave system
 * Triggered during timesheet approval workflow
 */

// ==================== INCOMPLETE WEEK DETECTION ====================

/**
 * Checks if a work week is incomplete (missing clock-in days)
 *
 * @param {string} employeeName - Employee name to check
 * @param {Date} weekEnding - Week ending date (Friday)
 * @returns {Object} Result with incomplete flag and missing days info
 */
function checkIncompleteWeek(employeeName, weekEnding) {
  try {
    Logger.log('\n========== CHECK INCOMPLETE WEEK ==========');
    Logger.log('Employee: ' + employeeName);
    Logger.log('Week Ending: ' + weekEnding);

    // Get missing days for this week
    const missingDaysResult = getMissingDaysForWeek(employeeName, weekEnding);

    if (!missingDaysResult.success) {
      throw new Error(missingDaysResult.error);
    }

    const missingDays = missingDaysResult.data.missingDays;
    const isIncomplete = missingDays.length > 0;

    Logger.log('Is Incomplete: ' + isIncomplete);
    Logger.log('Missing Days Count: ' + missingDays.length);

    if (isIncomplete) {
      Logger.log('Missing Days: ' + JSON.stringify(missingDays));
    }

    Logger.log('========== CHECK INCOMPLETE WEEK COMPLETE ==========\n');

    return {
      success: true,
      data: {
        isIncomplete: isIncomplete,
        missingDays: missingDays,
        employeeName: employeeName,
        weekEnding: weekEnding
      }
    };

  } catch (error) {
    Logger.log('❌ ERROR in checkIncompleteWeek: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    return { success: false, error: error.message };
  }
}

// ==================== GET MISSING DAYS ====================

/**
 * Identifies specific dates that are missing clock-in data for a work week
 * Checks against RAW_CLOCK_DATA and existing leave records
 *
 * @param {string} employeeName - Employee name to check
 * @param {Date} weekEnding - Week ending date (Friday)
 * @returns {Object} Result with array of missing day objects
 */
function getMissingDaysForWeek(employeeName, weekEnding) {
  try {
    Logger.log('\n========== GET MISSING DAYS FOR WEEK ==========');
    Logger.log('Employee: ' + employeeName);
    Logger.log('Week Ending: ' + weekEnding);

    const sheets = getSheets();
    const clockDataSheet = sheets.rawClockData;
    const leaveSheet = sheets.leave;

    if (!clockDataSheet) {
      throw new Error('RAW_CLOCK_DATA sheet not found');
    }

    if (!leaveSheet) {
      throw new Error('Leave sheet not found');
    }

    // Calculate the Monday of this week (work week starts Monday)
    const weekEndingDate = new Date(weekEnding);
    const monday = new Date(weekEndingDate);
    monday.setDate(weekEndingDate.getDate() - 4); // Friday - 4 = Monday

    // Generate expected work days (Mon-Fri)
    const expectedDays = [];
    for (let i = 0; i < 5; i++) {
      const day = new Date(monday);
      day.setDate(monday.getDate() + i);
      expectedDays.push({
        date: day,
        dayName: getDayOfWeek(day),
        dayOfWeek: day.getDay(), // 1=Monday, 5=Friday
        dateString: formatDateForSheet(day)
      });
    }

    Logger.log('Expected work days: ' + JSON.stringify(expectedDays.map(d => d.dateString)));

    // Get clock-in data for this employee and week
    const clockData = clockDataSheet.getDataRange().getValues();
    const clockHeaders = clockData[0];

    const empNameCol = findColumnIndex(clockHeaders, 'EMPLOYEE_NAME');
    const punchDateCol = findColumnIndex(clockHeaders, 'PUNCH_DATE');
    const deviceNameCol = findColumnIndex(clockHeaders, 'DEVICE_NAME');

    // Get all punch dates for this employee in this week (main clock-ins only, not bathroom)
    const punchDates = new Set();

    for (let i = 1; i < clockData.length; i++) {
      const row = clockData[i];
      const rowEmpName = row[empNameCol];
      const punchDate = row[punchDateCol];
      const deviceName = row[deviceNameCol];

      if (rowEmpName === employeeName && punchDate && deviceName === 'Clock In') {
        const punchDateObj = new Date(punchDate);
        const punchDateString = formatDateForSheet(punchDateObj);
        punchDates.add(punchDateString);
      }
    }

    Logger.log('Days with clock-in: ' + JSON.stringify([...punchDates]));

    // Get leave records for this employee in this week
    const leaveData = leaveSheet.getDataRange().getValues();
    const leaveHeaders = leaveData[0];

    const leaveEmpNameCol = findColumnIndex(leaveHeaders, 'EMPLOYEE NAME');
    const leaveStartCol = findColumnIndex(leaveHeaders, 'STARTDATE.LEAVE');
    const leaveReturnCol = findColumnIndex(leaveHeaders, 'RETURNDATE.LEAVE');

    // Build set of dates covered by leave
    const leaveDates = new Set();

    for (let i = 1; i < leaveData.length; i++) {
      const row = leaveData[i];
      const rowEmpName = row[leaveEmpNameCol];
      const startDate = row[leaveStartCol];
      const returnDate = row[leaveReturnCol];

      if (rowEmpName === employeeName && startDate && returnDate) {
        const start = new Date(startDate);
        const end = new Date(returnDate);

        // Add all dates in the leave range
        const current = new Date(start);
        while (current <= end) {
          const dateString = formatDateForSheet(current);
          leaveDates.add(dateString);
          current.setDate(current.getDate() + 1);
        }
      }
    }

    Logger.log('Days with leave records: ' + JSON.stringify([...leaveDates]));

    // Find missing days (not in clock-in and not in leave)
    const missingDays = [];

    for (const expectedDay of expectedDays) {
      const dateString = expectedDay.dateString;
      const hasClockin = punchDates.has(dateString);
      const hasLeave = leaveDates.has(dateString);

      if (!hasClockin && !hasLeave) {
        missingDays.push({
          date: expectedDay.date,
          dateString: dateString,
          dayName: expectedDay.dayName,
          dayOfWeek: expectedDay.dayOfWeek
        });
      }
    }

    Logger.log('Missing days: ' + JSON.stringify(missingDays.map(d => d.dateString + ' (' + d.dayName + ')')));
    Logger.log('========== GET MISSING DAYS COMPLETE ==========\n');

    return {
      success: true,
      data: {
        missingDays: missingDays,
        expectedDays: expectedDays,
        daysWithClockin: [...punchDates],
        daysWithLeave: [...leaveDates]
      }
    };

  } catch (error) {
    Logger.log('❌ ERROR in getMissingDaysForWeek: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    return { success: false, error: error.message };
  }
}

// ==================== ADD MISSING DAYS TO LEAVE ====================

/**
 * Bulk creates leave records for missing days
 *
 * @param {string} employeeName - Employee name
 * @param {Array} missingDays - Array of missing day objects from getMissingDaysForWeek
 * @param {string} reason - Leave reason (AWOL, Sick Leave, Annual Leave, Unpaid Leave)
 * @param {string} notes - Notes for the leave records
 * @returns {Object} Result with array of created leave records
 */
function addMissingDaysToLeave(employeeName, missingDays, reason, notes) {
  try {
    Logger.log('\n========== ADD MISSING DAYS TO LEAVE ==========');
    Logger.log('Employee: ' + employeeName);
    Logger.log('Missing Days Count: ' + missingDays.length);
    Logger.log('Reason: ' + reason);

    if (!missingDays || missingDays.length === 0) {
      throw new Error('No missing days provided');
    }

    // Validate reason
    const validReasons = ['SICK LEAVE - UNPAID', 'SICK LEAVE - PAID', 'AWOL', 'PAIDLEAVE', 'UNPAID LEAVE', 'FAMILY RESPONSIBILITY'];
    if (!validReasons.includes(reason)) {
      throw new Error('Invalid leave reason. Must be one of: ' + validReasons.join(', '));
    }

    const createdRecords = [];
    const errors = [];

    // Create leave record for each missing day
    for (const missingDay of missingDays) {
      const leaveData = {
        employeeName: employeeName,
        startDate: missingDay.dateString,
        returnDate: missingDay.dateString, // Same day
        reason: reason,
        notes: notes || ('Auto-added from incomplete week detection - ' + missingDay.dayName),
        imageUrl: '' // No supporting document
      };

      Logger.log('Creating leave record for: ' + missingDay.dateString + ' (' + missingDay.dayName + ')');

      const result = addLeave(leaveData);

      if (result.success) {
        createdRecords.push({
          date: missingDay.dateString,
          dayName: missingDay.dayName,
          leaveRecord: result.data
        });
        Logger.log('✅ Leave record created for ' + missingDay.dateString);
      } else {
        errors.push({
          date: missingDay.dateString,
          dayName: missingDay.dayName,
          error: result.error
        });
        Logger.log('❌ Failed to create leave record for ' + missingDay.dateString + ': ' + result.error);
      }
    }

    Logger.log('Created: ' + createdRecords.length + ' records');
    Logger.log('Errors: ' + errors.length);
    Logger.log('========== ADD MISSING DAYS TO LEAVE COMPLETE ==========\n');

    return {
      success: errors.length === 0,
      data: {
        created: createdRecords,
        errors: errors,
        totalProcessed: missingDays.length
      }
    };

  } catch (error) {
    Logger.log('❌ ERROR in addMissingDaysToLeave: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    return { success: false, error: error.message };
  }
}

// ==================== APPROVE WITH INCOMPLETE WEEK HANDLING ====================

/**
 * Modified approval function that checks for incomplete weeks
 * Returns special response if week is incomplete, prompting user action
 *
 * @param {string} id - Timesheet ID
 * @returns {Object} Result with success flag and incomplete week data if applicable
 */
function checkTimesheetBeforeApproval(id) {
  try {
    Logger.log('\n========== CHECK TIMESHEET BEFORE APPROVAL ==========');
    Logger.log('Timesheet ID: ' + id);

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

    for (let i = 1; i < sheetData.length; i++) {
      if (sheetData[i][idCol] === id) {
        timesheetRecord = buildObjectFromRow(sheetData[i], headers);
        break;
      }
    }

    if (!timesheetRecord) {
      throw new Error('Pending timesheet not found: ' + id);
    }

    // Calculate week ending date
    const importDate = timesheetRecord['IMPORTED_DATE'] ? new Date(timesheetRecord['IMPORTED_DATE']) : new Date();
    const dayOfWeek = importDate.getDay();
    let daysToFriday = 5 - dayOfWeek;
    if (daysToFriday < 0) {
      daysToFriday += 7;
    }
    const weekEndingDate = new Date(importDate);
    weekEndingDate.setDate(importDate.getDate() + daysToFriday);

    // Check for incomplete week
    const employeeName = timesheetRecord['EMPLOYEE NAME'];
    const incompleteCheck = checkIncompleteWeek(employeeName, weekEndingDate);

    if (!incompleteCheck.success) {
      throw new Error(incompleteCheck.error);
    }

    Logger.log('========== CHECK TIMESHEET BEFORE APPROVAL COMPLETE ==========\n');

    // Format missing days for frontend (remove Date objects, keep only serializable data)
    const formattedMissingDays = incompleteCheck.data.missingDays.map(day => ({
      dateString: day.dateString,
      dayName: day.dayName,
      dayOfWeek: day.dayOfWeek
    }));

    return {
      success: true,
      data: {
        timesheetId: id,
        employeeName: employeeName,
        weekEnding: formatDateForSheet(weekEndingDate), // Convert to string
        isIncomplete: incompleteCheck.data.isIncomplete,
        missingDays: formattedMissingDays
      }
    };

  } catch (error) {
    Logger.log('❌ ERROR in checkTimesheetBeforeApproval: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    return { success: false, error: error.message };
  }
}

/**
 * Approves timesheet AND creates leave records for missing days in one transaction
 *
 * @param {string} id - Timesheet ID
 * @param {Array} missingDays - Array of missing day objects
 * @param {string} leaveReason - Leave reason for missing days
 * @param {string} leaveNotes - Notes for leave records
 * @returns {Object} Result with approval and leave creation results
 */
function approveTimesheetWithLeave(id, missingDays, leaveReason, leaveNotes) {
  try {
    Logger.log('\n========== APPROVE TIMESHEET WITH LEAVE ==========');
    Logger.log('Timesheet ID: ' + id);
    Logger.log('Missing Days: ' + missingDays.length);
    Logger.log('Leave Reason: ' + leaveReason);

    // First, get employee name from timesheet
    const sheets = getSheets();
    const pendingSheet = sheets.pending;
    const sheetData = pendingSheet.getDataRange().getValues();
    const headers = sheetData[0];
    const idCol = findColumnIndex(headers, 'ID');

    let employeeName = null;

    for (let i = 1; i < sheetData.length; i++) {
      if (sheetData[i][idCol] === id) {
        const timesheetRecord = buildObjectFromRow(sheetData[i], headers);
        employeeName = timesheetRecord['EMPLOYEE NAME'];
        break;
      }
    }

    if (!employeeName) {
      throw new Error('Timesheet not found: ' + id);
    }

    // Create leave records for missing days
    let leaveResult = { success: true, data: { created: [], errors: [] } };

    if (missingDays && missingDays.length > 0) {
      Logger.log('Creating leave records for missing days...');
      leaveResult = addMissingDaysToLeave(employeeName, missingDays, leaveReason, leaveNotes);

      if (!leaveResult.success) {
        throw new Error('Failed to create leave records: ' + leaveResult.error);
      }

      if (leaveResult.data.errors.length > 0) {
        Logger.log('⚠️ Some leave records failed to create');
      }
    }

    // Approve the timesheet
    Logger.log('Approving timesheet...');
    const approvalResult = approveTimesheet(id);

    if (!approvalResult.success) {
      // Rollback is not possible for leave records already created
      // Log the issue and return error
      Logger.log('❌ Timesheet approval failed after creating leave records');
      throw new Error('Timesheet approval failed: ' + approvalResult.error);
    }

    Logger.log('✅ Timesheet approved and leave records created');
    Logger.log('========== APPROVE TIMESHEET WITH LEAVE COMPLETE ==========\n');

    return {
      success: true,
      data: {
        timesheetId: id,
        payslipRecordNumber: approvalResult.data.payslipRecordNumber,
        leaveRecordsCreated: leaveResult.data.created.length,
        leaveRecordErrors: leaveResult.data.errors.length,
        leaveRecords: leaveResult.data.created
      }
    };

  } catch (error) {
    Logger.log('❌ ERROR in approveTimesheetWithLeave: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    return { success: false, error: error.message };
  }
}

// ==================== HELPER FUNCTIONS ====================

/**
 * Formats a date object for sheet comparison (YYYY-MM-DD)
 */
function formatDateForSheet(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return year + '-' + month + '-' + day;
}
