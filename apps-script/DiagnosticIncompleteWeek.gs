/**
 * DiagnosticIncompleteWeek.gs
 *
 * Diagnostic functions to help troubleshoot incomplete week detection
 * Run these from Apps Script editor to see what data exists
 */

/**
 * Diagnostic: Check what device names exist in RAW_CLOCK_DATA
 * This helps verify if the device name filter is correct
 */
function diagnosticCheckDeviceNames() {
  try {
    Logger.log('\n========== DIAGNOSTIC: CHECK DEVICE NAMES ==========');

    const sheets = getSheets();
    const clockDataSheet = sheets.rawClockData;

    if (!clockDataSheet) {
      Logger.log('❌ RAW_CLOCK_DATA sheet not found');
      return;
    }

    const clockData = clockDataSheet.getDataRange().getValues();
    const clockHeaders = clockData[0];
    const deviceNameCol = findColumnIndex(clockHeaders, 'DEVICE_NAME');

    // Get unique device names
    const deviceNames = new Set();
    for (let i = 1; i < clockData.length; i++) {
      const deviceName = clockData[i][deviceNameCol];
      if (deviceName) {
        deviceNames.add(deviceName);
      }
    }

    Logger.log('Found ' + deviceNames.size + ' unique device names:');
    deviceNames.forEach(name => {
      Logger.log('  - "' + name + '"');
    });

    Logger.log('\n✅ Check complete. Look for the clock-in device name above.');
    Logger.log('========== DIAGNOSTIC COMPLETE ==========\n');

  } catch (error) {
    Logger.log('❌ ERROR: ' + error.message);
    Logger.log('Stack: ' + error.stack);
  }
}

/**
 * Diagnostic: Check clock-in data for a specific employee and week
 *
 * @param {string} employeeName - Employee name to check
 * @param {string} weekEndingString - Week ending date as string (YYYY-MM-DD)
 */
function diagnosticCheckEmployeeWeek(employeeName, weekEndingString) {
  try {
    Logger.log('\n========== DIAGNOSTIC: CHECK EMPLOYEE WEEK ==========');
    Logger.log('Employee: ' + employeeName);
    Logger.log('Week Ending: ' + weekEndingString);

    const sheets = getSheets();
    const clockDataSheet = sheets.rawClockData;
    const leaveSheet = sheets.leave;

    if (!clockDataSheet) {
      Logger.log('❌ RAW_CLOCK_DATA sheet not found');
      return;
    }

    // Parse week ending and calculate Monday
    const weekEnding = new Date(weekEndingString);
    const monday = new Date(weekEnding);
    monday.setDate(weekEnding.getDate() - 4);

    Logger.log('\nWeek range:');
    Logger.log('  Monday: ' + formatDateForSheet(monday));
    Logger.log('  Friday: ' + formatDateForSheet(weekEnding));

    // Get clock data
    const clockData = clockDataSheet.getDataRange().getValues();
    const clockHeaders = clockData[0];

    const empNameCol = findColumnIndex(clockHeaders, 'EMPLOYEE_NAME');
    const punchDateCol = findColumnIndex(clockHeaders, 'PUNCH_DATE');
    const deviceNameCol = findColumnIndex(clockHeaders, 'DEVICE_NAME');
    const punchTimeCol = findColumnIndex(clockHeaders, 'PUNCH_TIME');

    Logger.log('\n📊 Clock-in records for this employee in this week:');

    let recordCount = 0;
    const punchDates = new Set();

    for (let i = 1; i < clockData.length; i++) {
      const row = clockData[i];
      const rowEmpName = row[empNameCol];
      const punchDate = row[punchDateCol];
      const deviceName = row[deviceNameCol];
      const punchTime = row[punchTimeCol];

      if (rowEmpName === employeeName && punchDate) {
        const punchDateObj = new Date(punchDate);
        const punchDateString = formatDateForSheet(punchDateObj);

        // Check if in our week range
        if (punchDateObj >= monday && punchDateObj <= weekEnding) {
          recordCount++;
          Logger.log('  Row ' + (i+1) + ': ' + punchDateString + ' | Device: "' + deviceName + '" | Time: ' + punchTime);

          if (deviceName === 'Clock In') {
            punchDates.add(punchDateString);
          }
        }
      }
    }

    Logger.log('\nTotal records found: ' + recordCount);
    Logger.log('Days with "Clock In" device: ' + punchDates.size);
    Logger.log('Dates: ' + JSON.stringify([...punchDates]));

    // Check leave records
    if (leaveSheet) {
      const leaveData = leaveSheet.getDataRange().getValues();
      const leaveHeaders = leaveData[0];

      const leaveEmpNameCol = findColumnIndex(leaveHeaders, 'EMPLOYEE NAME');
      const leaveStartCol = findColumnIndex(leaveHeaders, 'STARTDATE.LEAVE');
      const leaveReturnCol = findColumnIndex(leaveHeaders, 'RETURNDATE.LEAVE');
      const leaveReasonCol = findColumnIndex(leaveHeaders, 'REASON');

      Logger.log('\n📋 Leave records for this employee:');

      let leaveCount = 0;
      for (let i = 1; i < leaveData.length; i++) {
        const row = leaveData[i];
        const rowEmpName = row[leaveEmpNameCol];
        const startDate = row[leaveStartCol];
        const returnDate = row[leaveReturnCol];
        const reason = row[leaveReasonCol];

        if (rowEmpName === employeeName && startDate && returnDate) {
          const start = new Date(startDate);
          const end = new Date(returnDate);

          // Check if overlaps with our week
          if (start <= weekEnding && end >= monday) {
            leaveCount++;
            Logger.log('  Row ' + (i+1) + ': ' + formatDateForSheet(start) + ' to ' + formatDateForSheet(end) + ' | Reason: ' + reason);
          }
        }
      }

      Logger.log('\nTotal leave records overlapping this week: ' + leaveCount);
    }

    // Expected days
    Logger.log('\n📅 Expected work days (Mon-Fri):');
    for (let i = 0; i < 5; i++) {
      const day = new Date(monday);
      day.setDate(monday.getDate() + i);
      const dayName = getDayOfWeek(day);
      const dateString = formatDateForSheet(day);
      const hasClockIn = punchDates.has(dateString);
      Logger.log('  ' + dayName + ' ' + dateString + ': ' + (hasClockIn ? '✅ Has clock-in' : '❌ Missing'));
    }

    Logger.log('\n========== DIAGNOSTIC COMPLETE ==========\n');

  } catch (error) {
    Logger.log('❌ ERROR: ' + error.message);
    Logger.log('Stack: ' + error.stack);
  }
}

/**
 * Diagnostic: Test the incomplete week check on a specific timesheet
 *
 * @param {string} timesheetId - The timesheet ID to test
 */
function diagnosticTestTimesheetCheck(timesheetId) {
  try {
    Logger.log('\n========== DIAGNOSTIC: TEST TIMESHEET CHECK ==========');
    Logger.log('Timesheet ID: ' + timesheetId);

    const result = checkTimesheetBeforeApproval(timesheetId);

    Logger.log('\n📊 Result:');
    Logger.log('Success: ' + result.success);

    if (result.success) {
      Logger.log('Employee: ' + result.data.employeeName);
      Logger.log('Week Ending: ' + result.data.weekEnding);
      Logger.log('Is Incomplete: ' + result.data.isIncomplete);
      Logger.log('Missing Days: ' + JSON.stringify(result.data.missingDays));
    } else {
      Logger.log('Error: ' + result.error);
    }

    Logger.log('\n========== DIAGNOSTIC COMPLETE ==========\n');

    return result;

  } catch (error) {
    Logger.log('❌ ERROR: ' + error.message);
    Logger.log('Stack: ' + error.stack);
  }
}
