/**
 * TRIGGERS.GS - Trigger Management
 * HR Payroll System - CRITICAL FOR AUTO-SYNC
 *
 * This file manages the onChange trigger that automatically syncs loan transactions
 * when payslips are created or edited.
 *
 * CRITICAL FUNCTION:
 * - onChange(e): Fired when MASTERSALARY sheet changes
 * - Detects loan activity and calls syncLoanForPayslip()
 */

/**
 * Install the onChange trigger on MASTERSALARY sheet
 * This should be called once during system initialization
 *
 * @returns {Object} Result object
 */
function installOnChangeTrigger() {
  try {
    Logger.log('\n========== INSTALLING ONCHANGE TRIGGER ==========');

    // First, remove any existing onChange triggers to avoid duplicates
    const existingTriggers = ScriptApp.getProjectTriggers();
    for (let trigger of existingTriggers) {
      if (trigger.getHandlerFunction() === 'onChange') {
        Logger.log('ℹ️ Removing existing onChange trigger');
        ScriptApp.deleteTrigger(trigger);
      }
    }

    // Install new trigger
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    ScriptApp.newTrigger('onChange')
      .forSpreadsheet(ss)
      .onChange()
      .create();

    Logger.log('✅ onChange trigger installed successfully');
    Logger.log('========== TRIGGER INSTALLATION COMPLETE ==========\n');

    return { success: true, message: 'Trigger installed successfully' };

  } catch (error) {
    Logger.log('❌ ERROR installing trigger: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== TRIGGER INSTALLATION FAILED ==========\n');

    return { success: false, error: error.message };
  }
}

/**
 * CRITICAL: onChange event handler
 * This function is automatically called when the spreadsheet changes
 *
 * @param {Object} e - Event object from onChange trigger
 */
function onChange(e) {
  try {
    Logger.log('\n========== ONCHANGE TRIGGER FIRED ==========');
    Logger.log('ℹ️ Event type: ' + (e ? e.changeType : 'unknown'));

    // Use lock to prevent concurrent execution
    const lock = LockService.getScriptLock();

    // Try to acquire lock (wait up to 30 seconds)
    if (!lock.tryLock(30000)) {
      Logger.log('⚠️ Could not acquire lock - another sync in progress');
      Logger.log('========== ONCHANGE ABORTED ==========\n');
      return;
    }

    try {
      // Get sheets
      const sheets = getSheets();
      const salarySheet = sheets.salary;

      if (!salarySheet) {
        Logger.log('⚠️ Salary sheet not found');
        return;
      }

      // Get all data
      const allData = salarySheet.getDataRange().getValues();
      const headers = allData[0];
      const rows = allData.slice(1);

      // Find columns
      const recordNumCol = findColumnIndex(headers, 'RECORDNUMBER');
      const timestampCol = findColumnIndex(headers, 'TIMESTAMP');
      const lastModifiedCol = findColumnIndex(headers, 'LAST_MODIFIED');
      const loanDeductionCol = findColumnIndex(headers, 'LoanDeductionThisWeek');
      const newLoanCol = findColumnIndex(headers, 'NewLoanThisWeek');
      const loanLoggedCol = findColumnIndex(headers, 'LoanRepaymentLogged');

      if (recordNumCol === -1) {
        Logger.log('⚠️ Required columns not found');
        return;
      }

      // Look for recent changes (within last 5 minutes)
      const now = new Date();
      const fiveMinutesAgo = new Date(now.getTime() - CHANGE_DETECTION_WINDOW);

      Logger.log('🔍 Scanning for changes after: ' + formatDate(fiveMinutesAgo));

      let processedCount = 0;

      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const recordNumber = row[recordNumCol];

        if (!recordNumber) continue;

        // Check if modified recently
        const lastModified = row[lastModifiedCol] ? new Date(row[lastModifiedCol]) : null;
        const timestamp = row[timestampCol] ? new Date(row[timestampCol]) : null;

        const relevantDate = lastModified || timestamp;

        if (!relevantDate || relevantDate < fiveMinutesAgo) {
          continue; // Skip old records
        }

        // Check for loan activity
        const loanDeduction = parseFloat(row[loanDeductionCol]) || 0;
        const newLoan = parseFloat(row[newLoanCol]) || 0;

        if (loanDeduction === 0 && newLoan === 0) {
          continue; // No loan activity
        }

        // Check if already logged
        const alreadyLogged = row[loanLoggedCol] === true || 
                              row[loanLoggedCol] === 'TRUE' || 
                              row[loanLoggedCol] === 'true';

        if (alreadyLogged) {
          continue; // Already processed
        }

        // Process this payslip
        Logger.log('📋 Processing payslip #' + recordNumber);
        Logger.log('  Loan Deduction: ' + formatCurrency(loanDeduction));
        Logger.log('  New Loan: ' + formatCurrency(newLoan));

        // Call sync function
        const syncResult = syncLoanForPayslip(recordNumber);

        if (syncResult.success) {
          // Mark as logged
          const sheetRowIndex = i + 2; // +1 for header, +1 for 1-based indexing

          if (loanLoggedCol >= 0) {
            salarySheet.getRange(sheetRowIndex, loanLoggedCol + 1).setValue(true);
          }

          processedCount++;
          Logger.log('✅ Loan sync completed for payslip #' + recordNumber);
        } else {
          Logger.log('❌ Loan sync failed for payslip #' + recordNumber + ': ' + syncResult.error);
        }
      }

      SpreadsheetApp.flush();

      Logger.log('✅ onChange processing complete. Processed ' + processedCount + ' payslip(s)');

    } finally {
      // Always release the lock
      lock.releaseLock();
    }

    Logger.log('========== ONCHANGE COMPLETE ==========\n');

  } catch (error) {
    Logger.log('❌ ERROR in onChange: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== ONCHANGE FAILED ==========\n');
  }
}

/**
 * Uninstall all project triggers
 * Use this for cleanup or troubleshooting
 *
 * @returns {Object} Result object
 */
function uninstallTriggers() {
  try {
    Logger.log('\n========== UNINSTALLING TRIGGERS ==========');

    const triggers = ScriptApp.getProjectTriggers();

    Logger.log('ℹ️ Found ' + triggers.length + ' trigger(s)');

    for (let trigger of triggers) {
      Logger.log('  Deleting trigger: ' + trigger.getHandlerFunction());
      ScriptApp.deleteTrigger(trigger);
    }

    Logger.log('✅ All triggers uninstalled');
    Logger.log('========== UNINSTALL COMPLETE ==========\n');

    return { success: true, message: 'All triggers uninstalled' };

  } catch (error) {
    Logger.log('❌ ERROR uninstalling triggers: ' + error.message);
    Logger.log('========== UNINSTALL FAILED ==========\n');

    return { success: false, error: error.message };
  }
}

/**
 * List all installed triggers (for debugging)
 *
 * @returns {Object} Result object with trigger list
 */
function listTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();

    const triggerList = triggers.map(trigger => ({
      handlerFunction: trigger.getHandlerFunction(),
      eventType: trigger.getEventType().toString(),
      triggerSource: trigger.getTriggerSource().toString()
    }));

    Logger.log('📋 Installed Triggers:');
    triggerList.forEach(t => {
      Logger.log('  - ' + t.handlerFunction + ' (' + t.eventType + ')');
    });

    return { success: true, data: triggerList };

  } catch (error) {
    Logger.log('❌ ERROR listing triggers: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Test the onChange trigger manually
 * This simulates the trigger firing
 */
function test_onChangeTrigger() {
  Logger.log('\n========== TEST: ONCHANGE TRIGGER ==========');

  try {
    // Call onChange directly (simulates trigger)
    onChange({ changeType: 'EDIT' });

    Logger.log('✅ TEST PASSED: onChange executed successfully');

  } catch (error) {
    Logger.log('❌ TEST FAILED: ' + error.message);
  }

  Logger.log('========== TEST COMPLETE ==========\n');
}

/**
 * Force sync all payslips with loan activity
 * Use this if auto-sync missed some records
 */
function forceSyncAllLoans() {
  try {
    Logger.log('\n========== FORCE SYNC ALL LOANS ==========');

    const sheets = getSheets();
    const salarySheet = sheets.salary;

    if (!salarySheet) {
      throw new Error('Salary sheet not found');
    }

    const allData = salarySheet.getDataRange().getValues();
    const headers = allData[0];
    const rows = allData.slice(1);

    const recordNumCol = findColumnIndex(headers, 'RECORDNUMBER');
    const loanDeductionCol = findColumnIndex(headers, 'LoanDeductionThisWeek');
    const newLoanCol = findColumnIndex(headers, 'NewLoanThisWeek');

    let processedCount = 0;
    let successCount = 0;

    for (let row of rows) {
      const recordNumber = row[recordNumCol];
      if (!recordNumber) continue;

      const loanDeduction = parseFloat(row[loanDeductionCol]) || 0;
      const newLoan = parseFloat(row[newLoanCol]) || 0;

      if (loanDeduction === 0 && newLoan === 0) continue;

      processedCount++;
      Logger.log('📋 Processing payslip #' + recordNumber);

      const result = syncLoanForPayslip(recordNumber);

      if (result.success) {
        successCount++;
      } else {
        Logger.log('❌ Failed: ' + result.error);
      }
    }

    Logger.log('✅ Force sync complete');
    Logger.log('  Processed: ' + processedCount);
    Logger.log('  Successful: ' + successCount);
    Logger.log('========== FORCE SYNC COMPLETE ==========\n');

    return {
      success: true,
      data: {
        processed: processedCount,
        successful: successCount
      }
    };

  } catch (error) {
    Logger.log('❌ ERROR in forceSyncAllLoans: ' + error.message);
    Logger.log('========== FORCE SYNC FAILED ==========\n');

    return { success: false, error: error.message };
  }
}
