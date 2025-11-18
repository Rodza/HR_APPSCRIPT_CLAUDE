/**
 * MIGRATION SCRIPT: Fix PENDING_TIMESHEETS column order
 *
 * PROBLEM: createEnhancedPendingTimesheet() was inserting data in wrong column order
 * This script fixes existing records that were inserted with buggy column order
 *
 * BEFORE RUNNING:
 * 1. Make a backup of your PENDING_TIMESHEETS sheet!
 * 2. Run this ONCE only
 *
 * HOW TO USE:
 * 1. Open Apps Script Editor
 * 2. Select function: migratePendingTimesheetsData
 * 3. Click Run
 * 4. Check execution log for results
 */

function migratePendingTimesheetsData() {
  try {
    Logger.log('\n========== MIGRATE PENDING TIMESHEETS DATA ==========');
    Logger.log('⚠️ This script fixes column order for existing records');
    Logger.log('⚠️ Make sure you have a backup before running!');

    const sheets = getSheets();
    const pendingSheet = sheets.pending;

    if (!pendingSheet) {
      throw new Error('PENDING_TIMESHEETS sheet not found');
    }

    const data = pendingSheet.getDataRange().getValues();
    const headers = data[0];

    // Expected header positions (Config.gs:326-351)
    const expectedHeaders = [
      'ID',                      // 0
      'RAW_DATA_IMPORT_ID',      // 1
      'EMPLOYEE_ID',             // 2
      'EMPLOYEE NAME',           // 3
      'EMPLOYEE_CLOCK_REF',      // 4
      'WEEKENDING',              // 5
      'CALCULATED_TOTAL_HOURS',  // 6
      'CALCULATED_TOTAL_MINUTES',// 7
      'HOURS',                   // 8
      'MINUTES',                 // 9
      'OVERTIMEHOURS',           // 10
      'OVERTIMEMINUTES',         // 11
      'LUNCH_DEDUCTION_MIN',     // 12
      'BATHROOM_TIME_MIN',       // 13
      'RECON_DETAILS',           // 14
      'WARNINGS',                // 15
      'NOTES',                   // 16
      'STATUS',                  // 17
      'IMPORTED_BY',             // 18
      'IMPORTED_DATE',           // 19
      'REVIEWED_BY',             // 20
      'REVIEWED_DATE',           // 21
      'PAYSLIP_ID',              // 22
      'IS_LOCKED',               // 23
      'LOCKED_DATE'              // 24
    ];

    // Verify headers match
    Logger.log('📋 Verifying headers...');
    let headerMismatch = false;
    for (let i = 0; i < expectedHeaders.length; i++) {
      if (headers[i] !== expectedHeaders[i]) {
        Logger.log('⚠️ Header mismatch at column ' + (i+1) + ': Expected "' + expectedHeaders[i] + '", Found "' + headers[i] + '"');
        headerMismatch = true;
      }
    }

    if (headerMismatch) {
      throw new Error('Headers do not match expected format. Please check your sheet structure.');
    }

    Logger.log('✅ Headers verified');
    Logger.log('📊 Total rows: ' + data.length);

    if (data.length === 1) {
      Logger.log('ℹ️ No data rows to migrate (sheet only has headers)');
      Logger.log('========== MIGRATION COMPLETE ==========\n');
      return { success: true, message: 'No data to migrate' };
    }

    // Check if migration needed by looking for 'Pending' in NOTES column (col 16)
    let needsMigration = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][16] === 'Pending' || data[i][16] === 'Approved' || data[i][16] === 'Rejected') {
        needsMigration = true;
        break;
      }
    }

    if (!needsMigration) {
      Logger.log('✅ Data appears to already be in correct format (no STATUS values found in NOTES column)');
      Logger.log('========== MIGRATION NOT NEEDED ==========\n');
      return { success: true, message: 'Data already migrated or correct' };
    }

    Logger.log('⚠️ Migration needed! Found STATUS values in NOTES column');
    Logger.log('🔄 Starting migration...');

    let migratedCount = 0;

    // Process each row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowNum = i + 1;

      // Skip empty rows
      if (!row[0]) continue;

      // OLD buggy order was:
      // 16: STATUS value
      // 17: IMPORTED_DATE value
      // 18: IMPORTED_BY value
      // 19: '' (REVIEWED_DATE)
      // 20: '' (REVIEWED_BY)
      // 21: '' (PAYSLIP_ID)
      // 22: NOTES value
      // 23: IS_LOCKED
      // 24: LOCKED_DATE

      // NEW correct order:
      // 16: NOTES
      // 17: STATUS
      // 18: IMPORTED_BY
      // 19: IMPORTED_DATE
      // 20: REVIEWED_BY
      // 21: REVIEWED_DATE
      // 22: PAYSLIP_ID
      // 23: IS_LOCKED
      // 24: LOCKED_DATE

      // Check if this row needs migration (has STATUS in NOTES position)
      if (row[16] === 'Pending' || row[16] === 'Approved' || row[16] === 'Rejected') {
        Logger.log('  Migrating row ' + rowNum + '...');

        // Extract values from OLD positions
        const oldNotes = row[22] || '';            // Was in position 22
        const oldStatus = row[16];                 // Was in position 16
        const oldImportedBy = row[18];             // Correct position
        const oldImportedDate = row[17];           // Was in position 17
        const oldReviewedBy = row[20] || '';       // Was in position 20
        const oldReviewedDate = row[19] || '';     // Was in position 19
        const oldPayslipId = row[21] || '';        // Was in position 21
        const oldIsLocked = row[23];               // Correct position
        const oldLockedDate = row[24] || '';       // Correct position

        // Write to NEW correct positions
        pendingSheet.getRange(rowNum, 17).setValue(oldNotes);        // NOTES
        pendingSheet.getRange(rowNum, 18).setValue(oldStatus);       // STATUS
        pendingSheet.getRange(rowNum, 19).setValue(oldImportedBy);   // IMPORTED_BY
        pendingSheet.getRange(rowNum, 20).setValue(oldImportedDate); // IMPORTED_DATE
        pendingSheet.getRange(rowNum, 21).setValue(oldReviewedBy);   // REVIEWED_BY
        pendingSheet.getRange(rowNum, 22).setValue(oldReviewedDate); // REVIEWED_DATE
        pendingSheet.getRange(rowNum, 23).setValue(oldPayslipId);    // PAYSLIP_ID
        pendingSheet.getRange(rowNum, 24).setValue(oldIsLocked);     // IS_LOCKED
        pendingSheet.getRange(rowNum, 25).setValue(oldLockedDate);   // LOCKED_DATE

        migratedCount++;
      }
    }

    SpreadsheetApp.flush();

    Logger.log('✅ Migration complete!');
    Logger.log('   Rows migrated: ' + migratedCount);
    Logger.log('   Total rows: ' + (data.length - 1));
    Logger.log('========== MIGRATION COMPLETE ==========\n');

    return {
      success: true,
      message: 'Migrated ' + migratedCount + ' rows successfully'
    };

  } catch (error) {
    Logger.log('❌ ERROR: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== MIGRATION FAILED ==========\n');
    return { success: false, error: error.message };
  }
}

/**
 * Verify migration was successful
 */
function verifyMigration() {
  try {
    Logger.log('\n========== VERIFY MIGRATION ==========');

    const sheets = getSheets();
    const pendingSheet = sheets.pending;

    if (!pendingSheet) {
      throw new Error('PENDING_TIMESHEETS sheet not found');
    }

    const data = pendingSheet.getDataRange().getValues();
    Logger.log('📊 Total rows: ' + data.length);

    if (data.length === 1) {
      Logger.log('ℹ️ No data rows to verify');
      return { success: true, message: 'No data' };
    }

    let issuesFound = 0;

    // Check each row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowNum = i + 1;

      // Skip empty rows
      if (!row[0]) continue;

      // Column 17 should be STATUS (Pending/Approved/Rejected)
      const status = row[17];
      if (status !== 'Pending' && status !== 'Approved' && status !== 'Rejected' && status !== '') {
        Logger.log('⚠️ Row ' + rowNum + ': STATUS column has unexpected value: ' + status);
        issuesFound++;
      }

      // Column 16 should be NOTES (empty or text, NOT a status)
      const notes = row[16];
      if (notes === 'Pending' || notes === 'Approved' || notes === 'Rejected') {
        Logger.log('❌ Row ' + rowNum + ': NOTES column still has STATUS value: ' + notes);
        issuesFound++;
      }
    }

    if (issuesFound > 0) {
      Logger.log('⚠️ Found ' + issuesFound + ' issues');
      Logger.log('========== VERIFICATION FAILED ==========\n');
      return { success: false, message: issuesFound + ' issues found' };
    } else {
      Logger.log('✅ All rows verified successfully!');
      Logger.log('========== VERIFICATION COMPLETE ==========\n');
      return { success: true, message: 'All rows correct' };
    }

  } catch (error) {
    Logger.log('❌ ERROR: ' + error.message);
    Logger.log('========== VERIFICATION FAILED ==========\n');
    return { success: false, error: error.message };
  }
}
