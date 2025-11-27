/**
 * MIGRATION SCRIPT: Add "OTHER INCOME TEXT" column to MASTERSALARY sheet
 *
 * PROBLEM: Added new field "OTHER INCOME TEXT" to Config.gs but column doesn't exist in spreadsheet
 * This script adds the column in the correct position (after OTHERINCOME)
 *
 * BEFORE RUNNING:
 * 1. Make a backup of your MASTERSALARY sheet!
 * 2. Run this ONCE only
 *
 * HOW TO USE:
 * 1. Open Apps Script Editor
 * 2. Select function: addOtherIncomeTextColumn
 * 3. Click Run
 * 4. Check execution log for results
 */

function addOtherIncomeTextColumn() {
  try {
    Logger.log('\n========== ADD OTHER INCOME TEXT COLUMN ==========');
    Logger.log('⚠️ This script adds a new column to the payslip sheet');
    Logger.log('⚠️ Make sure you have a backup before running!');

    const sheets = getSheets();
    const salarySheet = sheets.salary;

    if (!salarySheet) {
      throw new Error('Salary sheet not found (MASTERSALARY/salaryrecords/payroll)');
    }

    // Get current headers
    const data = salarySheet.getDataRange().getValues();
    const headers = data[0];

    Logger.log('📋 Current sheet: ' + salarySheet.getName());
    Logger.log('📊 Total rows: ' + data.length);
    Logger.log('📊 Current columns: ' + headers.length);

    // Find the OTHERINCOME column
    let otherIncomeIndex = -1;
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] === 'OTHERINCOME') {
        otherIncomeIndex = i;
        break;
      }
    }

    if (otherIncomeIndex === -1) {
      throw new Error('OTHERINCOME column not found. Please check your sheet structure.');
    }

    Logger.log('✅ Found OTHERINCOME at column ' + (otherIncomeIndex + 1));

    // Check if OTHER INCOME TEXT already exists
    let otherIncomeTextExists = false;
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] === 'OTHER INCOME TEXT') {
        otherIncomeTextExists = true;
        Logger.log('⚠️ OTHER INCOME TEXT column already exists at column ' + (i + 1));
        break;
      }
    }

    if (otherIncomeTextExists) {
      Logger.log('ℹ️ Migration not needed - column already exists');
      Logger.log('========== MIGRATION COMPLETE ==========\n');
      return { success: true, message: 'Column already exists' };
    }

    // Insert new column after OTHERINCOME
    const insertPosition = otherIncomeIndex + 2; // +2 because columns are 1-indexed
    Logger.log('🔄 Inserting new column at position ' + insertPosition + ' (after OTHERINCOME)...');

    salarySheet.insertColumnAfter(otherIncomeIndex + 1);

    // Set the header for the new column
    salarySheet.getRange(1, insertPosition).setValue('OTHER INCOME TEXT');

    SpreadsheetApp.flush();

    Logger.log('✅ Column added successfully!');
    Logger.log('   Column name: OTHER INCOME TEXT');
    Logger.log('   Position: ' + insertPosition);
    Logger.log('   After: OTHERINCOME (column ' + (otherIncomeIndex + 1) + ')');
    Logger.log('========== MIGRATION COMPLETE ==========\n');

    return {
      success: true,
      message: 'OTHER INCOME TEXT column added at position ' + insertPosition
    };

  } catch (error) {
    Logger.log('❌ ERROR: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== MIGRATION FAILED ==========\n');
    return { success: false, error: error.message };
  }
}

/**
 * Verify the column was added successfully
 */
function verifyOtherIncomeTextColumn() {
  try {
    Logger.log('\n========== VERIFY OTHER INCOME TEXT COLUMN ==========');

    const sheets = getSheets();
    const salarySheet = sheets.salary;

    if (!salarySheet) {
      throw new Error('Salary sheet not found');
    }

    const headers = salarySheet.getRange(1, 1, 1, salarySheet.getLastColumn()).getValues()[0];

    Logger.log('📊 Total columns: ' + headers.length);

    // Find both columns
    let otherIncomeIndex = -1;
    let otherIncomeTextIndex = -1;

    for (let i = 0; i < headers.length; i++) {
      if (headers[i] === 'OTHERINCOME') {
        otherIncomeIndex = i;
      }
      if (headers[i] === 'OTHER INCOME TEXT') {
        otherIncomeTextIndex = i;
      }
    }

    if (otherIncomeIndex === -1) {
      Logger.log('❌ OTHERINCOME column not found!');
      return { success: false, message: 'OTHERINCOME column missing' };
    }

    if (otherIncomeTextIndex === -1) {
      Logger.log('❌ OTHER INCOME TEXT column not found!');
      return { success: false, message: 'OTHER INCOME TEXT column missing' };
    }

    // Verify it's in the correct position (right after OTHERINCOME)
    if (otherIncomeTextIndex !== otherIncomeIndex + 1) {
      Logger.log('⚠️ WARNING: OTHER INCOME TEXT is not immediately after OTHERINCOME');
      Logger.log('   OTHERINCOME at column ' + (otherIncomeIndex + 1));
      Logger.log('   OTHER INCOME TEXT at column ' + (otherIncomeTextIndex + 1));
      Logger.log('   Expected position: ' + (otherIncomeIndex + 2));
      return {
        success: false,
        message: 'Column exists but in wrong position'
      };
    }

    Logger.log('✅ Verification successful!');
    Logger.log('   OTHERINCOME at column ' + (otherIncomeIndex + 1));
    Logger.log('   OTHER INCOME TEXT at column ' + (otherIncomeTextIndex + 1));
    Logger.log('========== VERIFICATION COMPLETE ==========\n');

    return { success: true, message: 'Column verified in correct position' };

  } catch (error) {
    Logger.log('❌ ERROR: ' + error.message);
    Logger.log('========== VERIFICATION FAILED ==========\n');
    return { success: false, error: error.message };
  }
}
