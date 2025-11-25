/**
 * DIAGNOSTIC SCRIPT: PaidtoAccount Column Issue
 *
 * This script helps identify why the PaidtoAccount column shows 0 in reports.
 * Run test_PaidtoAccount_Diagnostic() from the Apps Script Editor.
 */

/**
 * Main diagnostic function - Run this to diagnose the PaidtoAccount issue
 */
function test_PaidtoAccount_Diagnostic() {
  Logger.log('\n========== PAID TO ACCOUNT DIAGNOSTIC ==========\n');

  try {
    // Test 1: Check MASTERSALARY sheet headers
    Logger.log('TEST 1: Checking MASTERSALARY Sheet Headers');
    Logger.log('─'.repeat(60));
    checkMasterSalaryHeaders();

    // Test 2: Check recent payslip data
    Logger.log('\nTEST 2: Checking Recent Payslip Data');
    Logger.log('─'.repeat(60));
    checkRecentPayslipData();

    // Test 3: Test the listPayslips function
    Logger.log('\nTEST 3: Testing listPayslips Function');
    Logger.log('─'.repeat(60));
    testListPayslipsFunction();

    // Test 4: Simulate report data retrieval
    Logger.log('\nTEST 4: Simulating Report Data Retrieval');
    Logger.log('─'.repeat(60));
    simulateReportRetrieval();

    Logger.log('\n========== DIAGNOSTIC COMPLETE ==========\n');
    Logger.log('📋 Review the logs above to identify the issue.');
    Logger.log('📋 Look for any column name mismatches or unexpected values.\n');

  } catch (error) {
    Logger.log('❌ ERROR in diagnostic: ' + error.message);
    Logger.log('Stack: ' + error.stack);
  }
}

/**
 * Test 1: Check the actual column headers in MASTERSALARY sheet
 */
function checkMasterSalaryHeaders() {
  try {
    const sheets = getSheets();
    const salarySheet = sheets.salary;

    if (!salarySheet) {
      Logger.log('❌ MASTERSALARY sheet not found!');
      return;
    }

    const headers = salarySheet.getRange(1, 1, 1, salarySheet.getLastColumn()).getValues()[0];

    Logger.log('✅ Found ' + headers.length + ' columns in MASTERSALARY sheet\n');

    // Look for PaidtoAccount column (case-sensitive search)
    let paidtoAccountIndex = -1;
    let similarColumns = [];

    for (let i = 0; i < headers.length; i++) {
      const header = String(headers[i]);

      // Exact match
      if (header === 'PaidtoAccount') {
        paidtoAccountIndex = i;
        Logger.log('✅ FOUND exact match: Column ' + (i + 1) + ' = "' + header + '"');
      }

      // Look for similar column names (case-insensitive)
      if (header.toLowerCase().includes('paid') && header.toLowerCase().includes('account')) {
        similarColumns.push({
          index: i + 1,
          name: header,
          exact: header === 'PaidtoAccount'
        });
      }
    }

    if (paidtoAccountIndex === -1) {
      Logger.log('❌ ISSUE FOUND: "PaidtoAccount" column NOT found (exact match)');
      Logger.log('');

      if (similarColumns.length > 0) {
        Logger.log('⚠️  Found similar column names:');
        similarColumns.forEach(col => {
          Logger.log('   - Column ' + col.index + ': "' + col.name + '"');
          Logger.log('     ASCII codes: ' + getAsciiCodes(col.name));
        });
        Logger.log('');
        Logger.log('💡 SOLUTION: The column name might have:');
        Logger.log('   1. Different capitalization (e.g., "PaidToAccount" vs "PaidtoAccount")');
        Logger.log('   2. Extra spaces (e.g., "Paid to Account" or "PaidtoAccount ")');
        Logger.log('   3. Hidden characters');
        Logger.log('');
        Logger.log('📝 Expected: "PaidtoAccount" (lowercase "t" in "to")');
      } else {
        Logger.log('❌ No similar column found! The PaidtoAccount column might be missing entirely.');
      }
    } else {
      Logger.log('✅ Column "PaidtoAccount" found at position ' + (paidtoAccountIndex + 1));
    }

    Logger.log('\n📋 All column headers:');
    headers.forEach((header, index) => {
      Logger.log('   ' + (index + 1) + '. "' + header + '"');
    });

  } catch (error) {
    Logger.log('❌ Error checking headers: ' + error.message);
  }
}

/**
 * Test 2: Check recent payslip records for PaidtoAccount values
 */
function checkRecentPayslipData() {
  try {
    const sheets = getSheets();
    const salarySheet = sheets.salary;

    if (!salarySheet) {
      Logger.log('❌ MASTERSALARY sheet not found!');
      return;
    }

    const allData = salarySheet.getDataRange().getValues();
    const headers = allData[0];
    const rows = allData.slice(1);

    if (rows.length === 0) {
      Logger.log('⚠️  No payslip data found in MASTERSALARY sheet');
      return;
    }

    // Find relevant column indices
    const recordNumCol = findColumnIndex(headers, 'RECORDNUMBER');
    const empNameCol = findColumnIndex(headers, 'EMPLOYEE NAME');
    const weekEndingCol = findColumnIndex(headers, 'WEEKENDING');
    const netSalaryCol = findColumnIndex(headers, 'NETTSALARY');

    // Try to find PaidtoAccount column
    let paidtoAccountCol = -1;
    try {
      paidtoAccountCol = findColumnIndex(headers, 'PaidtoAccount');
    } catch (e) {
      Logger.log('❌ Cannot find "PaidtoAccount" column using findColumnIndex');
    }

    // Manually search for any "paid" + "account" column
    let manualPaidCol = -1;
    for (let i = 0; i < headers.length; i++) {
      const header = String(headers[i]).toLowerCase();
      if (header.includes('paid') && header.includes('account')) {
        manualPaidCol = i;
        Logger.log('📍 Found potential column at index ' + i + ': "' + headers[i] + '"');
        break;
      }
    }

    // Get last 5 payslips
    const recentRows = rows.slice(-5);

    Logger.log('\n📋 Last ' + recentRows.length + ' payslip records:\n');

    recentRows.forEach((row, index) => {
      Logger.log('Record #' + row[recordNumCol]);
      Logger.log('  Employee: ' + row[empNameCol]);
      Logger.log('  Week Ending: ' + formatDate(row[weekEndingCol]));
      Logger.log('  Net Salary: R' + (row[netSalaryCol] || 0).toFixed(2));

      if (paidtoAccountCol >= 0) {
        Logger.log('  PaidtoAccount (using findColumnIndex): R' + (row[paidtoAccountCol] || 0).toFixed(2));
      } else {
        Logger.log('  PaidtoAccount (using findColumnIndex): ❌ Column not found');
      }

      if (manualPaidCol >= 0) {
        Logger.log('  Value in column "' + headers[manualPaidCol] + '": R' + (row[manualPaidCol] || 0).toFixed(2));
      }

      Logger.log('');
    });

  } catch (error) {
    Logger.log('❌ Error checking payslip data: ' + error.message);
    Logger.log('Stack: ' + error.stack);
  }
}

/**
 * Test 3: Test the listPayslips function
 */
function testListPayslipsFunction() {
  try {
    // Get most recent week ending date
    const sheets = getSheets();
    const salarySheet = sheets.salary;
    const allData = salarySheet.getDataRange().getValues();
    const headers = allData[0];
    const rows = allData.slice(1);

    if (rows.length === 0) {
      Logger.log('⚠️  No payslip data found');
      return;
    }

    const weekEndingCol = findColumnIndex(headers, 'WEEKENDING');
    const lastRow = rows[rows.length - 1];
    const weekEnding = lastRow[weekEndingCol];

    Logger.log('Testing listPayslips for week ending: ' + formatDate(weekEnding));

    const result = listPayslips({ weekEnding: weekEnding });

    if (!result.success) {
      Logger.log('❌ listPayslips failed: ' + result.error);
      return;
    }

    const payslips = result.data;
    Logger.log('✅ listPayslips returned ' + payslips.length + ' payslips\n');

    if (payslips.length > 0) {
      const firstPayslip = payslips[0];

      Logger.log('📋 First payslip object keys:');
      const keys = Object.keys(firstPayslip);
      keys.forEach(key => {
        if (key.toLowerCase().includes('paid') || key.toLowerCase().includes('account')) {
          Logger.log('   🔍 "' + key + '" = ' + firstPayslip[key]);
        }
      });

      Logger.log('\n📋 Checking PaidtoAccount field:');
      if (firstPayslip.hasOwnProperty('PaidtoAccount')) {
        Logger.log('   ✅ "PaidtoAccount" exists');
        Logger.log('   Value: ' + firstPayslip['PaidtoAccount']);
        Logger.log('   Type: ' + typeof firstPayslip['PaidtoAccount']);
      } else {
        Logger.log('   ❌ "PaidtoAccount" field NOT found in payslip object');
        Logger.log('   ⚠️  This is why the report shows 0!');
      }

      Logger.log('\n📋 Sample payslip data:');
      Logger.log('   Employee: ' + firstPayslip['EMPLOYEE NAME']);
      Logger.log('   Net Salary: ' + firstPayslip['NETTSALARY']);
      Logger.log('   PaidtoAccount: ' + (firstPayslip['PaidtoAccount'] || 'UNDEFINED'));
    }

  } catch (error) {
    Logger.log('❌ Error testing listPayslips: ' + error.message);
    Logger.log('Stack: ' + error.stack);
  }
}

/**
 * Test 4: Simulate how the report retrieves PaidtoAccount
 */
function simulateReportRetrieval() {
  try {
    Logger.log('Simulating Weekly Payroll Summary Report data retrieval...\n');

    // Get most recent week ending date
    const sheets = getSheets();
    const salarySheet = sheets.salary;
    const allData = salarySheet.getDataRange().getValues();
    const headers = allData[0];
    const rows = allData.slice(1);

    if (rows.length === 0) {
      Logger.log('⚠️  No payslip data found');
      return;
    }

    const weekEndingCol = findColumnIndex(headers, 'WEEKENDING');
    const lastRow = rows[rows.length - 1];
    const weekEnd = lastRow[weekEndingCol];

    Logger.log('Simulating report for week ending: ' + formatDate(weekEnd));

    // This is exactly what the report does
    const payslipsResult = listPayslips({ weekEnding: weekEnd });

    if (!payslipsResult.success) {
      Logger.log('❌ Failed to get payslips: ' + payslipsResult.error);
      return;
    }

    const payslips = payslipsResult.data;
    Logger.log('✅ Retrieved ' + payslips.length + ' payslips\n');

    let totalPaidToAccounts = 0;

    Logger.log('📊 Processing payslips (like the report does):\n');

    for (let i = 0; i < Math.min(3, payslips.length); i++) {
      const p = payslips[i];
      const paidToAccount = p['PaidtoAccount'] || 0;  // This is the line in Reports.gs:376

      totalPaidToAccounts += paidToAccount;

      Logger.log('Payslip ' + (i + 1) + ':');
      Logger.log('  Employee: ' + p['EMPLOYEE NAME']);
      Logger.log('  Net Salary: R' + (p['NETTSALARY'] || 0).toFixed(2));
      Logger.log('  p[\'PaidtoAccount\']: ' + paidToAccount);
      Logger.log('  Running Total: R' + totalPaidToAccounts.toFixed(2));
      Logger.log('');
    }

    if (payslips.length > 3) {
      Logger.log('... (' + (payslips.length - 3) + ' more payslips)\n');

      // Calculate full total
      for (let i = 3; i < payslips.length; i++) {
        totalPaidToAccounts += payslips[i]['PaidtoAccount'] || 0;
      }
    }

    Logger.log('═'.repeat(60));
    Logger.log('📊 REPORT TOTALS:');
    Logger.log('   Total Employees: ' + payslips.length);
    Logger.log('   Total Paid to Accounts: R' + totalPaidToAccounts.toFixed(2));
    Logger.log('═'.repeat(60));

    if (totalPaidToAccounts === 0) {
      Logger.log('\n❌ CONFIRMED: Total Paid to Accounts is 0!');
      Logger.log('');
      Logger.log('💡 This means p[\'PaidtoAccount\'] is returning 0 or undefined');
      Logger.log('   for all payslips. Check the column name issue identified in TEST 1.');
    } else {
      Logger.log('\n✅ Total is NOT 0. The issue might be with a specific week or filters.');
    }

  } catch (error) {
    Logger.log('❌ Error simulating report: ' + error.message);
    Logger.log('Stack: ' + error.stack);
  }
}

/**
 * Helper: Get ASCII codes for a string (to identify hidden characters)
 */
function getAsciiCodes(str) {
  let codes = [];
  for (let i = 0; i < str.length; i++) {
    codes.push(str.charCodeAt(i));
  }
  return codes.join(',');
}

/**
 * Helper: Format date for display
 */
function formatDate(date) {
  if (!date) return 'N/A';
  const d = new Date(date);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return year + '-' + month + '-' + day;
}
