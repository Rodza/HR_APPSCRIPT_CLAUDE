/**
 * DIAGNOSTIC SCRIPT FOR PAYROLL MODULE
 * This script helps identify where null/undefined errors are occurring
 */

/**
 * Test payslip creation with full error tracking
 */
function diagnosticTestCreatePayslip() {
  console.log('========== DIAGNOSTIC TEST: CREATE PAYSLIP ==========');

  try {
    // Test data
    const testData = {
      employeeName: 'John Doe',  // Change to a real employee name from your sheet
      weekEnding: '2025-11-15',
      hours: 40,
      minutes: 0,
      overtimeHours: 5,
      overtimeMinutes: 30,
      leavePay: 0,
      bonusPay: 0,
      otherIncome: 0,
      otherDeductions: 0,
      otherDeductionsText: '',
      loanDeductionThisWeek: 0,
      newLoanThisWeek: 0,
      loanDisbursementType: 'Separate',
      notes: 'Diagnostic test'
    };

    console.log('✓ Test data created:', JSON.stringify(testData));

    // Step 1: Test getEmployeeByName
    console.log('\n--- STEP 1: Testing getEmployeeByName ---');
    const empResult = safeCall(() => getEmployeeByName(testData.employeeName), 'getEmployeeByName');

    if (!empResult) {
      console.error('❌ FATAL: getEmployeeByName returned null/undefined');
      return { success: false, error: 'getEmployeeByName returned null', step: 1 };
    }

    if (!empResult.success) {
      console.error('❌ FATAL: Employee not found:', testData.employeeName);
      console.error('Available employees:');
      listAvailableEmployees();
      return { success: false, error: 'Employee not found: ' + testData.employeeName, step: 1 };
    }

    console.log('✓ Employee found:', empResult.data.REFNAME);
    console.log('  - ID:', empResult.data.id);
    console.log('  - EMPLOYER:', empResult.data.EMPLOYER);
    console.log('  - EMPLOYMENT STATUS:', empResult.data['EMPLOYMENT STATUS']);
    console.log('  - HOURLY RATE:', empResult.data['HOURLY RATE']);

    const employee = empResult.data;

    // Step 2: Test getCurrentLoanBalance
    console.log('\n--- STEP 2: Testing getCurrentLoanBalance ---');
    const loanBalance = safeCall(() => getCurrentLoanBalance(employee.id), 'getCurrentLoanBalance');

    if (loanBalance === null || loanBalance === undefined) {
      console.warn('⚠️ WARNING: getCurrentLoanBalance returned null/undefined, using 0');
      testData.CurrentLoanBalance = 0;
    } else {
      console.log('✓ Loan balance retrieved:', loanBalance);
      testData.CurrentLoanBalance = loanBalance;
    }

    // Step 3: Test field mapping
    console.log('\n--- STEP 3: Testing field mapping ---');
    const mappedData = mapUIFieldsToInternal(testData);
    console.log('✓ Fields mapped successfully');

    // Step 4: Enrich with employee data
    console.log('\n--- STEP 4: Enriching with employee data ---');
    mappedData.id = employee.id;
    mappedData['EMPLOYEE NAME'] = testData.employeeName;
    mappedData.EMPLOYER = employee.EMPLOYER;
    mappedData['EMPLOYMENT STATUS'] = employee['EMPLOYMENT STATUS'];
    mappedData.HOURLYRATE = employee['HOURLY RATE'];
    console.log('✓ Data enriched with employee details');

    // Step 5: Test calculatePayslip
    console.log('\n--- STEP 5: Testing calculatePayslip ---');
    const calculations = safeCall(() => calculatePayslip(mappedData), 'calculatePayslip');

    if (!calculations) {
      console.error('❌ FATAL: calculatePayslip returned null/undefined');
      return { success: false, error: 'calculatePayslip returned null', step: 5 };
    }

    console.log('✓ Calculations completed:', JSON.stringify(calculations));

    // Step 6: Test sheet operations
    console.log('\n--- STEP 6: Testing sheet access ---');
    const sheets = safeCall(() => getSheets(), 'getSheets');

    if (!sheets) {
      console.error('❌ FATAL: getSheets returned null/undefined');
      return { success: false, error: 'getSheets returned null', step: 6 };
    }

    if (!sheets.salary) {
      console.error('❌ FATAL: Salary sheet not found in sheets object');
      console.error('Available sheets:', Object.keys(sheets));
      return { success: false, error: 'Salary sheet not found', step: 6 };
    }

    console.log('✓ Salary sheet found:', sheets.salary.getName());

    // Step 7: Test header access
    console.log('\n--- STEP 7: Testing sheet headers ---');
    const salarySheet = sheets.salary;
    const allData = salarySheet.getDataRange().getValues();
    const headers = allData[0];

    console.log('✓ Headers retrieved. Column count:', headers.length);
    console.log('  Headers:', headers.join(', '));

    // Step 8: Test utility functions
    console.log('\n--- STEP 8: Testing utility functions ---');

    const testRound = safeCall(() => roundTo(123.456789, 2), 'roundTo');
    console.log('✓ roundTo(123.456789, 2):', testRound);

    const testFormat = safeCall(() => formatCurrency(1234.56), 'formatCurrency');
    console.log('✓ formatCurrency(1234.56):', testFormat);

    const testAudit = {};
    safeCall(() => addAuditFields(testAudit, true), 'addAuditFields');
    console.log('✓ addAuditFields:', JSON.stringify(testAudit));

    const testRow = safeCall(() => objectToRow(mappedData, headers), 'objectToRow');
    console.log('✓ objectToRow: Generated row with', testRow ? testRow.length : 0, 'columns');

    console.log('\n========== ALL DIAGNOSTIC TESTS PASSED ==========');
    return {
      success: true,
      message: 'All diagnostic tests passed',
      data: {
        employee: employee.REFNAME,
        calculations: calculations,
        sheetColumns: headers.length
      }
    };

  } catch (error) {
    console.error('\n❌ DIAGNOSTIC TEST FAILED');
    console.error('Error:', error.toString());
    console.error('Stack:', error.stack);
    return {
      success: false,
      error: error.toString(),
      stack: error.stack
    };
  }
}

/**
 * Safe function call wrapper with error tracking
 */
function safeCall(fn, functionName) {
  try {
    console.log(`  Calling ${functionName}...`);
    const result = fn();

    if (result === null) {
      console.warn(`  ⚠️ ${functionName} returned NULL`);
    } else if (result === undefined) {
      console.warn(`  ⚠️ ${functionName} returned UNDEFINED`);
    } else {
      console.log(`  ✓ ${functionName} completed successfully`);
    }

    return result;
  } catch (error) {
    console.error(`  ❌ ${functionName} threw an error:`, error.toString());
    console.error(`  Stack:`, error.stack);
    throw error;
  }
}

/**
 * Map UI field names to internal field names
 */
function mapUIFieldsToInternal(data) {
  const mapped = Object.assign({}, data);

  if (data.hours !== undefined) mapped.HOURS = data.hours;
  if (data.minutes !== undefined) mapped.MINUTES = data.minutes;
  if (data.overtimeHours !== undefined) mapped.OVERTIMEHOURS = data.overtimeHours;
  if (data.overtimeMinutes !== undefined) mapped.OVERTIMEMINUTES = data.overtimeMinutes;
  if (data.leavePay !== undefined) mapped['LEAVE PAY'] = data.leavePay;
  if (data.bonusPay !== undefined) mapped['BONUS PAY'] = data.bonusPay;
  if (data.otherIncome !== undefined) mapped.OTHERINCOME = data.otherIncome;
  if (data.otherDeductions !== undefined) mapped['OTHER DEDUCTIONS'] = data.otherDeductions;
  if (data.otherDeductionsText !== undefined) mapped['OTHER DEDUCTIONS TEXT'] = data.otherDeductionsText;
  if (data.weekEnding !== undefined) mapped.WEEKENDING = data.weekEnding;
  if (data.notes !== undefined) mapped.NOTES = data.notes;
  if (data.loanDeductionThisWeek !== undefined) mapped.LoanDeductionThisWeek = data.loanDeductionThisWeek;
  if (data.newLoanThisWeek !== undefined) mapped.NewLoanThisWeek = data.newLoanThisWeek;
  if (data.loanDisbursementType !== undefined) mapped.LoanDisbursementType = data.loanDisbursementType;

  return mapped;
}

/**
 * List all available employees
 */
function listAvailableEmployees() {
  try {
    const result = listEmployees({ activeOnly: true });
    if (result && result.success && result.data) {
      console.log('Available active employees:');
      result.data.forEach((emp, index) => {
        console.log(`  ${index + 1}. ${emp.REFNAME} (ID: ${emp.id})`);
      });
    } else {
      console.log('Could not retrieve employee list');
    }
  } catch (error) {
    console.error('Error listing employees:', error.toString());
  }
}

/**
 * Test specific function calls
 */
function diagnosticTestFunctions() {
  console.log('========== TESTING INDIVIDUAL FUNCTIONS ==========\n');

  // Test 1: getSheets
  console.log('--- Test 1: getSheets() ---');
  try {
    const sheets = getSheets();
    if (!sheets) {
      console.error('❌ getSheets returned null/undefined');
    } else {
      console.log('✓ getSheets returned:', Object.keys(sheets));
      console.log('  - employee:', sheets.employee ? sheets.employee.getName() : 'NULL');
      console.log('  - salary:', sheets.salary ? sheets.salary.getName() : 'NULL');
      console.log('  - loans:', sheets.loans ? sheets.loans.getName() : 'NULL');
    }
  } catch (error) {
    console.error('❌ Error:', error.toString());
  }

  // Test 2: listEmployees
  console.log('\n--- Test 2: listEmployees() ---');
  try {
    const result = listEmployees({ activeOnly: true });
    if (!result) {
      console.error('❌ listEmployees returned null/undefined');
    } else if (!result.success) {
      console.error('❌ listEmployees failed:', result.error);
    } else {
      console.log('✓ listEmployees returned', result.data.length, 'employees');
    }
  } catch (error) {
    console.error('❌ Error:', error.toString());
  }

  // Test 3: Utility functions
  console.log('\n--- Test 3: Utility functions ---');
  try {
    console.log('roundTo(123.456, 2):', roundTo(123.456, 2));
    console.log('formatCurrency(1234.56):', formatCurrency(1234.56));

    const testObj = {};
    addAuditFields(testObj, true);
    console.log('addAuditFields result:', JSON.stringify(testObj));

    const testRow = objectToRow({ name: 'Test', value: 123 }, ['name', 'value', 'extra']);
    console.log('objectToRow result:', testRow);

    console.log('✓ All utility functions working');
  } catch (error) {
    console.error('❌ Error:', error.toString());
  }

  console.log('\n========== FUNCTION TESTS COMPLETE ==========');
}

/**
 * Trace function calls in createPayslip
 */
function diagnosticTraceCreatePayslip() {
  console.log('========== TRACING createPayslip EXECUTION ==========\n');

  // Wrap all function calls with logging
  const originalGetEmployeeByName = getEmployeeByName;
  const originalGetCurrentLoanBalance = getCurrentLoanBalance;
  const originalCalculatePayslip = calculatePayslip;

  // Override with traced versions
  globalThis.getEmployeeByName = function(name) {
    console.log('→ getEmployeeByName called with:', name);
    const result = originalGetEmployeeByName(name);
    console.log('← getEmployeeByName returned:', result ? (result.success ? 'SUCCESS' : 'FAILURE') : 'NULL');
    return result;
  };

  globalThis.getCurrentLoanBalance = function(id) {
    console.log('→ getCurrentLoanBalance called with:', id);
    const result = originalGetCurrentLoanBalance(id);
    console.log('← getCurrentLoanBalance returned:', result);
    return result;
  };

  globalThis.calculatePayslip = function(data) {
    console.log('→ calculatePayslip called');
    const result = originalCalculatePayslip(data);
    console.log('← calculatePayslip returned:', result ? 'OBJECT' : 'NULL');
    return result;
  };

  try {
    const testData = {
      employeeName: 'John Doe',  // Change to real employee name
      weekEnding: '2025-11-15',
      hours: 40,
      minutes: 0,
      overtimeHours: 0,
      overtimeMinutes: 0,
      leavePay: 0,
      bonusPay: 0,
      otherIncome: 0,
      otherDeductions: 0,
      loanDeductionThisWeek: 0,
      newLoanThisWeek: 0,
      loanDisbursementType: 'Separate',
      notes: 'Trace test'
    };

    console.log('Calling createPayslip...\n');
    const result = createPayslip(testData);
    console.log('\ncreatePay slip completed:', result.success ? 'SUCCESS' : 'FAILURE');

  } finally {
    // Restore original functions
    globalThis.getEmployeeByName = originalGetEmployeeByName;
    globalThis.getCurrentLoanBalance = originalGetCurrentLoanBalance;
    globalThis.calculatePayslip = originalCalculatePayslip;
  }

  console.log('\n========== TRACE COMPLETE ==========');
}
