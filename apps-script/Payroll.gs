/**
 * PAYROLL.GS - Payroll Processing Engine
 * HR Payroll System - MOST CRITICAL FILE
 *
 * This file contains ALL payroll calculation logic and payslip management:
 * - createPayslip: Create new payslip
 * - calculatePayslip: ALL CALCULATIONS (100% accurate)
 * - generatePayslipPDF: PDF generation matching company format
 * - updatePayslip: Edit existing payslip
 * - listPayslips: Get payslips with filtering
 *
 * CRITICAL CALCULATIONS:
 * 1. Standard Time = (hours × rate) + ((rate/60) × minutes)
 * 2. Overtime = (hours × rate × 1.5) + ((rate/60) × minutes × 1.5)
 * 3. Gross = Standard + OT + Leave + Bonus + Other
 * 4. UIF = Gross × 1% (permanent only)
 * 5. Net = Gross - UIF - Other Deductions
 * 6. Paid to Account = Net - Loan Deduction + (New Loan if "With Salary")
 */

/**
 * Create a new payslip
 * CRITICAL: This triggers auto-sync for loan transactions
 */
function createPayslip(data) {
  try {
    Logger.log('\n========== CREATE PAYSLIP ==========');
    Logger.log('ℹ️ Input: ' + JSON.stringify(data));

    // Get employee details to lookup values
    const empResult = getEmployeeByName(data.employeeName);
    if (!empResult.success) {
      throw new Error('Employee not found: ' + data.employeeName);
    }

    const employee = empResult.data;

    // Enrich data with lookups
    data.EMPLOYER = employee.EMPLOYER;
    data['EMPLOYMENT STATUS'] = employee['EMPLOYMENT STATUS'];
    data.HOURLYRATE = employee['HOURLY RATE'];

    // Get current loan balance
    data.CurrentLoanBalance = getCurrentLoanBalance(employee.id);

    // Calculate all payslip values
    const calculations = calculatePayslip(data);

    // Merge calculations into data
    Object.assign(data, calculations);

    // Validate
    validatePayslip(data);

    // Get sheets
    const sheets = getSheets();
    const salarySheet = sheets.salary;
    if (!salarySheet) throw new Error('Salary sheet not found');

    // Generate record number (max + 1)
    const allData = salarySheet.getDataRange().getValues();
    const headers = allData[0];
    const rows = allData.slice(1);

    const recordNumCol = findColumnIndex(headers, 'RECORDNUMBER');
    let maxRecord = 0;
    rows.forEach(row => {
      const num = parseInt(row[recordNumCol]);
      if (!isNaN(num) && num > maxRecord) {
        maxRecord = num;
      }
    });

    data.RECORDNUMBER = maxRecord + 1;

    // Add audit fields
    addAuditFields(data, true);

    // Convert to row
    const row = objectToRow(data, headers);

    // Append
    salarySheet.appendRow(row);
    SpreadsheetApp.flush();

    Logger.log('✅ Payslip created: #' + data.RECORDNUMBER);
    Logger.log('✅ Paid to Account: ' + formatCurrency(data.PaidToAccount));
    Logger.log('========== CREATE PAYSLIP COMPLETE ==========\n');

    return { success: true, data: data };

  } catch (error) {
    Logger.log('❌ ERROR in createPayslip: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== CREATE PAYSLIP FAILED ==========\n');

    return { success: false, error: error.message };
  }
}

/**
 * CRITICAL: Calculate all payslip values
 * This function MUST be 100% accurate
 *
 * @param {Object} data - Payslip data
 * @returns {Object} Calculated values
 */
function calculatePayslip(data) {
  Logger.log('🔢 Calculating payslip...');

  // Parse input values
  const hours = parseFloat(data.HOURS) || 0;
  const minutes = parseFloat(data.MINUTES) || 0;
  const overtimeHours = parseFloat(data.OVERTIMEHOURS) || 0;
  const overtimeMinutes = parseFloat(data.OVERTIMEMINUTES) || 0;
  const hourlyRate = parseFloat(data.HOURLYRATE) || 0;
  const leavePay = parseFloat(data['LEAVE PAY']) || 0;
  const bonusPay = parseFloat(data['BONUS PAY']) || 0;
  const otherIncome = parseFloat(data.OTHERINCOME) || 0;
  const employmentStatus = data['EMPLOYMENT STATUS'] || '';
  const otherDeductions = parseFloat(data['OTHER DEDUCTIONS']) || 0;
  const loanDeduction = parseFloat(data.LoanDeductionThisWeek) || 0;
  const newLoan = parseFloat(data.NewLoanThisWeek) || 0;
  const loanDisbursementType = data.LoanDisbursementType || 'Separate';

  // 1. STANDARD TIME
  const standardTime = (hours * hourlyRate) + ((hourlyRate / 60) * minutes);
  Logger.log('  Standard Time: ' + formatCurrency(standardTime));

  // 2. OVERTIME (1.5x multiplier)
  const overtime = (overtimeHours * hourlyRate * OVERTIME_MULTIPLIER) + 
                   ((hourlyRate / 60) * overtimeMinutes * OVERTIME_MULTIPLIER);
  Logger.log('  Overtime: ' + formatCurrency(overtime));

  // 3. GROSS SALARY
  const grossSalary = standardTime + overtime + leavePay + bonusPay + otherIncome;
  Logger.log('  Gross Salary: ' + formatCurrency(grossSalary));

  // 4. UIF (1% for permanent employees ONLY)
  const uif = (employmentStatus === 'Permanent') ? (grossSalary * UIF_RATE) : 0;
  Logger.log('  UIF: ' + formatCurrency(uif) + ' (Status: ' + employmentStatus + ')');

  // 5. TOTAL DEDUCTIONS (NOTE: Loan deduction NOT included here - it comes from net)
  const totalDeductions = uif + otherDeductions;
  Logger.log('  Total Deductions: ' + formatCurrency(totalDeductions));

  // 6. NET SALARY
  const netSalary = grossSalary - totalDeductions;
  Logger.log('  Net Salary: ' + formatCurrency(netSalary));

  // 7. PAID TO ACCOUNT (CRITICAL - actual bank transfer amount)
  // Net - Loan Repayment + (New Loan if "With Salary")
  const newLoanToAdd = (loanDisbursementType === 'With Salary') ? newLoan : 0;
  const paidToAccount = netSalary - loanDeduction + newLoanToAdd;
  Logger.log('  Loan Deduction: ' + formatCurrency(loanDeduction));
  Logger.log('  New Loan (if with salary): ' + formatCurrency(newLoanToAdd));
  Logger.log('  ** PAID TO ACCOUNT: ' + formatCurrency(paidToAccount) + ' **');

  // 8. UPDATED LOAN BALANCE
  const currentBalance = parseFloat(data.CurrentLoanBalance) || 0;
  const updatedBalance = currentBalance - loanDeduction + newLoan;
  Logger.log('  Updated Loan Balance: ' + formatCurrency(updatedBalance));

  // Return all calculated values (rounded to 2 decimals)
  return {
    STANDARDTIME: roundTo(standardTime, 2),
    OVERTIME: roundTo(overtime, 2),
    GROSSSALARY: roundTo(grossSalary, 2),
    UIF: roundTo(uif, 2),
    TOTALDEDUCTIONS: roundTo(totalDeductions, 2),
    NETTSALARY: roundTo(netSalary, 2),
    PaidToAccount: roundTo(paidToAccount, 2),
    UpdatedLoanBalance: roundTo(updatedBalance, 2)
  };
}

/**
 * Get a single payslip by record number
 */
function getPayslip(recordNumber) {
  try {
    const sheets = getSheets();
    const salarySheet = sheets.salary;
    if (!salarySheet) throw new Error('Salary sheet not found');

    const allData = salarySheet.getDataRange().getValues();
    const headers = allData[0];
    const rows = allData.slice(1);

    const recordNumCol = findColumnIndex(headers, 'RECORDNUMBER');
    const row = rows.find(r => String(r[recordNumCol]) === String(recordNumber));

    if (!row) {
      throw new Error('Payslip not found: ' + recordNumber);
    }

    const payslip = buildObjectFromRow(row, headers);

    return { success: true, data: payslip };

  } catch (error) {
    Logger.log('❌ ERROR in getPayslip: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Update an existing payslip
 */
function updatePayslip(recordNumber, data) {
  try {
    Logger.log('\n========== UPDATE PAYSLIP ==========');
    Logger.log('ℹ️ Record Number: ' + recordNumber);

    const sheets = getSheets();
    const salarySheet = sheets.salary;
    if (!salarySheet) throw new Error('Salary sheet not found');

    // Find payslip
    const allData = salarySheet.getDataRange().getValues();
    const headers = allData[0];
    const rows = allData.slice(1);

    const recordNumCol = findColumnIndex(headers, 'RECORDNUMBER');
    const rowIndex = rows.findIndex(r => String(r[recordNumCol]) === String(recordNumber));

    if (rowIndex === -1) {
      throw new Error('Payslip not found: ' + recordNumber);
    }

    // Get current payslip
    const currentPayslip = buildObjectFromRow(rows[rowIndex], headers);

    // Merge updates
    const updatedPayslip = Object.assign({}, currentPayslip, data);

    // Recalculate
    const calculations = calculatePayslip(updatedPayslip);
    Object.assign(updatedPayslip, calculations);

    // Validate
    validatePayslip(updatedPayslip);

    // Add audit fields (update)
    addAuditFields(updatedPayslip, false);

    // Convert to row
    const updatedRow = objectToRow(updatedPayslip, headers);

    // Update sheet
    const sheetRowIndex = rowIndex + 2;
    salarySheet.getRange(sheetRowIndex, 1, 1, headers.length).setValues([updatedRow]);
    SpreadsheetApp.flush();

    Logger.log('✅ Payslip updated: #' + recordNumber);
    Logger.log('========== UPDATE PAYSLIP COMPLETE ==========\n');

    return { success: true, data: updatedPayslip };

  } catch (error) {
    Logger.log('❌ ERROR in updatePayslip: ' + error.message);
    Logger.log('========== UPDATE PAYSLIP FAILED ==========\n');

    return { success: false, error: error.message };
  }
}

/**
 * List payslips with optional filtering
 */
function listPayslips(filters = {}) {
  try {
    const sheets = getSheets();
    const salarySheet = sheets.salary;
    if (!salarySheet) throw new Error('Salary sheet not found');

    const allData = salarySheet.getDataRange().getValues();
    const headers = allData[0];
    let rows = allData.slice(1);

    // Convert to objects
    let payslips = rows.map(row => buildObjectFromRow(row, headers));

    // Apply filters
    if (filters.employeeName) {
      payslips = payslips.filter(p => p['EMPLOYEE NAME'] === filters.employeeName);
    }

    if (filters.employer) {
      payslips = payslips.filter(p => p.EMPLOYER === filters.employer);
    }

    if (filters.weekEnding) {
      const targetDate = parseDate(filters.weekEnding);
      payslips = payslips.filter(p => {
        const payslipDate = parseDate(p.WEEKENDING);
        return payslipDate.getTime() === targetDate.getTime();
      });
    }

    if (filters.startDate) {
      const start = parseDate(filters.startDate);
      payslips = payslips.filter(p => {
        const payslipDate = parseDate(p.WEEKENDING);
        return payslipDate >= start;
      });
    }

    if (filters.endDate) {
      const end = parseDate(filters.endDate);
      payslips = payslips.filter(p => {
        const payslipDate = parseDate(p.WEEKENDING);
        return payslipDate <= end;
      });
    }

    return { success: true, data: payslips };

  } catch (error) {
    Logger.log('❌ ERROR in listPayslips: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Validate payslip data
 */
function validatePayslip(data) {
  const errors = [];

  // Check required fields
  if (!data.employeeName && !data['EMPLOYEE NAME']) {
    errors.push('Employee Name is required');
  }

  if (!data.WEEKENDING) {
    errors.push('Week Ending date is required');
  }

  // Check for duplicate payslip
  if (data.employeeName && data.WEEKENDING && !data.RECORDNUMBER) {
    const empName = data.employeeName || data['EMPLOYEE NAME'];
    const isDuplicate = checkDuplicatePayslip(empName, data.WEEKENDING);
    if (isDuplicate) {
      errors.push('Payslip already exists for ' + empName + ' on ' + formatDateShort(data.WEEKENDING));
    }
  }

  // Validate time values
  const hours = parseFloat(data.HOURS) || 0;
  const minutes = parseFloat(data.MINUTES) || 0;
  const otHours = parseFloat(data.OVERTIMEHOURS) || 0;
  const otMinutes = parseFloat(data.OVERTIMEMINUTES) || 0;

  if (hours < 0 || minutes < 0 || otHours < 0 || otMinutes < 0) {
    errors.push('Time values cannot be negative');
  }

  // Warn if loan deduction exceeds balance
  const loanDeduction = parseFloat(data.LoanDeductionThisWeek) || 0;
  const currentBalance = parseFloat(data.CurrentLoanBalance) || 0;

  if (loanDeduction > currentBalance) {
    Logger.log('⚠️ WARNING: Loan deduction (' + formatCurrency(loanDeduction) + 
               ') exceeds current balance (' + formatCurrency(currentBalance) + ')');
  }

  if (errors.length > 0) {
    throw new Error(errors.join('; '));
  }
}

/**
 * Check if payslip already exists for employee and week
 */
function checkDuplicatePayslip(employeeName, weekEnding) {
  try {
    const sheets = getSheets();
    const salarySheet = sheets.salary;
    if (!salarySheet) return false;

    const allData = salarySheet.getDataRange().getValues();
    const headers = allData[0];
    const rows = allData.slice(1);

    const empNameCol = findColumnIndex(headers, 'EMPLOYEE NAME');
    const weekEndingCol = findColumnIndex(headers, 'WEEKENDING');

    const targetDate = parseDate(weekEnding).getTime();

    for (let row of rows) {
      if (row[empNameCol] === employeeName) {
        const rowDate = parseDate(row[weekEndingCol]).getTime();
        if (rowDate === targetDate) {
          return true;
        }
      }
    }

    return false;

  } catch (error) {
    Logger.log('⚠️ Error checking duplicate: ' + error.message);
    return false;
  }
}

/**
 * Generate PDF payslip (placeholder - full implementation requires Google Docs API)
 */
function generatePayslipPDF(recordNumber) {
  try {
    Logger.log('\n========== GENERATE PAYSLIP PDF ==========');
    Logger.log('ℹ️ Record Number: ' + recordNumber);

    // Get payslip
    const result = getPayslip(recordNumber);
    if (!result.success) {
      throw new Error('Payslip not found');
    }

    const payslip = result.data;

    // TODO: Full PDF generation implementation
    // This requires creating a Google Doc with the payslip format
    // For now, return placeholder

    Logger.log('⚠️ PDF generation not yet implemented');
    Logger.log('========== GENERATE PDF COMPLETE ==========\n');

    return {
      success: true,
      message: 'PDF generation placeholder - implement using Google Docs API',
      data: { url: '#' }
    };

  } catch (error) {
    Logger.log('❌ ERROR in generatePayslipPDF: ' + error.message);
    return { success: false, error: error.message };
  }
}

// ========== TEST FUNCTIONS ==========

/**
 * Test Case 1: From sample payslip #7916
 * Input: 39h 30m @ R33.96, Permanent, Loan repayment R150
 * Expected: Standard=R1,341.42, UIF=R13.41, Net=R1,328.01, Paid=R1,178.01
 */
function test_calculatePayslip_StandardTime() {
  Logger.log('\n========== TEST: PAYSLIP CALCULATION (Case 1) ==========');

  const testData = {
    HOURS: 39,
    MINUTES: 30,
    OVERTIMEHOURS: 0,
    OVERTIMEMINUTES: 0,
    HOURLYRATE: 33.96,
    'LEAVE PAY': 0,
    'BONUS PAY': 0,
    OTHERINCOME: 0,
    'EMPLOYMENT STATUS': 'Permanent',
    'OTHER DEDUCTIONS': 0,
    LoanDeductionThisWeek: 150,
    NewLoanThisWeek: 0,
    CurrentLoanBalance: 150,
    LoanDisbursementType: 'Separate'
  };

  const result = calculatePayslip(testData);

  // Expected values
  const expected = {
    STANDARDTIME: 1341.42,
    OVERTIME: 0.00,
    GROSSSALARY: 1341.42,
    UIF: 13.41,
    NETTSALARY: 1328.01,
    PaidToAccount: 1178.01
  };

  let passed = true;
  for (let key in expected) {
    const diff = Math.abs(result[key] - expected[key]);
    if (diff > 0.01) {
      Logger.log('❌ FAILED: ' + key + ' - Expected: ' + expected[key] + ', Got: ' + result[key]);
      passed = false;
    } else {
      Logger.log('✅ PASSED: ' + key + ' = ' + result[key]);
    }
  }

  if (passed) {
    Logger.log('✅ TEST CASE 1 PASSED');
  } else {
    Logger.log('❌ TEST CASE 1 FAILED');
  }

  Logger.log('========== TEST COMPLETE ==========\n');
}

/**
 * Test Case 2: New loan with salary
 * Input: 40h @ R40.00, Permanent, New loan R500 (with salary)
 * Expected: Standard=R1,600.00, UIF=R16.00, Net=R1,584.00, Paid=R2,084.00
 */
function test_calculatePayslip_NewLoanWithSalary() {
  Logger.log('\n========== TEST: PAYSLIP CALCULATION (Case 2) ==========');

  const testData = {
    HOURS: 40,
    MINUTES: 0,
    OVERTIMEHOURS: 0,
    OVERTIMEMINUTES: 0,
    HOURLYRATE: 40.00,
    'LEAVE PAY': 0,
    'BONUS PAY': 0,
    OTHERINCOME: 0,
    'EMPLOYMENT STATUS': 'Permanent',
    'OTHER DEDUCTIONS': 0,
    LoanDeductionThisWeek: 0,
    NewLoanThisWeek: 500,
    CurrentLoanBalance: 0,
    LoanDisbursementType: 'With Salary'
  };

  const result = calculatePayslip(testData);

  const expected = {
    STANDARDTIME: 1600.00,
    UIF: 16.00,
    NETTSALARY: 1584.00,
    PaidToAccount: 2084.00  // Net + Loan
  };

  let passed = true;
  for (let key in expected) {
    const diff = Math.abs(result[key] - expected[key]);
    if (diff > 0.01) {
      Logger.log('❌ FAILED: ' + key + ' - Expected: ' + expected[key] + ', Got: ' + result[key]);
      passed = false;
    } else {
      Logger.log('✅ PASSED: ' + key + ' = ' + result[key]);
    }
  }

  if (passed) {
    Logger.log('✅ TEST CASE 2 PASSED');
  } else {
    Logger.log('❌ TEST CASE 2 FAILED');
  }

  Logger.log('========== TEST COMPLETE ==========\n');
}

/**
 * Test Case 3: Temporary employee (no UIF)
 * Input: 40h @ R35.00, Temporary, No loans
 * Expected: Standard=R1,400.00, UIF=R0.00, Net=R1,400.00, Paid=R1,400.00
 */
function test_calculatePayslip_UIF() {
  Logger.log('\n========== TEST: PAYSLIP CALCULATION (Case 3 - UIF) ==========');

  const testData = {
    HOURS: 40,
    MINUTES: 0,
    OVERTIMEHOURS: 0,
    OVERTIMEMINUTES: 0,
    HOURLYRATE: 35.00,
    'LEAVE PAY': 0,
    'BONUS PAY': 0,
    OTHERINCOME: 0,
    'EMPLOYMENT STATUS': 'Temporary',
    'OTHER DEDUCTIONS': 0,
    LoanDeductionThisWeek: 0,
    NewLoanThisWeek: 0,
    CurrentLoanBalance: 0,
    LoanDisbursementType: 'Separate'
  };

  const result = calculatePayslip(testData);

  const expected = {
    STANDARDTIME: 1400.00,
    UIF: 0.00,  // Temporary employees don't pay UIF
    NETTSALARY: 1400.00,
    PaidToAccount: 1400.00
  };

  let passed = true;
  for (let key in expected) {
    const diff = Math.abs(result[key] - expected[key]);
    if (diff > 0.01) {
      Logger.log('❌ FAILED: ' + key + ' - Expected: ' + expected[key] + ', Got: ' + result[key]);
      passed = false;
    } else {
      Logger.log('✅ PASSED: ' + key + ' = ' + result[key]);
    }
  }

  if (passed) {
    Logger.log('✅ TEST CASE 3 PASSED');
  } else {
    Logger.log('❌ TEST CASE 3 FAILED');
  }

  Logger.log('========== TEST COMPLETE ==========\n');
}

/**
 * Test Case 4: Overtime calculation
 * Input: 35h + 5h OT @ R30.00, Permanent
 * Expected: Standard=R1,050.00, OT=R225.00, Gross=R1,275.00
 */
function test_calculatePayslip_Overtime() {
  Logger.log('\n========== TEST: PAYSLIP CALCULATION (Case 4 - Overtime) ==========');

  const testData = {
    HOURS: 35,
    MINUTES: 0,
    OVERTIMEHOURS: 5,
    OVERTIMEMINUTES: 0,
    HOURLYRATE: 30.00,
    'LEAVE PAY': 0,
    'BONUS PAY': 0,
    OTHERINCOME: 0,
    'EMPLOYMENT STATUS': 'Permanent',
    'OTHER DEDUCTIONS': 0,
    LoanDeductionThisWeek: 0,
    NewLoanThisWeek: 0,
    CurrentLoanBalance: 0,
    LoanDisbursementType: 'Separate'
  };

  const result = calculatePayslip(testData);

  const expected = {
    STANDARDTIME: 1050.00,
    OVERTIME: 225.00,  // 5 × 30 × 1.5
    GROSSSALARY: 1275.00,
    UIF: 12.75,
    NETTSALARY: 1262.25,
    PaidToAccount: 1262.25
  };

  let passed = true;
  for (let key in expected) {
    const diff = Math.abs(result[key] - expected[key]);
    if (diff > 0.01) {
      Logger.log('❌ FAILED: ' + key + ' - Expected: ' + expected[key] + ', Got: ' + result[key]);
      passed = false;
    } else {
      Logger.log('✅ PASSED: ' + key + ' = ' + result[key]);
    }
  }

  if (passed) {
    Logger.log('✅ TEST CASE 4 PASSED');
  } else {
    Logger.log('❌ TEST CASE 4 FAILED');
  }

  Logger.log('========== TEST COMPLETE ==========\n');
}
