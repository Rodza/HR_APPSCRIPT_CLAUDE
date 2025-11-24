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
 */
function createPayslip(data) {
  try {
    Logger.log('\n========== CREATE PAYSLIP ==========');
    Logger.log('ℹ️ Input: ' + JSON.stringify(data));

    // Map UI field names to internal field names (if needed)
    if (data.hours !== undefined) data.HOURS = data.hours;
    if (data.minutes !== undefined) data.MINUTES = data.minutes;
    if (data.overtimeHours !== undefined) data.OVERTIMEHOURS = data.overtimeHours;
    if (data.overtimeMinutes !== undefined) data.OVERTIMEMINUTES = data.overtimeMinutes;
    if (data.leavePay !== undefined) data['LEAVE PAY'] = data.leavePay;
    if (data.bonusPay !== undefined) data['BONUS PAY'] = data.bonusPay;
    if (data.otherIncome !== undefined) data.OTHERINCOME = data.otherIncome;
    if (data.otherDeductions !== undefined) data['OTHER DEDUCTIONS'] = data.otherDeductions;
    if (data.otherDeductionsText !== undefined) data['OTHER DEDUCTIONS TEXT'] = data.otherDeductionsText;
    if (data.weekEnding !== undefined) data.WEEKENDING = data.weekEnding;
    if (data.notes !== undefined) data.NOTES = data.notes;

    // Get employee details to lookup values
    const empResult = getEmployeeByName(data.employeeName);

    // Defensive null check
    if (!empResult) {
      throw new Error('getEmployeeByName returned null for employee: ' + data.employeeName);
    }

    if (!empResult.success) {
      throw new Error('Employee not found: ' + data.employeeName);
    }

    const employee = empResult.data;

    // Defensive null check for employee data
    if (!employee) {
      throw new Error('Employee data is null for: ' + data.employeeName);
    }

    // Enrich data with lookups
    data.id = employee.id;                                    // Link to employee record
    data['EMPLOYEE NAME'] = data.employeeName;
    data.EMPLOYER = employee.EMPLOYER;
    data['EMPLOYMENT STATUS'] = employee['EMPLOYMENT STATUS'];
    data.HOURLYRATE = employee['HOURLY RATE'];

    // Get current loan balance (handle null safely)
    const loanBalanceResult = getCurrentLoanBalance(employee.id);
    data.CurrentLoanBalance = (loanBalanceResult && loanBalanceResult.success && loanBalanceResult.data !== undefined)
      ? loanBalanceResult.data
      : 0;

    // Calculate all payslip values
    const calculations = calculatePayslip(data);

    // Defensive null check for calculations
    if (!calculations) {
      throw new Error('calculatePayslip returned null');
    }

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
    Logger.log('✅ Paid to Account: ' + formatCurrency(data.PaidtoAccount));
    Logger.log('========== CREATE PAYSLIP COMPLETE ==========\n');

    // Sanitize for web - convert Date objects to strings
    const sanitizedData = sanitizeForWeb(data);

    return { success: true, data: sanitizedData };

  } catch (error) {
    Logger.log('❌ ERROR in createPayslip: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== CREATE PAYSLIP FAILED ==========\n');

    return { success: false, error: error.message };
  }
}

/**
 * Calculate payslip preview (for UI)
 * Wrapper around calculatePayslip that returns {success, data} format
 *
 * @param {Object} data - Payslip input data
 * @returns {Object} Response with {success, data, error}
 */
function calculatePayslipPreview(data) {
  try {
    // Get employee details to populate missing fields
    if (data.employeeName) {
      const empResult = getEmployeeByName(data.employeeName);

      // Defensive null check
      if (empResult && empResult.success && empResult.data) {
        const employee = empResult.data;
        data.EMPLOYER = employee.EMPLOYER;
        data['EMPLOYMENT STATUS'] = employee['EMPLOYMENT STATUS'];
        data.HOURLYRATE = employee['HOURLY RATE'];
      } else {
        Logger.log('⚠️ Warning: Could not get employee details for preview: ' + data.employeeName);
      }
    }

    // Map UI field names to internal field names
    data.HOURS = data.hours || 0;
    data.MINUTES = data.minutes || 0;
    data.OVERTIMEHOURS = data.overtimeHours || 0;
    data.OVERTIMEMINUTES = data.overtimeMinutes || 0;
    data['LEAVE PAY'] = data.leavePay || 0;
    data['BONUS PAY'] = data.bonusPay || 0;
    data.OTHERINCOME = data.otherIncome || 0;
    data['OTHER DEDUCTIONS'] = data.otherDeductions || 0;
    data.LoanDeductionThisWeek = data.loanDeductionThisWeek || 0;
    data.NewLoanThisWeek = data.newLoanThisWeek || 0;

    const calculations = calculatePayslip(data);

    // Defensive null check for calculations
    if (!calculations) {
      throw new Error('calculatePayslip returned null in preview');
    }

    return {
      success: true,
      data: {
        standardTime: calculations.STANDARDTIME,
        overtime: calculations.OVERTIME,
        grossSalary: calculations.GROSSSALARY,
        uif: calculations.UIF,
        totalDeductions: calculations.TOTALDEDUCTIONS,
        netSalary: calculations.NETTSALARY,
        paidToAccount: calculations.PaidtoAccount
      },
      error: null
    };
  } catch (error) {
    Logger.log('❌ ERROR in calculatePayslipPreview: ' + error.message);
    return {
      success: false,
      data: null,
      error: error.message
    };
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
    PaidtoAccount: roundTo(paidToAccount, 2),
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

    // Sanitize for web - convert Date objects to strings
    const sanitizedPayslip = sanitizeForWeb(payslip);

    return { success: true, data: sanitizedPayslip };

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

    // Map UI field names to internal field names (if needed)
    if (data.hours !== undefined) data.HOURS = data.hours;
    if (data.minutes !== undefined) data.MINUTES = data.minutes;
    if (data.overtimeHours !== undefined) data.OVERTIMEHOURS = data.overtimeHours;
    if (data.overtimeMinutes !== undefined) data.OVERTIMEMINUTES = data.overtimeMinutes;
    if (data.leavePay !== undefined) data['LEAVE PAY'] = data.leavePay;
    if (data.bonusPay !== undefined) data['BONUS PAY'] = data.bonusPay;
    if (data.otherIncome !== undefined) data.OTHERINCOME = data.otherIncome;
    if (data.otherDeductions !== undefined) data['OTHER DEDUCTIONS'] = data.otherDeductions;
    if (data.otherDeductionsText !== undefined) data['OTHER DEDUCTIONS TEXT'] = data.otherDeductionsText;
    if (data.weekEnding !== undefined) data.WEEKENDING = data.weekEnding;
    if (data.notes !== undefined) data.NOTES = data.notes;
    if (data.employeeName !== undefined) data['EMPLOYEE NAME'] = data.employeeName;

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

    // Check if payslip is still editable
    if (!isPayslipEditable(currentPayslip.WEEKENDING)) {
      throw new Error('Cannot update payslip: The editing period has ended (past Friday midnight of the payslip week)');
    }

    // Merge updates
    const updatedPayslip = Object.assign({}, currentPayslip, data);

    // If employee name changed, update employee ID and related fields
    if (data.employeeName && data.employeeName !== currentPayslip['EMPLOYEE NAME']) {
      const empResult = getEmployeeByName(data.employeeName);
      if (empResult.success) {
        const employee = empResult.data;
        updatedPayslip.id = employee.id;
        updatedPayslip.EMPLOYER = employee.EMPLOYER;
        updatedPayslip['EMPLOYMENT STATUS'] = employee['EMPLOYMENT STATUS'];
        updatedPayslip.HOURLYRATE = employee['HOURLY RATE'];
      }
    }

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

    // Sanitize for web - convert Date objects to strings
    const sanitizedPayslip = sanitizeForWeb(updatedPayslip);

    return { success: true, data: sanitizedPayslip };

  } catch (error) {
    Logger.log('❌ ERROR in updatePayslip: ' + error.message);
    Logger.log('========== UPDATE PAYSLIP FAILED ==========\n');

    return { success: false, error: error.message };
  }
}

/**
 * List payslips with optional filtering
 */
function listPayslips(filters) {
  try {
    // Ensure filters is an object (handle undefined/null)
    if (!filters) {
      filters = {};
    }

    Logger.log('📋 listPayslips called with filters:', JSON.stringify(filters));

    const sheets = getSheets();
    const salarySheet = sheets.salary;
    if (!salarySheet) throw new Error('Salary sheet not found');

    const allData = salarySheet.getDataRange().getValues();
    const headers = allData[0];
    let rows = allData.slice(1);

    Logger.log('📋 Found ' + rows.length + ' payslip rows');

    // Convert to objects
    let payslips = rows.map(row => buildObjectFromRow(row, headers));

    Logger.log('📋 Converted to ' + payslips.length + ' payslip objects');

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

    Logger.log('📋 Returning ' + payslips.length + ' payslips (after filters)');

    // Sanitize data for web - convert Date objects to strings
    const sanitizedPayslips = payslips.map(function(payslip) {
      return sanitizeForWeb(payslip);
    });

    Logger.log('📋 Sanitized ' + sanitizedPayslips.length + ' payslips for web');

    const result = { success: true, data: sanitizedPayslips };
    Logger.log('📋 listPayslips result type:', typeof result);
    Logger.log('📋 listPayslips result.success:', result.success);
    return result;

  } catch (error) {
    Logger.log('❌ ERROR in listPayslips: ' + error.message);
    Logger.log('❌ Error stack:', error.stack);
    const errorResult = { success: false, error: error.message };
    Logger.log('❌ Returning error result:', JSON.stringify(errorResult));
    return errorResult;
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

    // Create PDF filename
    const docName = `Payslip_${payslip.RECORDNUMBER}_${payslip['EMPLOYEE NAME']}_${formatDateForFilename(payslip.WEEKENDING)}`;

    // HTML template for payslip
    const htmlTemplate = `<html><head><meta content="text/html; charset=UTF-8" http-equiv="content-type"><style type="text/css">@page { size: A4 landscape; } ol{margin:0;padding:0}table td,table th{padding:1pt 5pt 1pt 5pt}.c36{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#ffffff;border-top-width:0pt;border-right-width:0pt;border-left-color:#ffffff;vertical-align:top;border-right-color:#ffffff;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:0pt;width:12pt;border-top-color:#ffffff;border-bottom-style:solid}.c35{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1.5pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:180pt;border-top-color:#000000;border-bottom-style:dotted}.c26{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:150pt;border-top-color:#000000;border-bottom-style:solid}.c24{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:830pt;border-top-color:#000000;border-bottom-style:solid}.c37{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1.5pt;width:165pt;border-top-color:#000000;border-bottom-style:solid}.c62{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:0pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:150pt;border-top-color:#000000;border-bottom-style:solid}.c27{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:80pt;border-top-color:#000000;border-bottom-style:solid}.c8{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1.5pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:200pt;border-top-color:#000000;border-bottom-style:solid}.c51{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1.5pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1.5pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:200pt;border-top-color:#000000;border-bottom-style:solid}.c60{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:0pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:350pt;border-top-color:#000000;border-bottom-style:solid}.c59{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:0pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:200pt;border-top-color:#000000;border-bottom-style:solid}.c0{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#ffffff;border-top-width:1pt;border-right-width:1pt;border-left-color:#ffffff;vertical-align:top;border-right-color:#ffffff;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:291pt;border-top-color:#ffffff;border-bottom-style:solid}.c61{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:0pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:150pt;border-top-color:#000000;border-bottom-style:solid}.c57{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#ffffff;border-top-width:0pt;border-right-width:0pt;border-left-color:#ffffff;vertical-align:top;border-right-color:#ffffff;border-left-width:0pt;border-top-style:solid;border-left-style:solid;border-bottom-width:0pt;width:12pt;border-top-color:#ffffff;border-bottom-style:solid}.c54{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1.5pt;width:180pt;border-top-color:#000000;border-bottom-style:solid}.c22{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:220pt;border-top-color:#000000;border-bottom-style:solid}.c34{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1.5pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:180pt;border-top-color:#000000;border-bottom-style:solid}.c45{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:0pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:700pt;border-top-color:#000000;border-bottom-style:solid}.c25{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:0pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:200pt;border-top-color:#000000;border-bottom-style:solid}.c52{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#ffffff;border-top-width:1pt;border-right-width:1pt;border-left-color:#ffffff;vertical-align:top;border-right-color:#ffffff;border-left-width:0pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:225.8pt;border-top-color:#ffffff;border-bottom-style:solid}.c29{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:0pt;width:414pt;border-top-color:#000000;border-bottom-style:solid}.c40{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1.5pt;border-right-width:1.5pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:180pt;border-top-color:#000000;border-bottom-style:solid}.c20{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:80pt;border-top-color:#000000;border-bottom-style:solid}.c14{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:165pt;border-top-color:#000000;border-bottom-style:solid}.c5{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#ffffff;border-top-width:0pt;border-right-width:0pt;border-left-color:#ffffff;vertical-align:top;border-right-color:#ffffff;border-left-width:0pt;border-top-style:solid;border-left-style:solid;border-bottom-width:0pt;width:204.8pt;border-top-color:#ffffff;border-bottom-style:solid}.c58{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:0pt;border-right-width:1.5pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:0pt;width:12.8pt;border-top-color:#000000;border-bottom-style:solid}.c17{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:120pt;border-top-color:#000000;border-bottom-style:solid}.c16{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:240pt;border-top-color:#000000;border-bottom-style:solid}.c11{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:180pt;border-top-color:#000000;border-bottom-style:solid}.c48{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:320.2pt;border-top-color:#000000;border-bottom-style:solid}.c42{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:12.8pt;border-top-color:#000000;border-bottom-style:solid}.c39{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1.5pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:200pt;border-top-color:#000000;border-bottom-style:dotted}.c15{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1.5pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:0pt;width:12.8pt;border-top-color:#000000;border-bottom-style:solid}.c32{border-right-style:solid;border-bottom-color:#ffffff;border-top-width:1pt;border-right-width:0pt;border-left-color:#ffffff;vertical-align:top;border-right-color:#ffffff;border-left-width:0pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:291pt;border-top-color:#ffffff;border-bottom-style:solid}.c12{color:#000000;font-weight:400;text-decoration:none;vertical-align:baseline;font-size:1pt;font-family:"Arial";font-style:normal}.c18{color:#000000;font-weight:400;text-decoration:none;vertical-align:baseline;font-size:2pt;font-family:"Arial";font-style:normal}.c28{color:#000000;font-weight:400;text-decoration:none;vertical-align:baseline;font-size:9pt;font-family:"Arial";font-style:normal}.c53{padding-top:8pt;padding-bottom:4pt;line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}.c2{color:#000000;font-weight:400;text-decoration:none;vertical-align:baseline;font-size:11pt;font-family:"Arial";font-style:normal}.c6{color:#000000;text-decoration:none;vertical-align:baseline;font-size:11pt;font-family:"Arial";font-style:normal}.c21{padding-top:0pt;padding-bottom:0pt;line-height:1.15;orphans:2;widows:2;text-align:right}.c30{color:#000000;font-weight:400;text-decoration:none;vertical-align:baseline;font-family:"Arial";font-style:normal}.c3{padding-top:0pt;padding-bottom:0pt;line-height:1.15;orphans:2;widows:2;text-align:left}.c1{padding-top:0pt;padding-bottom:0pt;line-height:1.0;text-align:left;height:11pt}.c31{color:#000000;text-decoration:none;vertical-align:baseline;font-family:"Arial";font-style:normal}.c9{border-spacing:0;border-collapse:collapse;margin-right:auto}.c44{padding-top:0pt;padding-bottom:0pt;line-height:1.0;text-align:right}.c10{padding-top:0pt;padding-bottom:0pt;line-height:1.0;text-align:left}.c41{padding-top:0pt;padding-bottom:0pt;line-height:1.0;text-align:center}.c49{background-color:#ffffff;max-width:none;padding:0pt 72pt 0pt 72pt;page-break-after: avoid;}.c43{color:inherit;text-decoration:inherit}.c19{height:14pt}.c56{font-size:9pt}.c38{font-size:19pt}.c47{height:22pt}.c7{height:0pt}.c23{font-size:10pt}.c33{font-size:17pt}.c13{height:8pt}.c50{height:14pt}.c4{font-weight:700}.c46{font-size:1pt}.c55{font-size:13pt}.title{padding-top:0pt;color:#000000;font-size:26pt;padding-bottom:3pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}.subtitle{padding-top:0pt;color:#666666;font-size:15pt;padding-bottom:16pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}li{color:#000000;font-size:11pt;font-family:"Arial"}p{margin:0;color:#000000;font-size:11pt;font-family:"Arial"}h1{padding-top:20pt;color:#000000;font-size:20pt;padding-bottom:6pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h2{padding-top:18pt;color:#000000;font-size:16pt;padding-bottom:6pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h3{padding-top:16pt;color:#434343;font-size:14pt;padding-bottom:4pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h4{padding-top:14pt;color:#666666;font-size:12pt;padding-bottom:4pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h5{padding-top:12pt;color:#666666;font-size:11pt;padding-bottom:4pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h6{padding-top:12pt;color:#666666;font-size:11pt;padding-bottom:4pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;font-style:italic;orphans:2;widows:2;text-align:left}</style></head><body class="c49 doc-content"><h1 class="c53"><span class="c31 c33 c4">WEEKLY PAY REMITTANCE</span></h1><table class="c9"><tr class="c19"><td class="c0" colspan="1" rowspan="1"><p class="c10"><span class="c4">PAYSLIP NUMBER:</span><span class="c2">&nbsp;{{RECORDNUMBER}}</span></p></td><td class="c36" colspan="1" rowspan="1"><p class="c1"><span class="c2"></span></p></td><td class="c5" colspan="1" rowspan="1"><p class="c1"><span class="c2"></span></p></td><td class="c52" colspan="1" rowspan="1"><p class="c3"><span class="c4 c33">{{EMPLOYER}}</span></p></td></tr><tr class="c7"><td class="c32" colspan="1" rowspan="1"><p class="c1"><span class="c12"></span></p></td><td class="c57" colspan="1" rowspan="1"><p class="c1"><span class="c12"></span></p></td><td class="c5" colspan="1" rowspan="2"><p class="c1"><span class="c28"></span></p></td><td class="c52" colspan="1" rowspan="2"><p class="c10"><span class="c28">18 DODGE STREET</span></p><p class="c10"><span class="c28">AUREUS</span></p><p class="c10"><span class="c28">RANDFONTEIN</span></p><p class="c10"><span class="c28">(011) 693 4278 / 083 338 5609</span></p><p class="c10"><span class="c56"><a class="c43" href="mailto:INFO@SAGRINDING.CO.ZA">INFO@SAGRINDING.CO.ZA</a></span><span class="c28">&nbsp;/ INFO@SCORPIOABRASIVES.CO.ZA</span></p></td></tr><tr class="c47"><td class="c0" colspan="1" rowspan="1"><p class="c10"><span class="c4">EMPLOYEE:</span><span class="c2">&nbsp;{{EMPLOYEE NAME}}</span></p><p class="c1"><span class="c12"></span></p><p class="c10"><span class="c4">WEEKENDING:</span><span class="c2">&nbsp;{{WEEKENDING}}</span></p><p class="c1"><span class="c12"></span></p><p class="c10"><span class="c4">EMPLOYMENT STATUS:</span><span>&nbsp;{{EMPLOYMENT STATUS}}</span></p></td><td class="c36" colspan="1" rowspan="1"><p class="c1"><span class="c2"></span></p></td></tr></table><p class="c1"><span class="c18"></span></p><table class="c9"><tr class="c19"><td class="c29" colspan="5" rowspan="1"><p class="c41"><span class="c31 c4 c55">EARNINGS</span></p></td><td class="c48" colspan="3" rowspan="1"><p class="c41"><span class="c31 c4 c55">DEDUCTIONS</span></p></td></tr><tr class="c19"><td class="c61" colspan="1" rowspan="1"><p class="c1"><span class="c6 c4"></span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">HOURS</span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">MINUTES</span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">RATE</span></p></td><td class="c17" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">AMOUNT</span></p></td><td class="c22" colspan="2" rowspan="3"><p class="c10"><span class="c6 c4">UIF</span></p><p class="c1"><span class="c31 c4 c46"></span></p><p class="c1"><span class="c31 c4 c46"></span></p><p class="c10"><span class="c6 c4">OTHER DEDUCTIONS</span></p><p class="c1"><span class="c6 c4"></span></p><p class="c10"><span class="c6 c4">OTHER DEDUCTIONS NOTES</span></p></td><td class="c11" colspan="1" rowspan="3"><p class="c3"><span class="c2">R{{UIF}}</span></p><p class="c1"><span class="c12"></span></p><p class="c1"><span class="c12"></span></p><p class="c3"><span>R</span><span class="c2">{{OTHER DEDUCTIONS}}<br></span></p><p class="c3"><span class="c2">{{OTHER DEDUCTIONS TEXT}}</span></p></td></tr><tr class="c19"><td class="c26" colspan="1" rowspan="1"><p class="c10"><span class="c4">NORMAL TIME</span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c3"><span class="c2">{{HOURS}}</span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c3"><span class="c2">{{MINUTES}}</span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c10"><span class="c2">R{{HOURLYRATE}}</span></p></td><td class="c17" colspan="1" rowspan="1"><p class="c10"><span class="c2">R{{STANDARDTIME}}</span></p></td></tr><tr class="c50"><td class="c26" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">OVER TIME</span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c3"><span class="c2">{{OVERTIMEHOURS}}</span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c3"><span class="c2">{{OVERTIMEMINUTES}}</span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c1"><span class="c2"></span></p></td><td class="c17" colspan="1" rowspan="1"><p class="c10"><span class="c2">R{{OVERTIME}}</span></p></td></tr><tr class="c19"><td class="c26" colspan="1" rowspan="3"><p class="c10"><span class="c6 c4">ADDITIONAL PAY</span></p></td><td class="c16" colspan="3" rowspan="1"><p class="c10"><span class="c6 c4">BONUS</span></p></td><td class="c17" colspan="1" rowspan="1"><p class="c3"><span class="c2">R{{BONUS PAY}}</span></p></td><td class="c15" colspan="1" rowspan="1"><p class="c1"><span class="c2"></span></p></td><td class="c51" colspan="1" rowspan="1"><p class="c10"><span class="c31 c23 c4">LOANS OPENING BALANCE</span></p></td><td class="c40" colspan="1" rowspan="1"><p class="c3"><span>R</span><span>{{CurrentLoanBalance}}</span></p></td></tr><tr class="c19"><td class="c16" colspan="3" rowspan="1"><p class="c10"><span class="c6 c4">LEAVE</span></p></td><td class="c17" colspan="1" rowspan="1"><p class="c3"><span class="c2">R{{LEAVE PAY}}</span></p></td><td class="c58" colspan="1" rowspan="1"><p class="c1"><span class="c2"></span></p></td><td class="c8" colspan="1" rowspan="1"><p class="c10"><span class="c4 c23">LOAN/REPAYMENT:<br></span><span>{{LoanDisbursementType}}</span></p></td><td class="c34" colspan="1" rowspan="1"><p class="c3"><span>R</span><span>{{LoanDeductionThisWeek}} {{NewLoanThisWeek}}</span></p></td></tr><tr class="c19"><td class="c16" colspan="3" rowspan="1"><p class="c10"><span class="c6 c4">OTHER</span></p></td><td class="c17" colspan="1" rowspan="1"><p class="c3"><span class="c2">R{{OTHERINCOME}}</span></p></td><td class="c58" colspan="1" rowspan="1"><p class="c1"><span class="c2"></span></p></td><td class="c39" colspan="1" rowspan="1"><p class="c10"><span class="c31 c23 c4">LOANS CLOSING BALANCE</span></p></td><td class="c35" colspan="1" rowspan="1"><p class="c3"><span>R</span><span>{{UpdatedLoanBalance}}</span></p></td></tr></table><p class="c1"><span class="c12"></span></p><table class="c9"><tr class="c7"><td class="c27" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">NOTES</span></p></td><td class="c24" colspan="1" rowspan="1"><p class="c3"><span class="c2">{{NOTES}}</span></p></td></tr></table><p class="c1"><span class="c12"></span></p><table class="c9"><tr class="c7"><td class="c60" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">GROSS PAY</span></p></td><td class="c62" colspan="1" rowspan="1"><p class="c44"><span class="c2">R{{GROSSSALARY}}</span></p></td><td class="c59" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">TOTAL DEDUCTIONS</span></p></td><td class="c25" colspan="1" rowspan="1"><p class="c21"><span class="c2">R{{TOTALDEDUCTIONS}}</span></p></td></tr></table><p class="c1"><span class="c18"></span></p><table class="c9"><tr class="c7"><td class="c45" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">NETT PAY</span></p></td><td class="c25" colspan="1" rowspan="1"><p class="c21"><span class="c2">R{{NETTSALARY}}</span></p></td></tr><tr class="c7"><td class="c45" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">AMOUNT PAID TO ACCOUNT </span></p></td><td class="c25" colspan="1" rowspan="1"><p class="c21"><span class="c2">R{{PaidtoAccount}}</span></p></td></tr></table><p class="c1"><span class="c2"></span></p><p class="c3 c13"><span class="c2"></span></p><p class="c3 c13"><span class="c2"></span></p></body></html>`;

    // Replace placeholders with actual values
    let htmlContent = htmlTemplate
      .replace(/\{\{RECORDNUMBER\}\}/g, payslip.RECORDNUMBER || '')
      .replace(/\{\{EMPLOYER\}\}/g, payslip.EMPLOYER || 'SA Grinding Wheels')
      .replace(/\{\{EMPLOYEE NAME\}\}/g, payslip['EMPLOYEE NAME'] || '')
      .replace(/\{\{WEEKENDING\}\}/g, formatDateForDisplay(payslip.WEEKENDING))
      .replace(/\{\{EMPLOYMENT STATUS\}\}/g, payslip['EMPLOYMENT STATUS'] || '')
      .replace(/\{\{HOURS\}\}/g, payslip.HOURS || '0')
      .replace(/\{\{MINUTES\}\}/g, payslip.MINUTES || '0')
      .replace(/\{\{HOURLYRATE\}\}/g, formatAmount(payslip.HOURLYRATE))
      .replace(/\{\{STANDARDTIME\}\}/g, formatAmount(payslip.STANDARDTIME))
      .replace(/\{\{OVERTIMEHOURS\}\}/g, payslip.OVERTIMEHOURS || '0')
      .replace(/\{\{OVERTIMEMINUTES\}\}/g, payslip.OVERTIMEMINUTES || '0')
      .replace(/\{\{OVERTIME\}\}/g, formatAmount(payslip.OVERTIME))
      .replace(/\{\{BONUS PAY\}\}/g, formatAmount(payslip['BONUS PAY']))
      .replace(/\{\{LEAVE PAY\}\}/g, formatAmount(payslip['LEAVE PAY']))
      .replace(/\{\{OTHERINCOME\}\}/g, formatAmount(payslip.OTHERINCOME))
      .replace(/\{\{UIF\}\}/g, formatAmount(payslip.UIF))
      .replace(/\{\{OTHER DEDUCTIONS\}\}/g, formatAmount(payslip['OTHER DEDUCTIONS']))
      .replace(/\{\{OTHER DEDUCTIONS TEXT\}\}/g, payslip['OTHER DEDUCTIONS TEXT'] || '')
      .replace(/\{\{CurrentLoanBalance\}\}/g, formatAmount(payslip.CurrentLoanBalance))
      .replace(/\{\{LoanDisbursementType\}\}/g, payslip.LoanDisbursementType || 'Separate')
      .replace(/\{\{LoanDeductionThisWeek\}\}/g, formatAmount(payslip.LoanDeductionThisWeek))
      .replace(/\{\{NewLoanThisWeek\}\}/g, payslip.NewLoanThisWeek > 0 ? formatAmount(payslip.NewLoanThisWeek) : '')
      .replace(/\{\{UpdatedLoanBalance\}\}/g, formatAmount(payslip.UpdatedLoanBalance))
      .replace(/\{\{NOTES\}\}/g, payslip.NOTES || '')
      .replace(/\{\{GROSSSALARY\}\}/g, formatAmount(payslip.GROSSSALARY))
      .replace(/\{\{TOTALDEDUCTIONS\}\}/g, formatAmount(payslip.TOTALDEDUCTIONS))
      .replace(/\{\{NETTSALARY\}\}/g, formatAmount(payslip.NETTSALARY))
      .replace(/\{\{PaidtoAccount\}\}/g, formatAmount(payslip.PaidtoAccount));

    // Create a temporary Google Doc from HTML
    const doc = DocumentApp.create(docName);
    const docId = doc.getId();
    const body = doc.getBody();

    // Clear default content
    body.clear();

    // We need to use a different approach - create HTML file and convert
    doc.saveAndClose();

    // Create HTML blob and convert to PDF via Google Drive
    const htmlBlob = Utilities.newBlob(htmlContent, 'text/html', docName + '.html');

    // Create temporary HTML file
    const htmlFile = DriveApp.createFile(htmlBlob);

    // Get the file and convert to PDF using export
    const pdfBlob = htmlFile.getAs('application/pdf');
    pdfBlob.setName(docName + '.pdf');

    // Save PDF to Drive
    const pdfFile = DriveApp.createFile(pdfBlob);
    const pdfUrl = pdfFile.getUrl();

    // Delete temporary files
    htmlFile.setTrashed(true);
    DriveApp.getFileById(docId).setTrashed(true);

    // Update payslip record with PDF link
    const sheets = getSheets();
    const salarySheet = sheets.salary;
    const allData = salarySheet.getDataRange().getValues();
    const headers = allData[0];
    const rows = allData.slice(1);

    const recordNumCol = findColumnIndex(headers, 'RECORDNUMBER');
    const fileNameCol = findColumnIndex(headers, 'FILENAME');
    const fileLinkCol = findColumnIndex(headers, 'FILELINK');

    const rowIndex = rows.findIndex(r => String(r[recordNumCol]) === String(recordNumber));
    if (rowIndex >= 0) {
      const sheetRowIndex = rowIndex + 2; // +1 for header, +1 for 0-index
      if (fileNameCol >= 0) {
        salarySheet.getRange(sheetRowIndex, fileNameCol + 1).setValue(pdfFile.getName());
      }
      if (fileLinkCol >= 0) {
        salarySheet.getRange(sheetRowIndex, fileLinkCol + 1).setValue(pdfUrl);
      }
    }

    Logger.log('✅ PDF generated successfully: ' + pdfFile.getName());
    Logger.log('✅ PDF URL: ' + pdfUrl);
    Logger.log('========== GENERATE PDF COMPLETE ==========\n');

    return {
      success: true,
      message: 'PDF generated successfully',
      data: {
        url: pdfUrl,
        filename: pdfFile.getName()
      }
    };

  } catch (error) {
    Logger.log('❌ ERROR in generatePayslipPDF: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    return { success: false, error: error.message };
  }
}

/**
 * Helper function to format dates for display
 */
function formatDateForDisplay(dateValue) {
  if (!dateValue) return '';
  const d = new Date(dateValue);
  const options = { year: 'numeric', month: 'long', day: 'numeric' };
  return d.toLocaleDateString('en-ZA', options);
}

/**
 * Helper function to format dates for filenames
 */
function formatDateForFilename(dateValue) {
  if (!dateValue) return '';
  const d = new Date(dateValue);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * Helper function to format amounts for display
 */
function formatAmount(amount) {
  if (amount === null || amount === undefined || amount === '') {
    return '0.00';
  }
  return parseFloat(amount).toFixed(2);
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
    PaidtoAccount: 1178.01
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
    PaidtoAccount: 2084.00  // Net + Loan
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
    PaidtoAccount: 1400.00
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
 * Delete a payslip by record number
 * Only allowed if the payslip is still editable (before end of Friday)
 */
function deletePayslip(recordNumber) {
  try {
    Logger.log('\n========== DELETE PAYSLIP ==========');
    Logger.log('ℹ️ Record Number: ' + recordNumber);

    const sheets = getSheets();
    const salarySheet = sheets.salary;
    if (!salarySheet) throw new Error('Salary sheet not found');

    // Find payslip
    const allData = salarySheet.getDataRange().getValues();
    const headers = allData[0];
    const rows = allData.slice(1);

    const recordNumCol = findColumnIndex(headers, 'RECORDNUMBER');
    const weekEndingCol = findColumnIndex(headers, 'WEEKENDING');
    const rowIndex = rows.findIndex(r => String(r[recordNumCol]) === String(recordNumber));

    if (rowIndex === -1) {
      throw new Error('Payslip not found: ' + recordNumber);
    }

    // Check if payslip is still editable
    const weekEnding = rows[rowIndex][weekEndingCol];
    if (!isPayslipEditable(weekEnding)) {
      throw new Error('Cannot delete payslip: The editing period has ended (past Friday midnight of the payslip week)');
    }

    // Delete the row (add 2 to account for header and 0-index)
    const sheetRowIndex = rowIndex + 2;
    salarySheet.deleteRow(sheetRowIndex);
    SpreadsheetApp.flush();

    Logger.log('✅ Payslip deleted: #' + recordNumber);
    Logger.log('========== DELETE PAYSLIP COMPLETE ==========\n');

    return { success: true, message: 'Payslip deleted successfully' };

  } catch (error) {
    Logger.log('❌ ERROR in deletePayslip: ' + error.message);
    Logger.log('========== DELETE PAYSLIP FAILED ==========\n');
    return { success: false, error: error.message };
  }
}

/**
 * Check if a payslip is still editable based on week ending date
 * Payslips are editable until the end of Friday of the payslip week
 *
 * @param {Date|string} weekEnding - The week ending date of the payslip
 * @returns {boolean} True if editable, false otherwise
 */
function isPayslipEditable(weekEnding) {
  if (!weekEnding) return false;

  const weekEndingDate = new Date(weekEnding);
  const now = new Date();

  // Get the Friday of the payslip week (week ending is always Friday)
  // Set to end of Friday (23:59:59)
  const editDeadline = new Date(weekEndingDate);
  editDeadline.setHours(23, 59, 59, 999);

  return now <= editDeadline;
}

/**
 * Get payslip editability status
 * Returns both the status and the deadline for display purposes
 */
function getPayslipEditStatus(recordNumber) {
  try {
    const result = getPayslip(recordNumber);
    if (!result.success) {
      return { success: false, error: result.error };
    }

    const payslip = result.data;
    const isEditable = isPayslipEditable(payslip.WEEKENDING);

    return {
      success: true,
      data: {
        isEditable: isEditable,
        weekEnding: payslip.WEEKENDING,
        message: isEditable
          ? 'Editable until end of Friday ' + formatDateForDisplay(payslip.WEEKENDING)
          : 'Locked - editing period ended'
      }
    };

  } catch (error) {
    Logger.log('❌ ERROR in getPayslipEditStatus: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Generate PDFs for multiple payslips (batch operation)
 *
 * @param {Array} recordNumbers - Array of record numbers to generate PDFs for
 * @returns {Object} Result with success count and any errors
 */
function generateBatchPayslipPDFs(recordNumbers) {
  try {
    Logger.log('\n========== BATCH GENERATE PAYSLIP PDFs ==========');
    Logger.log('ℹ️ Record Numbers: ' + JSON.stringify(recordNumbers));

    if (!recordNumbers || recordNumbers.length === 0) {
      throw new Error('No payslips selected');
    }

    const results = {
      success: true,
      total: recordNumbers.length,
      succeeded: 0,
      failed: 0,
      details: []
    };

    for (const recordNumber of recordNumbers) {
      try {
        const pdfResult = generatePayslipPDF(recordNumber);
        if (pdfResult.success) {
          results.succeeded++;
          results.details.push({
            recordNumber: recordNumber,
            success: true,
            url: pdfResult.data.url
          });
        } else {
          results.failed++;
          results.details.push({
            recordNumber: recordNumber,
            success: false,
            error: pdfResult.error
          });
        }
      } catch (error) {
        results.failed++;
        results.details.push({
          recordNumber: recordNumber,
          success: false,
          error: error.message
        });
      }
    }

    Logger.log('✅ Batch complete: ' + results.succeeded + '/' + results.total + ' succeeded');
    Logger.log('========== BATCH GENERATE COMPLETE ==========\n');

    return results;

  } catch (error) {
    Logger.log('❌ ERROR in generateBatchPayslipPDFs: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Update loan payment for a payslip (inline edit from list)
 *
 * @param {number} recordNumber - The payslip record number
 * @param {Object} loanData - Loan data {loanPayment, paymentType}
 * @returns {Object} Result
 */
function updatePayslipLoanPayment(recordNumber, loanData) {
  try {
    Logger.log('★★★ PAYROLL.GS VERSION: SYNC-2025-11-21-D ★★★');
    Logger.log('\n========== UPDATE LOAN PAYMENT ==========');
    Logger.log('ℹ️ Record Number: ' + recordNumber);
    Logger.log('ℹ️ Loan Data: ' + JSON.stringify(loanData));

    const sheets = getSheets();
    const salarySheet = sheets.salary;
    if (!salarySheet) throw new Error('Salary sheet not found');

    // Find payslip
    const allData = salarySheet.getDataRange().getValues();
    const headers = allData[0];
    const rows = allData.slice(1);

    const recordNumCol = findColumnIndex(headers, 'RECORDNUMBER');
    const weekEndingCol = findColumnIndex(headers, 'WEEKENDING');
    const rowIndex = rows.findIndex(r => String(r[recordNumCol]) === String(recordNumber));

    if (rowIndex === -1) {
      throw new Error('Payslip not found: ' + recordNumber);
    }

    // Check if payslip is still editable
    const weekEnding = rows[rowIndex][weekEndingCol];
    if (!isPayslipEditable(weekEnding)) {
      throw new Error('Cannot update payslip: The editing period has ended');
    }

    // Get current payslip
    const currentPayslip = buildObjectFromRow(rows[rowIndex], headers);

    // Get the payment type (use existing if not provided)
    const paymentType = loanData.paymentType || currentPayslip.LoanDisbursementType || 'Separate';
    const amount = loanData.loanPayment !== undefined
      ? parseFloat(loanData.loanPayment) || 0
      : (paymentType === 'Repayment'
          ? (parseFloat(currentPayslip.LoanDeductionThisWeek) || 0)
          : (parseFloat(currentPayslip.NewLoanThisWeek) || 0));

    // Update loan fields based on payment type
    // Repayment: Deduct from net pay, deduct from loan balance
    // With Salary: Add to paid to account, add to loan balance
    // Separate: Does not affect paid to account, add to loan balance

    if (paymentType === 'Repayment') {
      // This is a loan repayment - employee paying back
      currentPayslip.LoanDeductionThisWeek = amount;
      currentPayslip.NewLoanThisWeek = 0;
      currentPayslip.LoanDisbursementType = 'Repayment';
      Logger.log('📝 Set as Repayment: ' + amount);
    } else if (paymentType === 'With Salary') {
      // This is a new loan disbursed with salary
      currentPayslip.NewLoanThisWeek = amount;
      currentPayslip.LoanDeductionThisWeek = 0;
      currentPayslip.LoanDisbursementType = 'With Salary';
      Logger.log('📝 Set as New Loan With Salary: ' + amount);
    } else {
      // Separate - new loan disbursed separately
      currentPayslip.NewLoanThisWeek = amount;
      currentPayslip.LoanDeductionThisWeek = 0;
      currentPayslip.LoanDisbursementType = 'Separate';
      Logger.log('📝 Set as New Loan Separate: ' + amount);
    }

    // Recalculate payslip
    const calculations = calculatePayslip(currentPayslip);
    Object.assign(currentPayslip, calculations);

    // Add audit fields
    addAuditFields(currentPayslip, false);

    // Convert to row and update
    const updatedRow = objectToRow(currentPayslip, headers);
    const sheetRowIndex = rowIndex + 2;
    salarySheet.getRange(sheetRowIndex, 1, 1, headers.length).setValues([updatedRow]);
    SpreadsheetApp.flush();

    Logger.log('✅ Loan payment updated for payslip #' + recordNumber);
    Logger.log('✅ Paid to Account: ' + formatCurrency(currentPayslip.PaidtoAccount));

    // Sync loan transaction to EmployeeLoans sheet
    Logger.log('🔄🔄🔄 SYNC BLOCK START 🔄🔄🔄');
    try {
      Logger.log('🔄 Step 1: Preparing sync data from currentPayslip');
      Logger.log('🔄 currentPayslip.RECORDNUMBER: ' + currentPayslip.RECORDNUMBER);
      Logger.log('🔄 currentPayslip["EMPLOYEE NAME"]: ' + currentPayslip['EMPLOYEE NAME']);
      Logger.log('🔄 currentPayslip.id: ' + currentPayslip.id);
      Logger.log('🔄 currentPayslip.NewLoanThisWeek: ' + currentPayslip.NewLoanThisWeek);
      Logger.log('🔄 currentPayslip.LoanDeductionThisWeek: ' + currentPayslip.LoanDeductionThisWeek);
      Logger.log('🔄 currentPayslip.LoanDisbursementType: ' + currentPayslip.LoanDisbursementType);
      Logger.log('🔄 currentPayslip.UpdatedLoanBalance: ' + currentPayslip.UpdatedLoanBalance);

      const syncData = {
        employeeName: currentPayslip['EMPLOYEE NAME'],
        employeeId: currentPayslip.id,
        recordNumber: currentPayslip.RECORDNUMBER,
        newLoan: parseFloat(currentPayslip.NewLoanThisWeek) || 0,
        loanDeduction: parseFloat(currentPayslip.LoanDeductionThisWeek) || 0,
        disbursementType: currentPayslip.LoanDisbursementType || 'Separate',
        updatedBalance: parseFloat(currentPayslip.UpdatedLoanBalance) || 0,
        weekEnding: currentPayslip.WEEKENDING
      };

      Logger.log('🔄 Step 2: Sync data prepared: ' + JSON.stringify(syncData));

      syncLoanTransactionFromPayslip(syncData);
      Logger.log('✅ Loan transaction synced to EmployeeLoans sheet');
    } catch (syncError) {
      Logger.log('❌ ERROR: Failed to sync loan transaction: ' + syncError.message);
      Logger.log('❌ Stack: ' + syncError.stack);
    }
    Logger.log('🔄🔄🔄 SYNC BLOCK END 🔄🔄🔄');

    Logger.log('========== UPDATE LOAN PAYMENT COMPLETE ==========\n');

    // Sanitize for web
    const sanitizedPayslip = sanitizeForWeb(currentPayslip);
    return { success: true, data: sanitizedPayslip };

  } catch (error) {
    Logger.log('❌ ERROR in updatePayslipLoanPayment: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Generate a combined PDF with multiple payslips (batch print)
 *
 * @param {Array} recordNumbers - Array of record numbers to include
 * @returns {Object} Result with PDF URL
 */
function generateCombinedPayslipPDF(recordNumbers) {
  try {
    Logger.log('\n========== GENERATE COMBINED PAYSLIP PDF ==========');
    Logger.log('ℹ️ Record Numbers: ' + JSON.stringify(recordNumbers));

    if (!recordNumbers || recordNumbers.length === 0) {
      throw new Error('No payslips selected');
    }

    // Get all payslips
    const payslips = [];
    for (const recordNumber of recordNumbers) {
      const result = getPayslip(recordNumber);
      if (result.success) {
        payslips.push(result.data);
      } else {
        Logger.log('⚠️ Could not get payslip #' + recordNumber + ': ' + result.error);
      }
    }

    if (payslips.length === 0) {
      throw new Error('No valid payslips found');
    }

    // Create a Google Doc with all payslips
    const docName = `Payslips_Batch_${formatDateForFilename(new Date())}_${payslips.length}_records`;
    const doc = DocumentApp.create(docName);
    const body = doc.getBody();
    body.clear();

    // Add each payslip to the document
    payslips.forEach((payslip, index) => {
      // Add page break between payslips (not before first one)
      if (index > 0) {
        body.appendPageBreak();
      }

      // Header
      body.appendParagraph(payslip.EMPLOYER || 'SA Grinding Wheels')
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .setBold(true)
        .setFontSize(14);

      body.appendParagraph('PAYSLIP')
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .setFontSize(18)
        .setBold(true);

      body.appendParagraph('');

      // Employee info
      body.appendParagraph(`Payslip Number: #${payslip.RECORDNUMBER}`);
      body.appendParagraph(`Employee: ${payslip['EMPLOYEE NAME'] || ''}`);
      body.appendParagraph(`Employment Status: ${payslip['EMPLOYMENT STATUS'] || ''}`);
      body.appendParagraph(`Week Ending: ${formatDateForDisplay(payslip.WEEKENDING)}`);
      body.appendParagraph(`Hourly Rate: R${formatAmount(payslip.HOURLYRATE)}`);
      body.appendParagraph('');

      // Earnings table
      const table = body.appendTable();

      // Header row
      const headerRow = table.appendTableRow();
      headerRow.appendTableCell('Description').setBackgroundColor('#f2f2f2').setBold(true);
      headerRow.appendTableCell('Amount (R)').setBackgroundColor('#f2f2f2').setBold(true);

      // Standard time
      const row1 = table.appendTableRow();
      row1.appendTableCell(`Standard Time (${payslip.HOURS || 0}h ${payslip.MINUTES || 0}m)`);
      row1.appendTableCell(formatAmount(payslip.STANDARDTIME));

      // Overtime
      const row2 = table.appendTableRow();
      row2.appendTableCell(`Overtime (${payslip.OVERTIMEHOURS || 0}h ${payslip.OVERTIMEMINUTES || 0}m)`);
      row2.appendTableCell(formatAmount(payslip.OVERTIME));

      // Leave Pay (if any)
      if (payslip['LEAVE PAY'] > 0) {
        const rowLeave = table.appendTableRow();
        rowLeave.appendTableCell('Leave Pay');
        rowLeave.appendTableCell(formatAmount(payslip['LEAVE PAY']));
      }

      // Bonus Pay (if any)
      if (payslip['BONUS PAY'] > 0) {
        const rowBonus = table.appendTableRow();
        rowBonus.appendTableCell('Bonus Pay');
        rowBonus.appendTableCell(formatAmount(payslip['BONUS PAY']));
      }

      // Other Income (if any)
      if (payslip.OTHERINCOME > 0) {
        const rowOther = table.appendTableRow();
        rowOther.appendTableCell('Other Income');
        rowOther.appendTableCell(formatAmount(payslip.OTHERINCOME));
      }

      // Gross
      const row3 = table.appendTableRow();
      row3.appendTableCell('Gross Salary').setBold(true);
      row3.appendTableCell(formatAmount(payslip.GROSSSALARY)).setBold(true);

      // UIF
      const row4 = table.appendTableRow();
      row4.appendTableCell('UIF (1%)');
      row4.appendTableCell(`(${formatAmount(payslip.UIF)})`);

      // Other Deductions (if any)
      if (payslip['OTHER DEDUCTIONS'] > 0) {
        const rowDed = table.appendTableRow();
        const dedText = payslip['OTHER DEDUCTIONS TEXT']
          ? `Other Deductions - ${payslip['OTHER DEDUCTIONS TEXT']}`
          : 'Other Deductions';
        rowDed.appendTableCell(dedText);
        rowDed.appendTableCell(`(${formatAmount(payslip['OTHER DEDUCTIONS'])})`);
      }

      // Total Deductions
      const rowTotalDed = table.appendTableRow();
      rowTotalDed.appendTableCell('Total Deductions').setBold(true);
      rowTotalDed.appendTableCell(`(${formatAmount(payslip.TOTALDEDUCTIONS)})`).setBold(true);

      // Net
      const row5 = table.appendTableRow();
      row5.appendTableCell('Net Salary').setBold(true);
      row5.appendTableCell(formatAmount(payslip.NETTSALARY)).setBold(true);

      // Loan Deduction (if any)
      if (payslip.LoanDeductionThisWeek > 0) {
        const rowLoan = table.appendTableRow();
        rowLoan.appendTableCell('Loan Deduction');
        rowLoan.appendTableCell(`(${formatAmount(payslip.LoanDeductionThisWeek)})`);
      }

      // New Loan with Salary (if any)
      if (payslip.NewLoanThisWeek > 0 && payslip.LoanDisbursementType === 'With Salary') {
        const rowNewLoan = table.appendTableRow();
        rowNewLoan.appendTableCell('New Loan (With Salary)');
        rowNewLoan.appendTableCell(formatAmount(payslip.NewLoanThisWeek));
      }

      // Paid to Account
      const row6 = table.appendTableRow();
      row6.appendTableCell('PAID TO ACCOUNT').setBold(true);
      row6.appendTableCell(formatAmount(payslip.PaidtoAccount)).setBold(true).setBackgroundColor('#e8f5e9');

      // Notes (if any)
      if (payslip.NOTES) {
        body.appendParagraph('');
        body.appendParagraph('Notes: ' + payslip.NOTES).setItalic(true);
      }
    });

    // Footer
    body.appendParagraph('');
    body.appendParagraph(`Generated on ${formatDateForDisplay(new Date())} - ${payslips.length} payslip(s)`)
      .setFontSize(10)
      .setForegroundColor('#666666');

    doc.saveAndClose();

    // Convert to PDF
    const docId = doc.getId();
    const docFile = DriveApp.getFileById(docId);
    const pdfBlob = docFile.getAs('application/pdf');
    pdfBlob.setName(docName + '.pdf');

    // Save PDF to Drive
    const pdfFile = DriveApp.createFile(pdfBlob);
    const pdfUrl = pdfFile.getUrl();

    // Delete the temporary doc
    docFile.setTrashed(true);

    Logger.log('✅ Combined PDF generated: ' + pdfFile.getName());
    Logger.log('✅ Contains ' + payslips.length + ' payslips');
    Logger.log('✅ PDF URL: ' + pdfUrl);
    Logger.log('========== GENERATE COMBINED PDF COMPLETE ==========\n');

    return {
      success: true,
      data: {
        url: pdfUrl,
        filename: pdfFile.getName(),
        count: payslips.length
      }
    };

  } catch (error) {
    Logger.log('❌ ERROR in generateCombinedPayslipPDF: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    return { success: false, error: error.message };
  }
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
    PaidtoAccount: 1262.25
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

/**
 * Sync loan transaction from payslip to EmployeeLoans table
 * With comprehensive debugging to track failures
 *
 * @param {Object} data - Sync data from payslip
 */
function syncLoanTransactionFromPayslip(data) {
  Logger.log('🔄🔄🔄 syncLoanTransactionFromPayslip START 🔄🔄🔄');
  Logger.log('🔄 Input data: ' + JSON.stringify(data));

  try {
    // Step 1: Validate input
    Logger.log('🔄 Step 1: Validating input data...');
    if (!data) {
      throw new Error('No data provided to syncLoanTransactionFromPayslip');
    }
    if (!data.employeeName) {
      throw new Error('Missing employeeName in sync data');
    }
    if (!data.employeeId) {
      throw new Error('Missing employeeId in sync data');
    }
    Logger.log('✅ Step 1: Input validation passed');

    // Step 2: Check if there's actually a loan transaction
    Logger.log('🔄 Step 2: Checking for loan transaction...');
    const newLoan = parseFloat(data.newLoan) || 0;
    const loanDeduction = parseFloat(data.loanDeduction) || 0;
    Logger.log('🔄 newLoan: ' + newLoan);
    Logger.log('🔄 loanDeduction: ' + loanDeduction);

    if (newLoan === 0 && loanDeduction === 0) {
      Logger.log('ℹ️ No loan transaction to sync (both newLoan and loanDeduction are 0)');
      Logger.log('🔄🔄🔄 syncLoanTransactionFromPayslip END (no transaction) 🔄🔄🔄');
      return { success: true, message: 'No loan transaction to sync' };
    }
    Logger.log('✅ Step 2: Loan transaction found');

    // Step 3: Get the EmployeeLoans sheet
    Logger.log('🔄 Step 3: Getting EmployeeLoans sheet...');
    const sheets = getSheets();
    Logger.log('🔄 sheets object keys: ' + Object.keys(sheets).join(', '));

    const loansSheet = sheets.loans;
    if (!loansSheet) {
      Logger.log('❌ Available sheets: ' + Object.keys(sheets).join(', '));
      throw new Error('EmployeeLoans sheet not found in sheets object');
    }
    Logger.log('✅ Step 3: EmployeeLoans sheet found: ' + loansSheet.getName());

    // Step 4: Get headers and existing data
    Logger.log('🔄 Step 4: Getting sheet headers and data...');
    const allData = loansSheet.getDataRange().getValues();
    Logger.log('🔄 Total rows in sheet: ' + allData.length);

    if (allData.length === 0) {
      throw new Error('EmployeeLoans sheet is empty (no headers)');
    }

    const headers = allData[0];
    Logger.log('🔄 Headers: ' + JSON.stringify(headers));
    Logger.log('✅ Step 4: Headers retrieved (' + headers.length + ' columns)');

    // Step 5: Find column indices
    Logger.log('🔄 Step 5: Finding column indices...');
    const colIndices = {
      loanId: findColumnIndex(headers, 'LoanID'),
      employeeName: findColumnIndex(headers, 'Employee Name'),
      employeeId: findColumnIndex(headers, 'Employee ID'),
      timestamp: findColumnIndex(headers, 'Timestamp'),
      transactionDate: findColumnIndex(headers, 'TransactionDate'),
      loanAmount: findColumnIndex(headers, 'LoanAmount'),
      loanType: findColumnIndex(headers, 'LoanType'),
      disbursementMode: findColumnIndex(headers, 'DisbursementMode'),
      salaryLink: findColumnIndex(headers, 'SalaryLink'),
      notes: findColumnIndex(headers, 'Notes'),
      balanceBefore: findColumnIndex(headers, 'BalanceBefore'),
      balanceAfter: findColumnIndex(headers, 'BalanceAfter')
    };
    Logger.log('🔄 Column indices: ' + JSON.stringify(colIndices));

    // Check for missing columns
    for (const [key, value] of Object.entries(colIndices)) {
      if (value === -1) {
        Logger.log('⚠️ Warning: Column not found: ' + key);
      }
    }
    Logger.log('✅ Step 5: Column indices found');

    // Step 6: Check for existing record with same SalaryLink (payslip)
    Logger.log('🔄 Step 6: Checking for existing record for payslip #' + data.recordNumber + '...');
    let existingRowIndex = -1;
    let existingLoanId = null;
    const rows = allData.slice(1);

    if (colIndices.salaryLink !== -1) {
      for (let i = 0; i < rows.length; i++) {
        const salaryLinkVal = String(rows[i][colIndices.salaryLink]);
        if (salaryLinkVal === String(data.recordNumber)) {
          existingRowIndex = i + 2; // +2 because: +1 for header, +1 for 1-based index
          existingLoanId = rows[i][colIndices.loanId];
          Logger.log('🔄 Found existing record at row ' + existingRowIndex + ' with LoanID ' + existingLoanId);
          break;
        }
      }
    }

    // Generate new LoanID only if no existing record
    let loanId;
    if (existingRowIndex === -1) {
      let maxLoanId = 0;
      rows.forEach((row) => {
        const loanIdVal = row[colIndices.loanId];
        if (loanIdVal && typeof loanIdVal === 'number' && loanIdVal > maxLoanId) {
          maxLoanId = loanIdVal;
        }
      });
      loanId = maxLoanId + 1;
      Logger.log('🔄 No existing record found. New LoanID: ' + loanId);
    } else {
      loanId = existingLoanId;
      Logger.log('🔄 Will UPDATE existing record with LoanID: ' + loanId);
    }
    Logger.log('✅ Step 6: Record check complete');

    // Step 7: Calculate balances and transaction details
    Logger.log('🔄 Step 7: Calculating transaction details...');
    let loanAmount, loanType, disbursementMode;
    const updatedBalance = parseFloat(data.updatedBalance) || 0;

    if (newLoan > 0) {
      // Disbursement
      loanAmount = newLoan;
      loanType = 'Disbursement';
      disbursementMode = data.disbursementType || 'Separate';
    } else {
      // Repayment
      loanAmount = -loanDeduction; // Negative for repayment
      loanType = 'Repayment';
      disbursementMode = 'N/A';
    }

    const balanceBefore = updatedBalance - loanAmount;
    const balanceAfter = updatedBalance;

    Logger.log('🔄 loanAmount: ' + loanAmount);
    Logger.log('🔄 loanType: ' + loanType);
    Logger.log('🔄 disbursementMode: ' + disbursementMode);
    Logger.log('🔄 balanceBefore: ' + balanceBefore);
    Logger.log('🔄 balanceAfter: ' + balanceAfter);
    Logger.log('✅ Step 7: Transaction details calculated');

    // Step 8: Build the row data
    Logger.log('🔄 Step 8: Building row data...');
    const rowData = new Array(headers.length).fill('');

    if (colIndices.loanId !== -1) rowData[colIndices.loanId] = loanId;
    if (colIndices.employeeName !== -1) rowData[colIndices.employeeName] = data.employeeName;
    if (colIndices.employeeId !== -1) rowData[colIndices.employeeId] = data.employeeId;
    if (colIndices.timestamp !== -1) rowData[colIndices.timestamp] = new Date();
    if (colIndices.transactionDate !== -1) rowData[colIndices.transactionDate] = data.weekEnding || new Date();
    if (colIndices.loanAmount !== -1) rowData[colIndices.loanAmount] = loanAmount;
    if (colIndices.loanType !== -1) rowData[colIndices.loanType] = loanType;
    if (colIndices.disbursementMode !== -1) rowData[colIndices.disbursementMode] = disbursementMode;
    if (colIndices.salaryLink !== -1) rowData[colIndices.salaryLink] = data.recordNumber || '';
    if (colIndices.notes !== -1) rowData[colIndices.notes] = 'Auto-synced from payslip #' + (data.recordNumber || 'unknown');
    if (colIndices.balanceBefore !== -1) rowData[colIndices.balanceBefore] = balanceBefore;
    if (colIndices.balanceAfter !== -1) rowData[colIndices.balanceAfter] = balanceAfter;

    Logger.log('🔄 Row data: ' + JSON.stringify(rowData));
    Logger.log('✅ Step 8: Row built');

    // Step 9: Update existing row OR append new row
    if (existingRowIndex !== -1) {
      Logger.log('🔄 Step 9: UPDATING existing row ' + existingRowIndex + '...');
      loansSheet.getRange(existingRowIndex, 1, 1, headers.length).setValues([rowData]);
      SpreadsheetApp.flush();
      Logger.log('✅ Step 9: Row UPDATED successfully');
    } else {
      Logger.log('🔄 Step 9: APPENDING new row to EmployeeLoans sheet...');
      loansSheet.appendRow(rowData);
      SpreadsheetApp.flush();
      Logger.log('✅ Step 9: Row APPENDED successfully');
    }

    // Step 10: Verify the row was added
    Logger.log('🔄 Step 10: Verifying row was added...');
    const newData = loansSheet.getDataRange().getValues();
    Logger.log('🔄 New total rows: ' + newData.length);
    Logger.log('✅ Step 10: Verification complete');

    Logger.log('✅✅✅ LOAN TRANSACTION SYNCED SUCCESSFULLY ✅✅✅');
    Logger.log('🔄🔄🔄 syncLoanTransactionFromPayslip END 🔄🔄🔄');

    return { success: true, loanId: newLoanId };

  } catch (error) {
    Logger.log('❌❌❌ ERROR in syncLoanTransactionFromPayslip ❌❌❌');
    Logger.log('❌ Error message: ' + error.message);
    Logger.log('❌ Error stack: ' + error.stack);
    Logger.log('🔄🔄🔄 syncLoanTransactionFromPayslip END (with error) 🔄🔄🔄');
    throw error;
  }
}

/**
 * TEST FUNCTION - Run this directly from Apps Script editor
 * to verify code version and sync functionality
 */
function test_VerifyCodeVersion() {
  Logger.log('========== CODE VERSION VERIFICATION TEST ==========');
  Logger.log('');
  Logger.log('Expected version markers:');
  Logger.log('  - PAYROLL.GS: SYNC-2025-11-21-D');
  Logger.log('  - UTILS.GS: SYNC-2025-11-21-D');
  Logger.log('');

  // Check if syncLoanTransactionFromPayslip exists
  Logger.log('Checking if syncLoanTransactionFromPayslip exists...');
  if (typeof syncLoanTransactionFromPayslip === 'function') {
    Logger.log('✅ syncLoanTransactionFromPayslip function EXISTS');
  } else {
    Logger.log('❌ syncLoanTransactionFromPayslip function NOT FOUND');
  }

  // Check updatePayslipLoanPayment
  Logger.log('');
  Logger.log('Checking updatePayslipLoanPayment function...');
  const funcStr = updatePayslipLoanPayment.toString();

  if (funcStr.includes('SYNC BLOCK START')) {
    Logger.log('✅ Sync block EXISTS in updatePayslipLoanPayment');
  } else {
    Logger.log('❌ Sync block NOT FOUND in updatePayslipLoanPayment');
  }

  if (funcStr.includes('PAYROLL.GS VERSION')) {
    Logger.log('✅ Version marker EXISTS in updatePayslipLoanPayment');
  } else {
    Logger.log('❌ Version marker NOT FOUND in updatePayslipLoanPayment');
  }

  Logger.log('');
  Logger.log('========== TEST COMPLETE ==========');
}
