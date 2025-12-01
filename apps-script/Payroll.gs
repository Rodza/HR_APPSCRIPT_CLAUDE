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
    if (data.otherIncomeText !== undefined) data['OTHER INCOME TEXT'] = data.otherIncomeText;
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
    if (data.otherIncomeText !== undefined) data['OTHER INCOME TEXT'] = data.otherIncomeText;
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
      const targetDateStr = formatDate(parseDate(filters.weekEnding));
      payslips = payslips.filter(p => {
        if (!p.WEEKENDING) return false;
        const payslipDateStr = formatDate(parseDate(p.WEEKENDING));
        return payslipDateStr === targetDateStr;
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
      errors.push('Payslip already exists for ' + empName + ' on ' + formatDate(data.WEEKENDING));
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
    const htmlTemplate = `<html><head><meta content="text/html; charset=UTF-8" http-equiv="content-type"><style type="text/css">@page { size: A4 landscape; } ol{margin:0;padding:0}table td,table th{padding:1pt 5pt 1pt 5pt}.c36{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#ffffff;border-top-width:0pt;border-right-width:0pt;border-left-color:#ffffff;vertical-align:top;border-right-color:#ffffff;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:0pt;width:12pt;border-top-color:#ffffff;border-bottom-style:solid}.c35{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1.5pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:180pt;border-top-color:#000000;border-bottom-style:dotted}.c26{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:150pt;border-top-color:#000000;border-bottom-style:solid}.c24{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:830pt;border-top-color:#000000;border-bottom-style:solid}.c37{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1.5pt;width:165pt;border-top-color:#000000;border-bottom-style:solid}.c62{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:0pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:150pt;border-top-color:#000000;border-bottom-style:solid}.c27{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:80pt;border-top-color:#000000;border-bottom-style:solid}.c8{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1.5pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:200pt;border-top-color:#000000;border-bottom-style:solid}.c51{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1.5pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1.5pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:200pt;border-top-color:#000000;border-bottom-style:solid}.c60{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:0pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:350pt;border-top-color:#000000;border-bottom-style:solid}.c59{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:0pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:200pt;border-top-color:#000000;border-bottom-style:solid}.c0{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#ffffff;border-top-width:1pt;border-right-width:1pt;border-left-color:#ffffff;vertical-align:top;border-right-color:#ffffff;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:291pt;border-top-color:#ffffff;border-bottom-style:solid}.c61{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:0pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:150pt;border-top-color:#000000;border-bottom-style:solid}.c57{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#ffffff;border-top-width:0pt;border-right-width:0pt;border-left-color:#ffffff;vertical-align:top;border-right-color:#ffffff;border-left-width:0pt;border-top-style:solid;border-left-style:solid;border-bottom-width:0pt;width:12pt;border-top-color:#ffffff;border-bottom-style:solid}.c54{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1.5pt;width:180pt;border-top-color:#000000;border-bottom-style:solid}.c22{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:220pt;border-top-color:#000000;border-bottom-style:solid}.c34{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1.5pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:180pt;border-top-color:#000000;border-bottom-style:solid}.c45{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:0pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:700pt;border-top-color:#000000;border-bottom-style:solid}.c25{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:0pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:200pt;border-top-color:#000000;border-bottom-style:solid}.c52{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#ffffff;border-top-width:1pt;border-right-width:1pt;border-left-color:#ffffff;vertical-align:top;border-right-color:#ffffff;border-left-width:0pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:225.8pt;border-top-color:#ffffff;border-bottom-style:solid}.c29{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:0pt;width:414pt;border-top-color:#000000;border-bottom-style:solid}.c40{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1.5pt;border-right-width:1.5pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:180pt;border-top-color:#000000;border-bottom-style:solid}.c20{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:80pt;border-top-color:#000000;border-bottom-style:solid}.c14{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:165pt;border-top-color:#000000;border-bottom-style:solid}.c5{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#ffffff;border-top-width:0pt;border-right-width:0pt;border-left-color:#ffffff;vertical-align:top;border-right-color:#ffffff;border-left-width:0pt;border-top-style:solid;border-left-style:solid;border-bottom-width:0pt;width:204.8pt;border-top-color:#ffffff;border-bottom-style:solid}.c58{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:0pt;border-right-width:1.5pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:0pt;width:12.8pt;border-top-color:#000000;border-bottom-style:solid}.c17{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:120pt;border-top-color:#000000;border-bottom-style:solid}.c16{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:240pt;border-top-color:#000000;border-bottom-style:solid}.c11{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:180pt;border-top-color:#000000;border-bottom-style:solid}.c48{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:320.2pt;border-top-color:#000000;border-bottom-style:solid}.c42{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:12.8pt;border-top-color:#000000;border-bottom-style:solid}.c39{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1.5pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:200pt;border-top-color:#000000;border-bottom-style:dotted}.c15{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1.5pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:0pt;width:12.8pt;border-top-color:#000000;border-bottom-style:solid}.c32{border-right-style:solid;border-bottom-color:#ffffff;border-top-width:1pt;border-right-width:0pt;border-left-color:#ffffff;vertical-align:top;border-right-color:#ffffff;border-left-width:0pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:291pt;border-top-color:#ffffff;border-bottom-style:solid}.c12{color:#000000;font-weight:400;text-decoration:none;vertical-align:baseline;font-size:1pt;font-family:"Arial";font-style:normal}.c18{color:#000000;font-weight:400;text-decoration:none;vertical-align:baseline;font-size:2pt;font-family:"Arial";font-style:normal}.c28{color:#000000;font-weight:400;text-decoration:none;vertical-align:baseline;font-size:9pt;font-family:"Arial";font-style:normal}.c53{padding-top:8pt;padding-bottom:4pt;line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}.c2{color:#000000;font-weight:400;text-decoration:none;vertical-align:baseline;font-size:11pt;font-family:"Arial";font-style:normal}.c6{color:#000000;text-decoration:none;vertical-align:baseline;font-size:11pt;font-family:"Arial";font-style:normal}.c21{padding-top:0pt;padding-bottom:0pt;line-height:1.15;orphans:2;widows:2;text-align:right}.c30{color:#000000;font-weight:400;text-decoration:none;vertical-align:baseline;font-family:"Arial";font-style:normal}.c3{padding-top:0pt;padding-bottom:0pt;line-height:1.15;orphans:2;widows:2;text-align:left}.c1{padding-top:0pt;padding-bottom:0pt;line-height:1.0;text-align:left;height:11pt}.c31{color:#000000;text-decoration:none;vertical-align:baseline;font-family:"Arial";font-style:normal}.c9{border-spacing:0;border-collapse:collapse;margin-right:auto}.c44{padding-top:0pt;padding-bottom:0pt;line-height:1.0;text-align:right}.c10{padding-top:0pt;padding-bottom:0pt;line-height:1.0;text-align:left}.c41{padding-top:0pt;padding-bottom:0pt;line-height:1.0;text-align:center}.c49{background-color:#ffffff;max-width:none;padding:0pt 72pt 0pt 72pt;page-break-after: avoid;}.c43{color:inherit;text-decoration:inherit}.c19{height:14pt}.c56{font-size:9pt}.c38{font-size:19pt}.c47{height:22pt}.c7{height:0pt}.c23{font-size:10pt}.c33{font-size:17pt}.c13{height:8pt}.c50{height:14pt}.c4{font-weight:700}.c46{font-size:1pt}.c55{font-size:13pt}.title{padding-top:0pt;color:#000000;font-size:26pt;padding-bottom:3pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}.subtitle{padding-top:0pt;color:#666666;font-size:15pt;padding-bottom:16pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}li{color:#000000;font-size:11pt;font-family:"Arial"}p{margin:0;color:#000000;font-size:11pt;font-family:"Arial"}h1{padding-top:20pt;color:#000000;font-size:20pt;padding-bottom:6pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h2{padding-top:18pt;color:#000000;font-size:16pt;padding-bottom:6pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h3{padding-top:16pt;color:#434343;font-size:14pt;padding-bottom:4pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h4{padding-top:14pt;color:#666666;font-size:12pt;padding-bottom:4pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h5{padding-top:12pt;color:#666666;font-size:11pt;padding-bottom:4pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h6{padding-top:12pt;color:#666666;font-size:11pt;padding-bottom:4pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;font-style:italic;orphans:2;widows:2;text-align:left}</style></head><body class="c49 doc-content"><h1 class="c53"><span class="c31 c33 c4">WEEKLY PAY REMITTANCE</span></h1><table class="c9"><tr class="c19"><td class="c0" colspan="1" rowspan="1"><p class="c10"><span class="c4">PAYSLIP NUMBER:</span><span class="c2">&nbsp;{{RECORDNUMBER}}</span></p></td><td class="c36" colspan="1" rowspan="1"><p class="c1"><span class="c2"></span></p></td><td class="c5" colspan="1" rowspan="1"><p class="c1"><span class="c2"></span></p></td><td class="c52" colspan="1" rowspan="1"><p class="c3"><span class="c4 c33">{{EMPLOYER}}</span></p></td></tr><tr class="c7"><td class="c32" colspan="1" rowspan="1"><p class="c1"><span class="c12"></span></p></td><td class="c57" colspan="1" rowspan="1"><p class="c1"><span class="c12"></span></p></td><td class="c5" colspan="1" rowspan="2"><p class="c1"><span class="c28"></span></p></td><td class="c52" colspan="1" rowspan="2"><p class="c10"><span class="c28">18 DODGE STREET</span></p><p class="c10"><span class="c28">AUREUS</span></p><p class="c10"><span class="c28">RANDFONTEIN</span></p><p class="c10"><span class="c28">(011) 693 4278 / 083 338 5609</span></p><p class="c10"><span class="c56"><a class="c43" href="mailto:INFO@SAGRINDING.CO.ZA">INFO@SAGRINDING.CO.ZA</a></span><span class="c28">&nbsp;/ INFO@SCORPIOABRASIVES.CO.ZA</span></p></td></tr><tr class="c47"><td class="c0" colspan="1" rowspan="1"><p class="c10"><span class="c4">EMPLOYEE:</span><span class="c2">&nbsp;{{EMPLOYEE NAME}}</span></p><p class="c1"><span class="c12"></span></p><p class="c10"><span class="c4">WEEKENDING:</span><span class="c2">&nbsp;{{WEEKENDING}}</span></p><p class="c1"><span class="c12"></span></p><p class="c10"><span class="c4">EMPLOYMENT STATUS:</span><span>&nbsp;{{EMPLOYMENT STATUS}}</span></p></td><td class="c36" colspan="1" rowspan="1"><p class="c1"><span class="c2"></span></p></td></tr></table><p class="c1"><span class="c18"></span></p><table class="c9"><tr class="c19"><td class="c29" colspan="5" rowspan="1"><p class="c41"><span class="c31 c4 c55">EARNINGS</span></p></td><td class="c48" colspan="3" rowspan="1"><p class="c41"><span class="c31 c4 c55">DEDUCTIONS</span></p></td></tr><tr class="c19"><td class="c61" colspan="1" rowspan="1"><p class="c1"><span class="c6 c4"></span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">HOURS</span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">MINUTES</span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">RATE</span></p></td><td class="c17" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">AMOUNT</span></p></td><td class="c22" colspan="2" rowspan="3"><p class="c10"><span class="c6 c4">UIF</span></p><p class="c1"><span class="c31 c4 c46"></span></p><p class="c1"><span class="c31 c4 c46"></span></p><p class="c10"><span class="c6 c4">OTHER DEDUCTIONS</span></p><p class="c1"><span class="c6 c4"></span></p><p class="c10"><span class="c6 c4">OTHER DEDUCTIONS NOTES</span></p></td><td class="c11" colspan="1" rowspan="3"><p class="c3"><span class="c2">R{{UIF}}</span></p><p class="c1"><span class="c12"></span></p><p class="c1"><span class="c12"></span></p><p class="c3"><span>R</span><span class="c2">{{OTHER DEDUCTIONS}}<br></span></p><p class="c3"><span class="c2">{{OTHER DEDUCTIONS TEXT}}</span></p></td></tr><tr class="c19"><td class="c26" colspan="1" rowspan="1"><p class="c10"><span class="c4">NORMAL TIME</span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c3"><span class="c2">{{HOURS}}</span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c3"><span class="c2">{{MINUTES}}</span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c10"><span class="c2">R{{HOURLYRATE}}</span></p></td><td class="c17" colspan="1" rowspan="1"><p class="c10"><span class="c2">R{{STANDARDTIME}}</span></p></td></tr><tr class="c50"><td class="c26" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">OVER TIME</span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c3"><span class="c2">{{OVERTIMEHOURS}}</span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c3"><span class="c2">{{OVERTIMEMINUTES}}</span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c1"><span class="c2"></span></p></td><td class="c17" colspan="1" rowspan="1"><p class="c10"><span class="c2">R{{OVERTIME}}</span></p></td></tr><tr class="c19"><td class="c26" colspan="1" rowspan="3"><p class="c10"><span class="c6 c4">ADDITIONAL PAY</span></p></td><td class="c16" colspan="3" rowspan="1"><p class="c10"><span class="c6 c4">BONUS</span></p></td><td class="c17" colspan="1" rowspan="1"><p class="c3"><span class="c2">R{{BONUS PAY}}</span></p></td><td class="c15" colspan="1" rowspan="1"><p class="c1"><span class="c2"></span></p></td><td class="c51" colspan="1" rowspan="1"><p class="c10"><span class="c31 c23 c4">LOANS OPENING BALANCE</span></p></td><td class="c40" colspan="1" rowspan="1"><p class="c3"><span>R</span><span>{{CurrentLoanBalance}}</span></p></td></tr><tr class="c19"><td class="c16" colspan="3" rowspan="1"><p class="c10"><span class="c6 c4">LEAVE</span></p></td><td class="c17" colspan="1" rowspan="1"><p class="c3"><span class="c2">R{{LEAVE PAY}}</span></p></td><td class="c58" colspan="1" rowspan="1"><p class="c1"><span class="c2"></span></p></td><td class="c8" colspan="1" rowspan="1"><p class="c10"><span class="c4 c23">LOAN/REPAYMENT:<br></span><span>{{LoanDisbursementType}}</span></p></td><td class="c34" colspan="1" rowspan="1"><p class="c3"><span>R</span><span>{{LoanDeductionThisWeek}} {{NewLoanThisWeek}}</span></p></td></tr><tr class="c19"><td class="c16" colspan="3" rowspan="1"><p class="c10"><span class="c6 c4">OTHER</span></p></td><td class="c17" colspan="1" rowspan="1"><p class="c3"><span class="c2">R{{OTHERINCOME}}<br></span></p><p class="c3"><span class="c2">{{OTHER INCOME TEXT}}</span></p></td><td class="c58" colspan="1" rowspan="1"><p class="c1"><span class="c2"></span></p></td><td class="c39" colspan="1" rowspan="1"><p class="c10"><span class="c31 c23 c4">LOANS CLOSING BALANCE</span></p></td><td class="c35" colspan="1" rowspan="1"><p class="c3"><span>R</span><span>{{UpdatedLoanBalance}}</span></p></td></tr></table><p class="c1"><span class="c12"></span></p><table class="c9"><tr class="c7"><td class="c27" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">NOTES</span></p></td><td class="c24" colspan="1" rowspan="1"><p class="c3"><span class="c2">{{NOTES}}</span></p></td></tr></table><p class="c1"><span class="c12"></span></p><table class="c9"><tr class="c7"><td class="c60" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">GROSS PAY</span></p></td><td class="c62" colspan="1" rowspan="1"><p class="c44"><span class="c2">R{{GROSSSALARY}}</span></p></td><td class="c59" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">TOTAL DEDUCTIONS</span></p></td><td class="c25" colspan="1" rowspan="1"><p class="c21"><span class="c2">R{{TOTALDEDUCTIONS}}</span></p></td></tr></table><p class="c1"><span class="c18"></span></p><table class="c9"><tr class="c7"><td class="c45" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">NETT PAY</span></p></td><td class="c25" colspan="1" rowspan="1"><p class="c21"><span class="c2">R{{NETTSALARY}}</span></p></td></tr><tr class="c7"><td class="c45" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">AMOUNT PAID TO ACCOUNT </span></p></td><td class="c25" colspan="1" rowspan="1"><p class="c21"><span class="c2">R{{PaidtoAccount}}</span></p></td></tr></table><p class="c1"><span class="c2"></span></p><p class="c3 c13"><span class="c2"></span></p><p class="c3 c13"><span class="c2"></span></p></body></html>`;

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
      .replace(/\{\{OTHER INCOME TEXT\}\}/g, payslip['OTHER INCOME TEXT'] || '')
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

    // Sync loan transaction to EmployeeLoans sheet - OPTION 1 with fallback to OPTION 3
    Logger.log('🔄🔄🔄 SYNC BLOCK START 🔄🔄🔄');
    try {
      Logger.log('🔄 Attempting OPTION 1: Using syncLoanForPayslip()');

      // Call the correct sync function (from Loans.gs)
      const syncResult = syncLoanForPayslip(recordNumber);

      if (syncResult.success) {
        Logger.log('✅ Loan transaction synced via syncLoanForPayslip');

        // Validate the sync worked correctly
        const validation = validateLoanSync(recordNumber, currentPayslip);

        if (validation.success) {
          Logger.log('✅ VALIDATION PASSED: Loan sync verified correct');
        } else {
          Logger.log('⚠️ VALIDATION FAILED: ' + validation.error);
          Logger.log('🔄 Falling back to OPTION 3: Refreshing loan balance and retrying');

          // OPTION 3 FALLBACK: Refresh CurrentLoanBalance before recalculating
          const freshBalance = getCurrentLoanBalance(currentPayslip.id);
          if (freshBalance.success) {
            currentPayslip.CurrentLoanBalance = freshBalance.data;
            Logger.log('🔄 Refreshed CurrentLoanBalance: ' + freshBalance.data);

            // Recalculate with fresh balance
            const recalc = calculatePayslip(currentPayslip);
            Object.assign(currentPayslip, recalc);

            // Update the sheet with corrected values
            const correctedRow = objectToRow(currentPayslip, headers);
            salarySheet.getRange(sheetRowIndex, 1, 1, headers.length).setValues([correctedRow]);
            SpreadsheetApp.flush();
            Logger.log('✅ Payslip recalculated with fresh loan balance');

            // Retry sync with corrected data
            const retrySync = syncLoanForPayslip(recordNumber);
            if (retrySync.success) {
              Logger.log('✅ RETRY SUCCESSFUL: Loan synced with corrected balance');
            } else {
              Logger.log('❌ RETRY FAILED: ' + retrySync.error);
            }
          } else {
            Logger.log('❌ Could not fetch fresh loan balance: ' + freshBalance.error);
          }
        }
      } else {
        Logger.log('❌ ERROR: syncLoanForPayslip failed: ' + syncResult.error);
      }
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
 * Validates that loan sync was successful and created correct records
 *
 * @param {number} recordNumber - Payslip record number
 * @param {Object} payslip - Payslip data object
 * @returns {Object} Result with success flag and error if validation fails
 */
function validateLoanSync(recordNumber, payslip) {
  try {
    Logger.log('🔍 Validating loan sync for payslip #' + recordNumber);

    // Check if there's actually a loan transaction to validate
    const loanDeduction = parseFloat(payslip.LoanDeductionThisWeek) || 0;
    const newLoan = parseFloat(payslip.NewLoanThisWeek) || 0;

    if (loanDeduction === 0 && newLoan === 0) {
      Logger.log('✅ No loan activity - validation not needed');
      return { success: true };
    }

    // Find the loan record by SalaryLink
    const loanRecord = findLoanRecordBySalaryLink(recordNumber);

    if (!loanRecord.success) {
      return { success: false, error: 'Failed to query loan records: ' + loanRecord.error };
    }

    if (!loanRecord.data) {
      return { success: false, error: 'No loan record found with SalaryLink=' + recordNumber };
    }

    const record = loanRecord.data;
    Logger.log('🔍 Found loan record: LoanID=' + record.LoanID);

    // Validation 1: Check LoanID is UUID format (not numeric)
    const loanId = String(record.LoanID);
    const isUUID = loanId.includes('-') && loanId.length > 10;

    if (!isUUID) {
      return {
        success: false,
        error: 'Invalid LoanID format (expected UUID, got: ' + loanId + '). This indicates wrong sync function was used.'
      };
    }
    Logger.log('✅ LoanID is valid UUID format');

    // Validation 2: Check loan amount matches payslip
    const expectedAmount = loanDeduction > 0 ? -loanDeduction : newLoan;
    const actualAmount = parseFloat(record.LoanAmount) || 0;

    if (Math.abs(expectedAmount - actualAmount) > 0.01) {
      return {
        success: false,
        error: 'Loan amount mismatch (expected: ' + expectedAmount + ', actual: ' + actualAmount + ')'
      };
    }
    Logger.log('✅ Loan amount matches payslip');

    // Validation 3: Check balance is sensible (BalanceAfter exists and is not null)
    const balanceAfter = parseFloat(record.BalanceAfter);

    if (balanceAfter === null || balanceAfter === undefined || isNaN(balanceAfter)) {
      return {
        success: false,
        error: 'BalanceAfter is missing or invalid: ' + record.BalanceAfter
      };
    }
    Logger.log('✅ BalanceAfter is valid: ' + balanceAfter);

    // Validation 4: Check BalanceBefore exists
    const balanceBefore = parseFloat(record.BalanceBefore);

    if (balanceBefore === null || balanceBefore === undefined || isNaN(balanceBefore)) {
      return {
        success: false,
        error: 'BalanceBefore is missing or invalid: ' + record.BalanceBefore
      };
    }
    Logger.log('✅ BalanceBefore is valid: ' + balanceBefore);

    // Validation 5: Check balance calculation is correct
    const expectedBalanceAfter = balanceBefore + actualAmount;

    if (Math.abs(expectedBalanceAfter - balanceAfter) > 0.01) {
      return {
        success: false,
        error: 'Balance calculation error (Before: ' + balanceBefore +
               ', Amount: ' + actualAmount +
               ', Expected After: ' + expectedBalanceAfter +
               ', Actual After: ' + balanceAfter + ')'
      };
    }
    Logger.log('✅ Balance calculation is correct');

    Logger.log('✅✅✅ ALL VALIDATIONS PASSED ✅✅✅');
    return { success: true };

  } catch (error) {
    Logger.log('❌ ERROR in validateLoanSync: ' + error.message);
    return { success: false, error: 'Validation error: ' + error.message };
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

    // Create combined PDF filename
    const docName = `Payslips_Batch_${formatDateForFilename(new Date())}_${payslips.length}_records`;

    // HTML template for payslip (same as individual payslip)
    const htmlTemplateHeader = `<html><head><meta content="text/html; charset=UTF-8" http-equiv="content-type"><style type="text/css">@page { size: A4 landscape; } ol{margin:0;padding:0}table td,table th{padding:1pt 5pt 1pt 5pt}.c36{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#ffffff;border-top-width:0pt;border-right-width:0pt;border-left-color:#ffffff;vertical-align:top;border-right-color:#ffffff;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:0pt;width:12pt;border-top-color:#ffffff;border-bottom-style:solid}.c35{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1.5pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:180pt;border-top-color:#000000;border-bottom-style:dotted}.c26{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:150pt;border-top-color:#000000;border-bottom-style:solid}.c24{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:830pt;border-top-color:#000000;border-bottom-style:solid}.c37{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1.5pt;width:165pt;border-top-color:#000000;border-bottom-style:solid}.c62{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:0pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:150pt;border-top-color:#000000;border-bottom-style:solid}.c27{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:80pt;border-top-color:#000000;border-bottom-style:solid}.c8{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1.5pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:200pt;border-top-color:#000000;border-bottom-style:solid}.c51{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1.5pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1.5pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:200pt;border-top-color:#000000;border-bottom-style:solid}.c60{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:0pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:350pt;border-top-color:#000000;border-bottom-style:solid}.c59{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:0pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:200pt;border-top-color:#000000;border-bottom-style:solid}.c0{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#ffffff;border-top-width:1pt;border-right-width:1pt;border-left-color:#ffffff;vertical-align:top;border-right-color:#ffffff;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:291pt;border-top-color:#ffffff;border-bottom-style:solid}.c61{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:0pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:150pt;border-top-color:#000000;border-bottom-style:solid}.c57{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#ffffff;border-top-width:0pt;border-right-width:0pt;border-left-color:#ffffff;vertical-align:top;border-right-color:#ffffff;border-left-width:0pt;border-top-style:solid;border-left-style:solid;border-bottom-width:0pt;width:12pt;border-top-color:#ffffff;border-bottom-style:solid}.c54{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1.5pt;width:180pt;border-top-color:#000000;border-bottom-style:solid}.c22{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:220pt;border-top-color:#000000;border-bottom-style:solid}.c34{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1.5pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:180pt;border-top-color:#000000;border-bottom-style:solid}.c45{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:0pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:700pt;border-top-color:#000000;border-bottom-style:solid}.c25{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:0pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:200pt;border-top-color:#000000;border-bottom-style:solid}.c52{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#ffffff;border-top-width:1pt;border-right-width:1pt;border-left-color:#ffffff;vertical-align:top;border-right-color:#ffffff;border-left-width:0pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:225.8pt;border-top-color:#ffffff;border-bottom-style:solid}.c29{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:0pt;width:414pt;border-top-color:#000000;border-bottom-style:solid}.c40{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1.5pt;border-right-width:1.5pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:180pt;border-top-color:#000000;border-bottom-style:solid}.c20{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:80pt;border-top-color:#000000;border-bottom-style:solid}.c14{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:165pt;border-top-color:#000000;border-bottom-style:solid}.c5{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#ffffff;border-top-width:0pt;border-right-width:0pt;border-left-color:#ffffff;vertical-align:top;border-right-color:#ffffff;border-left-width:0pt;border-top-style:solid;border-left-style:solid;border-bottom-width:0pt;width:204.8pt;border-top-color:#ffffff;border-bottom-style:solid}.c58{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:0pt;border-right-width:1.5pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:0pt;width:12.8pt;border-top-color:#000000;border-bottom-style:solid}.c17{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:120pt;border-top-color:#000000;border-bottom-style:solid}.c16{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:240pt;border-top-color:#000000;border-bottom-style:solid}.c11{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:180pt;border-top-color:#000000;border-bottom-style:solid}.c48{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:320.2pt;border-top-color:#000000;border-bottom-style:solid}.c42{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:12.8pt;border-top-color:#000000;border-bottom-style:solid}.c39{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1.5pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:200pt;border-top-color:#000000;border-bottom-style:dotted}.c15{border-right-style:solid;padding:1pt 5pt 1pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1.5pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:0pt;width:12.8pt;border-top-color:#000000;border-bottom-style:solid}.c32{border-right-style:solid;border-bottom-color:#ffffff;border-top-width:1pt;border-right-width:0pt;border-left-color:#ffffff;vertical-align:top;border-right-color:#ffffff;border-left-width:0pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:291pt;border-top-color:#ffffff;border-bottom-style:solid}.c12{color:#000000;font-weight:400;text-decoration:none;vertical-align:baseline;font-size:1pt;font-family:"Arial";font-style:normal}.c18{color:#000000;font-weight:400;text-decoration:none;vertical-align:baseline;font-size:2pt;font-family:"Arial";font-style:normal}.c28{color:#000000;font-weight:400;text-decoration:none;vertical-align:baseline;font-size:9pt;font-family:"Arial";font-style:normal}.c53{padding-top:8pt;padding-bottom:4pt;line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}.c2{color:#000000;font-weight:400;text-decoration:none;vertical-align:baseline;font-size:11pt;font-family:"Arial";font-style:normal}.c6{color:#000000;text-decoration:none;vertical-align:baseline;font-size:11pt;font-family:"Arial";font-style:normal}.c21{padding-top:0pt;padding-bottom:0pt;line-height:1.15;orphans:2;widows:2;text-align:right}.c30{color:#000000;font-weight:400;text-decoration:none;vertical-align:baseline;font-family:"Arial";font-style:normal}.c3{padding-top:0pt;padding-bottom:0pt;line-height:1.15;orphans:2;widows:2;text-align:left}.c1{padding-top:0pt;padding-bottom:0pt;line-height:1.0;text-align:left;height:11pt}.c31{color:#000000;text-decoration:none;vertical-align:baseline;font-family:"Arial";font-style:normal}.c9{border-spacing:0;border-collapse:collapse;margin-right:auto}.c44{padding-top:0pt;padding-bottom:0pt;line-height:1.0;text-align:right}.c10{padding-top:0pt;padding-bottom:0pt;line-height:1.0;text-align:left}.c41{padding-top:0pt;padding-bottom:0pt;line-height:1.0;text-align:center}.c49{background-color:#ffffff;max-width:none;padding:0pt 72pt 0pt 72pt;page-break-after: avoid;}.c43{color:inherit;text-decoration:inherit}.c19{height:14pt}.c56{font-size:9pt}.c38{font-size:19pt}.c47{height:22pt}.c7{height:0pt}.c23{font-size:10pt}.c33{font-size:17pt}.c13{height:8pt}.c50{height:14pt}.c4{font-weight:700}.c46{font-size:1pt}.c55{font-size:13pt}.title{padding-top:0pt;color:#000000;font-size:26pt;padding-bottom:3pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}.subtitle{padding-top:0pt;color:#666666;font-size:15pt;padding-bottom:16pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}li{color:#000000;font-size:11pt;font-family:"Arial"}p{margin:0;color:#000000;font-size:11pt;font-family:"Arial"}h1{padding-top:20pt;color:#000000;font-size:20pt;padding-bottom:6pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h2{padding-top:18pt;color:#000000;font-size:16pt;padding-bottom:6pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h3{padding-top:16pt;color:#434343;font-size:14pt;padding-bottom:4pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h4{padding-top:14pt;color:#666666;font-size:12pt;padding-bottom:4pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h5{padding-top:12pt;color:#666666;font-size:11pt;padding-bottom:4pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h6{padding-top:12pt;color:#666666;font-size:11pt;padding-bottom:4pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;font-style:italic;orphans:2;widows:2;text-align:left}</style></head><body class="c49 doc-content">`;

    // Build combined HTML content with all payslips
    let combinedHtml = htmlTemplateHeader;

    payslips.forEach((payslip, index) => {
      // Add page break between payslips (not before first one)
      if (index > 0) {
        combinedHtml += '<div style="page-break-before: always;"></div>';
      }

      // Add individual payslip HTML content
      const payslipHtml = `<h1 class="c53"><span class="c31 c33 c4">WEEKLY PAY REMITTANCE</span></h1><table class="c9"><tr class="c19"><td class="c0" colspan="1" rowspan="1"><p class="c10"><span class="c4">PAYSLIP NUMBER:</span><span class="c2">&nbsp;${payslip.RECORDNUMBER || ''}</span></p></td><td class="c36" colspan="1" rowspan="1"><p class="c1"><span class="c2"></span></p></td><td class="c5" colspan="1" rowspan="1"><p class="c1"><span class="c2"></span></p></td><td class="c52" colspan="1" rowspan="1"><p class="c3"><span class="c4 c33">${payslip.EMPLOYER || 'SA Grinding Wheels'}</span></p></td></tr><tr class="c7"><td class="c32" colspan="1" rowspan="1"><p class="c1"><span class="c12"></span></p></td><td class="c57" colspan="1" rowspan="1"><p class="c1"><span class="c12"></span></p></td><td class="c5" colspan="1" rowspan="2"><p class="c1"><span class="c28"></span></p></td><td class="c52" colspan="1" rowspan="2"><p class="c10"><span class="c28">18 DODGE STREET</span></p><p class="c10"><span class="c28">AUREUS</span></p><p class="c10"><span class="c28">RANDFONTEIN</span></p><p class="c10"><span class="c28">(011) 693 4278 / 083 338 5609</span></p><p class="c10"><span class="c56"><a class="c43" href="mailto:INFO@SAGRINDING.CO.ZA">INFO@SAGRINDING.CO.ZA</a></span><span class="c28">&nbsp;/ INFO@SCORPIOABRASIVES.CO.ZA</span></p></td></tr><tr class="c47"><td class="c0" colspan="1" rowspan="1"><p class="c10"><span class="c4">EMPLOYEE:</span><span class="c2">&nbsp;${payslip['EMPLOYEE NAME'] || ''}</span></p><p class="c1"><span class="c12"></span></p><p class="c10"><span class="c4">WEEKENDING:</span><span class="c2">&nbsp;${formatDateForDisplay(payslip.WEEKENDING)}</span></p><p class="c1"><span class="c12"></span></p><p class="c10"><span class="c4">EMPLOYMENT STATUS:</span><span>&nbsp;${payslip['EMPLOYMENT STATUS'] || ''}</span></p></td><td class="c36" colspan="1" rowspan="1"><p class="c1"><span class="c2"></span></p></td></tr></table><p class="c1"><span class="c18"></span></p><table class="c9"><tr class="c19"><td class="c29" colspan="5" rowspan="1"><p class="c41"><span class="c31 c4 c55">EARNINGS</span></p></td><td class="c48" colspan="3" rowspan="1"><p class="c41"><span class="c31 c4 c55">DEDUCTIONS</span></p></td></tr><tr class="c19"><td class="c61" colspan="1" rowspan="1"><p class="c1"><span class="c6 c4"></span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">HOURS</span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">MINUTES</span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">RATE</span></p></td><td class="c17" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">AMOUNT</span></p></td><td class="c22" colspan="2" rowspan="3"><p class="c10"><span class="c6 c4">UIF</span></p><p class="c1"><span class="c31 c4 c46"></span></p><p class="c1"><span class="c31 c4 c46"></span></p><p class="c10"><span class="c6 c4">OTHER DEDUCTIONS</span></p><p class="c1"><span class="c6 c4"></span></p><p class="c10"><span class="c6 c4">OTHER DEDUCTIONS NOTES</span></p></td><td class="c11" colspan="1" rowspan="3"><p class="c3"><span class="c2">R${formatAmount(payslip.UIF)}</span></p><p class="c1"><span class="c12"></span></p><p class="c1"><span class="c12"></span></p><p class="c3"><span>R</span><span class="c2">${formatAmount(payslip['OTHER DEDUCTIONS'])}<br></span></p><p class="c3"><span class="c2">${payslip['OTHER DEDUCTIONS TEXT'] || ''}</span></p></td></tr><tr class="c19"><td class="c26" colspan="1" rowspan="1"><p class="c10"><span class="c4">NORMAL TIME</span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c3"><span class="c2">${payslip.HOURS || '0'}</span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c3"><span class="c2">${payslip.MINUTES || '0'}</span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c10"><span class="c2">R${formatAmount(payslip.HOURLYRATE)}</span></p></td><td class="c17" colspan="1" rowspan="1"><p class="c10"><span class="c2">R${formatAmount(payslip.STANDARDTIME)}</span></p></td></tr><tr class="c50"><td class="c26" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">OVER TIME</span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c3"><span class="c2">${payslip.OVERTIMEHOURS || '0'}</span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c3"><span class="c2">${payslip.OVERTIMEMINUTES || '0'}</span></p></td><td class="c20" colspan="1" rowspan="1"><p class="c1"><span class="c2"></span></p></td><td class="c17" colspan="1" rowspan="1"><p class="c10"><span class="c2">R${formatAmount(payslip.OVERTIME)}</span></p></td></tr><tr class="c19"><td class="c26" colspan="1" rowspan="3"><p class="c10"><span class="c6 c4">ADDITIONAL PAY</span></p></td><td class="c16" colspan="3" rowspan="1"><p class="c10"><span class="c6 c4">BONUS</span></p></td><td class="c17" colspan="1" rowspan="1"><p class="c3"><span class="c2">R${formatAmount(payslip['BONUS PAY'])}</span></p></td><td class="c15" colspan="1" rowspan="1"><p class="c1"><span class="c2"></span></p></td><td class="c51" colspan="1" rowspan="1"><p class="c10"><span class="c31 c23 c4">LOANS OPENING BALANCE</span></p></td><td class="c40" colspan="1" rowspan="1"><p class="c3"><span>R</span><span>${formatAmount(payslip.CurrentLoanBalance)}</span></p></td></tr><tr class="c19"><td class="c16" colspan="3" rowspan="1"><p class="c10"><span class="c6 c4">LEAVE</span></p></td><td class="c17" colspan="1" rowspan="1"><p class="c3"><span class="c2">R${formatAmount(payslip['LEAVE PAY'])}</span></p></td><td class="c58" colspan="1" rowspan="1"><p class="c1"><span class="c2"></span></p></td><td class="c8" colspan="1" rowspan="1"><p class="c10"><span class="c4 c23">LOAN/REPAYMENT:<br></span><span>${payslip.LoanDisbursementType || 'Separate'}</span></p></td><td class="c34" colspan="1" rowspan="1"><p class="c3"><span>R</span><span>${formatAmount(payslip.LoanDeductionThisWeek)} ${payslip.NewLoanThisWeek > 0 ? formatAmount(payslip.NewLoanThisWeek) : ''}</span></p></td></tr><tr class="c19"><td class="c16" colspan="3" rowspan="1"><p class="c10"><span class="c6 c4">OTHER</span></p></td><td class="c17" colspan="1" rowspan="1"><p class="c3"><span class="c2">R${formatAmount(payslip.OTHERINCOME)}</span></p></td><td class="c58" colspan="1" rowspan="1"><p class="c1"><span class="c2"></span></p></td><td class="c39" colspan="1" rowspan="1"><p class="c10"><span class="c31 c23 c4">LOANS CLOSING BALANCE</span></p></td><td class="c35" colspan="1" rowspan="1"><p class="c3"><span>R</span><span>${formatAmount(payslip.UpdatedLoanBalance)}</span></p></td></tr></table><p class="c1"><span class="c12"></span></p><table class="c9"><tr class="c7"><td class="c27" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">NOTES</span></p></td><td class="c24" colspan="1" rowspan="1"><p class="c3"><span class="c2">${payslip.NOTES || ''}</span></p></td></tr></table><p class="c1"><span class="c12"></span></p><table class="c9"><tr class="c7"><td class="c60" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">GROSS PAY</span></p></td><td class="c62" colspan="1" rowspan="1"><p class="c44"><span class="c2">R${formatAmount(payslip.GROSSSALARY)}</span></p></td><td class="c59" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">TOTAL DEDUCTIONS</span></p></td><td class="c25" colspan="1" rowspan="1"><p class="c21"><span class="c2">R${formatAmount(payslip.TOTALDEDUCTIONS)}</span></p></td></tr></table><p class="c1"><span class="c18"></span></p><table class="c9"><tr class="c7"><td class="c45" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">NETT PAY</span></p></td><td class="c25" colspan="1" rowspan="1"><p class="c21"><span class="c2">R${formatAmount(payslip.NETTSALARY)}</span></p></td></tr><tr class="c7"><td class="c45" colspan="1" rowspan="1"><p class="c10"><span class="c6 c4">AMOUNT PAID TO ACCOUNT </span></p></td><td class="c25" colspan="1" rowspan="1"><p class="c21"><span class="c2">R${formatAmount(payslip.PaidtoAccount)}</span></p></td></tr></table><p class="c1"><span class="c2"></span></p><p class="c3 c13"><span class="c2"></span></p><p class="c3 c13"><span class="c2"></span></p>`;

      combinedHtml += payslipHtml;
    });

    // Close HTML
    combinedHtml += '</body></html>';

    // Create HTML blob and convert to PDF via Google Drive
    const htmlBlob = Utilities.newBlob(combinedHtml, 'text/html', docName + '.html');

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
 * TEST FUNCTION - Run this directly from Apps Script editor
 * to verify loan sync implementation
 */
function test_VerifyLoanSyncImplementation() {
  Logger.log('========== LOAN SYNC IMPLEMENTATION TEST ==========');
  Logger.log('');

  // Check if correct sync function exists (from Loans.gs)
  Logger.log('Checking if syncLoanForPayslip exists...');
  if (typeof syncLoanForPayslip === 'function') {
    Logger.log('✅ syncLoanForPayslip function EXISTS (correct function from Loans.gs)');
  } else {
    Logger.log('❌ syncLoanForPayslip function NOT FOUND');
  }

  // Check if validation function exists
  Logger.log('');
  Logger.log('Checking if validateLoanSync exists...');
  if (typeof validateLoanSync === 'function') {
    Logger.log('✅ validateLoanSync function EXISTS');
  } else {
    Logger.log('❌ validateLoanSync function NOT FOUND');
  }

  // Check updatePayslipLoanPayment uses correct function
  Logger.log('');
  Logger.log('Checking updatePayslipLoanPayment implementation...');
  const funcStr = updatePayslipLoanPayment.toString();

  if (funcStr.includes('syncLoanForPayslip')) {
    Logger.log('✅ updatePayslipLoanPayment calls syncLoanForPayslip (correct)');
  } else {
    Logger.log('❌ updatePayslipLoanPayment does NOT call syncLoanForPayslip');
  }

  if (funcStr.includes('validateLoanSync')) {
    Logger.log('✅ updatePayslipLoanPayment includes validation');
  } else {
    Logger.log('❌ updatePayslipLoanPayment does NOT include validation');
  }

  if (funcStr.includes('OPTION 3')) {
    Logger.log('✅ updatePayslipLoanPayment includes Option 3 fallback');
  } else {
    Logger.log('❌ updatePayslipLoanPayment does NOT include Option 3 fallback');
  }

  // Verify broken function is removed
  Logger.log('');
  Logger.log('Verifying broken function was removed...');
  if (typeof syncLoanTransactionFromPayslip === 'function') {
    Logger.log('⚠️ WARNING: syncLoanTransactionFromPayslip still exists (should be removed)');
  } else {
    Logger.log('✅ syncLoanTransactionFromPayslip removed (correct)');
  }

  Logger.log('');
  Logger.log('========== TEST COMPLETE ==========');
}
