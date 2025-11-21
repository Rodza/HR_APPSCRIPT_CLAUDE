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

    // Sync loan transaction to EmployeeLoans sheet
    try {
      syncLoanTransactionFromPayslip(data);
      Logger.log('✅ Loan transaction synced to EmployeeLoans sheet');
    } catch (syncError) {
      Logger.log('❌ ERROR: Failed to sync loan transaction: ' + syncError.message);
      Logger.log('❌ Stack: ' + syncError.stack);
      // Don't fail the whole operation if sync fails, but log prominently
    }

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

    // Create PDF content as HTML
    const htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; }
          .header { text-align: center; margin-bottom: 30px; }
          .company { font-size: 18px; font-weight: bold; }
          .title { font-size: 24px; font-weight: bold; margin: 10px 0; }
          .info-section { margin: 20px 0; }
          .info-row { display: flex; margin: 5px 0; }
          .label { width: 200px; font-weight: bold; }
          .value { flex: 1; }
          table { width: 100%; border-collapse: collapse; margin: 20px 0; }
          th, td { padding: 10px; text-align: left; border: 1px solid #ddd; }
          th { background-color: #f2f2f2; font-weight: bold; }
          .amount { text-align: right; }
          .total-row { font-weight: bold; background-color: #f9f9f9; }
          .footer { margin-top: 30px; font-size: 12px; color: #666; }
        </style>
      </head>
      <body>
        <div class="header">
          <div class="company">${payslip.EMPLOYER || 'SA Grinding Wheels'}</div>
          <div class="title">PAYSLIP</div>
        </div>

        <div class="info-section">
          <div class="info-row">
            <div class="label">Payslip Number:</div>
            <div class="value">#${payslip.RECORDNUMBER}</div>
          </div>
          <div class="info-row">
            <div class="label">Employee:</div>
            <div class="value">${payslip['EMPLOYEE NAME'] || ''}</div>
          </div>
          <div class="info-row">
            <div class="label">Employment Status:</div>
            <div class="value">${payslip['EMPLOYMENT STATUS'] || ''}</div>
          </div>
          <div class="info-row">
            <div class="label">Week Ending:</div>
            <div class="value">${formatDateForDisplay(payslip.WEEKENDING)}</div>
          </div>
          <div class="info-row">
            <div class="label">Date Generated:</div>
            <div class="value">${formatDateForDisplay(new Date())}</div>
          </div>
        </div>

        <table>
          <thead>
            <tr>
              <th>Description</th>
              <th class="amount">Amount (R)</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>Standard Time (${payslip.HOURS || 0}h ${payslip.MINUTES || 0}m)</td>
              <td class="amount">${formatAmount(payslip.STANDARDTIME)}</td>
            </tr>
            <tr>
              <td>Overtime (${payslip.OVERTIMEHOURS || 0}h ${payslip.OVERTIMEMINUTES || 0}m)</td>
              <td class="amount">${formatAmount(payslip.OVERTIME)}</td>
            </tr>
            ${payslip['LEAVE PAY'] > 0 ? `<tr><td>Leave Pay</td><td class="amount">${formatAmount(payslip['LEAVE PAY'])}</td></tr>` : ''}
            ${payslip['BONUS PAY'] > 0 ? `<tr><td>Bonus Pay</td><td class="amount">${formatAmount(payslip['BONUS PAY'])}</td></tr>` : ''}
            ${payslip.OTHERINCOME > 0 ? `<tr><td>Other Income</td><td class="amount">${formatAmount(payslip.OTHERINCOME)}</td></tr>` : ''}
            <tr class="total-row">
              <td><strong>Gross Salary</strong></td>
              <td class="amount"><strong>${formatAmount(payslip.GROSSSALARY)}</strong></td>
            </tr>
            <tr>
              <td>UIF (1%)</td>
              <td class="amount">(${formatAmount(payslip.UIF)})</td>
            </tr>
            ${payslip['OTHER DEDUCTIONS'] > 0 ? `<tr><td>Other Deductions${payslip['OTHER DEDUCTIONS TEXT'] ? ' - ' + payslip['OTHER DEDUCTIONS TEXT'] : ''}</td><td class="amount">(${formatAmount(payslip['OTHER DEDUCTIONS'])})</td></tr>` : ''}
            <tr class="total-row">
              <td><strong>Total Deductions</strong></td>
              <td class="amount"><strong>(${formatAmount(payslip.TOTALDEDUCTIONS)})</strong></td>
            </tr>
            <tr class="total-row">
              <td><strong>Net Salary</strong></td>
              <td class="amount"><strong>${formatAmount(payslip.NETTSALARY)}</strong></td>
            </tr>
            ${payslip.LoanDeductionThisWeek > 0 ? `<tr><td>Loan Deduction</td><td class="amount">(${formatAmount(payslip.LoanDeductionThisWeek)})</td></tr>` : ''}
            ${payslip.NewLoanThisWeek > 0 && payslip.LoanDisbursementType === 'With Salary' ? `<tr><td>New Loan (With Salary)</td><td class="amount">${formatAmount(payslip.NewLoanThisWeek)}</td></tr>` : ''}
            <tr class="total-row" style="background-color: #e8f5e9;">
              <td><strong>PAID TO ACCOUNT</strong></td>
              <td class="amount"><strong>${formatAmount(payslip.PaidtoAccount)}</strong></td>
            </tr>
          </tbody>
        </table>

        ${payslip.NOTES ? `<div class="info-section"><div class="label">Notes:</div><div class="value">${payslip.NOTES}</div></div>` : ''}

        <div class="footer">
          Generated on ${formatDateForDisplay(new Date())} by HR Payroll System
        </div>
      </body>
      </html>
    `;

    // Create a temporary Google Doc
    const docName = `Payslip_${payslip.RECORDNUMBER}_${payslip['EMPLOYEE NAME']}_${formatDateForFilename(payslip.WEEKENDING)}`;
    const doc = DocumentApp.create(docName);
    const docId = doc.getId();

    // Clear default content and add HTML
    const body = doc.getBody();
    body.clear();
    body.appendParagraph(payslip.EMPLOYER || 'SA Grinding Wheels').setAlignment(DocumentApp.HorizontalAlignment.CENTER).setBold(true);
    body.appendParagraph('PAYSLIP').setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(18).setBold(true);
    body.appendParagraph('');
    body.appendParagraph(`Payslip Number: #${payslip.RECORDNUMBER}`);
    body.appendParagraph(`Employee: ${payslip['EMPLOYEE NAME'] || ''}`);
    body.appendParagraph(`Employment Status: ${payslip['EMPLOYMENT STATUS'] || ''}`);
    body.appendParagraph(`Week Ending: ${formatDateForDisplay(payslip.WEEKENDING)}`);
    body.appendParagraph('');

    // Add earnings table
    const table = body.appendTable();
    const headerRow = table.appendTableRow();
    headerRow.appendTableCell('Description').setBackgroundColor('#f2f2f2');
    headerRow.appendTableCell('Amount (R)').setBackgroundColor('#f2f2f2');

    // Standard time
    const row1 = table.appendTableRow();
    row1.appendTableCell(`Standard Time (${payslip.HOURS || 0}h ${payslip.MINUTES || 0}m)`);
    row1.appendTableCell(formatAmount(payslip.STANDARDTIME));

    // Overtime
    const row2 = table.appendTableRow();
    row2.appendTableCell(`Overtime (${payslip.OVERTIMEHOURS || 0}h ${payslip.OVERTIMEMINUTES || 0}m)`);
    row2.appendTableCell(formatAmount(payslip.OVERTIME));

    // Gross
    const row3 = table.appendTableRow();
    row3.appendTableCell('Gross Salary').setBold(true);
    row3.appendTableCell(formatAmount(payslip.GROSSSALARY)).setBold(true);

    // UIF
    const row4 = table.appendTableRow();
    row4.appendTableCell('UIF (1%)');
    row4.appendTableCell(`(${formatAmount(payslip.UIF)})`);

    // Net
    const row5 = table.appendTableRow();
    row5.appendTableCell('Net Salary').setBold(true);
    row5.appendTableCell(formatAmount(payslip.NETTSALARY)).setBold(true);

    // Paid to Account
    const row6 = table.appendTableRow();
    row6.appendTableCell('PAID TO ACCOUNT').setBold(true);
    row6.appendTableCell(formatAmount(payslip.PaidtoAccount)).setBold(true).setBackgroundColor('#e8f5e9');

    doc.saveAndClose();

    // Convert to PDF
    const docFile = DriveApp.getFileById(docId);
    const pdfBlob = docFile.getAs('application/pdf');
    pdfBlob.setName(docName + '.pdf');

    // Save PDF to Drive (in root or specific folder)
    const pdfFile = DriveApp.createFile(pdfBlob);
    const pdfUrl = pdfFile.getUrl();

    // Delete the temporary doc
    docFile.setTrashed(true);

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
    Logger.log('>>> PAYROLL_GS_VERSION: 2025-11-21-A <<<');
    Logger.log('\n========== UPDATE LOAN PAYMENT ==========');
    Logger.log('Record Number: ' + recordNumber);
    Logger.log('Loan Data: ' + JSON.stringify(loanData));

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
    Logger.log('');
    Logger.log('🔄🔄🔄 SYNC BLOCK START (v2025-11-21-B) 🔄🔄🔄');
    Logger.log('🔄 Payslip RECORDNUMBER: ' + currentPayslip.RECORDNUMBER);
    Logger.log('🔄 Payslip id (Employee ID): ' + currentPayslip.id);
    Logger.log('🔄 Payslip EMPLOYEE NAME: ' + currentPayslip['EMPLOYEE NAME']);
    Logger.log('🔄 LoanDisbursementType: ' + currentPayslip.LoanDisbursementType);
    Logger.log('🔄 NewLoanThisWeek: ' + currentPayslip.NewLoanThisWeek);
    Logger.log('🔄 LoanDeductionThisWeek: ' + currentPayslip.LoanDeductionThisWeek);
    Logger.log('🔄 About to call syncLoanTransactionFromPayslip...');
    try {
      syncLoanTransactionFromPayslip(currentPayslip);
      Logger.log('✅ Loan transaction synced to EmployeeLoans sheet');
    } catch (syncError) {
      Logger.log('❌ ERROR: Failed to sync loan transaction: ' + syncError.message);
      Logger.log('❌ Stack: ' + (syncError.stack || 'No stack'));
      // Don't fail the whole operation if sync fails, but log prominently
    }
    Logger.log('🔄🔄🔄 SYNC BLOCK END 🔄🔄🔄');
    Logger.log('');

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
 * Sync loan transaction from payslip to EmployeeLoans sheet
 * Handles creating/updating/deleting loan transactions based on payslip data
 *
 * @param {Object} payslip - The payslip data
 */
function syncLoanTransactionFromPayslip(payslip) {
  Logger.log('\n========== SYNC LOAN TRANSACTION ==========');
  Logger.log('📋 Payslip #' + payslip.RECORDNUMBER);
  Logger.log('📋 Employee: ' + payslip['EMPLOYEE NAME']);
  Logger.log('📋 Employee ID: ' + payslip.id);
  Logger.log('📋 Payment Type: ' + (payslip.LoanDisbursementType || 'Separate'));
  Logger.log('📋 LoanDeductionThisWeek: ' + payslip.LoanDeductionThisWeek);
  Logger.log('📋 NewLoanThisWeek: ' + payslip.NewLoanThisWeek);

  const sheets = getSheets();

  // Log available sheets for debugging
  Logger.log('🔍 Available sheets: ' + Object.keys(sheets).join(', '));

  const loanSheet = sheets.loans;

  if (!loanSheet) {
    Logger.log('❌ ERROR: EmployeeLoans sheet not found in sheets object');
    Logger.log('❌ sheets.loans is: ' + sheets.loans);
    throw new Error('EmployeeLoans sheet not found. Please ensure a sheet with "loan" in its name exists.');
  }

  Logger.log('✅ Found loans sheet: ' + loanSheet.getName());

  const loanData = loanSheet.getDataRange().getValues();
  const headers = loanData[0];

  // Find column indexes
  const salaryLinkCol = findColumnIndex(headers, 'SalaryLink');
  const loanIdCol = findColumnIndex(headers, 'LoanID');

  if (salaryLinkCol === -1) {
    Logger.log('⚠️ SalaryLink column not found - skipping sync');
    return;
  }

  // Find and delete existing transactions linked to this payslip
  const recordNumber = String(payslip.RECORDNUMBER);
  const rowsToDelete = [];

  for (let i = 1; i < loanData.length; i++) {
    if (String(loanData[i][salaryLinkCol]) === recordNumber) {
      rowsToDelete.push(i + 1); // Sheet row index (1-based)
    }
  }

  // Delete rows from bottom to top to avoid index shifting
  rowsToDelete.sort((a, b) => b - a);
  for (const rowIndex of rowsToDelete) {
    loanSheet.deleteRow(rowIndex);
    Logger.log('🗑️ Deleted existing loan transaction at row ' + rowIndex);
  }

  // Get employee ID
  const employeeId = payslip.id;
  if (!employeeId) {
    Logger.log('⚠️ No employee ID in payslip - skipping loan sync');
    return;
  }

  // Get current loan balance for this employee (after deletions)
  SpreadsheetApp.flush();
  const balanceResult = getCurrentLoanBalance(employeeId);
  const currentBalance = (balanceResult && balanceResult.success) ? balanceResult.data : 0;

  const paymentType = payslip.LoanDisbursementType || 'Separate';
  const transactionDate = payslip.WEEKENDING || new Date();

  // Create loan transaction based on payment type
  // Row order must match: LoanID, Employee Name, Employee ID, Timestamp, TransactionDate,
  //                       LoanAmount, LoanType, DisbursementMode, SalaryLink, Notes,
  //                       BalanceBefore, BalanceAfter

  if (paymentType === 'Repayment' && payslip.LoanDeductionThisWeek > 0) {
    // Repayment - negative amount (reduces balance)
    const amount = -Math.abs(parseFloat(payslip.LoanDeductionThisWeek));
    const balanceAfter = currentBalance + amount;

    const rowData = [
      generateFullUUID(),                           // LoanID
      payslip['EMPLOYEE NAME'],                     // Employee Name
      employeeId,                                   // Employee ID
      new Date(),                                   // Timestamp
      transactionDate,                              // TransactionDate
      amount,                                       // LoanAmount
      'Repayment',                                  // LoanType
      'Repayment',                                  // DisbursementMode
      recordNumber,                                 // SalaryLink
      'Auto-synced from payslip #' + recordNumber,  // Notes
      currentBalance,                               // BalanceBefore
      balanceAfter                                  // BalanceAfter
    ];

    loanSheet.appendRow(rowData);
    Logger.log('✅ Created Repayment transaction: ' + formatCurrency(amount));

  } else if ((paymentType === 'With Salary' || paymentType === 'Separate') && payslip.NewLoanThisWeek > 0) {
    // New loan - positive amount (increases balance)
    const amount = Math.abs(parseFloat(payslip.NewLoanThisWeek));
    const balanceAfter = currentBalance + amount;

    const rowData = [
      generateFullUUID(),                           // LoanID
      payslip['EMPLOYEE NAME'],                     // Employee Name
      employeeId,                                   // Employee ID
      new Date(),                                   // Timestamp
      transactionDate,                              // TransactionDate
      amount,                                       // LoanAmount
      'Disbursement',                               // LoanType
      paymentType,                                  // DisbursementMode
      recordNumber,                                 // SalaryLink
      'Auto-synced from payslip #' + recordNumber,  // Notes
      currentBalance,                               // BalanceBefore
      balanceAfter                                  // BalanceAfter
    ];

    loanSheet.appendRow(rowData);
    Logger.log('✅ Created Disbursement transaction: ' + formatCurrency(amount) + ' (' + paymentType + ')');

  } else {
    Logger.log('ℹ️ No loan transaction to create (amount is 0)');
  }

  SpreadsheetApp.flush();
  Logger.log('========== SYNC LOAN TRANSACTION COMPLETE ==========\n');
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
 * Test Case 5: Loan Sync to EmployeeLoans Table
 * Tests that syncLoanTransactionFromPayslip correctly syncs data to EmployeeLoans sheet
 *
 * Run this test from Apps Script editor: test_LoanSync()
 *
 * IMPORTANT: This test requires a valid employee ID from your EMPLOYEE DETAILS sheet.
 * Replace TEST_EMPLOYEE_ID with a real employee ID before running.
 */
function test_LoanSync() {
  Logger.log('\n========== TEST: LOAN SYNC TO EMPLOYEELOANS TABLE ==========');
  Logger.log('⏱️ Started at: ' + new Date().toISOString());

  // ==============================
  // CONFIGURATION
  // ==============================
  const TEST_EMPLOYEE_ID = 'e85a1e99';
  const TEST_EMPLOYEE_NAME = 'Manfid Bowker';

  const sheets = getSheets();
  const loanSheet = sheets.loans;

  if (!loanSheet) {
    Logger.log('❌ FATAL: EmployeeLoans sheet not found!');
    return;
  }

  let allTestsPassed = true;
  let testResults = [];

  // Helper function to get loan records for a specific SalaryLink
  function getLoanRecordBySalaryLink(salaryLink) {
    const data = loanSheet.getDataRange().getValues();
    const headers = data[0];
    const salaryLinkCol = findColumnIndex(headers, 'SalaryLink');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][salaryLinkCol]) === String(salaryLink)) {
        return buildObjectFromRow(data[i], headers);
      }
    }
    return null;
  }

  // Helper function to count loan records for a SalaryLink
  function countLoanRecordsForSalaryLink(salaryLink) {
    const data = loanSheet.getDataRange().getValues();
    const headers = data[0];
    const salaryLinkCol = findColumnIndex(headers, 'SalaryLink');
    let count = 0;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][salaryLinkCol]) === String(salaryLink)) {
        count++;
      }
    }
    return count;
  }

  // Clean up any test records before starting
  function cleanupTestRecords(recordNumbers) {
    const data = loanSheet.getDataRange().getValues();
    const headers = data[0];
    const salaryLinkCol = findColumnIndex(headers, 'SalaryLink');
    const rowsToDelete = [];

    for (let i = 1; i < data.length; i++) {
      if (recordNumbers.includes(String(data[i][salaryLinkCol]))) {
        rowsToDelete.push(i + 1);
      }
    }

    // Delete from bottom to top
    rowsToDelete.sort((a, b) => b - a);
    for (const rowIndex of rowsToDelete) {
      loanSheet.deleteRow(rowIndex);
    }
    SpreadsheetApp.flush();
  }

  // Test record numbers (use unique numbers that won't conflict with real data)
  const testRecordNumbers = ['TEST-9999-A', 'TEST-9999-B', 'TEST-9999-C', 'TEST-9999-D'];

  // Clean up before tests
  Logger.log('🧹 Cleaning up any existing test records...');
  cleanupTestRecords(testRecordNumbers);

  // ==============================
  // TEST 1: New Loan Disbursement (With Salary)
  // ==============================
  Logger.log('\n--- TEST 1: New Loan Disbursement (With Salary) ---');

  const testPayslip1 = {
    RECORDNUMBER: testRecordNumbers[0],
    'EMPLOYEE NAME': TEST_EMPLOYEE_NAME,
    id: TEST_EMPLOYEE_ID,
    WEEKENDING: new Date(),
    LoanDisbursementType: 'With Salary',
    NewLoanThisWeek: 500,
    LoanDeductionThisWeek: 0
  };

  try {
    syncLoanTransactionFromPayslip(testPayslip1);
    SpreadsheetApp.flush();

    const record1 = getLoanRecordBySalaryLink(testRecordNumbers[0]);

    if (!record1) {
      Logger.log('❌ TEST 1 FAILED: No loan record created');
      allTestsPassed = false;
      testResults.push({ test: 1, passed: false, reason: 'No record created' });
    } else {
      let test1Passed = true;

      // Verify loan amount (positive for disbursement)
      if (record1.LoanAmount !== 500) {
        Logger.log('❌ LoanAmount: Expected 500, Got ' + record1.LoanAmount);
        test1Passed = false;
      } else {
        Logger.log('✅ LoanAmount: 500');
      }

      // Verify loan type
      if (record1.LoanType !== 'Disbursement') {
        Logger.log('❌ LoanType: Expected "Disbursement", Got "' + record1.LoanType + '"');
        test1Passed = false;
      } else {
        Logger.log('✅ LoanType: Disbursement');
      }

      // Verify disbursement mode
      if (record1.DisbursementMode !== 'With Salary') {
        Logger.log('❌ DisbursementMode: Expected "With Salary", Got "' + record1.DisbursementMode + '"');
        test1Passed = false;
      } else {
        Logger.log('✅ DisbursementMode: With Salary');
      }

      // Verify SalaryLink
      if (String(record1.SalaryLink) !== testRecordNumbers[0]) {
        Logger.log('❌ SalaryLink: Expected "' + testRecordNumbers[0] + '", Got "' + record1.SalaryLink + '"');
        test1Passed = false;
      } else {
        Logger.log('✅ SalaryLink: ' + testRecordNumbers[0]);
      }

      // Verify Employee ID
      if (record1['Employee ID'] !== TEST_EMPLOYEE_ID) {
        Logger.log('❌ Employee ID: Expected "' + TEST_EMPLOYEE_ID + '", Got "' + record1['Employee ID'] + '"');
        test1Passed = false;
      } else {
        Logger.log('✅ Employee ID: ' + TEST_EMPLOYEE_ID);
      }

      // Verify balance calculation
      const expectedBalanceAfter = (record1.BalanceBefore || 0) + 500;
      if (Math.abs(record1.BalanceAfter - expectedBalanceAfter) > 0.01) {
        Logger.log('❌ BalanceAfter: Expected ' + expectedBalanceAfter + ', Got ' + record1.BalanceAfter);
        test1Passed = false;
      } else {
        Logger.log('✅ BalanceAfter: ' + record1.BalanceAfter + ' (BalanceBefore: ' + record1.BalanceBefore + ')');
      }

      if (test1Passed) {
        Logger.log('✅ TEST 1 PASSED: New Loan Disbursement (With Salary)');
        testResults.push({ test: 1, passed: true });
      } else {
        Logger.log('❌ TEST 1 FAILED');
        allTestsPassed = false;
        testResults.push({ test: 1, passed: false, reason: 'Field validation failed' });
      }
    }
  } catch (error) {
    Logger.log('❌ TEST 1 ERROR: ' + error.message);
    allTestsPassed = false;
    testResults.push({ test: 1, passed: false, reason: error.message });
  }

  // ==============================
  // TEST 2: New Loan Disbursement (Separate)
  // ==============================
  Logger.log('\n--- TEST 2: New Loan Disbursement (Separate) ---');

  const testPayslip2 = {
    RECORDNUMBER: testRecordNumbers[1],
    'EMPLOYEE NAME': TEST_EMPLOYEE_NAME,
    id: TEST_EMPLOYEE_ID,
    WEEKENDING: new Date(),
    LoanDisbursementType: 'Separate',
    NewLoanThisWeek: 750,
    LoanDeductionThisWeek: 0
  };

  try {
    syncLoanTransactionFromPayslip(testPayslip2);
    SpreadsheetApp.flush();

    const record2 = getLoanRecordBySalaryLink(testRecordNumbers[1]);

    if (!record2) {
      Logger.log('❌ TEST 2 FAILED: No loan record created');
      allTestsPassed = false;
      testResults.push({ test: 2, passed: false, reason: 'No record created' });
    } else {
      let test2Passed = true;

      if (record2.LoanAmount !== 750) {
        Logger.log('❌ LoanAmount: Expected 750, Got ' + record2.LoanAmount);
        test2Passed = false;
      } else {
        Logger.log('✅ LoanAmount: 750');
      }

      if (record2.DisbursementMode !== 'Separate') {
        Logger.log('❌ DisbursementMode: Expected "Separate", Got "' + record2.DisbursementMode + '"');
        test2Passed = false;
      } else {
        Logger.log('✅ DisbursementMode: Separate');
      }

      if (test2Passed) {
        Logger.log('✅ TEST 2 PASSED: New Loan Disbursement (Separate)');
        testResults.push({ test: 2, passed: true });
      } else {
        Logger.log('❌ TEST 2 FAILED');
        allTestsPassed = false;
        testResults.push({ test: 2, passed: false, reason: 'Field validation failed' });
      }
    }
  } catch (error) {
    Logger.log('❌ TEST 2 ERROR: ' + error.message);
    allTestsPassed = false;
    testResults.push({ test: 2, passed: false, reason: error.message });
  }

  // ==============================
  // TEST 3: Loan Repayment
  // ==============================
  Logger.log('\n--- TEST 3: Loan Repayment ---');

  const testPayslip3 = {
    RECORDNUMBER: testRecordNumbers[2],
    'EMPLOYEE NAME': TEST_EMPLOYEE_NAME,
    id: TEST_EMPLOYEE_ID,
    WEEKENDING: new Date(),
    LoanDisbursementType: 'Repayment',
    NewLoanThisWeek: 0,
    LoanDeductionThisWeek: 200
  };

  try {
    syncLoanTransactionFromPayslip(testPayslip3);
    SpreadsheetApp.flush();

    const record3 = getLoanRecordBySalaryLink(testRecordNumbers[2]);

    if (!record3) {
      Logger.log('❌ TEST 3 FAILED: No loan record created');
      allTestsPassed = false;
      testResults.push({ test: 3, passed: false, reason: 'No record created' });
    } else {
      let test3Passed = true;

      // Verify loan amount (negative for repayment)
      if (record3.LoanAmount !== -200) {
        Logger.log('❌ LoanAmount: Expected -200, Got ' + record3.LoanAmount);
        test3Passed = false;
      } else {
        Logger.log('✅ LoanAmount: -200 (negative for repayment)');
      }

      // Verify loan type
      if (record3.LoanType !== 'Repayment') {
        Logger.log('❌ LoanType: Expected "Repayment", Got "' + record3.LoanType + '"');
        test3Passed = false;
      } else {
        Logger.log('✅ LoanType: Repayment');
      }

      // Verify balance calculation (should decrease)
      const expectedBalanceAfter3 = (record3.BalanceBefore || 0) - 200;
      if (Math.abs(record3.BalanceAfter - expectedBalanceAfter3) > 0.01) {
        Logger.log('❌ BalanceAfter: Expected ' + expectedBalanceAfter3 + ', Got ' + record3.BalanceAfter);
        test3Passed = false;
      } else {
        Logger.log('✅ BalanceAfter: ' + record3.BalanceAfter + ' (decreased by 200)');
      }

      if (test3Passed) {
        Logger.log('✅ TEST 3 PASSED: Loan Repayment');
        testResults.push({ test: 3, passed: true });
      } else {
        Logger.log('❌ TEST 3 FAILED');
        allTestsPassed = false;
        testResults.push({ test: 3, passed: false, reason: 'Field validation failed' });
      }
    }
  } catch (error) {
    Logger.log('❌ TEST 3 ERROR: ' + error.message);
    allTestsPassed = false;
    testResults.push({ test: 3, passed: false, reason: error.message });
  }

  // ==============================
  // TEST 4: Update Existing Loan (Delete & Recreate)
  // ==============================
  Logger.log('\n--- TEST 4: Update Existing Loan (Delete & Recreate) ---');

  // Re-sync testPayslip1 with different amount - should delete old and create new
  const testPayslip4 = {
    RECORDNUMBER: testRecordNumbers[0], // Same as test 1
    'EMPLOYEE NAME': TEST_EMPLOYEE_NAME,
    id: TEST_EMPLOYEE_ID,
    WEEKENDING: new Date(),
    LoanDisbursementType: 'With Salary',
    NewLoanThisWeek: 800, // Changed from 500 to 800
    LoanDeductionThisWeek: 0
  };

  try {
    syncLoanTransactionFromPayslip(testPayslip4);
    SpreadsheetApp.flush();

    // Verify only one record exists for this SalaryLink
    const recordCount = countLoanRecordsForSalaryLink(testRecordNumbers[0]);
    if (recordCount !== 1) {
      Logger.log('❌ Record count: Expected 1, Got ' + recordCount + ' (duplicate records!)');
      allTestsPassed = false;
      testResults.push({ test: 4, passed: false, reason: 'Duplicate records found' });
    } else {
      const record4 = getLoanRecordBySalaryLink(testRecordNumbers[0]);

      let test4Passed = true;

      // Verify updated amount
      if (record4.LoanAmount !== 800) {
        Logger.log('❌ LoanAmount: Expected 800, Got ' + record4.LoanAmount);
        test4Passed = false;
      } else {
        Logger.log('✅ LoanAmount updated: 800 (was 500)');
      }

      Logger.log('✅ Only 1 record exists (old record deleted)');

      if (test4Passed) {
        Logger.log('✅ TEST 4 PASSED: Update Existing Loan');
        testResults.push({ test: 4, passed: true });
      } else {
        Logger.log('❌ TEST 4 FAILED');
        allTestsPassed = false;
        testResults.push({ test: 4, passed: false, reason: 'Field validation failed' });
      }
    }
  } catch (error) {
    Logger.log('❌ TEST 4 ERROR: ' + error.message);
    allTestsPassed = false;
    testResults.push({ test: 4, passed: false, reason: error.message });
  }

  // ==============================
  // TEST 5: No Loan (Should Not Create Record)
  // ==============================
  Logger.log('\n--- TEST 5: No Loan (Should Not Create Record) ---');

  const testPayslip5 = {
    RECORDNUMBER: testRecordNumbers[3],
    'EMPLOYEE NAME': TEST_EMPLOYEE_NAME,
    id: TEST_EMPLOYEE_ID,
    WEEKENDING: new Date(),
    LoanDisbursementType: 'Separate',
    NewLoanThisWeek: 0,
    LoanDeductionThisWeek: 0
  };

  try {
    syncLoanTransactionFromPayslip(testPayslip5);
    SpreadsheetApp.flush();

    const record5 = getLoanRecordBySalaryLink(testRecordNumbers[3]);

    if (record5) {
      Logger.log('❌ TEST 5 FAILED: Record was created but should not have been');
      Logger.log('   LoanAmount: ' + record5.LoanAmount);
      allTestsPassed = false;
      testResults.push({ test: 5, passed: false, reason: 'Unexpected record created' });
    } else {
      Logger.log('✅ TEST 5 PASSED: No record created (as expected)');
      testResults.push({ test: 5, passed: true });
    }
  } catch (error) {
    Logger.log('❌ TEST 5 ERROR: ' + error.message);
    allTestsPassed = false;
    testResults.push({ test: 5, passed: false, reason: error.message });
  }

  // ==============================
  // CLEANUP
  // ==============================
  Logger.log('\n🧹 Cleaning up test records...');
  cleanupTestRecords(testRecordNumbers);
  Logger.log('✅ Test records cleaned up');

  // ==============================
  // SUMMARY
  // ==============================
  Logger.log('\n========== TEST SUMMARY ==========');
  const passed = testResults.filter(r => r.passed).length;
  const total = testResults.length;

  Logger.log('Tests Passed: ' + passed + '/' + total);

  testResults.forEach(function(result) {
    const status = result.passed ? '✅' : '❌';
    const reason = result.reason ? ' - ' + result.reason : '';
    Logger.log('  Test ' + result.test + ': ' + status + reason);
  });

  if (allTestsPassed) {
    Logger.log('\n🎉 ALL TESTS PASSED!');
  } else {
    Logger.log('\n⚠️ SOME TESTS FAILED - Review errors above');
  }

  Logger.log('⏱️ Completed at: ' + new Date().toISOString());
  Logger.log('========== TEST COMPLETE ==========\n');

  return {
    success: allTestsPassed,
    passed: passed,
    total: total,
    results: testResults
  };
}

/**
 * Quick test helper - run this to test with mock data
 * This version doesn't require a real employee ID
 */
function test_LoanSync_Verification() {
  Logger.log('\n========== LOAN SYNC VERIFICATION TEST ==========');
  Logger.log('This test verifies the syncLoanTransactionFromPayslip function exists and is callable.');
  Logger.log('For full integration testing, run test_LoanSync() with a valid employee ID.');

  // Verify function exists
  if (typeof syncLoanTransactionFromPayslip === 'function') {
    Logger.log('✅ syncLoanTransactionFromPayslip function exists');
  } else {
    Logger.log('❌ syncLoanTransactionFromPayslip function not found!');
    return { success: false };
  }

  // Verify helper functions exist
  const helpers = ['getSheets', 'getCurrentLoanBalance', 'findColumnIndex', 'buildObjectFromRow', 'generateFullUUID'];
  let allHelpersExist = true;

  helpers.forEach(function(helper) {
    if (typeof eval(helper) === 'function') {
      Logger.log('✅ ' + helper + ' function exists');
    } else {
      Logger.log('❌ ' + helper + ' function not found!');
      allHelpersExist = false;
    }
  });

  // Verify EmployeeLoans sheet exists
  const sheets = getSheets();
  if (sheets.loans) {
    Logger.log('✅ EmployeeLoans sheet found');

    // Verify column structure
    const headers = sheets.loans.getRange(1, 1, 1, 12).getValues()[0];
    Logger.log('📋 EmployeeLoans columns: ' + headers.join(', '));
  } else {
    Logger.log('❌ EmployeeLoans sheet not found!');
    return { success: false };
  }

  Logger.log('\n✅ Verification complete - all components are in place');
  Logger.log('📝 To run full tests, update TEST_EMPLOYEE_ID in test_LoanSync() and run it.');
  Logger.log('========== VERIFICATION COMPLETE ==========\n');

  return { success: allHelpersExist };
}
