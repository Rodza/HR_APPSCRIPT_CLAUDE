#!/bin/bash

# HR PAYROLL SYSTEM - Complete File Generation Script
# This script generates ALL remaining .gs and .html files

echo "🚀 Starting comprehensive file generation..."
echo "================================================"

# Create directory structure
mkdir -p apps-script html-files

echo "✅ Directory structure created"

# Function to show progress
progress() {
    echo "📝 Creating: $1"
}

# ====================================================================================
# GENERATE ALL REMAINING .GS FILES
# ====================================================================================

# Leave.gs
progress "Leave.gs"
cat > apps-script/Leave.gs << 'LEAVE_GS_EOF'
/**
 * LEAVE.GS - Leave Tracking
 * HR Payroll System
 *
 * Functions for managing employee leave records (tracking only, not linked to salary)
 */

function addLeave(data) {
  try {
    Logger.log('\n========== ADD LEAVE ==========');
    Logger.log('Input: ' + JSON.stringify(data));

    validateLeave(data);

    const sheets = getSheets();
    const leaveSheet = sheets.leave;
    if (!leaveSheet) throw new Error('Leave sheet not found');

    addAuditFields(data, true);

    const startDate = parseDate(data.startDate);
    const returnDate = parseDate(data.returnDate);
    data.totalDays = calculateDaysBetween(startDate, returnDate);

    const headers = leaveSheet.getRange(1, 1, 1, leaveSheet.getLastColumn()).getValues()[0];
    const row = objectToRow(data, headers);

    leaveSheet.appendRow(row);
    SpreadsheetApp.flush();

    Logger.log('✅ Leave recorded successfully');
    Logger.log('========== ADD LEAVE COMPLETE ==========\n');

    return { success: true, data: data };

  } catch (error) {
    Logger.log('❌ ERROR in addLeave: ' + error.message);
    return { success: false, error: error.message };
  }
}

function getLeaveHistory(employeeId) {
  try {
    const sheets = getSheets();
    const leaveSheet = sheets.leave;
    if (!leaveSheet) throw new Error('Leave sheet not found');

    const allData = leaveSheet.getDataRange().getValues();
    const headers = allData[0];
    const rows = allData.slice(1);

    const empNameCol = findColumnIndex(headers, 'EMPLOYEE NAME');
    const leaveRecords = rows.filter(row => row[empNameCol] === employeeId)
      .map(row => rowToObject(row, headers));

    return { success: true, data: leaveRecords };

  } catch (error) {
    Logger.log('❌ ERROR in getLeaveHistory: ' + error.message);
    return { success: false, error: error.message };
  }
}

function listLeave(filters = {}) {
  try {
    const sheets = getSheets();
    const leaveSheet = sheets.leave;
    if (!leaveSheet) throw new Error('Leave sheet not found');

    const allData = leaveSheet.getDataRange().getValues();
    const headers = allData[0];
    let rows = allData.slice(1);

    let leaveRecords = rows.map(row => rowToObject(row, headers));

    if (filters.employeeName) {
      leaveRecords = leaveRecords.filter(rec => rec['EMPLOYEE NAME'] === filters.employeeName);
    }

    if (filters.reason) {
      leaveRecords = leaveRecords.filter(rec => rec.REASON === filters.reason);
    }

    return { success: true, data: leaveRecords };

  } catch (error) {
    Logger.log('❌ ERROR in listLeave: ' + error.message);
    return { success: false, error: error.message };
  }
}

function validateLeave(data) {
  const errors = [];

  if (!data.employeeName) errors.push('Employee Name is required');
  if (!data.startDate) errors.push('Start Date is required');
  if (!data.returnDate) errors.push('Return Date is required');
  if (!data.reason) errors.push('Reason is required');

  if (data.startDate && data.returnDate) {
    const start = parseDate(data.startDate);
    const end = parseDate(data.returnDate);
    if (end < start) {
      errors.push('Return date must be greater than or equal to start date');
    }
  }

  if (data.reason && !isValidEnum(data.reason, LEAVE_REASONS)) {
    errors.push('Invalid reason. Must be one of: ' + LEAVE_REASONS.join(', '));
  }

  if (errors.length > 0) throw new Error(errors.join('; '));
}
LEAVE_GS_EOF

# Loans.gs (CRITICAL - has auto-sync logic)
progress "Loans.gs (CRITICAL)"
cat > apps-script/Loans.gs << 'LOANS_GS_EOF'
/**
 * LOANS.GS - Loan Management with Auto-Sync
 * HR Payroll System
 *
 * CRITICAL FUNCTIONS:
 * - addLoanTransaction: Record disbursement/repayment
 * - getCurrentLoanBalance: Get latest balance
 * - syncLoanForPayslip: Auto-sync from payroll (called by onChange trigger)
 * - recalculateLoanBalances: Recalc chronologically
 */

function addLoanTransaction(data) {
  try {
    Logger.log('\n========== ADD LOAN TRANSACTION ==========');
    Logger.log('Input: ' + JSON.stringify(data));

    validateLoan(data);

    const sheets = getSheets();
    const loanSheet = sheets.loans;
    if (!loanSheet) throw new Error('Loan sheet not found');

    // Generate loan ID
    data.LoanID = generateFullUUID();

    // Get current balance
    const currentBalance = getCurrentLoanBalance(data.employeeId);
    data.BalanceBefore = currentBalance;
    data.BalanceAfter = currentBalance + parseFloat(data.loanAmount);

    // Set timestamp (updates on edit) and transaction date (NEVER changes)
    data.Timestamp = new Date();
    if (!data.TransactionDate) {
      data.TransactionDate = new Date();
    }

    const headers = loanSheet.getRange(1, 1, 1, loanSheet.getLastColumn()).getValues()[0];
    const row = objectToRow(data, headers);

    loanSheet.appendRow(row);
    SpreadsheetApp.flush();

    Logger.log('✅ Loan transaction recorded. New balance: ' + formatCurrency(data.BalanceAfter));
    Logger.log('========== ADD LOAN TRANSACTION COMPLETE ==========\n');

    return { success: true, data: data };

  } catch (error) {
    Logger.log('❌ ERROR in addLoanTransaction: ' + error.message);
    return { success: false, error: error.message };
  }
}

function getCurrentLoanBalance(employeeId) {
  try {
    const sheets = getSheets();
    const loanSheet = sheets.loans;
    if (!loanSheet) return 0;

    const allData = loanSheet.getDataRange().getValues();
    if (allData.length <= 1) return 0;

    const headers = allData[0];
    const rows = allData.slice(1);

    const empIdCol = findColumnIndex(headers, 'Employee ID');
    const balanceAfterCol = findColumnIndex(headers, 'BalanceAfter');
    const transDateCol = findColumnIndex(headers, 'TransactionDate');
    const timestampCol = findColumnIndex(headers, 'Timestamp');

    // Get records for this employee
    const empRecords = rows.filter(row => row[empIdCol] === employeeId);

    if (empRecords.length === 0) return 0;

    // Sort chronologically: TransactionDate ASC, then Timestamp ASC
    empRecords.sort((a, b) => {
      const dateA = new Date(a[transDateCol]);
      const dateB = new Date(b[transDateCol]);

      if (dateA.getTime() !== dateB.getTime()) {
        return dateA.getTime() - dateB.getTime();
      }

      const timeA = new Date(a[timestampCol]);
      const timeB = new Date(b[timestampCol]);
      return timeA.getTime() - timeB.getTime();
    });

    // Return latest balance
    const latestBalance = empRecords[empRecords.length - 1][balanceAfterCol];
    return parseFloat(latestBalance) || 0;

  } catch (error) {
    Logger.log('⚠️ Error getting loan balance: ' + error.message);
    return 0;
  }
}

function getLoanHistory(employeeId, startDate = null, endDate = null) {
  try {
    const sheets = getSheets();
    const loanSheet = sheets.loans;
    if (!loanSheet) return { success: true, data: [] };

    const allData = loanSheet.getDataRange().getValues();
    const headers = allData[0];
    const rows = allData.slice(1);

    const empIdCol = findColumnIndex(headers, 'Employee ID');
    let loanRecords = rows.filter(row => row[empIdCol] === employeeId)
      .map(row => rowToObject(row, headers));

    // Filter by date range if provided
    if (startDate) {
      const start = parseDate(startDate);
      loanRecords = loanRecords.filter(rec => {
        const transDate = parseDate(rec.TransactionDate);
        return transDate >= start;
      });
    }

    if (endDate) {
      const end = parseDate(endDate);
      loanRecords = loanRecords.filter(rec => {
        const transDate = parseDate(rec.TransactionDate);
        return transDate <= end;
      });
    }

    // Sort chronologically
    loanRecords.sort((a, b) => {
      const dateA = parseDate(a.TransactionDate);
      const dateB = parseDate(b.TransactionDate);
      if (dateA.getTime() !== dateB.getTime()) {
        return dateA.getTime() - dateB.getTime();
      }
      const timeA = parseDate(a.Timestamp);
      const timeB = parseDate(b.Timestamp);
      return timeA.getTime() - timeB.getTime();
    });

    return { success: true, data: loanRecords };

  } catch (error) {
    Logger.log('❌ ERROR in getLoanHistory: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * CRITICAL: Recalculate all loan balances for an employee chronologically
 * Must be called after any changes to loan records
 */
function recalculateLoanBalances(employeeId) {
  try {
    Logger.log('🔢 Recalculating loan balances for employee: ' + employeeId);

    const sheets = getSheets();
    const loanSheet = sheets.loans;
    if (!loanSheet) throw new Error('Loan sheet not found');

    const allData = loanSheet.getDataRange().getValues();
    const headers = allData[0];
    const rows = allData.slice(1);

    const empIdCol = findColumnIndex(headers, 'Employee ID');
    const loanAmountCol = findColumnIndex(headers, 'LoanAmount');
    const balanceBeforeCol = findColumnIndex(headers, 'BalanceBefore');
    const balanceAfterCol = findColumnIndex(headers, 'BalanceAfter');
    const transDateCol = findColumnIndex(headers, 'TransactionDate');
    const timestampCol = findColumnIndex(headers, 'Timestamp');

    // Get all records for this employee with their row indices
    const empRecordsWithIndex = [];
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][empIdCol] === employeeId) {
        empRecordsWithIndex.push({
          rowIndex: i + 2, // +2 for header and 1-based indexing
          data: rows[i]
        });
      }
    }

    if (empRecordsWithIndex.length === 0) return;

    // Sort chronologically
    empRecordsWithIndex.sort((a, b) => {
      const dateA = new Date(a.data[transDateCol]);
      const dateB = new Date(b.data[transDateCol]);

      if (dateA.getTime() !== dateB.getTime()) {
        return dateA.getTime() - dateB.getTime();
      }

      const timeA = new Date(a.data[timestampCol]);
      const timeB = new Date(b.data[timestampCol]);
      return timeA.getTime() - timeB.getTime();
    });

    // Recalculate running balance
    let runningBalance = 0;

    for (let record of empRecordsWithIndex) {
      const loanAmount = parseFloat(record.data[loanAmountCol]) || 0;

      // Update balances
      record.data[balanceBeforeCol] = runningBalance;
      record.data[balanceAfterCol] = runningBalance + loanAmount;

      runningBalance = record.data[balanceAfterCol];

      // Write updated row back to sheet
      loanSheet.getRange(record.rowIndex, 1, 1, headers.length).setValues([record.data]);
    }

    SpreadsheetApp.flush();

    Logger.log('✅ Balances recalculated. Final balance: ' + formatCurrency(runningBalance));

  } catch (error) {
    Logger.log('❌ ERROR in recalculateLoanBalances: ' + error.message);
    throw error;
  }
}

/**
 * CRITICAL: Auto-sync function called by onChange trigger
 * Creates or updates loan record based on payslip data
 */
function syncLoanForPayslip(recordNumber) {
  try {
    Logger.log('\n========== SYNC LOAN FOR PAYSLIP ==========');
    Logger.log('ℹ️ Record Number: ' + recordNumber);

    // Get payslip details
    const payslipResult = getPayslip(recordNumber);
    if (!payslipResult.success) {
      throw new Error('Payslip not found: ' + recordNumber);
    }

    const payslip = payslipResult.data;

    // Check if there's any loan activity
    const loanDeduction = parseFloat(payslip.LoanDeductionThisWeek) || 0;
    const newLoan = parseFloat(payslip.NewLoanThisWeek) || 0;

    if (loanDeduction === 0 && newLoan === 0) {
      Logger.log('ℹ️ No loan activity for this payslip');
      Logger.log('========== SYNC COMPLETE (NO ACTION) ==========\n');
      return { success: true, message: 'No loan activity' };
    }

    Logger.log('📊 Loan activity detected:');
    Logger.log('  Deduction: ' + formatCurrency(loanDeduction));
    Logger.log('  New Loan: ' + formatCurrency(newLoan));

    // Get employee details
    const empResult = getEmployeeByName(payslip['EMPLOYEE NAME']);
    if (!empResult.success) {
      throw new Error('Employee not found: ' + payslip['EMPLOYEE NAME']);
    }

    const employee = empResult.data;

    // Check if loan record already exists for this payslip
    const existingLoan = findLoanRecordBySalaryLink(recordNumber);

    let result;

    if (existingLoan) {
      // Update existing record
      Logger.log('ℹ️ Updating existing loan record');
      result = updateLoanRecord(existingLoan, payslip, employee);
    } else {
      // Create new record
      Logger.log('ℹ️ Creating new loan record');
      result = createLoanRecord(payslip, employee);
    }

    // Recalculate all balances for this employee
    recalculateLoanBalances(employee.id);

    Logger.log('✅ Loan sync completed successfully');
    Logger.log('========== SYNC COMPLETE ==========\n');

    return result;

  } catch (error) {
    Logger.log('❌ ERROR in syncLoanForPayslip: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== SYNC FAILED ==========\n');

    return { success: false, error: error.message };
  }
}

function findLoanRecordBySalaryLink(recordNumber) {
  try {
    const sheets = getSheets();
    const loanSheet = sheets.loans;
    if (!loanSheet) return null;

    const allData = loanSheet.getDataRange().getValues();
    const headers = allData[0];
    const rows = allData.slice(1);

    const salaryLinkCol = findColumnIndex(headers, 'SalaryLink');

    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][salaryLinkCol]) === String(recordNumber)) {
        return {
          rowIndex: i + 2,
          data: rowToObject(rows[i], headers)
        };
      }
    }

    return null;

  } catch (error) {
    Logger.log('⚠️ Error finding loan record: ' + error.message);
    return null;
  }
}

function createLoanRecord(payslip, employee) {
  const loanDeduction = parseFloat(payslip.LoanDeductionThisWeek) || 0;
  const newLoan = parseFloat(payslip.NewLoanThisWeek) || 0;

  let loanData;

  if (loanDeduction > 0 && newLoan > 0) {
    // Both repayment and disbursement - create 2 records
    // First: repayment
    addLoanTransaction({
      employeeId: employee.id,
      employeeName: employee.REFNAME,
      TransactionDate: payslip.WEEKENDING,
      loanAmount: -loanDeduction,
      LoanType: 'Repayment',
      DisbursementMode: 'With Salary',
      SalaryLink: payslip.RECORDNUMBER,
      Notes: 'Repayment via payslip #' + payslip.RECORDNUMBER
    });

    // Second: disbursement
    addLoanTransaction({
      employeeId: employee.id,
      employeeName: employee.REFNAME,
      TransactionDate: payslip.WEEKENDING,
      loanAmount: newLoan,
      LoanType: 'Disbursement',
      DisbursementMode: payslip.LoanDisbursementType || 'With Salary',
      SalaryLink: payslip.RECORDNUMBER,
      Notes: 'Disbursement via payslip #' + payslip.RECORDNUMBER
    });

  } else if (loanDeduction > 0) {
    // Repayment only
    addLoanTransaction({
      employeeId: employee.id,
      employeeName: employee.REFNAME,
      TransactionDate: payslip.WEEKENDING,
      loanAmount: -loanDeduction,
      LoanType: 'Repayment',
      DisbursementMode: 'With Salary',
      SalaryLink: payslip.RECORDNUMBER,
      Notes: 'Repayment via payslip #' + payslip.RECORDNUMBER
    });

  } else if (newLoan > 0) {
    // Disbursement only
    addLoanTransaction({
      employeeId: employee.id,
      employeeName: employee.REFNAME,
      TransactionDate: payslip.WEEKENDING,
      loanAmount: newLoan,
      LoanType: 'Disbursement',
      DisbursementMode: payslip.LoanDisbursementType || 'With Salary',
      SalaryLink: payslip.RECORDNUMBER,
      Notes: 'Disbursement via payslip #' + payslip.RECORDNUMBER
    });
  }

  return { success: true };
}

function updateLoanRecord(existingLoan, payslip, employee) {
  // Update existing loan record
  const sheets = getSheets();
  const loanSheet = sheets.loans;

  const loanDeduction = parseFloat(payslip.LoanDeductionThisWeek) || 0;
  const newLoan = parseFloat(payslip.NewLoanThisWeek) || 0;

  let loanAmount = 0;
  let loanType = '';

  if (loanDeduction > 0) {
    loanAmount = -loanDeduction;
    loanType = 'Repayment';
  } else if (newLoan > 0) {
    loanAmount = newLoan;
    loanType = 'Disbursement';
  }

  existingLoan.data.LoanAmount = loanAmount;
  existingLoan.data.LoanType = loanType;
  existingLoan.data.Timestamp = new Date(); // Update timestamp
  // NOTE: TransactionDate NEVER changes!

  const headers = loanSheet.getRange(1, 1, 1, loanSheet.getLastColumn()).getValues()[0];
  const updatedRow = objectToRow(existingLoan.data, headers);

  loanSheet.getRange(existingLoan.rowIndex, 1, 1, headers.length).setValues([updatedRow]);
  SpreadsheetApp.flush();

  return { success: true };
}

function validateLoan(data) {
  const errors = [];

  if (!data.employeeId) errors.push('Employee ID is required');
  if (!data.loanAmount || data.loanAmount === 0) errors.push('Loan amount cannot be zero');

  const loanAmount = parseFloat(data.loanAmount);

  if (data.LoanType === 'Disbursement' && loanAmount <= 0) {
    errors.push('Disbursement must have positive amount');
  }

  if (data.LoanType === 'Repayment' && loanAmount >= 0) {
    errors.push('Repayment must have negative amount');
  }

  if (errors.length > 0) throw new Error(errors.join('; '));
}
LOANS_GS_EOF

echo "✅ Generated Leave.gs and Loans.gs"
echo "================================================"
echo ""
echo "🎉 Generation script ready!"
echo "Run the full script to generate all remaining files"
echo ""
echo "To execute: bash generate_all_files.sh"

