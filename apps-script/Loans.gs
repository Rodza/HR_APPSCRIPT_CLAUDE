/**
 * LOANS.GS - Loan Management Module
 * HR Payroll System for SA Grinding Wheels & Scorpio Abrasives
 *
 * Handles employee loan tracking with automatic synchronization from payroll.
 * CRITICAL: Auto-sync triggered by onChange when payslips have loan activity.
 *
 * Sheet: EmployeeLoans
 * Columns: LoanID, Employee ID, Employee Name, Timestamp, TransactionDate,
 *          LoanAmount, LoanType, DisbursementMode, SalaryLink, BalanceBefore,
 *          BalanceAfter, Notes
 *
 * IMPORTANT RULES:
 * - Positive LoanAmount = Disbursement (new loan)
 * - Negative LoanAmount = Repayment
 * - TransactionDate NEVER changes after creation (critical for chronological sorting)
 * - Timestamp updates on edit (for deduplication)
 * - Balance calculation ALWAYS chronological by TransactionDate, then Timestamp
 */

// ==================== ADD LOAN TRANSACTION ====================

/**
 * Records a new loan transaction (disbursement or repayment)
 *
 * @param {Object} data - Loan transaction data
 * @param {string} data.employeeId - Employee ID (required)
 * @param {Date|string} data.transactionDate - Date of transaction (required)
 * @param {number} data.loanAmount - Amount (positive=disbursement, negative=repayment) (required)
 * @param {string} data.loanType - "Disbursement" or "Repayment" (required)
 * @param {string} data.disbursementMode - "With Salary", "Separate", or "Manual Entry" (required)
 * @param {string} [data.salaryLink] - RECORDNUMBER from MASTERSALARY (optional)
 * @param {string} [data.notes] - Transaction notes (optional)
 *
 * @returns {Object} Result with success flag and loan record data
 *
 * @example
 * const result = addLoanTransaction({
 *   employeeId: 'e0b6115a',
 *   transactionDate: '2025-10-17',
 *   loanAmount: 500.00,
 *   loanType: 'Disbursement',
 *   disbursementMode: 'With Salary',
 *   salaryLink: '7916',
 *   notes: 'Emergency loan'
 * });
 */
function addLoanTransaction(data) {
  try {
    Logger.log('\n========== ADD LOAN TRANSACTION ==========');
    Logger.log('Input: ' + JSON.stringify(data));

    // Validate input
    validateLoan(data);

    // Get sheets
    const sheets = getSheets();
    const loanSheet = sheets.loans;

    if (!loanSheet) {
      throw new Error('EmployeeLoans sheet not found');
    }

    // Get employee name
    const empResult = getEmployeeById(data.employeeId);
    if (!empResult.success) {
      throw new Error('Employee not found: ' + data.employeeId);
    }
    const employeeName = empResult.data.REFNAME;

    // Calculate balance before and after
    const currentBalance = getCurrentLoanBalance(data.employeeId);
    const balanceBefore = currentBalance.success ? currentBalance.data : 0;
    const balanceAfter = balanceBefore + data.loanAmount;

    Logger.log('💰 Balance before: R' + balanceBefore.toFixed(2));
    Logger.log('💵 Loan amount: R' + data.loanAmount.toFixed(2));
    Logger.log('💰 Balance after: R' + balanceAfter.toFixed(2));

    // Generate loan ID
    const loanId = generateFullUUID();
    const timestamp = new Date();
    const transactionDate = parseDate(data.transactionDate);

    // Prepare row data
    const rowData = [
      loanId,                             // LoanID
      data.employeeId,                    // Employee ID
      employeeName,                       // Employee Name
      timestamp,                          // Timestamp
      transactionDate,                    // TransactionDate
      data.loanAmount,                    // LoanAmount
      data.loanType,                      // LoanType
      data.disbursementMode,              // DisbursementMode
      data.salaryLink || '',              // SalaryLink
      balanceBefore,                      // BalanceBefore
      balanceAfter,                       // BalanceAfter
      data.notes || ''                    // Notes
    ];

    // Append to sheet
    loanSheet.appendRow(rowData);

    const result = {
      loanId: loanId,
      employeeId: data.employeeId,
      employeeName: employeeName,
      timestamp: timestamp.toISOString(),
      transactionDate: formatDate(transactionDate),
      loanAmount: data.loanAmount,
      loanType: data.loanType,
      disbursementMode: data.disbursementMode,
      salaryLink: data.salaryLink || '',
      balanceBefore: balanceBefore,
      balanceAfter: balanceAfter,
      notes: data.notes || ''
    };

    Logger.log('✅ Loan transaction added successfully');
    Logger.log('========== ADD LOAN TRANSACTION COMPLETE ==========\n');

    return { success: true, data: result };

  } catch (error) {
    Logger.log('❌ ERROR in addLoanTransaction: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== ADD LOAN TRANSACTION FAILED ==========\n');
    return { success: false, error: error.message };
  }
}

// ==================== GET CURRENT LOAN BALANCE ====================

/**
 * Gets the current loan balance for an employee
 * Returns the BalanceAfter from the most recent transaction
 *
 * @param {string} employeeId - Employee ID
 * @returns {Object} Result with success flag and current balance
 */
function getCurrentLoanBalance(employeeId) {
  try {
    Logger.log('\n========== GET CURRENT LOAN BALANCE ==========');
    Logger.log('Employee ID: ' + employeeId);

    const sheets = getSheets();
    const loanSheet = sheets.loans;

    if (!loanSheet) {
      throw new Error('EmployeeLoans sheet not found');
    }

    const data = loanSheet.getDataRange().getValues();
    const headers = data[0];

    // Find column indexes
    const empIdCol = findColumnIndex(headers, 'Employee ID');
    const transDateCol = findColumnIndex(headers, 'TransactionDate');
    const timestampCol = findColumnIndex(headers, 'Timestamp');
    const balanceAfterCol = findColumnIndex(headers, 'BalanceAfter');

    if (empIdCol === -1 || transDateCol === -1 || balanceAfterCol === -1) {
      throw new Error('Required columns not found in EmployeeLoans sheet');
    }

    // Filter records for this employee
    const employeeRecords = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[empIdCol] === employeeId) {
        try {
          // Skip if critical data is missing
          if (!row[transDateCol] || !row[timestampCol]) {
            Logger.log('⚠️ Skipping row ' + (i + 1) + ' - missing date fields');
            continue;
          }

          const balanceAfter = row[balanceAfterCol];
          if (balanceAfter === null || balanceAfter === undefined || balanceAfter === '') {
            Logger.log('⚠️ Skipping row ' + (i + 1) + ' - missing balance after');
            continue;
          }

          employeeRecords.push({
            transactionDate: parseDate(row[transDateCol]),
            timestamp: parseDate(row[timestampCol]),
            balanceAfter: parseFloat(balanceAfter)
          });
        } catch (dateError) {
          Logger.log('⚠️ Skipping row ' + (i + 1) + ' - invalid date: ' + dateError.message);
          continue;
        }
      }
    }

    if (employeeRecords.length === 0) {
      Logger.log('ℹ️ No loan records found for employee');
      Logger.log('========== GET CURRENT LOAN BALANCE COMPLETE ==========\n');
      return { success: true, data: 0 };
    }

    // Sort chronologically: TransactionDate first, then Timestamp
    employeeRecords.sort((a, b) => {
      if (a.transactionDate.getTime() !== b.transactionDate.getTime()) {
        return b.transactionDate - a.transactionDate;  // Descending
      }
      return b.timestamp - a.timestamp;  // Descending
    });

    const currentBalance = employeeRecords[0].balanceAfter;

    Logger.log('✅ Current balance: R' + currentBalance.toFixed(2));
    Logger.log('========== GET CURRENT LOAN BALANCE COMPLETE ==========\n');

    return { success: true, data: currentBalance };

  } catch (error) {
    Logger.log('❌ ERROR in getCurrentLoanBalance: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== GET CURRENT LOAN BALANCE FAILED ==========\n');
    return { success: false, error: error.message };
  }
}

// ==================== GET LOAN HISTORY ====================

/**
 * Gets loan transaction history for an employee with optional date filtering
 *
 * @param {string} employeeId - Employee ID
 * @param {Date|string} [startDate] - Filter start date (optional)
 * @param {Date|string} [endDate] - Filter end date (optional)
 *
 * @returns {Object} Result with success flag and transaction records
 */
function getLoanHistory(employeeId, startDate, endDate) {
  try {
    Logger.log('\n========== GET LOAN HISTORY ==========');
    Logger.log('Looking for Employee ID: ' + employeeId);
    Logger.log('Employee ID type: ' + typeof employeeId);
    if (startDate) Logger.log('Start Date: ' + startDate);
    if (endDate) Logger.log('End Date: ' + endDate);

    const sheets = getSheets();
    const loanSheet = sheets.loans;

    if (!loanSheet) {
      throw new Error('EmployeeLoans sheet not found');
    }

    const data = loanSheet.getDataRange().getValues();
    const headers = data[0];

    Logger.log('📊 Total rows in sheet: ' + (data.length - 1));
    Logger.log('📋 Headers: ' + headers.join(', '));

    // Find column indexes
    const empIdCol = findColumnIndex(headers, 'Employee ID');
    const transDateCol = findColumnIndex(headers, 'TransactionDate');

    if (empIdCol === -1) {
      throw new Error('Employee ID column not found in EmployeeLoans sheet');
    }

    Logger.log('Employee ID column index: ' + empIdCol);

    // Log first few employee IDs to see what's in the sheet
    Logger.log('\n📝 Employee IDs in sheet:');
    const uniqueEmpIds = [];
    for (let i = 1; i < Math.min(data.length, 6); i++) {
      const rowEmpId = data[i][empIdCol];
      Logger.log('  Row ' + (i + 1) + ': "' + rowEmpId + '" (type: ' + typeof rowEmpId + ')');
      if (rowEmpId && !uniqueEmpIds.includes(rowEmpId)) {
        uniqueEmpIds.push(rowEmpId);
      }
    }

    // Parse filter dates if provided
    let filterStartDate = startDate ? parseDate(startDate) : null;
    let filterEndDate = endDate ? parseDate(endDate) : null;

    // Filter records
    let records = [];
    let matchedRows = 0;
    let skippedNoDate = 0;
    let skippedError = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowEmpId = row[empIdCol];

      if (rowEmpId === employeeId) {
        matchedRows++;
        Logger.log('✓ Row ' + (i + 1) + ' matches employee ID');

        try {
          // Skip if no transaction date
          if (!row[transDateCol]) {
            Logger.log('  ⚠️ Skipping - missing transaction date');
            skippedNoDate++;
            continue;
          }

          const record = buildObjectFromRow(row, headers);
          Logger.log('  ✓ Record built successfully');

          // Apply date filters
          if (filterStartDate || filterEndDate) {
            const transDate = parseDate(row[transDateCol]);

            if (filterStartDate && transDate < filterStartDate) {
              Logger.log('  ⚠️ Skipping - before start date');
              continue;
            }
            if (filterEndDate && transDate > filterEndDate) {
              Logger.log('  ⚠️ Skipping - after end date');
              continue;
            }
          }

          records.push(record);
          Logger.log('  ✅ Record added to results');
        } catch (error) {
          Logger.log('  ❌ Error processing row: ' + error.message);
          skippedError++;
          continue;
        }
      }
    }

    Logger.log('\n📊 Summary:');
    Logger.log('  Total rows checked: ' + (data.length - 1));
    Logger.log('  Rows matching employee ID: ' + matchedRows);
    Logger.log('  Records added: ' + records.length);
    Logger.log('  Skipped (no date): ' + skippedNoDate);
    Logger.log('  Skipped (error): ' + skippedError);

    // Sort chronologically: TransactionDate first, then Timestamp
    records.sort((a, b) => {
      const dateA = parseDate(a['TransactionDate']);
      const dateB = parseDate(b['TransactionDate']);

      if (dateA.getTime() !== dateB.getTime()) {
        return dateA - dateB;  // Ascending (oldest first)
      }

      const timeA = parseDate(a['Timestamp']);
      const timeB = parseDate(b['Timestamp']);
      return timeA - timeB;  // Ascending
    });

    // Sanitize records for web serialization
    const sanitizedRecords = records.map(function(record) {
      return sanitizeForWeb(record);
    });

    Logger.log('✅ Returning ' + sanitizedRecords.length + ' loan records');
    Logger.log('========== GET LOAN HISTORY COMPLETE ==========\n');

    return { success: true, data: sanitizedRecords };

  } catch (error) {
    Logger.log('❌ ERROR in getLoanHistory: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== GET LOAN HISTORY FAILED ==========\n');
    return { success: false, error: error.message };
  }
}

// ==================== RECALCULATE LOAN BALANCES ====================

/**
 * Recalculates all loan balances for an employee chronologically
 * CRITICAL: Must be called after any loan record is created/updated/deleted
 *
 * @param {string} employeeId - Employee ID
 * @returns {Object} Result with success flag
 */
function recalculateLoanBalances(employeeId) {
  try {
    Logger.log('\n========== RECALCULATE LOAN BALANCES ==========');
    Logger.log('Employee ID: ' + employeeId);

    const sheets = getSheets();
    const loanSheet = sheets.loans;

    if (!loanSheet) {
      throw new Error('EmployeeLoans sheet not found');
    }

    const data = loanSheet.getDataRange().getValues();
    const headers = data[0];

    // Find column indexes
    const empIdCol = findColumnIndex(headers, 'Employee ID');
    const transDateCol = findColumnIndex(headers, 'TransactionDate');
    const timestampCol = findColumnIndex(headers, 'Timestamp');
    const loanAmountCol = findColumnIndex(headers, 'LoanAmount');
    const balanceBeforeCol = findColumnIndex(headers, 'BalanceBefore');
    const balanceAfterCol = findColumnIndex(headers, 'BalanceAfter');

    if (empIdCol === -1 || transDateCol === -1 || loanAmountCol === -1) {
      throw new Error('Required columns not found in EmployeeLoans sheet');
    }

    // Collect all records for this employee with row numbers
    const employeeRecords = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[empIdCol] === employeeId) {
        employeeRecords.push({
          rowIndex: i + 1,  // Sheet row number (1-based)
          transactionDate: parseDate(row[transDateCol]),
          timestamp: parseDate(row[timestampCol]),
          loanAmount: row[loanAmountCol]
        });
      }
    }

    if (employeeRecords.length === 0) {
      Logger.log('ℹ️ No loan records to recalculate');
      Logger.log('========== RECALCULATE LOAN BALANCES COMPLETE ==========\n');
      return { success: true, data: 'No records to recalculate' };
    }

    // Sort chronologically: TransactionDate first, then Timestamp
    employeeRecords.sort((a, b) => {
      if (a.transactionDate.getTime() !== b.transactionDate.getTime()) {
        return a.transactionDate - b.transactionDate;  // Ascending
      }
      return a.timestamp - b.timestamp;  // Ascending
    });

    // Recalculate balances
    let runningBalance = 0;

    for (let i = 0; i < employeeRecords.length; i++) {
      const record = employeeRecords[i];
      const balanceBefore = runningBalance;
      const balanceAfter = runningBalance + record.loanAmount;

      // Update sheet
      loanSheet.getRange(record.rowIndex, balanceBeforeCol + 1).setValue(balanceBefore);
      loanSheet.getRange(record.rowIndex, balanceAfterCol + 1).setValue(balanceAfter);

      Logger.log('  Row ' + record.rowIndex + ': ' +
                'Before=R' + balanceBefore.toFixed(2) + ', ' +
                'Amount=R' + record.loanAmount.toFixed(2) + ', ' +
                'After=R' + balanceAfter.toFixed(2));

      runningBalance = balanceAfter;
    }

    SpreadsheetApp.flush();

    Logger.log('✅ Recalculated ' + employeeRecords.length + ' loan records');
    Logger.log('Final balance: R' + runningBalance.toFixed(2));
    Logger.log('========== RECALCULATE LOAN BALANCES COMPLETE ==========\n');

    return { success: true, data: runningBalance };

  } catch (error) {
    Logger.log('❌ ERROR in recalculateLoanBalances: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== RECALCULATE LOAN BALANCES FAILED ==========\n');
    return { success: false, error: error.message };
  }
}

// ==================== AUTO-SYNC FROM PAYROLL ====================

/**
 * Synchronizes loan record from payslip
 * CRITICAL: Called by onChange trigger when payslip has loan activity
 *
 * @param {string|number} recordNumber - MASTERSALARY RECORDNUMBER
 * @returns {Object} Result with success flag
 */
function syncLoanForPayslip(recordNumber) {
  try {
    Logger.log('\n========== SYNC LOAN FOR PAYSLIP ==========');
    Logger.log('Record Number: ' + recordNumber);

    // Get payslip details
    const payslipResult = getPayslip(recordNumber);
    if (!payslipResult.success) {
      throw new Error('Payslip not found: ' + recordNumber);
    }

    const payslip = payslipResult.data;

    // Check if there's loan activity
    const hasLoanDeduction = payslip['LoanDeductionThisWeek'] && payslip['LoanDeductionThisWeek'] > 0;
    const hasNewLoan = payslip['NewLoanThisWeek'] && payslip['NewLoanThisWeek'] > 0;

    if (!hasLoanDeduction && !hasNewLoan) {
      Logger.log('ℹ️ No loan activity in this payslip');
      Logger.log('========== SYNC LOAN FOR PAYSLIP COMPLETE ==========\n');
      return { success: true, data: 'No loan activity' };
    }

    // Get employee details
    const empResult = getEmployeeByName(payslip['EMPLOYEE NAME']);
    if (!empResult.success) {
      throw new Error('Employee not found: ' + payslip['EMPLOYEE NAME']);
    }

    const employee = empResult.data;

    // Check if loan record already exists
    const existingLoan = findLoanRecordBySalaryLink(recordNumber);

    if (existingLoan.success && existingLoan.data) {
      // Update existing loan record
      Logger.log('📝 Updating existing loan record');
      return updateLoanRecord(existingLoan.data, payslip, employee);
    } else {
      // Create new loan record
      Logger.log('📝 Creating new loan record');
      return createLoanRecord(payslip, employee);
    }

  } catch (error) {
    Logger.log('❌ ERROR in syncLoanForPayslip: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== SYNC LOAN FOR PAYSLIP FAILED ==========\n');
    return { success: false, error: error.message };
  }
}

/**
 * Finds loan record by SalaryLink
 *
 * @param {string|number} recordNumber - MASTERSALARY RECORDNUMBER
 * @returns {Object} Result with loan record or null
 */
function findLoanRecordBySalaryLink(recordNumber) {
  try {
    const sheets = getSheets();
    const loanSheet = sheets.loans;

    if (!loanSheet) {
      throw new Error('EmployeeLoans sheet not found');
    }

    const data = loanSheet.getDataRange().getValues();
    const headers = data[0];

    const salaryLinkCol = findColumnIndex(headers, 'SalaryLink');

    if (salaryLinkCol === -1) {
      throw new Error('SalaryLink column not found');
    }

    // Search for matching SalaryLink
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const salaryLink = String(row[salaryLinkCol]);
      const searchValue = String(recordNumber);

      if (salaryLink === searchValue) {
        const record = buildObjectFromRow(row, headers);
        record._rowIndex = i + 1;  // Store row index for updates
        return { success: true, data: record };
      }
    }

    return { success: true, data: null };

  } catch (error) {
    Logger.log('❌ ERROR in findLoanRecordBySalaryLink: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Creates loan record from payslip data
 *
 * @param {Object} payslip - Payslip data
 * @param {Object} employee - Employee data
 * @returns {Object} Result with success flag
 */
function createLoanRecord(payslip, employee) {
  try {
    const loanDeduction = payslip['LoanDeductionThisWeek'] || 0;
    const newLoan = payslip['NewLoanThisWeek'] || 0;

    // Calculate net loan amount (negative = repayment, positive = disbursement)
    const netLoanAmount = newLoan - loanDeduction;

    // Determine loan type and disbursement mode
    let loanType, disbursementMode;

    if (netLoanAmount < 0) {
      loanType = 'Repayment';
      disbursementMode = 'With Salary';
    } else {
      loanType = 'Disbursement';
      disbursementMode = payslip['LoanDisbursementType'] || 'With Salary';
    }

    // Create loan transaction
    const loanData = {
      employeeId: employee.id,
      transactionDate: payslip['WEEKENDING'],
      loanAmount: netLoanAmount,
      loanType: loanType,
      disbursementMode: disbursementMode,
      salaryLink: String(payslip['RECORDNUMBER']),
      notes: 'Auto-synced from payslip #' + payslip['RECORDNUMBER'] +
             (loanDeduction > 0 ? ' (Repayment: R' + loanDeduction.toFixed(2) + ')' : '') +
             (newLoan > 0 ? ' (New Loan: R' + newLoan.toFixed(2) + ')' : '')
    };

    const result = addLoanTransaction(loanData);

    if (result.success) {
      // Recalculate all balances
      recalculateLoanBalances(employee.id);
    }

    return result;

  } catch (error) {
    Logger.log('❌ ERROR in createLoanRecord: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Updates existing loan record from payslip data
 *
 * @param {Object} existingLoan - Existing loan record
 * @param {Object} payslip - Payslip data
 * @param {Object} employee - Employee data
 * @returns {Object} Result with success flag
 */
function updateLoanRecord(existingLoan, payslip, employee) {
  try {
    const sheets = getSheets();
    const loanSheet = sheets.loans;

    if (!loanSheet) {
      throw new Error('EmployeeLoans sheet not found');
    }

    const loanDeduction = payslip['LoanDeductionThisWeek'] || 0;
    const newLoan = payslip['NewLoanThisWeek'] || 0;
    const netLoanAmount = newLoan - loanDeduction;

    // Determine loan type and disbursement mode
    let loanType, disbursementMode;

    if (netLoanAmount < 0) {
      loanType = 'Repayment';
      disbursementMode = 'With Salary';
    } else {
      loanType = 'Disbursement';
      disbursementMode = payslip['LoanDisbursementType'] || 'With Salary';
    }

    const notes = 'Auto-synced from payslip #' + payslip['RECORDNUMBER'] +
                 (loanDeduction > 0 ? ' (Repayment: R' + loanDeduction.toFixed(2) + ')' : '') +
                 (newLoan > 0 ? ' (New Loan: R' + newLoan.toFixed(2) + ')' : '') +
                 ' [Updated]';

    // Update the loan record
    const rowIndex = existingLoan._rowIndex;
    const headers = loanSheet.getDataRange().getValues()[0];

    const timestampCol = findColumnIndex(headers, 'Timestamp');
    const loanAmountCol = findColumnIndex(headers, 'LoanAmount');
    const loanTypeCol = findColumnIndex(headers, 'LoanType');
    const disbursementModeCol = findColumnIndex(headers, 'DisbursementMode');
    const notesCol = findColumnIndex(headers, 'Notes');

    // Update fields (except TransactionDate which must never change)
    loanSheet.getRange(rowIndex, timestampCol + 1).setValue(new Date());
    loanSheet.getRange(rowIndex, loanAmountCol + 1).setValue(netLoanAmount);
    loanSheet.getRange(rowIndex, loanTypeCol + 1).setValue(loanType);
    loanSheet.getRange(rowIndex, disbursementModeCol + 1).setValue(disbursementMode);
    loanSheet.getRange(rowIndex, notesCol + 1).setValue(notes);

    SpreadsheetApp.flush();

    Logger.log('✅ Loan record updated');

    // Recalculate all balances
    recalculateLoanBalances(employee.id);

    return { success: true, data: 'Loan record updated' };

  } catch (error) {
    Logger.log('❌ ERROR in updateLoanRecord: ' + error.message);
    return { success: false, error: error.message };
  }
}

// ==================== VALIDATION ====================

/**
 * Validates loan transaction data
 *
 * @param {Object} data - Loan data to validate
 * @throws {Error} If validation fails
 */
function validateLoan(data) {
  const errors = [];

  // Required fields
  if (!data.employeeId) {
    errors.push('Employee ID is required');
  }

  if (!data.transactionDate) {
    errors.push('Transaction Date is required');
  }

  if (data.loanAmount === undefined || data.loanAmount === null) {
    errors.push('Loan Amount is required');
  }

  if (!data.loanType) {
    errors.push('Loan Type is required');
  }

  if (!data.disbursementMode) {
    errors.push('Disbursement Mode is required');
  }

  // Validate loan amount
  if (data.loanAmount !== undefined && data.loanAmount === 0) {
    errors.push('Loan amount cannot be zero');
  }

  // Validate loan type matches amount sign
  if (data.loanType === 'Disbursement' && data.loanAmount <= 0) {
    errors.push('Disbursement must have positive amount');
  }

  if (data.loanType === 'Repayment' && data.loanAmount >= 0) {
    errors.push('Repayment must have negative amount');
  }

  // Validate loan type
  if (data.loanType && !LOAN_TYPES.includes(data.loanType)) {
    errors.push('Invalid loan type. Must be one of: ' + LOAN_TYPES.join(', '));
  }

  // Validate disbursement mode
  if (data.disbursementMode && !DISBURSEMENT_MODES.includes(data.disbursementMode)) {
    errors.push('Invalid disbursement mode. Must be one of: ' + DISBURSEMENT_MODES.join(', '));
  }

  // Check for duplicate SalaryLink if provided
  if (data.salaryLink) {
    const existing = findLoanRecordBySalaryLink(data.salaryLink);
    if (existing.success && existing.data) {
      errors.push('Loan record already exists for payslip #' + data.salaryLink);
    }
  }

  if (errors.length > 0) {
    throw new Error(errors.join('; '));
  }
}

// ==================== GET TOTAL OUTSTANDING LOANS ====================

/**
 * Gets total outstanding loan balance across all employees
 * Used for dashboard statistics
 *
 * @returns {Object} Result with success flag and total balance
 */
function getTotalOutstandingLoans() {
  try {
    Logger.log('\n========== GET TOTAL OUTSTANDING LOANS ==========');

    const sheets = getSheets();
    const loanSheet = sheets.loans;

    if (!loanSheet) {
      throw new Error('EmployeeLoans sheet not found');
    }

    const data = loanSheet.getDataRange().getValues();
    const headers = data[0];

    // Find column indexes
    const empIdCol = findColumnIndex(headers, 'Employee ID');
    const transDateCol = findColumnIndex(headers, 'TransactionDate');
    const timestampCol = findColumnIndex(headers, 'Timestamp');
    const balanceAfterCol = findColumnIndex(headers, 'BalanceAfter');

    if (empIdCol === -1 || balanceAfterCol === -1) {
      throw new Error('Required columns not found in EmployeeLoans sheet');
    }

    // Group records by employee
    const employeeBalances = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const empId = row[empIdCol];

      if (!empId) continue;

      try {
        // Skip if critical data is missing
        if (!row[transDateCol] || !row[timestampCol]) {
          continue;
        }

        const balanceAfter = row[balanceAfterCol];
        if (balanceAfter === null || balanceAfter === undefined || balanceAfter === '') {
          continue;
        }

        if (!employeeBalances[empId]) {
          employeeBalances[empId] = [];
        }

        employeeBalances[empId].push({
          transactionDate: parseDate(row[transDateCol]),
          timestamp: parseDate(row[timestampCol]),
          balanceAfter: parseFloat(balanceAfter)
        });
      } catch (error) {
        continue;
      }
    }

    // Get most recent balance for each employee and sum
    let totalOutstanding = 0;
    let employeeCount = 0;

    for (const empId in employeeBalances) {
      const records = employeeBalances[empId];

      // Sort chronologically: TransactionDate first, then Timestamp (descending)
      records.sort((a, b) => {
        if (a.transactionDate.getTime() !== b.transactionDate.getTime()) {
          return b.transactionDate - a.transactionDate;  // Descending
        }
        return b.timestamp - a.timestamp;  // Descending
      });

      const currentBalance = records[0].balanceAfter;
      if (currentBalance > 0) {
        totalOutstanding += currentBalance;
        employeeCount++;
      }
    }

    Logger.log('✅ Total outstanding: R' + totalOutstanding.toFixed(2));
    Logger.log('Employees with loans: ' + employeeCount);
    Logger.log('========== GET TOTAL OUTSTANDING LOANS COMPLETE ==========\n');

    return {
      success: true,
      data: {
        total: totalOutstanding,
        employeeCount: employeeCount
      }
    };

  } catch (error) {
    Logger.log('❌ ERROR in getTotalOutstandingLoans: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== GET TOTAL OUTSTANDING LOANS FAILED ==========\n');
    return { success: false, error: error.message };
  }
}

// ==================== TEST FUNCTIONS ====================

/**
 * Test function for loan operations
 */
function test_loanOperations() {
  Logger.log('\n========== TEST: LOAN OPERATIONS ==========');

  // You'll need to replace with actual employee ID
  const testEmployeeId = 'test-emp-id';

  // Test 1: Add disbursement
  const disbursement = addLoanTransaction({
    employeeId: testEmployeeId,
    transactionDate: new Date(),
    loanAmount: 500.00,
    loanType: 'Disbursement',
    disbursementMode: 'Manual Entry',
    notes: 'Test disbursement'
  });

  Logger.log('Disbursement result: ' + (disbursement.success ? 'SUCCESS' : 'FAILED'));

  // Test 2: Add repayment
  const repayment = addLoanTransaction({
    employeeId: testEmployeeId,
    transactionDate: new Date(),
    loanAmount: -150.00,
    loanType: 'Repayment',
    disbursementMode: 'Manual Entry',
    notes: 'Test repayment'
  });

  Logger.log('Repayment result: ' + (repayment.success ? 'SUCCESS' : 'FAILED'));

  // Test 3: Get current balance
  const balance = getCurrentLoanBalance(testEmployeeId);
  Logger.log('Current balance: ' + (balance.success ? 'R' + balance.data.toFixed(2) : 'FAILED'));

  Logger.log('========== TEST COMPLETE ==========\n');
}
