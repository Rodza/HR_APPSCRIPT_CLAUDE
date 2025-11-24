/**
 * REPORTS.GS - Report Generation Module
 * HR Payroll System for SA Grinding Wheels & Scorpio Abrasives
 *
 * Generates formatted Google Sheets reports:
 * 1. Outstanding Loans Report - All employees with loan balances > 0
 * 2. Individual Statement - Complete loan history for one employee
 * 3. Weekly Payroll Summary - Detailed breakdown for a specific week
 *
 * All reports are saved to "Payroll Reports" folder in Google Drive
 * with sharing set to "Anyone with link can view"
 */

// ==================== OUTSTANDING LOANS REPORT ====================

/**
 * Generates Outstanding Loans Report showing all employees with current balances > 0
 *
 * @param {Date|string} [asOfDate] - As-of date (default: today)
 * @returns {Object} Result with success flag and report URL
 *
 * @example
 * const result = generateOutstandingLoansReport();
 * // Returns: { success: true, data: { url: '...', fileName: '...' } }
 */
function generateOutstandingLoansReport(asOfDate) {
  try {
    Logger.log('\n========== GENERATE OUTSTANDING LOANS REPORT ==========');

    const reportDate = asOfDate ? parseDate(asOfDate) : new Date();
    Logger.log('Report as of: ' + formatDate(reportDate));

    // Get all employees
    const empResult = listEmployees({ activeOnly: true });
    if (!empResult.success) {
      throw new Error('Failed to get employees: ' + empResult.error);
    }

    const employees = empResult.data;
    Logger.log('📊 Processing ' + employees.length + ' employees');

    // Get loan balances for each employee
    const reportData = [];
    let totalOutstanding = 0;

    for (let i = 0; i < employees.length; i++) {
      const emp = employees[i];
      const balanceResult = getCurrentLoanBalance(emp.id);

      if (balanceResult.success && balanceResult.data > 0) {
        const balance = balanceResult.data;

        // Get last transaction date
        const historyResult = getLoanHistory(emp.id);
        let lastTransactionDate = '';

        if (historyResult.success && historyResult.data.length > 0) {
          const records = historyResult.data;
          const lastRecord = records[records.length - 1];
          lastTransactionDate = formatDate(lastRecord['TransactionDate']);
        }

        reportData.push({
          employeeName: emp.REFNAME,
          employer: emp.EMPLOYER,
          balance: balance,
          lastTransactionDate: lastTransactionDate
        });

        totalOutstanding += balance;
      }
    }

    Logger.log('💰 Total outstanding: R' + totalOutstanding.toFixed(2));
    Logger.log('👥 Employees with loans: ' + reportData.length);

    // Sort by balance (descending)
    reportData.sort((a, b) => b.balance - a.balance);

    // Create Google Sheet
    const fileName = 'Outstanding Loans Report - ' + formatDate(reportDate);
    const spreadsheet = SpreadsheetApp.create(fileName);
    const sheet = spreadsheet.getActiveSheet();

    // Header
    sheet.setName('Outstanding Loans');
    sheet.getRange('A1:E1').setValues([[
      'OUTSTANDING LOANS REPORT',
      '',
      '',
      'As of: ' + formatDate(reportDate),
      ''
    ]]);
    sheet.getRange('A1:E1').setFontWeight('bold').setFontSize(14);

    // Column headers
    sheet.getRange('A3:D3').setValues([[
      'Employee Name',
      'Employer',
      'Outstanding Balance',
      'Last Transaction Date'
    ]]);
    sheet.getRange('A3:D3').setFontWeight('bold').setBackground('#4CAF50').setFontColor('#FFFFFF');

    // Data rows
    let rowNum = 4;
    for (let i = 0; i < reportData.length; i++) {
      const row = reportData[i];
      sheet.getRange(rowNum, 1, 1, 4).setValues([[
        row.employeeName,
        row.employer,
        row.balance,
        row.lastTransactionDate
      ]]);
      rowNum++;
    }

    // Total row
    sheet.getRange(rowNum, 1, 1, 3).setValues([[
      '',
      'TOTAL OUTSTANDING:',
      totalOutstanding
    ]]);
    sheet.getRange(rowNum, 2, 1, 2).setFontWeight('bold').setBackground('#FFD700');

    // Format currency column
    const dataRange = sheet.getRange(4, 3, reportData.length + 1, 1);
    dataRange.setNumberFormat('"R"#,##0.00');

    // Auto-resize columns
    sheet.autoResizeColumns(1, 4);

    // Move to reports folder and set sharing
    const reportUrl = moveToReportsFolder(spreadsheet);

    Logger.log('✅ Report generated: ' + fileName);
    Logger.log('📎 URL: ' + reportUrl);
    Logger.log('========== GENERATE OUTSTANDING LOANS REPORT COMPLETE ==========\n');

    return {
      success: true,
      data: {
        url: reportUrl,
        fileName: fileName,
        totalOutstanding: totalOutstanding,
        employeeCount: reportData.length
      }
    };

  } catch (error) {
    Logger.log('❌ ERROR in generateOutstandingLoansReport: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== GENERATE OUTSTANDING LOANS REPORT FAILED ==========\n');
    return { success: false, error: error.message };
  }
}

// ==================== INDIVIDUAL STATEMENT REPORT ====================

/**
 * Generates Individual Statement Report for one employee
 *
 * @param {string} employeeId - Employee ID
 * @param {Date|string} startDate - Statement start date
 * @param {Date|string} endDate - Statement end date
 * @returns {Object} Result with success flag and report URL
 *
 * @example
 * const result = generateIndividualStatementReport('e0b6115a', '2025-01-01', '2025-10-31');
 */
function generateIndividualStatementReport(employeeId, startDate, endDate) {
  try {
    Logger.log('\n========== GENERATE INDIVIDUAL STATEMENT REPORT ==========');
    Logger.log('Employee ID: ' + employeeId);

    const start = parseDate(startDate);
    const end = parseDate(endDate);

    Logger.log('Period: ' + formatDate(start) + ' to ' + formatDate(end));

    // Get employee details
    const empResult = getEmployeeById(employeeId);
    if (!empResult.success) {
      throw new Error('Employee not found: ' + employeeId);
    }

    const employee = empResult.data;
    Logger.log('Employee: ' + employee.REFNAME);

    // Get opening balance (as of day before start date)
    const dayBeforeStart = new Date(start);
    dayBeforeStart.setDate(dayBeforeStart.getDate() - 1);

    const historyBeforeResult = getLoanHistory(employeeId, null, dayBeforeStart);
    let openingBalance = 0;

    if (historyBeforeResult.success && historyBeforeResult.data.length > 0) {
      const records = historyBeforeResult.data;
      const lastRecord = records[records.length - 1];
      openingBalance = lastRecord['BalanceAfter'];
    }

    Logger.log('Opening balance: R' + openingBalance.toFixed(2));

    // Get transactions in period
    const historyResult = getLoanHistory(employeeId, start, end);
    if (!historyResult.success) {
      throw new Error('Failed to get loan history: ' + historyResult.error);
    }

    const transactions = historyResult.data;
    Logger.log('Transactions: ' + transactions.length);

    // Calculate closing balance
    const closingBalance = transactions.length > 0 ?
                          transactions[transactions.length - 1]['BalanceAfter'] :
                          openingBalance;

    Logger.log('Closing balance: R' + closingBalance.toFixed(2));

    // Create Google Sheet
    const fileName = 'Loan Statement - ' + employee.REFNAME + ' - ' +
                    formatDate(start) + ' to ' + formatDate(end);
    const spreadsheet = SpreadsheetApp.create(fileName);
    const sheet = spreadsheet.getActiveSheet();

    // Header
    sheet.setName('Loan Statement');
    sheet.getRange('A1:F1').setValues([[
      'EMPLOYEE LOAN STATEMENT',
      '',
      '',
      '',
      '',
      ''
    ]]);
    sheet.getRange('A1:F1').setFontWeight('bold').setFontSize(14);

    // Employee details
    sheet.getRange('A3:B3').setValues([['Employee:', employee.REFNAME]]);
    sheet.getRange('A4:B4').setValues([['Employer:', employee.EMPLOYER]]);
    sheet.getRange('A5:B5').setValues([['Period:', formatDate(start) + ' to ' + formatDate(end)]]);
    sheet.getRange('A3:A5').setFontWeight('bold');

    // Opening balance
    sheet.getRange('A7:B7').setValues([['Opening Balance:', openingBalance]]);
    sheet.getRange('A7:B7').setFontWeight('bold').setBackground('#E3F2FD');
    sheet.getRange('B7').setNumberFormat('"R"#,##0.00');

    // Transaction headers
    sheet.getRange('A9:F9').setValues([[
      'Date',
      'Type',
      'Amount',
      'Mode',
      'Balance',
      'Notes'
    ]]);
    sheet.getRange('A9:F9').setFontWeight('bold').setBackground('#4CAF50').setFontColor('#FFFFFF');

    // Transaction rows
    let rowNum = 10;
    for (let i = 0; i < transactions.length; i++) {
      const txn = transactions[i];
      sheet.getRange(rowNum, 1, 1, 6).setValues([[
        formatDate(txn['TransactionDate']),
        txn['LoanType'],
        txn['LoanAmount'],
        txn['DisbursementMode'],
        txn['BalanceAfter'],
        txn['Notes']
      ]]);
      rowNum++;
    }

    // Closing balance
    sheet.getRange(rowNum + 1, 1, 1, 2).setValues([['Closing Balance:', closingBalance]]);
    sheet.getRange(rowNum + 1, 1, 1, 2).setFontWeight('bold').setBackground('#FFD700');
    sheet.getRange(rowNum + 1, 2).setNumberFormat('"R"#,##0.00');

    // Format currency columns
    if (transactions.length > 0) {
      sheet.getRange(10, 3, transactions.length, 1).setNumberFormat('"R"#,##0.00');
      sheet.getRange(10, 5, transactions.length, 1).setNumberFormat('"R"#,##0.00');
    }

    // Auto-resize columns
    sheet.autoResizeColumns(1, 6);

    // Move to reports folder and set sharing
    const reportUrl = moveToReportsFolder(spreadsheet);

    Logger.log('✅ Report generated: ' + fileName);
    Logger.log('📎 URL: ' + reportUrl);
    Logger.log('========== GENERATE INDIVIDUAL STATEMENT REPORT COMPLETE ==========\n');

    return {
      success: true,
      data: {
        url: reportUrl,
        fileName: fileName,
        employeeName: employee.REFNAME,
        openingBalance: openingBalance,
        closingBalance: closingBalance,
        transactionCount: transactions.length
      }
    };

  } catch (error) {
    Logger.log('❌ ERROR in generateIndividualStatementReport: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== GENERATE INDIVIDUAL STATEMENT REPORT FAILED ==========\n');
    return { success: false, error: error.message };
  }
}

// ==================== WEEKLY PAYROLL SUMMARY REPORT ====================

/**
 * Generates Weekly Payroll Summary Report for a specific week
 *
 * @param {Date|string} weekEnding - Week ending date
 * @returns {Object} Result with success flag and report URL
 *
 * @example
 * const result = generateWeeklyPayrollSummaryReport('2025-10-17');
 */
function generateWeeklyPayrollSummaryReport(weekEnding) {
  try {
    Logger.log('\n========== GENERATE WEEKLY PAYROLL SUMMARY REPORT ==========');

    const weekEnd = parseDate(weekEnding);
    Logger.log('Week Ending: ' + formatDate(weekEnd));

    // Get all payslips for this week
    const payslipsResult = listPayslips({ weekEnding: weekEnd });
    if (!payslipsResult.success) {
      throw new Error('Failed to get payslips: ' + payslipsResult.error);
    }

    const payslips = payslipsResult.data;
    Logger.log('📊 Processing ' + payslips.length + ' payslips');

    if (payslips.length === 0) {
      throw new Error('No payslips found for week ending ' + formatDate(weekEnd));
    }

    // Calculate totals
    let totals = {
      employees: payslips.length,
      standardHours: 0,
      overtimeHours: 0,
      grossPay: 0,
      uif: 0,
      otherDeductions: 0,
      loanDeductions: 0,
      netPay: 0,
      paidToAccounts: 0
    };

    const byEmployer = {};

    for (let i = 0; i < payslips.length; i++) {
      const p = payslips[i];

      const totalHours = (p['HOURS'] || 0) + ((p['MINUTES'] || 0) / 60);
      const totalOvertimeHours = (p['OVERTIMEHOURS'] || 0) + ((p['OVERTIMEMINUTES'] || 0) / 60);

      totals.standardHours += totalHours;
      totals.overtimeHours += totalOvertimeHours;
      totals.grossPay += p['GROSSSALARY'] || 0;
      totals.uif += p['UIF'] || 0;
      totals.otherDeductions += p['OTHER DEDUCTIONS'] || 0;
      totals.loanDeductions += p['LoanDeductionThisWeek'] || 0;
      totals.netPay += p['NETTSALARY'] || 0;
      totals.paidToAccounts += p['PaidToAccount'] || 0;

      // Track by employer
      const employer = p['EMPLOYER'];
      if (!byEmployer[employer]) {
        byEmployer[employer] = {
          count: 0,
          grossPay: 0,
          netPay: 0,
          paidToAccounts: 0
        };
      }

      byEmployer[employer].count++;
      byEmployer[employer].grossPay += p['GROSSSALARY'] || 0;
      byEmployer[employer].netPay += p['NETTSALARY'] || 0;
      byEmployer[employer].paidToAccounts += p['PaidToAccount'] || 0;
    }

    // Create Google Sheet
    const fileName = 'Weekly Payroll Summary - ' + formatDate(weekEnd);
    const spreadsheet = SpreadsheetApp.create(fileName);

    // ===== TAB 1: Payroll Register =====
    const registerSheet = spreadsheet.getActiveSheet();
    registerSheet.setName('Payroll Register');

    // Header
    registerSheet.getRange('A1:K1').setValues([[
      'WEEKLY PAYROLL REGISTER - Week Ending: ' + formatDate(weekEnd),
      '', '', '', '', '', '', '', '', '', ''
    ]]);
    registerSheet.getRange('A1:K1').setFontWeight('bold').setFontSize(14);

    // Column headers
    registerSheet.getRange('A3:K3').setValues([[
      'Employee',
      'Employer',
      'Std Hours',
      'OT Hours',
      'Gross Pay',
      'UIF',
      'Other Ded.',
      'Loan Ded.',
      'Net Pay',
      'New Loan',
      'Paid to Account'
    ]]);
    registerSheet.getRange('A3:K3').setFontWeight('bold').setBackground('#4CAF50').setFontColor('#FFFFFF');

    // Data rows
    let rowNum = 4;
    for (let i = 0; i < payslips.length; i++) {
      const p = payslips[i];
      const stdHours = (p['HOURS'] || 0) + ((p['MINUTES'] || 0) / 60);
      const otHours = (p['OVERTIMEHOURS'] || 0) + ((p['OVERTIMEMINUTES'] || 0) / 60);

      registerSheet.getRange(rowNum, 1, 1, 11).setValues([[
        p['EMPLOYEE NAME'],
        p['EMPLOYER'],
        stdHours,
        otHours,
        p['GROSSSALARY'] || 0,
        p['UIF'] || 0,
        p['OTHER DEDUCTIONS'] || 0,
        p['LoanDeductionThisWeek'] || 0,
        p['NETTSALARY'] || 0,
        p['NewLoanThisWeek'] || 0,
        p['PaidToAccount'] || 0
      ]]);
      rowNum++;
    }

    // Totals row
    registerSheet.getRange(rowNum, 1, 1, 11).setValues([[
      'TOTALS',
      totals.employees + ' employees',
      totals.standardHours,
      totals.overtimeHours,
      totals.grossPay,
      totals.uif,
      totals.otherDeductions,
      totals.loanDeductions,
      totals.netPay,
      '',
      totals.paidToAccounts
    ]]);
    registerSheet.getRange(rowNum, 1, 1, 11).setFontWeight('bold').setBackground('#FFD700');

    // Format currency and number columns
    registerSheet.getRange(4, 3, payslips.length + 1, 1).setNumberFormat('0.00');  // Std Hours
    registerSheet.getRange(4, 4, payslips.length + 1, 1).setNumberFormat('0.00');  // OT Hours
    registerSheet.getRange(4, 5, payslips.length + 1, 7).setNumberFormat('"R"#,##0.00');  // Currency columns

    // Auto-resize
    registerSheet.autoResizeColumns(1, 11);

    // ===== TAB 2: Summary =====
    const summarySheet = spreadsheet.insertSheet('Summary');

    // Header
    summarySheet.getRange('A1:B1').setValues([['PAYROLL SUMMARY', '']]);
    summarySheet.getRange('A1:B1').setFontWeight('bold').setFontSize(14);

    summarySheet.getRange('A2:B2').setValues([['Week Ending:', formatDate(weekEnd)]]);

    // Overall summary
    summarySheet.getRange('A4:B4').setValues([['OVERALL SUMMARY', '']]);
    summarySheet.getRange('A4:B4').setFontWeight('bold').setBackground('#4CAF50').setFontColor('#FFFFFF');

    summarySheet.getRange('A5:B11').setValues([
      ['Total Employees:', totals.employees],
      ['Total Standard Hours:', totals.standardHours],
      ['Total Overtime Hours:', totals.overtimeHours],
      ['Total Gross Pay:', totals.grossPay],
      ['Total UIF:', totals.uif],
      ['Total Deductions:', totals.otherDeductions + totals.loanDeductions],
      ['Total Paid to Accounts:', totals.paidToAccounts]
    ]);

    summarySheet.getRange('A5:A11').setFontWeight('bold');
    summarySheet.getRange('B6:B11').setNumberFormat('"R"#,##0.00');

    // By employer
    summarySheet.getRange('A13:B13').setValues([['BY EMPLOYER', '']]);
    summarySheet.getRange('A13:B13').setFontWeight('bold').setBackground('#4CAF50').setFontColor('#FFFFFF');

    let summaryRow = 14;
    for (const employer in byEmployer) {
      const data = byEmployer[employer];
      summarySheet.getRange(summaryRow, 1, 1, 2).setValues([[employer, '']]);
      summarySheet.getRange(summaryRow, 1).setFontWeight('bold');
      summaryRow++;

      summarySheet.getRange(summaryRow, 1, 4, 2).setValues([
        ['  Employees:', data.count],
        ['  Gross Pay:', data.grossPay],
        ['  Net Pay:', data.netPay],
        ['  Paid to Accounts:', data.paidToAccounts]
      ]);

      summarySheet.getRange(summaryRow, 2, 3, 1).setNumberFormat('"R"#,##0.00');
      summaryRow += 5;
    }

    // Auto-resize
    summarySheet.autoResizeColumns(1, 2);

    // Move to reports folder and set sharing
    const reportUrl = moveToReportsFolder(spreadsheet);

    Logger.log('✅ Report generated: ' + fileName);
    Logger.log('📎 URL: ' + reportUrl);
    Logger.log('========== GENERATE WEEKLY PAYROLL SUMMARY REPORT COMPLETE ==========\n');

    return {
      success: true,
      data: {
        url: reportUrl,
        fileName: fileName,
        weekEnding: formatDate(weekEnd),
        totalEmployees: totals.employees,
        totalPaidToAccounts: totals.paidToAccounts
      }
    };

  } catch (error) {
    Logger.log('❌ ERROR in generateWeeklyPayrollSummaryReport: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== GENERATE WEEKLY PAYROLL SUMMARY REPORT FAILED ==========\n');
    return { success: false, error: error.message };
  }
}

// ==================== HELPER FUNCTIONS ====================

/**
 * Moves spreadsheet to "Payroll Reports" folder and sets sharing
 *
 * @param {Spreadsheet} spreadsheet - Spreadsheet to move
 * @returns {string} Spreadsheet URL
 */
function moveToReportsFolder(spreadsheet) {
  try {
    const file = DriveApp.getFileById(spreadsheet.getId());

    // Get or create "Payroll Reports" folder
    const reportsFolder = getOrCreateReportsFolder();

    // Move file to folder
    file.moveTo(reportsFolder);

    // Set sharing: Anyone with link can view
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return spreadsheet.getUrl();

  } catch (error) {
    Logger.log('⚠️ Warning: Could not move to reports folder: ' + error.message);
    return spreadsheet.getUrl();
  }
}

/**
 * Gets or creates the "Payroll Reports" folder
 *
 * @returns {Folder} Reports folder
 */
function getOrCreateReportsFolder() {
  const folderName = 'Payroll Reports';

  // Search for existing folder
  const folders = DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  }

  // Create new folder
  Logger.log('📁 Creating "Payroll Reports" folder');
  return DriveApp.createFolder(folderName);
}

// ==================== TEST FUNCTIONS ====================

/**
 * Test function for Outstanding Loans Report
 */
function test_outstandingLoansReport() {
  Logger.log('\n========== TEST: OUTSTANDING LOANS REPORT ==========');

  const result = generateOutstandingLoansReport();

  if (result.success) {
    Logger.log('✅ TEST PASSED');
    Logger.log('URL: ' + result.data.url);
    Logger.log('Total Outstanding: R' + result.data.totalOutstanding.toFixed(2));
    Logger.log('Employees with loans: ' + result.data.employeeCount);
  } else {
    Logger.log('❌ TEST FAILED: ' + result.error);
  }

  Logger.log('========== TEST COMPLETE ==========\n');
}

/**
 * Test function for Weekly Payroll Summary Report
 */
function test_weeklyPayrollReport() {
  Logger.log('\n========== TEST: WEEKLY PAYROLL SUMMARY REPORT ==========');

  // Use most recent week (update with actual date)
  const weekEnding = new Date('2025-10-17');

  const result = generateWeeklyPayrollSummaryReport(weekEnding);

  if (result.success) {
    Logger.log('✅ TEST PASSED');
    Logger.log('URL: ' + result.data.url);
    Logger.log('Total Employees: ' + result.data.totalEmployees);
    Logger.log('Total Paid: R' + result.data.totalPaidToAccounts.toFixed(2));
  } else {
    Logger.log('❌ TEST FAILED: ' + result.error);
  }

  Logger.log('========== TEST COMPLETE ==========\n');
}
