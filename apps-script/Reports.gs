/**
 * REPORTS.GS - Report Generation Module
 * HR Payroll System for SA Grinding Wheels & Scorpio Abrasives
 *
 * Generates formatted Google Sheets reports:
 * 1. Outstanding Loans Report - All employees with loan balances > 0
 * 2. Individual Statement - Complete loan history for one employee
 * 3. Weekly Payroll Summary - Detailed breakdown for a specific week
 * 4. Monthly Payroll Summary - Monthly totals from first to last Friday
 * 5. Employee Leave Trends - Analyzes leave patterns to identify potential abuse
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

    const employees = empResult.data.employees;
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

    const payslips = payslipsResult.data.payslips;
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
      newLoans: 0,
      netPay: 0,
      paidToAccounts: 0
    };

    const byEmployer = {};

    // DEBUG: Log first payslip to see what fields are available
    if (payslips.length > 0) {
      Logger.log('🔍 DEBUG: First payslip keys: ' + Object.keys(payslips[0]).join(', '));
      Logger.log('🔍 DEBUG: First payslip PaidtoAccount value: ' + payslips[0]['PaidtoAccount']);
      Logger.log('🔍 DEBUG: First payslip PaidtoAccount type: ' + typeof payslips[0]['PaidtoAccount']);
      Logger.log('🔍 DEBUG: First payslip NETTSALARY: ' + payslips[0]['NETTSALARY']);
    }

    for (let i = 0; i < payslips.length; i++) {
      const p = payslips[i];

      const totalHours = (p['HOURS'] || 0) + ((p['MINUTES'] || 0) / 60);
      const totalOvertimeHours = (p['OVERTIMEHOURS'] || 0) + ((p['OVERTIMEMINUTES'] || 0) / 60);

      // DEBUG: Log PaidtoAccount for first 3 payslips
      if (i < 3) {
        Logger.log('🔍 Payslip ' + (i + 1) + ' (' + p['EMPLOYEE NAME'] + '): PaidtoAccount = ' + p['PaidtoAccount'] + ', Net = ' + p['NETTSALARY']);
      }

      totals.standardHours += totalHours;
      totals.overtimeHours += totalOvertimeHours;
      totals.grossPay += p['GROSSSALARY'] || 0;
      totals.uif += p['UIF'] || 0;
      totals.otherDeductions += p['OTHER DEDUCTIONS'] || 0;
      totals.loanDeductions += p['LoanDeductionThisWeek'] || 0;
      totals.newLoans += p['NewLoanThisWeek'] || 0;
      totals.netPay += p['NETTSALARY'] || 0;

      const paidToAccountValue = p['PaidtoAccount'] || 0;
      totals.paidToAccounts += paidToAccountValue;

      // DEBUG: Log the running total
      if (i < 3) {
        Logger.log('🔍   Added ' + paidToAccountValue + ' to total, running total now: ' + totals.paidToAccounts);
      }

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
      byEmployer[employer].paidToAccounts += p['PaidtoAccount'] || 0;
    }

    // Create Google Sheet
    const fileName = 'Weekly Payroll Summary - ' + formatDate(weekEnd);
    const spreadsheet = SpreadsheetApp.create(fileName);

    // ===== TAB 1: Payroll Register =====
    const registerSheet = spreadsheet.getActiveSheet();
    registerSheet.setName('Payroll Register');

    // Header - set text first (merge will happen after setting widths)
    registerSheet.getRange('A1').setValue('WEEKLY PAYROLL REGISTER - Week Ending: ' + formatDate(weekEnd));
    registerSheet.setRowHeight(1, 30);

    // Column headers
    registerSheet.getRange('A3:K3').setValues([[
      'Employee',
      'Employer',
      'Std Hours',
      'OT Hours',
      'Gross Pay',
      'UIF',
      'Other Ded.',
      'Net Pay',
      'Loan Ded.',
      'New Loan',
      'Paid to Account'
    ]]);
    registerSheet.getRange('A3:K3').setFontWeight('bold').setBackground('#4CAF50').setFontColor('#FFFFFF').setHorizontalAlignment('center');
    registerSheet.setRowHeight(3, 25);

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
        p['NETTSALARY'] || 0,
        p['LoanDeductionThisWeek'] || 0,
        p['NewLoanThisWeek'] || 0,
        p['PaidtoAccount'] || 0
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
      totals.netPay,
      totals.loanDeductions,
      totals.newLoans,
      totals.paidToAccounts
    ]]);
    registerSheet.getRange(rowNum, 1, 1, 11).setFontWeight('bold').setBackground('#FFD700');

    // Format currency and number columns
    registerSheet.getRange(4, 3, payslips.length + 1, 1).setNumberFormat('0.00');  // Std Hours
    registerSheet.getRange(4, 4, payslips.length + 1, 1).setNumberFormat('0.00');  // OT Hours
    registerSheet.getRange(4, 5, payslips.length + 1, 7).setNumberFormat('"R"#,##0.00');  // Currency columns

    // Set hardcoded column widths for optimal layout
    registerSheet.setColumnWidth(1, 220);  // Column A - Employee
    registerSheet.setColumnWidth(2, 135);  // Column B - Employer
    registerSheet.setColumnWidth(3, 80);   // Column C - Std Hours
    registerSheet.setColumnWidth(4, 80);   // Column D - OT Hours
    registerSheet.setColumnWidth(5, 90);   // Column E - Gross Pay
    registerSheet.setColumnWidth(6, 80);   // Column F - UIF
    registerSheet.setColumnWidth(7, 110);  // Column G - Other Ded.
    registerSheet.setColumnWidth(8, 110);  // Column H - Net Pay
    registerSheet.setColumnWidth(9, 110);  // Column I - Loan Ded.
    registerSheet.setColumnWidth(10, 110); // Column J - New Loan
    registerSheet.setColumnWidth(11, 110); // Column K - Paid to Account

    // Merge header cells A1:K1 and apply formatting
    registerSheet.getRange('A1:K1').merge();
    registerSheet.getRange('A1').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

    // Add borders to the entire table (from column headers through last data row)
    const lastRegisterRow = rowNum;
    const registerTableRange = registerSheet.getRange('A3:K' + lastRegisterRow);
    registerTableRange.setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);

    // ===== TAB 2: Summary =====
    const summarySheet = spreadsheet.insertSheet('Summary');

    // Header
    summarySheet.getRange('A1:B1').setValues([['PAYROLL SUMMARY', '']]);
    summarySheet.getRange('A1:B1').setFontWeight('bold').setFontSize(14);

    summarySheet.getRange('A2:B2').setValues([['Week Ending:', formatDate(weekEnd)]]);

    // Overall summary
    summarySheet.getRange('A4:B4').setValues([['OVERALL SUMMARY', '']]);
    summarySheet.getRange('A4:B4').setFontWeight('bold').setBackground('#4CAF50').setFontColor('#FFFFFF');

    summarySheet.getRange('A5:B12').setValues([
      ['Total Employees:', totals.employees],
      ['Total Standard Hours:', totals.standardHours],
      ['Total Overtime Hours:', totals.overtimeHours],
      ['Total Gross Pay:', totals.grossPay],
      ['Total UIF:', totals.uif],
      ['Total Deductions:', totals.otherDeductions + totals.loanDeductions],
      ['Total Net Pay:', totals.netPay],
      ['Total Paid to Accounts:', totals.paidToAccounts]
    ]);

    summarySheet.getRange('A5:A12').setFontWeight('bold');
    // Format hours as numbers (not currency)
    summarySheet.getRange('B6:B7').setNumberFormat('0.00');
    // Format currency values
    summarySheet.getRange('B8:B12').setNumberFormat('"R"#,##0.00');

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

      // Format only currency values (Gross Pay, Net Pay, Paid to Accounts), skip Employees count
      summarySheet.getRange(summaryRow + 1, 2, 3, 1).setNumberFormat('"R"#,##0.00');
      summaryRow += 5;
    }

    // Auto-resize
    summarySheet.autoResizeColumns(1, 2);

    // ===== TAB 3: Payslip Received Register =====
    const receivedSheet = spreadsheet.insertSheet('Payslip Received Register');

    // Header - set text first (merge will happen after auto-resize)
    receivedSheet.getRange('A1').setValue('Payslip Received Register - Week Ending: ' + formatDate(weekEnd));
    receivedSheet.setRowHeight(1, 30); // Header row height

    // Column headers
    receivedSheet.getRange('A3:E3').setValues([[
      'Weekending',
      'Employer',
      'Employee Name',
      'Record Number',
      'Signature'
    ]]);
    receivedSheet.getRange('A3:E3').setFontWeight('bold').setBackground('#4CAF50').setFontColor('#FFFFFF').setHorizontalAlignment('center');
    receivedSheet.setRowHeight(3, 25); // Column header row height

    // Data rows
    let receivedRowNum = 4;
    for (let i = 0; i < payslips.length; i++) {
      const p = payslips[i];
      receivedSheet.getRange(receivedRowNum, 1, 1, 5).setValues([[
        formatDate(weekEnd),
        p['EMPLOYER'],
        p['EMPLOYEE NAME'],
        p['RECORDNUMBER'],
        '' // Blank for signature
      ]]);
      // Set row height to 50px to accommodate signatures
      receivedSheet.setRowHeight(receivedRowNum, 50);
      receivedRowNum++;
    }

    // Set hardcoded column widths for optimal A4 portrait layout
    receivedSheet.setColumnWidth(1, 85);   // Column A - Weekending
    receivedSheet.setColumnWidth(2, 135);  // Column B - Employer
    receivedSheet.setColumnWidth(3, 220);  // Column C - Employee Name
    receivedSheet.setColumnWidth(4, 110);  // Column D - Record Number
    receivedSheet.setColumnWidth(5, 220);  // Column E - Signature

    // Merge header cells A1:E1 and apply formatting
    receivedSheet.getRange('A1:E1').merge();
    receivedSheet.getRange('A1').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

    // Add borders to the entire table (from column headers through last data row)
    const lastRow = receivedRowNum - 1;
    const tableRange = receivedSheet.getRange('A3:E' + lastRow);
    tableRange.setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);

    // NOTE: For optimal A4 portrait printing, manually configure in Google Sheets:
    // File > Print > Set page orientation to Portrait, paper size to A4,
    // margins to Narrow, and scale to "Fit to width" = 1 page

    // Move to reports folder and set sharing
    const reportUrl = moveToReportsFolder(spreadsheet);

    // DEBUG: Log final totals before returning
    Logger.log('🔍 DEBUG: Final totals before return:');
    Logger.log('🔍   Total Employees: ' + totals.employees);
    Logger.log('🔍   Total Gross Pay: ' + totals.grossPay);
    Logger.log('🔍   Total Net Pay: ' + totals.netPay);
    Logger.log('🔍   Total Paid to Accounts: ' + totals.paidToAccounts);

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

// ==================== MONTHLY PAYROLL SUMMARY REPORT ====================

/**
 * Generates Monthly Payroll Summary Report for a specific month
 * Date range: First Friday to Last Friday of the month
 *
 * @param {Date|string} monthDate - Any date in the month to report on
 * @returns {Object} Result with success flag and report URL
 *
 * @example
 * const result = generateMonthlyPayrollSummaryReport('2025-11-15');
 * // Returns report for November 2025 (First Friday to Last Friday)
 */
function generateMonthlyPayrollSummaryReport(monthDate) {
  try {
    Logger.log('\n========== GENERATE MONTHLY PAYROLL SUMMARY REPORT ==========');

    const inputDate = parseDate(monthDate);
    Logger.log('Month: ' + (inputDate.getMonth() + 1) + '/' + inputDate.getFullYear());

    // Calculate first and last Friday of the month
    const firstFriday = getFirstFridayOfMonth(inputDate);
    const lastFriday = getLastFridayOfMonth(inputDate);

    Logger.log('Period: ' + formatDate(firstFriday) + ' to ' + formatDate(lastFriday));

    // Get all payslips for this date range
    const payslipsResult = listPayslips({
      dateRange: {
        start: firstFriday,
        end: lastFriday
      }
    });

    if (!payslipsResult.success) {
      throw new Error('Failed to get payslips: ' + payslipsResult.error);
    }

    const payslips = payslipsResult.data.payslips;
    Logger.log('📊 Processing ' + payslips.length + ' payslips');

    if (payslips.length === 0) {
      throw new Error('No payslips found for month ' + formatDate(inputDate));
    }

    // Group payslips by employee
    const employeeData = {};

    for (let i = 0; i < payslips.length; i++) {
      const p = payslips[i];
      const empName = p['EMPLOYEE NAME'];

      if (!employeeData[empName]) {
        employeeData[empName] = {
          employeeName: empName,
          employer: p['EMPLOYER'],
          employmentStatus: p['EMPLOYMENT STATUS'] || '',
          hourlyRate: p['HOURLYRATE'] || 0,
          standardHours: 0,
          overtimeHours: 0,
          grossPay: 0,
          leavePay: 0,
          bonusPay: 0,
          otherIncome: 0,
          otherIncomeNotes: [],
          uif: 0,
          otherDeductions: 0,
          netPay: 0,
          weeksWorked: 0
        };
      }

      const totalHours = (p['HOURS'] || 0) + ((p['MINUTES'] || 0) / 60);
      const totalOvertimeHours = (p['OVERTIMEHOURS'] || 0) + ((p['OVERTIMEMINUTES'] || 0) / 60);

      employeeData[empName].standardHours += totalHours;
      employeeData[empName].overtimeHours += totalOvertimeHours;
      employeeData[empName].grossPay += p['GROSSSALARY'] || 0;
      employeeData[empName].leavePay += p['LEAVE PAY'] || 0;
      employeeData[empName].bonusPay += p['BONUS PAY'] || 0;
      employeeData[empName].otherIncome += p['OTHERINCOME'] || 0;

      // Collect Other Income descriptions if present
      if ((p['OTHERINCOME'] || 0) > 0 && p['OTHER INCOME TEXT']) {
        employeeData[empName].otherIncomeNotes.push(p['OTHER INCOME TEXT']);
      }

      employeeData[empName].uif += p['UIF'] || 0;
      employeeData[empName].otherDeductions += p['OTHER DEDUCTIONS'] || 0;
      employeeData[empName].netPay += p['NETTSALARY'] || 0;
      employeeData[empName].weeksWorked++;
    }

    // Convert to array and sort by employee name
    const employeeArray = Object.values(employeeData);
    employeeArray.sort((a, b) => a.employeeName.localeCompare(b.employeeName));

    Logger.log('👥 Unique employees: ' + employeeArray.length);

    // Calculate totals
    let totals = {
      employees: employeeArray.length,
      standardHours: 0,
      overtimeHours: 0,
      grossPay: 0,
      leavePay: 0,
      bonusPay: 0,
      otherIncome: 0,
      uif: 0,
      otherDeductions: 0,
      netPay: 0
    };

    const byEmployer = {};

    for (let i = 0; i < employeeArray.length; i++) {
      const emp = employeeArray[i];

      totals.standardHours += emp.standardHours;
      totals.overtimeHours += emp.overtimeHours;
      totals.grossPay += emp.grossPay;
      totals.leavePay += emp.leavePay;
      totals.bonusPay += emp.bonusPay;
      totals.otherIncome += emp.otherIncome;
      totals.uif += emp.uif;
      totals.otherDeductions += emp.otherDeductions;
      totals.netPay += emp.netPay;

      // Track by employer
      const employer = emp.employer;
      if (!byEmployer[employer]) {
        byEmployer[employer] = {
          count: 0,
          grossPay: 0,
          netPay: 0
        };
      }

      byEmployer[employer].count++;
      byEmployer[employer].grossPay += emp.grossPay;
      byEmployer[employer].netPay += emp.netPay;
    }

    // Create Google Sheet
    const monthName = ['January', 'February', 'March', 'April', 'May', 'June',
                       'July', 'August', 'September', 'October', 'November', 'December'][inputDate.getMonth()];
    const fileName = 'Monthly Payroll Summary - ' + monthName + ' ' + inputDate.getFullYear();
    const spreadsheet = SpreadsheetApp.create(fileName);

    // ===== TAB 1: Payroll Register =====
    const registerSheet = spreadsheet.getActiveSheet();
    registerSheet.setName('Payroll Register');

    // Header
    registerSheet.getRange('A1').setValue('MONTHLY PAYROLL REGISTER - ' + monthName + ' ' + inputDate.getFullYear());
    registerSheet.setRowHeight(1, 30);

    // Period row
    registerSheet.getRange('A2').setValue('Period: ' + formatDate(firstFriday) + ' to ' + formatDate(lastFriday));
    registerSheet.getRange('A2').setFontStyle('italic');

    // Column headers
    registerSheet.getRange('A4:N4').setValues([[
      'Employee',
      'Employer',
      'Employment Status',
      'Hourly Rate',
      'Std Hours',
      'OT Hours',
      'Gross Pay',
      'Leave Pay',
      'Bonus Pay',
      'Other Income',
      'Other Income Notes',
      'UIF',
      'Other Ded.',
      'Net Pay'
    ]]);
    registerSheet.getRange('A4:N4').setFontWeight('bold').setBackground('#4CAF50').setFontColor('#FFFFFF').setHorizontalAlignment('center');
    registerSheet.setRowHeight(4, 25);

    // Data rows
    let rowNum = 5;
    for (let i = 0; i < employeeArray.length; i++) {
      const emp = employeeArray[i];

      // Join other income notes with commas
      const otherIncomeNotesText = emp.otherIncomeNotes.join(', ');

      // Convert hours to HH:MM format
      const stdHoursFormatted = decimalHoursToTime(emp.standardHours);
      const otHoursFormatted = decimalHoursToTime(emp.overtimeHours);

      registerSheet.getRange(rowNum, 1, 1, 14).setValues([[
        emp.employeeName,
        emp.employer,
        emp.employmentStatus,
        emp.hourlyRate,
        stdHoursFormatted,
        otHoursFormatted,
        emp.grossPay,
        emp.leavePay,
        emp.bonusPay,
        emp.otherIncome,
        otherIncomeNotesText,
        emp.uif,
        emp.otherDeductions,
        emp.netPay
      ]]);
      rowNum++;
    }

    // Totals row
    const totalStdHoursFormatted = decimalHoursToTime(totals.standardHours);
    const totalOtHoursFormatted = decimalHoursToTime(totals.overtimeHours);

    registerSheet.getRange(rowNum, 1, 1, 14).setValues([[
      'TOTALS',
      totals.employees + ' employees',
      '',  // No total for employment status
      '',  // No total for hourly rate
      totalStdHoursFormatted,
      totalOtHoursFormatted,
      totals.grossPay,
      totals.leavePay,
      totals.bonusPay,
      totals.otherIncome,
      '',  // No total for notes column
      totals.uif,
      totals.otherDeductions,
      totals.netPay
    ]]);
    registerSheet.getRange(rowNum, 1, 1, 14).setFontWeight('bold').setBackground('#FFD700');

    // Format currency and number columns
    registerSheet.getRange(5, 4, employeeArray.length, 1).setNumberFormat('"R"#,##0.00');  // Column D - Hourly Rate (not including totals row)
    // Columns E and F (Std Hours and OT Hours) are now text format (HH:MM), no number formatting needed
    registerSheet.getRange(5, 7, employeeArray.length + 1, 4).setNumberFormat('"R"#,##0.00');  // Currency columns G-J (Gross to Other Income)
    registerSheet.getRange(5, 12, employeeArray.length + 1, 3).setNumberFormat('"R"#,##0.00');  // Currency columns L-N (UIF to Net Pay)

    // Set hardcoded column widths for optimal layout
    registerSheet.setColumnWidth(1, 220);  // Column A - Employee
    registerSheet.setColumnWidth(2, 135);  // Column B - Employer
    registerSheet.setColumnWidth(3, 110);  // Column C - Employment Status
    registerSheet.setColumnWidth(4, 90);   // Column D - Hourly Rate
    registerSheet.setColumnWidth(5, 80);   // Column E - Std Hours
    registerSheet.setColumnWidth(6, 80);   // Column F - OT Hours
    registerSheet.setColumnWidth(7, 90);   // Column G - Gross Pay
    registerSheet.setColumnWidth(8, 90);   // Column H - Leave Pay
    registerSheet.setColumnWidth(9, 90);   // Column I - Bonus Pay
    registerSheet.setColumnWidth(10, 100); // Column J - Other Income
    registerSheet.setColumnWidth(11, 150); // Column K - Other Income Notes
    registerSheet.setColumnWidth(12, 80);  // Column L - UIF
    registerSheet.setColumnWidth(13, 110); // Column M - Other Ded.
    registerSheet.setColumnWidth(14, 110); // Column N - Net Pay

    // Merge header cells A1:N1 and apply formatting
    registerSheet.getRange('A1:N1').merge();
    registerSheet.getRange('A1').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

    // Add borders to the entire table (from column headers through last data row)
    const lastRegisterRow = rowNum;
    const registerTableRange = registerSheet.getRange('A4:N' + lastRegisterRow);
    registerTableRange.setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);

    // ===== TAB 2: Summary =====
    const summarySheet = spreadsheet.insertSheet('Summary');

    // Header
    summarySheet.getRange('A1:B1').setValues([['MONTHLY PAYROLL SUMMARY', '']]);
    summarySheet.getRange('A1:B1').setFontWeight('bold').setFontSize(14);

    summarySheet.getRange('A2:B2').setValues([['Month:', monthName + ' ' + inputDate.getFullYear()]]);
    summarySheet.getRange('A3:B3').setValues([['Period:', formatDate(firstFriday) + ' to ' + formatDate(lastFriday)]]);

    // Overall summary
    summarySheet.getRange('A5:B5').setValues([['OVERALL SUMMARY', '']]);
    summarySheet.getRange('A5:B5').setFontWeight('bold').setBackground('#4CAF50').setFontColor('#FFFFFF');

    summarySheet.getRange('A6:B14').setValues([
      ['Total Employees:', totals.employees],
      ['Total Standard Hours:', totals.standardHours],
      ['Total Overtime Hours:', totals.overtimeHours],
      ['Total Gross Pay:', totals.grossPay],
      ['Total Leave Pay:', totals.leavePay],
      ['Total Bonus Pay:', totals.bonusPay],
      ['Total Other Income:', totals.otherIncome],
      ['Total UIF:', totals.uif],
      ['Total Deductions:', totals.otherDeductions]
    ]);

    summarySheet.getRange('A6:A14').setFontWeight('bold');
    // Format hours as numbers (not currency)
    summarySheet.getRange('B7:B8').setNumberFormat('0.00');
    // Format currency values
    summarySheet.getRange('B9:B14').setNumberFormat('"R"#,##0.00');

    // By employer
    summarySheet.getRange('A16:B16').setValues([['BY EMPLOYER', '']]);
    summarySheet.getRange('A16:B16').setFontWeight('bold').setBackground('#4CAF50').setFontColor('#FFFFFF');

    let summaryRow = 17;
    for (const employer in byEmployer) {
      const data = byEmployer[employer];
      summarySheet.getRange(summaryRow, 1, 1, 2).setValues([[employer, '']]);
      summarySheet.getRange(summaryRow, 1).setFontWeight('bold');
      summaryRow++;

      summarySheet.getRange(summaryRow, 1, 3, 2).setValues([
        ['  Employees:', data.count],
        ['  Gross Pay:', data.grossPay],
        ['  Net Pay:', data.netPay]
      ]);

      // Format only currency values (Gross Pay and Net Pay), skip Employees count
      summarySheet.getRange(summaryRow + 1, 2, 2, 1).setNumberFormat('"R"#,##0.00');
      summaryRow += 4;
    }

    // Auto-resize
    summarySheet.autoResizeColumns(1, 2);

    // Move to reports folder and set sharing
    const reportUrl = moveToReportsFolder(spreadsheet);

    Logger.log('✅ Report generated: ' + fileName);
    Logger.log('📎 URL: ' + reportUrl);
    Logger.log('========== GENERATE MONTHLY PAYROLL SUMMARY REPORT COMPLETE ==========\n');

    return {
      success: true,
      data: {
        url: reportUrl,
        fileName: fileName,
        month: monthName + ' ' + inputDate.getFullYear(),
        periodStart: formatDate(firstFriday),
        periodEnd: formatDate(lastFriday),
        totalEmployees: totals.employees,
        totalNetPay: totals.netPay
      }
    };

  } catch (error) {
    Logger.log('❌ ERROR in generateMonthlyPayrollSummaryReport: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== GENERATE MONTHLY PAYROLL SUMMARY REPORT FAILED ==========\n');
    return { success: false, error: error.message };
  }
}

// ==================== EMPLOYEE LEAVE TRENDS REPORT ====================

/**
 * Generates Employee Leave Trends Report to identify leave patterns
 *
 * @param {Date|string} [startDate] - Report start date (default: 1 year ago)
 * @param {Date|string} [endDate] - Report end date (default: today)
 * @param {string} [employeeName] - Optional: filter by employee name (default: all employees)
 * @returns {Object} Result with success flag and report URL
 *
 * @example
 * const result = generateLeaveTrendsReport('2024-01-01', '2024-12-31', 'John Doe');
 * // Returns report analyzing leave patterns for John Doe in 2024
 */
function generateLeaveTrendsReport(startDate, endDate, employeeName) {
  try {
    Logger.log('\n========== GENERATE EMPLOYEE LEAVE TRENDS REPORT ==========');

    // Default to past year if no dates provided
    const end = endDate ? parseDate(endDate) : new Date();
    const start = startDate ? parseDate(startDate) : new Date(end.getFullYear() - 1, end.getMonth(), end.getDate());

    Logger.log('Period: ' + formatDate(start) + ' to ' + formatDate(end));
    if (employeeName) {
      Logger.log('Employee filter: ' + employeeName);
    } else {
      Logger.log('Analyzing all employees');
    }

    // Get leave data
    const filters = {
      startDate: start,
      endDate: end
    };
    if (employeeName) {
      filters.employeeName = employeeName;
    }

    const leaveResult = listLeave(filters);
    if (!leaveResult.success) {
      throw new Error('Failed to get leave data: ' + leaveResult.error);
    }

    const leaveRecords = leaveResult.data;
    Logger.log('📊 Processing ' + leaveRecords.length + ' leave records');

    if (leaveRecords.length === 0) {
      throw new Error('No leave records found for the specified period');
    }

    // Initialize analysis structures
    const employeeAnalysis = {};

    // Process each leave record
    for (let i = 0; i < leaveRecords.length; i++) {
      const record = leaveRecords[i];
      const empName = record['EMPLOYEE NAME'];
      const startLeave = parseDate(record['STARTDATE.LEAVE']);
      const returnLeave = parseDate(record['RETURNDATE.LEAVE']);
      const totalDays = record['TOTALDAYS.LEAVE'] || 1;
      const reason = record['REASON'] || 'Unknown';
      const weekDay = record['WEEK.DAY'] || getDayOfWeek(startLeave);

      // Initialize employee data if not exists
      if (!employeeAnalysis[empName]) {
        employeeAnalysis[empName] = {
          employeeName: empName,
          totalLeaves: 0,
          totalDays: 0,
          byDayOfWeek: { Monday: 0, Tuesday: 0, Wednesday: 0, Thursday: 0, Friday: 0, Saturday: 0, Sunday: 0 },
          byReason: {},
          singleDayLeaves: 0,
          multiDayLeaves: 0,
          lastWeekOfMonth: 0,
          restOfMonth: 0,
          firstMondayOfMonth: 0,
          mondayAfter25th: 0,
          fridayLeaves: 0,
          mondayLeaves: 0,
          adjacentToWeekend: 0
        };
      }

      const empData = employeeAnalysis[empName];

      // Total counts
      empData.totalLeaves++;
      empData.totalDays += totalDays;

      // Day of week analysis
      if (empData.byDayOfWeek[weekDay] !== undefined) {
        empData.byDayOfWeek[weekDay]++;
      }

      // Leave reason
      if (!empData.byReason[reason]) {
        empData.byReason[reason] = 0;
      }
      empData.byReason[reason]++;

      // Single vs Multi-day
      if (totalDays === 1) {
        empData.singleDayLeaves++;
      } else {
        empData.multiDayLeaves++;
      }

      // Month-end proximity (last 7 days of month)
      const daysInMonth = new Date(startLeave.getFullYear(), startLeave.getMonth() + 1, 0).getDate();
      const dayOfMonth = startLeave.getDate();
      if (dayOfMonth > daysInMonth - 7) {
        empData.lastWeekOfMonth++;
      } else {
        empData.restOfMonth++;
      }

      // First Monday of month pattern
      if (weekDay === 'Monday' && dayOfMonth <= 7) {
        empData.firstMondayOfMonth++;
      }

      // Monday after 25th pattern
      if (weekDay === 'Monday' && dayOfMonth > 25) {
        empData.mondayAfter25th++;
      }

      // Friday leaves (potential weekend extension)
      if (weekDay === 'Friday') {
        empData.fridayLeaves++;
      }

      // Monday leaves (potential weekend extension)
      if (weekDay === 'Monday') {
        empData.mondayLeaves++;
      }

      // Adjacent to weekend (Friday or Monday)
      if (weekDay === 'Friday' || weekDay === 'Monday') {
        empData.adjacentToWeekend++;
      }
    }

    // Convert to array and sort by total leaves (descending)
    const employeeArray = Object.values(employeeAnalysis);
    employeeArray.sort((a, b) => b.totalLeaves - a.totalLeaves);

    Logger.log('👥 Analyzed ' + employeeArray.length + ' employees');

    // Create Google Sheet
    const fileName = 'Employee Leave Trends Report - ' + formatDate(start) + ' to ' + formatDate(end) +
                    (employeeName ? ' - ' + employeeName : '');
    const spreadsheet = SpreadsheetApp.create(fileName);

    // ===== TAB 1: Leave Summary =====
    const summarySheet = spreadsheet.getActiveSheet();
    summarySheet.setName('Leave Summary');

    // Header
    summarySheet.getRange('A1').setValue('EMPLOYEE LEAVE TRENDS REPORT');
    summarySheet.setRowHeight(1, 30);

    // Period info
    summarySheet.getRange('A2').setValue('Period: ' + formatDate(start) + ' to ' + formatDate(end));
    if (employeeName) {
      summarySheet.getRange('A3').setValue('Employee: ' + employeeName);
    }
    summarySheet.getRange('A2:A3').setFontStyle('italic');

    // Column headers
    const summaryHeaders = [
      'Employee',
      'Total Leaves',
      'Total Days',
      'Single-Day',
      'Multi-Day',
      'Mon',
      'Tue',
      'Wed',
      'Thu',
      'Fri',
      'Sat',
      'Sun',
      'Last Week of Month',
      'Rest of Month',
      'First Mon of Month',
      'Mon after 25th'
    ];

    const headerRow = employeeName ? 4 : 5;
    summarySheet.getRange(headerRow, 1, 1, summaryHeaders.length).setValues([summaryHeaders]);
    summarySheet.getRange(headerRow, 1, 1, summaryHeaders.length)
      .setFontWeight('bold')
      .setBackground('#4CAF50')
      .setFontColor('#FFFFFF')
      .setHorizontalAlignment('center')
      .setWrap(true);
    summarySheet.setRowHeight(headerRow, 35);

    // Data rows
    let rowNum = headerRow + 1;
    for (let i = 0; i < employeeArray.length; i++) {
      const emp = employeeArray[i];

      summarySheet.getRange(rowNum, 1, 1, summaryHeaders.length).setValues([[
        emp.employeeName,
        emp.totalLeaves,
        emp.totalDays,
        emp.singleDayLeaves,
        emp.multiDayLeaves,
        emp.byDayOfWeek.Monday,
        emp.byDayOfWeek.Tuesday,
        emp.byDayOfWeek.Wednesday,
        emp.byDayOfWeek.Thursday,
        emp.byDayOfWeek.Friday,
        emp.byDayOfWeek.Saturday,
        emp.byDayOfWeek.Sunday,
        emp.lastWeekOfMonth,
        emp.restOfMonth,
        emp.firstMondayOfMonth,
        emp.mondayAfter25th
      ]]);
      rowNum++;
    }

    // Set column widths
    summarySheet.setColumnWidth(1, 220);  // Employee name
    for (let col = 2; col <= summaryHeaders.length; col++) {
      summarySheet.setColumnWidth(col, 90);
    }

    // Merge and format header
    summarySheet.getRange('A1:P1').merge();
    summarySheet.getRange('A1')
      .setFontWeight('bold')
      .setFontSize(14)
      .setHorizontalAlignment('center');

    // Add borders
    const lastSummaryRow = rowNum - 1;
    summarySheet.getRange(headerRow, 1, lastSummaryRow - headerRow + 1, summaryHeaders.length)
      .setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);

    // ===== TAB 2: Pattern Analysis =====
    const patternSheet = spreadsheet.insertSheet('Pattern Analysis');

    // Header
    patternSheet.getRange('A1').setValue('LEAVE PATTERN INDICATORS');
    patternSheet.getRange('A1')
      .setFontWeight('bold')
      .setFontSize(14)
      .setHorizontalAlignment('center');
    patternSheet.setRowHeight(1, 30);

    // Period info
    patternSheet.getRange('A2').setValue('Period: ' + formatDate(start) + ' to ' + formatDate(end));
    patternSheet.getRange('A2').setFontStyle('italic');

    // Column headers
    const patternHeaders = [
      'Employee',
      'Total Leaves',
      'Weekend Extension %',
      'Mon/Fri Leaves',
      'First Mon Pattern',
      'Mon after 25th',
      'Single-Day %',
      'Month-End %',
      'Top Leave Day',
      'Risk Score'
    ];

    patternSheet.getRange('A4:J4').setValues([patternHeaders]);
    patternSheet.getRange('A4:J4')
      .setFontWeight('bold')
      .setBackground('#FF9800')
      .setFontColor('#FFFFFF')
      .setHorizontalAlignment('center')
      .setWrap(true);
    patternSheet.setRowHeight(4, 35);

    // Calculate pattern indicators
    rowNum = 5;
    for (let i = 0; i < employeeArray.length; i++) {
      const emp = employeeArray[i];

      // Percentages
      const weekendExtensionPct = emp.totalLeaves > 0 ? (emp.adjacentToWeekend / emp.totalLeaves * 100).toFixed(1) : 0;
      const singleDayPct = emp.totalLeaves > 0 ? (emp.singleDayLeaves / emp.totalLeaves * 100).toFixed(1) : 0;
      const monthEndPct = emp.totalLeaves > 0 ? (emp.lastWeekOfMonth / emp.totalLeaves * 100).toFixed(1) : 0;

      // Find top leave day
      let topDay = '';
      let maxCount = 0;
      for (const day in emp.byDayOfWeek) {
        if (emp.byDayOfWeek[day] > maxCount) {
          maxCount = emp.byDayOfWeek[day];
          topDay = day;
        }
      }

      // Calculate risk score (0-100)
      let riskScore = 0;
      if (emp.totalLeaves > 0) {
        riskScore += (emp.adjacentToWeekend / emp.totalLeaves) * 40;  // 40 points for weekend extension
        riskScore += (emp.singleDayLeaves / emp.totalLeaves) * 30;    // 30 points for single-day pattern
        riskScore += (emp.lastWeekOfMonth / emp.totalLeaves) * 20;    // 20 points for month-end
        riskScore += (emp.firstMondayOfMonth / emp.totalLeaves) * 10; // 10 points for first Monday
      }
      riskScore = Math.round(riskScore);

      patternSheet.getRange(rowNum, 1, 1, 10).setValues([[
        emp.employeeName,
        emp.totalLeaves,
        weekendExtensionPct + '%',
        emp.adjacentToWeekend + ' (' + emp.mondayLeaves + 'M/' + emp.fridayLeaves + 'F)',
        emp.firstMondayOfMonth,
        emp.mondayAfter25th,
        singleDayPct + '%',
        monthEndPct + '%',
        topDay + ' (' + maxCount + ')',
        riskScore
      ]]);

      // Color-code risk score
      const riskCell = patternSheet.getRange(rowNum, 10);
      if (riskScore >= 70) {
        riskCell.setBackground('#FF5252').setFontColor('#FFFFFF'); // High risk - Red
      } else if (riskScore >= 40) {
        riskCell.setBackground('#FFC107'); // Medium risk - Orange
      } else {
        riskCell.setBackground('#4CAF50').setFontColor('#FFFFFF'); // Low risk - Green
      }

      rowNum++;
    }

    // Set column widths
    patternSheet.setColumnWidth(1, 220);  // Employee name
    for (let col = 2; col <= 10; col++) {
      patternSheet.setColumnWidth(col, 110);
    }

    // Merge header
    patternSheet.getRange('A1:J1').merge();

    // Add borders
    const lastPatternRow = rowNum - 1;
    patternSheet.getRange('A4:J' + lastPatternRow)
      .setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);

    // ===== TAB 3: Leave by Reason =====
    const reasonSheet = spreadsheet.insertSheet('Leave by Reason');

    // Header
    reasonSheet.getRange('A1').setValue('LEAVE ANALYSIS BY REASON');
    reasonSheet.getRange('A1')
      .setFontWeight('bold')
      .setFontSize(14)
      .setHorizontalAlignment('center');
    reasonSheet.setRowHeight(1, 30);

    // Get all unique reasons
    const allReasons = new Set();
    for (let i = 0; i < employeeArray.length; i++) {
      for (const reason in employeeArray[i].byReason) {
        allReasons.add(reason);
      }
    }
    const reasonArray = Array.from(allReasons).sort();

    // Column headers
    const reasonHeaders = ['Employee', 'Total Leaves'].concat(reasonArray);
    reasonSheet.getRange('A3:' + String.fromCharCode(65 + reasonHeaders.length - 1) + '3').setValues([reasonHeaders]);
    reasonSheet.getRange('A3:' + String.fromCharCode(65 + reasonHeaders.length - 1) + '3')
      .setFontWeight('bold')
      .setBackground('#2196F3')
      .setFontColor('#FFFFFF')
      .setHorizontalAlignment('center');
    reasonSheet.setRowHeight(3, 25);

    // Data rows
    rowNum = 4;
    for (let i = 0; i < employeeArray.length; i++) {
      const emp = employeeArray[i];
      const rowData = [emp.employeeName, emp.totalLeaves];

      for (let j = 0; j < reasonArray.length; j++) {
        const reason = reasonArray[j];
        rowData.push(emp.byReason[reason] || 0);
      }

      reasonSheet.getRange(rowNum, 1, 1, rowData.length).setValues([rowData]);
      rowNum++;
    }

    // Set column widths
    reasonSheet.setColumnWidth(1, 220);
    for (let col = 2; col <= reasonHeaders.length; col++) {
      reasonSheet.setColumnWidth(col, 110);
    }

    // Merge header
    reasonSheet.getRange('A1:' + String.fromCharCode(65 + reasonHeaders.length - 1) + '1').merge();

    // Add borders
    const lastReasonRow = rowNum - 1;
    reasonSheet.getRange('A3:' + String.fromCharCode(65 + reasonHeaders.length - 1) + lastReasonRow)
      .setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);

    // ===== TAB 4: Help & Interpretation =====
    const helpSheet = spreadsheet.insertSheet('How to Use');

    helpSheet.getRange('A1').setValue('HOW TO INTERPRET THIS REPORT');
    helpSheet.getRange('A1')
      .setFontWeight('bold')
      .setFontSize(14)
      .setHorizontalAlignment('center');

    helpSheet.getRange('A3').setValue('Risk Score Interpretation:');
    helpSheet.getRange('A3').setFontWeight('bold').setFontSize(12);

    helpSheet.getRange('A4:B7').setValues([
      ['Score Range', 'Interpretation'],
      ['0-39 (Green)', 'Low risk - Normal leave patterns'],
      ['40-69 (Orange)', 'Medium risk - Some patterns detected, worth monitoring'],
      ['70-100 (Red)', 'High risk - Strong patterns suggesting potential leave abuse']
    ]);
    helpSheet.getRange('A4:B4').setFontWeight('bold').setBackground('#E0E0E0');

    helpSheet.getRange('A9').setValue('Pattern Indicators:');
    helpSheet.getRange('A9').setFontWeight('bold').setFontSize(12);

    helpSheet.getRange('A10:B16').setValues([
      ['Indicator', 'What to Look For'],
      ['Weekend Extension %', 'High percentage of Monday/Friday leaves may indicate weekend extensions'],
      ['Single-Day %', 'High percentage of single-day leaves may indicate pattern abuse'],
      ['Month-End %', 'Leaves concentrated in last week of month'],
      ['First Mon of Month', 'Pattern of taking leave on first Monday of new month'],
      ['Mon after 25th', 'Pattern of taking Monday leaves after 25th of month'],
      ['Top Leave Day', 'Most frequent day for taking leave']
    ]);
    helpSheet.getRange('A10:B10').setFontWeight('bold').setBackground('#E0E0E0');

    helpSheet.getRange('A18').setValue('Tips:');
    helpSheet.getRange('A18').setFontWeight('bold').setFontSize(12);

    helpSheet.getRange('A19:A23').setValues([
      ['• Look for employees with high risk scores (70+) and investigate further'],
      ['• Check if weekend extension patterns correlate with single-day leaves'],
      ['• Compare leave patterns across different leave types (Annual vs Sick vs AWOL)'],
      ['• Consider total number of leaves - patterns are more concerning with frequent leaves'],
      ['• Use this as an initial screening tool, not definitive proof of abuse']
    ]);

    // Set column widths
    helpSheet.setColumnWidth(1, 200);
    helpSheet.setColumnWidth(2, 500);

    // Merge header
    helpSheet.getRange('A1:B1').merge();

    // Add borders to tables
    helpSheet.getRange('A4:B7').setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);
    helpSheet.getRange('A10:B16').setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);

    // Move to reports folder and set sharing
    const reportUrl = moveToReportsFolder(spreadsheet);

    Logger.log('✅ Report generated: ' + fileName);
    Logger.log('📎 URL: ' + reportUrl);
    Logger.log('========== GENERATE EMPLOYEE LEAVE TRENDS REPORT COMPLETE ==========\n');

    return {
      success: true,
      data: {
        url: reportUrl,
        fileName: fileName,
        periodStart: formatDate(start),
        periodEnd: formatDate(end),
        totalEmployees: employeeArray.length,
        totalLeaveRecords: leaveRecords.length
      }
    };

  } catch (error) {
    Logger.log('❌ ERROR in generateLeaveTrendsReport: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== GENERATE EMPLOYEE LEAVE TRENDS REPORT FAILED ==========\n');
    return { success: false, error: error.message };
  }
}

// ==================== HELPER FUNCTIONS ====================

/**
 * Converts decimal hours to HH:MM format
 *
 * @param {number} decimalHours - Hours in decimal format (e.g., 47.75)
 * @returns {string} Time in HH:MM format (e.g., "47:45")
 */
function decimalHoursToTime(decimalHours) {
  if (decimalHours === 0 || decimalHours === null || decimalHours === undefined) {
    return '0:00';
  }

  const hours = Math.floor(decimalHours);
  const minutes = Math.round((decimalHours - hours) * 60);

  // Pad minutes with leading zero if needed
  const minutesStr = minutes < 10 ? '0' + minutes : '' + minutes;

  return hours + ':' + minutesStr;
}

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

/**
 * Test function for Monthly Payroll Summary Report
 */
function test_monthlyPayrollReport() {
  Logger.log('\n========== TEST: MONTHLY PAYROLL SUMMARY REPORT ==========');

  // Use current month (or specify a month)
  const monthDate = new Date('2025-11-15');

  const result = generateMonthlyPayrollSummaryReport(monthDate);

  if (result.success) {
    Logger.log('✅ TEST PASSED');
    Logger.log('URL: ' + result.data.url);
    Logger.log('Month: ' + result.data.month);
    Logger.log('Period: ' + result.data.periodStart + ' to ' + result.data.periodEnd);
    Logger.log('Total Employees: ' + result.data.totalEmployees);
    Logger.log('Total Net Pay: R' + result.data.totalNetPay.toFixed(2));
  } else {
    Logger.log('❌ TEST FAILED: ' + result.error);
  }

  Logger.log('========== TEST COMPLETE ==========\n');
}
