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

    const payslips = payslipsResult.data;
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
    registerSheet.getRange('A4:L4').setValues([[
      'Employee',
      'Employer',
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
    registerSheet.getRange('A4:L4').setFontWeight('bold').setBackground('#4CAF50').setFontColor('#FFFFFF').setHorizontalAlignment('center');
    registerSheet.setRowHeight(4, 25);

    // Data rows
    let rowNum = 5;
    for (let i = 0; i < employeeArray.length; i++) {
      const emp = employeeArray[i];

      // Join other income notes with commas
      const otherIncomeNotesText = emp.otherIncomeNotes.join(', ');

      registerSheet.getRange(rowNum, 1, 1, 12).setValues([[
        emp.employeeName,
        emp.employer,
        emp.standardHours,
        emp.overtimeHours,
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
    registerSheet.getRange(rowNum, 1, 1, 12).setValues([[
      'TOTALS',
      totals.employees + ' employees',
      totals.standardHours,
      totals.overtimeHours,
      totals.grossPay,
      totals.leavePay,
      totals.bonusPay,
      totals.otherIncome,
      '',  // No total for notes column
      totals.uif,
      totals.otherDeductions,
      totals.netPay
    ]]);
    registerSheet.getRange(rowNum, 1, 1, 12).setFontWeight('bold').setBackground('#FFD700');

    // Format currency and number columns
    registerSheet.getRange(5, 3, employeeArray.length + 1, 1).setNumberFormat('0.00');  // Std Hours
    registerSheet.getRange(5, 4, employeeArray.length + 1, 1).setNumberFormat('0.00');  // OT Hours
    registerSheet.getRange(5, 5, employeeArray.length + 1, 4).setNumberFormat('"R"#,##0.00');  // Currency columns E-H (Gross to Other Income)
    registerSheet.getRange(5, 10, employeeArray.length + 1, 3).setNumberFormat('"R"#,##0.00');  // Currency columns J-L (UIF to Net Pay)

    // Set hardcoded column widths for optimal layout
    registerSheet.setColumnWidth(1, 220);  // Column A - Employee
    registerSheet.setColumnWidth(2, 135);  // Column B - Employer
    registerSheet.setColumnWidth(3, 80);   // Column C - Std Hours
    registerSheet.setColumnWidth(4, 80);   // Column D - OT Hours
    registerSheet.setColumnWidth(5, 90);   // Column E - Gross Pay
    registerSheet.setColumnWidth(6, 90);   // Column F - Leave Pay
    registerSheet.setColumnWidth(7, 90);   // Column G - Bonus Pay
    registerSheet.setColumnWidth(8, 100);  // Column H - Other Income
    registerSheet.setColumnWidth(9, 150);  // Column I - Other Income Notes
    registerSheet.setColumnWidth(10, 80);  // Column J - UIF
    registerSheet.setColumnWidth(11, 110); // Column K - Other Ded.
    registerSheet.setColumnWidth(12, 110); // Column L - Net Pay

    // Merge header cells A1:L1 and apply formatting
    registerSheet.getRange('A1:L1').merge();
    registerSheet.getRange('A1').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

    // Add borders to the entire table (from column headers through last data row)
    const lastRegisterRow = rowNum;
    const registerTableRange = registerSheet.getRange('A4:L' + lastRegisterRow);
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
    summarySheet.getRange('B7:B14').setNumberFormat('"R"#,##0.00');

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

      summarySheet.getRange(summaryRow, 2, 2, 1).setNumberFormat('"R"#,##0.00');
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
