/**
 * REPORTS.GS - Report Generation Module
 * HR Payroll System for SA Grinding Wheels & Scorpio Abrasives
 *
 * Generates formatted Google Sheets reports:
 * 1. Outstanding Loans Report - All employees with loan balances > 0
 * 2. Individual Statement - Complete loan history for one employee
 * 3. Loan Transaction Matrix - All loan transactions in a period (employees × dates)
 * 4. Weekly Payroll Summary - Detailed breakdown for a specific week
 * 5. Monthly Payroll Summary - Monthly totals from first to last Friday
 * 6. Employee Leave Trends - Analyzes leave patterns to identify potential abuse
 *
 * All reports are saved to "Payroll Reports" folder in Google Drive
 * with sharing set to "Anyone with link can view"
 */

// ==================== UTILITY FUNCTIONS ====================

/**
 * Gets list of employees for populating dropdown menus
 * Reads directly from Employee Details sheet (not web API)
 *
 * @returns {Object} Result with success flag and employees array
 *
 * @example
 * const result = getEmployeesForDropdown();
 * // Returns: { success: true, data: { employees: [{id, REFNAME}] } }
 */
function getEmployeesForDropdown() {
  try {
    Logger.log('\n========== GET EMPLOYEES FOR DROPDOWN ==========');

    const sheets = getSheets();
    const empSheet = sheets.empdetails;

    if (!empSheet) {
      throw new Error('Employee Details sheet not found');
    }

    const empData = empSheet.getDataRange().getValues();
    const empHeaders = empData[0];

    // Find column indices
    const idCol = empHeaders.indexOf('id');
    const refnameCol = empHeaders.indexOf('REFNAME');

    if (idCol === -1 || refnameCol === -1) {
      throw new Error('Required columns not found (id, REFNAME)');
    }

    const employees = [];

    // Build simple employee objects for dropdown
    for (let i = 1; i < empData.length; i++) {
      const row = empData[i];
      if (row[idCol]) {  // Skip empty rows
        employees.push({
          id: row[idCol],
          REFNAME: row[refnameCol] || 'Unknown'
        });
      }
    }

    // Sort by name
    employees.sort((a, b) => a.REFNAME.localeCompare(b.REFNAME));

    Logger.log('📊 Found ' + employees.length + ' employees for dropdown');

    return {
      success: true,
      data: {
        employees: employees
      }
    };

  } catch (error) {
    Logger.log('❌ ERROR in getEmployeesForDropdown: ' + error.message);
    return { success: false, error: error.message };
  }
}

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

    // Get ALL employees directly from sheet (don't use listEmployees API - it's for web calls)
    const sheets = getSheets();
    const empSheet = sheets.empdetails;

    if (!empSheet) {
      throw new Error('Employee Details sheet not found');
    }

    const empData = empSheet.getDataRange().getValues();
    const empHeaders = empData[0];
    const employees = [];

    // Build employee objects from sheet rows
    for (let i = 1; i < empData.length; i++) {
      if (empData[i][0]) {  // Skip empty rows
        try {
          const employee = buildObjectFromRow(empData[i], empHeaders);
          employees.push(employee);
        } catch (error) {
          // Skip invalid rows
          Logger.log('⚠️ Skipping invalid employee row ' + (i + 1));
        }
      }
    }

    Logger.log('📊 Processing ' + employees.length + ' employees');

    // Get loan balances for each employee
    const reportData = [];
    let totalOutstanding = 0;

    for (let i = 0; i < employees.length; i++) {
      const emp = employees[i];
      const balanceResult = getCurrentLoanBalance(emp.id, reportDate);

      if (balanceResult.success && balanceResult.data > 0) {
        const balance = balanceResult.data;

        // Get last transaction date (up to reportDate)
        const historyResult = getLoanHistory(emp.id, null, reportDate);
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

    // Get employee details directly from sheet (don't use getEmployeeById API - it's for web calls)
    const sheets = getSheets();
    const empSheet = sheets.empdetails;

    if (!empSheet) {
      throw new Error('Employee Details sheet not found');
    }

    const empData = empSheet.getDataRange().getValues();
    const empHeaders = empData[0];
    const idCol = empHeaders.indexOf('id');

    if (idCol === -1) {
      throw new Error('ID column not found in Employee Details sheet');
    }

    let employee = null;
    for (let i = 1; i < empData.length; i++) {
      if (empData[i][idCol] === employeeId) {
        employee = buildObjectFromRow(empData[i], empHeaders);
        break;
      }
    }

    if (!employee) {
      throw new Error('Employee not found: ' + employeeId);
    }

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

// ==================== LOAN TRANSACTION MATRIX REPORT ====================

/**
 * Generates Loan Transaction Matrix Report showing all transactions in a date range
 * Rows: Employee names
 * Columns: Transaction dates
 * Values: Transaction amounts (+/- for each transaction)
 * Last Column: Total outstanding balance
 *
 * @param {Date|string} startDate - Report start date
 * @param {Date|string} endDate - Report end date
 * @returns {Object} Result with success flag and report URL
 *
 * @example
 * const result = generateLoanTransactionMatrixReport('2025-01-01', '2025-12-31');
 */
function generateLoanTransactionMatrixReport(startDate, endDate) {
  try {
    Logger.log('\n========== GENERATE LOAN TRANSACTION MATRIX REPORT ==========');

    const start = parseDate(startDate);
    const end = parseDate(endDate);

    Logger.log('Period: ' + formatDate(start) + ' to ' + formatDate(end));

    // Get all employees directly from sheet (don't use listEmployees API - it's for web calls)
    const sheets = getSheets();
    const empSheet = sheets.empdetails;

    if (!empSheet) {
      throw new Error('Employee Details sheet not found');
    }

    const empData = empSheet.getDataRange().getValues();
    const empHeaders = empData[0];
    const employees = [];

    // Build employee objects from sheet rows
    for (let i = 1; i < empData.length; i++) {
      if (empData[i][0]) {  // Skip empty rows
        try {
          const employee = buildObjectFromRow(empData[i], empHeaders);
          employees.push(employee);
        } catch (error) {
          // Skip invalid rows
          Logger.log('⚠️ Skipping invalid employee row ' + (i + 1));
        }
      }
    }

    Logger.log('📊 Processing ' + employees.length + ' employees');

    // Data structure: { employeeId: { name, employer, transactions: { date: amount }, currentBalance } }
    const employeeData = {};
    const allTransactionDates = new Set();
    let employeesWithData = 0;

    // Collect transaction data for each employee
    for (let i = 0; i < employees.length; i++) {
      const emp = employees[i];

      // Get loan history for this employee within the date range
      const historyResult = getLoanHistory(emp.id, start, end);

      if (historyResult.success && historyResult.data.length > 0) {
        const transactions = historyResult.data;

        employeeData[emp.id] = {
          name: emp.REFNAME,
          employer: emp.EMPLOYER,
          transactions: {},
          currentBalance: 0
        };

        // Process each transaction
        for (let j = 0; j < transactions.length; j++) {
          const txn = transactions[j];
          const txnDate = formatDate(parseDate(txn['TransactionDate']));
          const amount = txn['LoanAmount'] || 0;

          // Store transaction amount for this date
          if (!employeeData[emp.id].transactions[txnDate]) {
            employeeData[emp.id].transactions[txnDate] = 0;
          }
          employeeData[emp.id].transactions[txnDate] += amount;

          // Add date to set of all dates
          allTransactionDates.add(txnDate);

          // Update current balance (from last transaction's BalanceAfter)
          if (j === transactions.length - 1) {
            employeeData[emp.id].currentBalance = txn['BalanceAfter'] || 0;
          }
        }

        employeesWithData++;
      }
    }

    Logger.log('👥 Employees with transactions: ' + employeesWithData);
    Logger.log('📅 Unique transaction dates: ' + allTransactionDates.size);

    if (employeesWithData === 0) {
      throw new Error('No loan transactions found for the specified period');
    }

    // Convert dates set to sorted array
    const dateColumns = Array.from(allTransactionDates).sort(function(a, b) {
      return parseDate(a) - parseDate(b);
    });

    // Convert employee data to sorted array
    const employeeArray = Object.values(employeeData);
    employeeArray.sort(function(a, b) {
      return a.name.localeCompare(b.name);
    });

    // Create Google Sheet
    const fileName = 'Loan Transaction Matrix Report - ' + formatDate(start) + ' to ' + formatDate(end);
    const spreadsheet = SpreadsheetApp.create(fileName);
    const sheet = spreadsheet.getActiveSheet();
    sheet.setName('Loan Transactions');

    // Freeze header rows and first two columns BEFORE setting any content
    // This prevents conflicts with merged cells
    sheet.setFrozenRows(4);
    sheet.setFrozenColumns(2);

    // Header
    sheet.getRange('A1').setValue('LOAN TRANSACTION MATRIX REPORT');
    sheet.setRowHeight(1, 30);

    // Period row
    sheet.getRange('A2').setValue('Period: ' + formatDate(start) + ' to ' + formatDate(end));
    sheet.getRange('A2').setFontStyle('italic');

    // Build column headers
    const headers = ['Employee', 'Employer'];
    headers.push.apply(headers, dateColumns);
    headers.push('Current Balance');

    const numColumns = headers.length;
    const headerRange = sheet.getRange(4, 1, 1, numColumns);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold')
              .setBackground('#4CAF50')
              .setFontColor('#FFFFFF')
              .setHorizontalAlignment('center')
              .setWrap(true);
    sheet.setRowHeight(4, 35);

    // Data rows
    let rowNum = 5;
    for (let i = 0; i < employeeArray.length; i++) {
      const emp = employeeArray[i];
      const rowData = [emp.name, emp.employer];

      // Add transaction amounts for each date column
      for (let j = 0; j < dateColumns.length; j++) {
        const date = dateColumns[j];
        const amount = emp.transactions[date] || 0;
        rowData.push(amount === 0 ? '' : amount); // Show blank for zero values
      }

      // Add current balance as last column
      rowData.push(emp.currentBalance);

      sheet.getRange(rowNum, 1, 1, numColumns).setValues([rowData]);
      rowNum++;
    }

    // Format columns
    // Employee name column
    sheet.setColumnWidth(1, 220);
    // Employer column
    sheet.setColumnWidth(2, 135);

    // Transaction date columns - set width and format as currency
    for (let col = 3; col <= numColumns - 1; col++) {
      sheet.setColumnWidth(col, 110);
    }

    // Current Balance column
    sheet.setColumnWidth(numColumns, 130);

    // Format all transaction and balance columns as currency
    if (employeeArray.length > 0) {
      const dataStartRow = 5;
      const dataEndRow = rowNum - 1;
      const currencyStartCol = 3; // First transaction column
      const currencyEndCol = numColumns; // Including current balance

      sheet.getRange(dataStartRow, currencyStartCol, dataEndRow - dataStartRow + 1, currencyEndCol - currencyStartCol + 1)
           .setNumberFormat('"R"#,##0.00;[Red]"R"-#,##0.00'); // Red for negative values
    }

    // Merge only the first two columns in the header (matching frozen columns)
    // This prevents conflict with frozen columns
    sheet.getRange('A1:B1').merge();
    sheet.getRange('A1')
         .setFontWeight('bold')
         .setFontSize(14)
         .setHorizontalAlignment('center');

    // Add borders to the entire table
    const lastRow = rowNum - 1;
    sheet.getRange(4, 1, lastRow - 3, numColumns)
         .setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);

    // Add totals row
    const totalsRow = rowNum;
    sheet.getRange(totalsRow, 1).setValue('TOTALS');
    sheet.getRange(totalsRow, 2).setValue(employeeArray.length + ' employees');

    // Calculate totals for each date column
    for (let col = 3; col <= numColumns - 1; col++) {
      const colLetter = String.fromCharCode(64 + col);
      const formula = '=SUM(' + colLetter + '5:' + colLetter + (totalsRow - 1) + ')';
      sheet.getRange(totalsRow, col).setFormula(formula);
    }

    // Calculate total outstanding balance
    const balanceColLetter = String.fromCharCode(64 + numColumns);
    const balanceFormula = '=SUM(' + balanceColLetter + '5:' + balanceColLetter + (totalsRow - 1) + ')';
    sheet.getRange(totalsRow, numColumns).setFormula(balanceFormula);

    // Format totals row
    sheet.getRange(totalsRow, 1, 1, numColumns)
         .setFontWeight('bold')
         .setBackground('#FFD700');

    // Format totals row currency columns
    sheet.getRange(totalsRow, 3, 1, numColumns - 2)
         .setNumberFormat('"R"#,##0.00;[Red]"R"-#,##0.00');

    // Move to reports folder and set sharing
    const reportUrl = moveToReportsFolder(spreadsheet);

    Logger.log('✅ Report generated: ' + fileName);
    Logger.log('📎 URL: ' + reportUrl);
    Logger.log('========== GENERATE LOAN TRANSACTION MATRIX REPORT COMPLETE ==========\n');

    return {
      success: true,
      data: {
        url: reportUrl,
        fileName: fileName,
        periodStart: formatDate(start),
        periodEnd: formatDate(end),
        employeeCount: employeesWithData,
        transactionDateCount: dateColumns.length
      }
    };

  } catch (error) {
    Logger.log('❌ ERROR in generateLoanTransactionMatrixReport: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== GENERATE LOAN TRANSACTION MATRIX REPORT FAILED ==========\n');
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
      leavePay: 0,
      bonusPay: 0,
      otherIncome: 0,
      grossPay: 0,
      uif: 0,
      otherDeductions: 0,
      loanDeductions: 0,
      newLoans: 0,
      netPay: 0,
      paidToAccounts: 0
    };

    // Group payslips by employment status
    const byEmploymentStatus = {};

    for (let i = 0; i < payslips.length; i++) {
      const p = payslips[i];
      const status = p['EMPLOYMENT STATUS'] || 'Unknown';

      if (!byEmploymentStatus[status]) {
        byEmploymentStatus[status] = [];
      }
      byEmploymentStatus[status].push(p);
    }

    // Create Google Sheet
    const fileName = 'Weekly Payroll Summary - ' + formatDate(weekEnd);
    const spreadsheet = SpreadsheetApp.create(fileName);

    // ===== Payroll Register =====
    const registerSheet = spreadsheet.getActiveSheet();
    registerSheet.setName('Payroll Register');

    // Main Header
    registerSheet.getRange('A1').setValue('WEEKLY PAYROLL REGISTER - Week Ending: ' + formatDate(weekEnd));
    registerSheet.setRowHeight(1, 30);

    // Set hardcoded column widths for optimal layout
    registerSheet.setColumnWidth(1, 220);  // Column A - Employee
    registerSheet.setColumnWidth(2, 135);  // Column B - Employer
    registerSheet.setColumnWidth(3, 110);  // Column C - Employment Status
    registerSheet.setColumnWidth(4, 90);   // Column D - Hourly Rate
    registerSheet.setColumnWidth(5, 80);   // Column E - Std Hours
    registerSheet.setColumnWidth(6, 80);   // Column F - OT Hours
    registerSheet.setColumnWidth(7, 90);   // Column G - Leave Pay
    registerSheet.setColumnWidth(8, 90);   // Column H - Bonus Pay
    registerSheet.setColumnWidth(9, 100);  // Column I - Other Income
    registerSheet.setColumnWidth(10, 150); // Column J - Other Income Notes
    registerSheet.setColumnWidth(11, 90);  // Column K - Gross Pay
    registerSheet.setColumnWidth(12, 80);  // Column L - UIF
    registerSheet.setColumnWidth(13, 110); // Column M - Other Ded.
    registerSheet.setColumnWidth(14, 110); // Column N - Net Pay
    registerSheet.setColumnWidth(15, 110); // Column O - Loan Ded.
    registerSheet.setColumnWidth(16, 110); // Column P - New Loan
    registerSheet.setColumnWidth(17, 110); // Column Q - Paid to Account

    // Merge header cells A1:Q1 and apply formatting
    registerSheet.getRange('A1:Q1').merge();
    registerSheet.getRange('A1').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

    let rowNum = 3;

    // Sort employment statuses to ensure consistent order
    const statusOrder = ['Permanent', 'Temporary', 'Resigned', 'Dismissed', 'Absconded', 'Unknown'];
    const sortedStatuses = Object.keys(byEmploymentStatus).sort((a, b) => {
      const indexA = statusOrder.indexOf(a);
      const indexB = statusOrder.indexOf(b);
      if (indexA === -1 && indexB === -1) return a.localeCompare(b);
      if (indexA === -1) return 1;
      if (indexB === -1) return -1;
      return indexA - indexB;
    });

    // Create a section for each employment status
    for (const status of sortedStatuses) {
      const statusPayslips = byEmploymentStatus[status];

      // Section header
      registerSheet.getRange(rowNum, 1).setValue('Weekly Report ' + status + ' Employees');
      registerSheet.getRange(rowNum, 1, 1, 17).merge();
      registerSheet.getRange(rowNum, 1).setFontWeight('bold').setFontSize(12);
      registerSheet.setRowHeight(rowNum, 25);
      rowNum++;

      // Column headers
      registerSheet.getRange(rowNum, 1, 1, 17).setValues([[
        'Employee',
        'Employer',
        'Employment Status',
        'Hourly Rate',
        'Std Hours',
        'OT Hours',
        'Leave Pay',
        'Bonus Pay',
        'Other Income',
        'Other Income Notes',
        'Gross Pay',
        'UIF',
        'Other Ded.',
        'Net Pay',
        'Loan Ded.',
        'New Loan',
        'Paid to Account'
      ]]);
      registerSheet.getRange(rowNum, 1, 1, 17).setFontWeight('bold').setBackground('#4CAF50').setFontColor('#FFFFFF').setHorizontalAlignment('center');
      registerSheet.setRowHeight(rowNum, 25);
      const headerRow = rowNum;
      rowNum++;

      // Calculate section totals
      let sectionTotals = {
        employees: statusPayslips.length,
        standardHours: 0,
        overtimeHours: 0,
        leavePay: 0,
        bonusPay: 0,
        otherIncome: 0,
        grossPay: 0,
        uif: 0,
        otherDeductions: 0,
        loanDeductions: 0,
        newLoans: 0,
        netPay: 0,
        paidToAccounts: 0
      };

      // Data rows
      const dataStartRow = rowNum;
      for (let i = 0; i < statusPayslips.length; i++) {
        const p = statusPayslips[i];
        const stdHours = (p['HOURS'] || 0) + ((p['MINUTES'] || 0) / 60);
        const otHours = (p['OVERTIMEHOURS'] || 0) + ((p['OVERTIMEMINUTES'] || 0) / 60);

        // Accumulate totals
        sectionTotals.standardHours += stdHours;
        sectionTotals.overtimeHours += otHours;
        sectionTotals.leavePay += p['LEAVE PAY'] || 0;
        sectionTotals.bonusPay += p['BONUS PAY'] || 0;
        sectionTotals.otherIncome += p['OTHERINCOME'] || 0;
        sectionTotals.grossPay += p['GROSSSALARY'] || 0;
        sectionTotals.uif += p['UIF'] || 0;
        sectionTotals.otherDeductions += p['OTHER DEDUCTIONS'] || 0;
        sectionTotals.loanDeductions += p['LoanDeductionThisWeek'] || 0;
        sectionTotals.newLoans += p['NewLoanThisWeek'] || 0;
        sectionTotals.netPay += p['NETTSALARY'] || 0;
        sectionTotals.paidToAccounts += p['PaidtoAccount'] || 0;

        // Collect Other Income notes if present
        const otherIncomeNotes = ((p['OTHERINCOME'] || 0) > 0 && p['OTHER INCOME TEXT']) ? p['OTHER INCOME TEXT'] : '';

        registerSheet.getRange(rowNum, 1, 1, 17).setValues([[
          p['EMPLOYEE NAME'],
          p['EMPLOYER'],
          p['EMPLOYMENT STATUS'] || '',
          p['HOURLYRATE'] || 0,
          decimalHoursToTime(stdHours),
          decimalHoursToTime(otHours),
          p['LEAVE PAY'] || 0,
          p['BONUS PAY'] || 0,
          p['OTHERINCOME'] || 0,
          otherIncomeNotes,
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
      registerSheet.getRange(rowNum, 1, 1, 17).setValues([[
        'TOTALS',
        sectionTotals.employees + ' employees',
        '',  // No total for employment status
        '',  // No total for hourly rate
        decimalHoursToTime(sectionTotals.standardHours),
        decimalHoursToTime(sectionTotals.overtimeHours),
        sectionTotals.leavePay,
        sectionTotals.bonusPay,
        sectionTotals.otherIncome,
        '',  // No total for notes column
        sectionTotals.grossPay,
        sectionTotals.uif,
        sectionTotals.otherDeductions,
        sectionTotals.netPay,
        sectionTotals.loanDeductions,
        sectionTotals.newLoans,
        sectionTotals.paidToAccounts
      ]]);
      registerSheet.getRange(rowNum, 1, 1, 17).setFontWeight('bold').setBackground('#FFD700');
      const totalsRow = rowNum;
      rowNum++;

      // Format currency and number columns for this section
      if (statusPayslips.length > 0) {
        registerSheet.getRange(dataStartRow, 4, statusPayslips.length, 1).setNumberFormat('"R"#,##0.00');  // Column D - Hourly Rate
        registerSheet.getRange(dataStartRow, 7, statusPayslips.length + 1, 3).setNumberFormat('"R"#,##0.00');  // Columns G-I (Leave Pay to Other Income)
        registerSheet.getRange(dataStartRow, 11, statusPayslips.length + 1, 7).setNumberFormat('"R"#,##0.00');  // Columns K-Q (Gross Pay through Paid to Account)
      }

      // Add borders to this section's table
      const sectionTableRange = registerSheet.getRange(headerRow, 1, rowNum - headerRow, 17);
      sectionTableRange.setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);

      // Add spacing between sections
      rowNum += 2;
    }

    // ===== Payslip Received Register =====
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

    // ===== Payments Schedule =====
    const paymentsSheet = spreadsheet.insertSheet('Payments Schedule');

    // Header
    paymentsSheet.getRange('A1').setValue('Payments Schedule - Week Ending: ' + formatDate(weekEnd));
    paymentsSheet.setRowHeight(1, 30);

    // Column headers
    paymentsSheet.getRange('A3:E3').setValues([[
      'Employee',
      'Net Pay',
      'Loan Ded.',
      'New Loan',
      'Paid to Account'
    ]]);
    paymentsSheet.getRange('A3:E3').setFontWeight('bold').setBackground('#4CAF50').setFontColor('#FFFFFF').setHorizontalAlignment('center');
    paymentsSheet.setRowHeight(3, 25);

    // Data rows - combine all employees (Permanent and Temporary)
    let paymentsRowNum = 4;

    // Sort payslips by employment status (Permanent first, then Temporary) and then by employee name
    const sortedPayslips = payslips.slice().sort((a, b) => {
      const statusA = a['EMPLOYMENT STATUS'] || 'Unknown';
      const statusB = b['EMPLOYMENT STATUS'] || 'Unknown';
      const statusOrder = ['Permanent', 'Temporary', 'Resigned', 'Dismissed', 'Absconded', 'Unknown'];

      const indexA = statusOrder.indexOf(statusA);
      const indexB = statusOrder.indexOf(statusB);

      // Compare by status first
      if (indexA !== indexB) {
        if (indexA === -1) return 1;
        if (indexB === -1) return -1;
        return indexA - indexB;
      }

      // If same status, compare by employee name
      const nameA = a['EMPLOYEE NAME'] || '';
      const nameB = b['EMPLOYEE NAME'] || '';
      return nameA.localeCompare(nameB);
    });

    for (let i = 0; i < sortedPayslips.length; i++) {
      const p = sortedPayslips[i];
      paymentsSheet.getRange(paymentsRowNum, 1, 1, 5).setValues([[
        p['EMPLOYEE NAME'],
        p['NETTSALARY'] || 0,
        p['LoanDeductionThisWeek'] || 0,
        p['NewLoanThisWeek'] || 0,
        p['PaidtoAccount'] || 0
      ]]);
      paymentsRowNum++;
    }

    // Totals row
    const totalNetPay = sortedPayslips.reduce((sum, p) => sum + (p['NETTSALARY'] || 0), 0);
    const totalLoanDed = sortedPayslips.reduce((sum, p) => sum + (p['LoanDeductionThisWeek'] || 0), 0);
    const totalNewLoan = sortedPayslips.reduce((sum, p) => sum + (p['NewLoanThisWeek'] || 0), 0);
    const totalPaidToAccount = sortedPayslips.reduce((sum, p) => sum + (p['PaidtoAccount'] || 0), 0);

    paymentsSheet.getRange(paymentsRowNum, 1, 1, 5).setValues([[
      'TOTALS',
      totalNetPay,
      totalLoanDed,
      totalNewLoan,
      totalPaidToAccount
    ]]);
    paymentsSheet.getRange(paymentsRowNum, 1, 1, 5).setFontWeight('bold').setBackground('#FFD700');

    // Set column widths
    paymentsSheet.setColumnWidth(1, 220);  // Column A - Employee
    paymentsSheet.setColumnWidth(2, 120);  // Column B - Net Pay
    paymentsSheet.setColumnWidth(3, 120);  // Column C - Loan Ded.
    paymentsSheet.setColumnWidth(4, 120);  // Column D - New Loan
    paymentsSheet.setColumnWidth(5, 140);  // Column E - Paid to Account

    // Merge header cells A1:E1 and apply formatting
    paymentsSheet.getRange('A1:E1').merge();
    paymentsSheet.getRange('A1').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

    // Format currency columns (B-E) for data rows and totals
    if (sortedPayslips.length > 0) {
      paymentsSheet.getRange(4, 2, sortedPayslips.length + 1, 4).setNumberFormat('"R"#,##0.00');
    }

    // Add borders to the entire table
    const paymentsLastRow = paymentsRowNum;
    const paymentsTableRange = paymentsSheet.getRange('A3:E' + paymentsLastRow);
    paymentsTableRange.setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);

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
        totalEmployees: payslips.length
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

    // Convert to array and group by employment status
    const employeeArray = Object.values(employeeData);
    employeeArray.sort((a, b) => a.employeeName.localeCompare(b.employeeName));

    Logger.log('👥 Unique employees: ' + employeeArray.length);

    // Group employees by employment status
    const byEmploymentStatus = {};
    for (let i = 0; i < employeeArray.length; i++) {
      const emp = employeeArray[i];
      const status = emp.employmentStatus || 'Unknown';

      if (!byEmploymentStatus[status]) {
        byEmploymentStatus[status] = [];
      }
      byEmploymentStatus[status].push(emp);
    }

    // Create Google Sheet
    const monthName = ['January', 'February', 'March', 'April', 'May', 'June',
                       'July', 'August', 'September', 'October', 'November', 'December'][inputDate.getMonth()];
    const fileName = 'Monthly Payroll Summary - ' + monthName + ' ' + inputDate.getFullYear();
    const spreadsheet = SpreadsheetApp.create(fileName);

    // ===== Payroll Register =====
    const registerSheet = spreadsheet.getActiveSheet();
    registerSheet.setName('Payroll Register');

    // Main Header
    registerSheet.getRange('A1').setValue('MONTHLY PAYROLL REGISTER - ' + monthName + ' ' + inputDate.getFullYear());
    registerSheet.setRowHeight(1, 30);

    // Period row
    registerSheet.getRange('A2').setValue('Period: ' + formatDate(firstFriday) + ' to ' + formatDate(lastFriday));
    registerSheet.getRange('A2').setFontStyle('italic');

    // Set hardcoded column widths for optimal layout
    registerSheet.setColumnWidth(1, 220);  // Column A - Employee
    registerSheet.setColumnWidth(2, 135);  // Column B - Employer
    registerSheet.setColumnWidth(3, 110);  // Column C - Employment Status
    registerSheet.setColumnWidth(4, 90);   // Column D - Hourly Rate
    registerSheet.setColumnWidth(5, 80);   // Column E - Std Hours
    registerSheet.setColumnWidth(6, 80);   // Column F - OT Hours
    registerSheet.setColumnWidth(7, 90);   // Column G - Leave Pay
    registerSheet.setColumnWidth(8, 90);   // Column H - Bonus Pay
    registerSheet.setColumnWidth(9, 100);  // Column I - Other Income
    registerSheet.setColumnWidth(10, 150); // Column J - Other Income Notes
    registerSheet.setColumnWidth(11, 90);  // Column K - Gross Pay
    registerSheet.setColumnWidth(12, 80);  // Column L - UIF
    registerSheet.setColumnWidth(13, 110); // Column M - Other Ded.
    registerSheet.setColumnWidth(14, 110); // Column N - Net Pay

    // Merge header cells A1:N1 and apply formatting
    registerSheet.getRange('A1:N1').merge();
    registerSheet.getRange('A1').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

    let rowNum = 4;

    // Sort employment statuses to ensure consistent order
    const statusOrder = ['Permanent', 'Temporary', 'Resigned', 'Dismissed', 'Absconded', 'Unknown'];
    const sortedStatuses = Object.keys(byEmploymentStatus).sort((a, b) => {
      const indexA = statusOrder.indexOf(a);
      const indexB = statusOrder.indexOf(b);
      if (indexA === -1 && indexB === -1) return a.localeCompare(b);
      if (indexA === -1) return 1;
      if (indexB === -1) return -1;
      return indexA - indexB;
    });

    // Create a section for each employment status
    for (const status of sortedStatuses) {
      const employees = byEmploymentStatus[status];

      // Section header
      registerSheet.getRange(rowNum, 1).setValue('Monthly Report ' + status + ' Employees');
      registerSheet.getRange(rowNum, 1, 1, 14).merge();
      registerSheet.getRange(rowNum, 1).setFontWeight('bold').setFontSize(12);
      registerSheet.setRowHeight(rowNum, 25);
      rowNum++;

      // Column headers
      registerSheet.getRange(rowNum, 1, 1, 14).setValues([[
        'Employee',
        'Employer',
        'Employment Status',
        'Hourly Rate',
        'Std Hours',
        'OT Hours',
        'Leave Pay',
        'Bonus Pay',
        'Other Income',
        'Other Income Notes',
        'Gross Pay',
        'UIF',
        'Other Ded.',
        'Net Pay'
      ]]);
      registerSheet.getRange(rowNum, 1, 1, 14).setFontWeight('bold').setBackground('#4CAF50').setFontColor('#FFFFFF').setHorizontalAlignment('center');
      registerSheet.setRowHeight(rowNum, 25);
      const headerRow = rowNum;
      rowNum++;

      // Calculate section totals
      let sectionTotals = {
        employees: employees.length,
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

      // Data rows
      const dataStartRow = rowNum;
      for (let i = 0; i < employees.length; i++) {
        const emp = employees[i];

        // Accumulate totals
        sectionTotals.standardHours += emp.standardHours;
        sectionTotals.overtimeHours += emp.overtimeHours;
        sectionTotals.grossPay += emp.grossPay;
        sectionTotals.leavePay += emp.leavePay;
        sectionTotals.bonusPay += emp.bonusPay;
        sectionTotals.otherIncome += emp.otherIncome;
        sectionTotals.uif += emp.uif;
        sectionTotals.otherDeductions += emp.otherDeductions;
        sectionTotals.netPay += emp.netPay;

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
          emp.leavePay,
          emp.bonusPay,
          emp.otherIncome,
          otherIncomeNotesText,
          emp.grossPay,
          emp.uif,
          emp.otherDeductions,
          emp.netPay
        ]]);
        rowNum++;
      }

      // Totals row
      const totalStdHoursFormatted = decimalHoursToTime(sectionTotals.standardHours);
      const totalOtHoursFormatted = decimalHoursToTime(sectionTotals.overtimeHours);

      registerSheet.getRange(rowNum, 1, 1, 14).setValues([[
        'TOTALS',
        sectionTotals.employees + ' employees',
        '',  // No total for employment status
        '',  // No total for hourly rate
        totalStdHoursFormatted,
        totalOtHoursFormatted,
        sectionTotals.leavePay,
        sectionTotals.bonusPay,
        sectionTotals.otherIncome,
        '',  // No total for notes column
        sectionTotals.grossPay,
        sectionTotals.uif,
        sectionTotals.otherDeductions,
        sectionTotals.netPay
      ]]);
      registerSheet.getRange(rowNum, 1, 1, 14).setFontWeight('bold').setBackground('#FFD700');
      const totalsRow = rowNum;
      rowNum++;

      // Format currency and number columns for this section
      if (employees.length > 0) {
        registerSheet.getRange(dataStartRow, 4, employees.length, 1).setNumberFormat('"R"#,##0.00');  // Column D - Hourly Rate
        registerSheet.getRange(dataStartRow, 7, employees.length + 1, 3).setNumberFormat('"R"#,##0.00');  // Columns G-I (Leave Pay to Other Income)
        registerSheet.getRange(dataStartRow, 11, employees.length + 1, 1).setNumberFormat('"R"#,##0.00');  // Column K (Gross Pay)
        registerSheet.getRange(dataStartRow, 12, employees.length + 1, 3).setNumberFormat('"R"#,##0.00');  // Columns L-N (UIF to Net Pay)
      }

      // Add borders to this section's table
      const sectionTableRange = registerSheet.getRange(headerRow, 1, rowNum - headerRow, 14);
      sectionTableRange.setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);

      // Add spacing between sections
      rowNum += 2;
    }

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
        totalEmployees: employeeArray.length
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

/**
 * Test function for Loan Transaction Matrix Report
 */
function test_loanTransactionMatrixReport() {
  Logger.log('\n========== TEST: LOAN TRANSACTION MATRIX REPORT ==========');

  // Use a date range (update with actual dates that have loan data)
  const startDate = new Date('2025-01-01');
  const endDate = new Date('2025-12-31');

  const result = generateLoanTransactionMatrixReport(startDate, endDate);

  if (result.success) {
    Logger.log('✅ TEST PASSED');
    Logger.log('URL: ' + result.data.url);
    Logger.log('Period: ' + result.data.periodStart + ' to ' + result.data.periodEnd);
    Logger.log('Employees with transactions: ' + result.data.employeeCount);
    Logger.log('Transaction dates: ' + result.data.transactionDateCount);
  } else {
    Logger.log('❌ TEST FAILED: ' + result.error);
  }

  Logger.log('========== TEST COMPLETE ==========\n');
}
