#!/usr/bin/env node

/**
 * DATA MIGRATION SCRIPT
 * Google Sheets → MySQL (Oracle Cloud)
 *
 * This script migrates all data from Google Sheets to MySQL database.
 * Run this ONCE during the initial migration.
 *
 * Prerequisites:
 * 1. MySQL database created (run schema.sql first)
 * 2. Google Sheets API credentials (service account JSON)
 * 3. Node.js packages installed (npm install)
 *
 * Usage:
 *   node migrate-data.js
 */

const { google } = require('googleapis');
const mysql = require('mysql2/promise');
const fs = require('fs');
const path = require('path');

// ============================================================
// CONFIGURATION
// ============================================================

const CONFIG = {
  // Google Sheets Configuration
  SPREADSHEET_ID: process.env.SPREADSHEET_ID || 'YOUR_SPREADSHEET_ID_HERE',
  SERVICE_ACCOUNT_FILE: process.env.SERVICE_ACCOUNT_FILE || './google-credentials.json',

  // MySQL Configuration (Oracle Cloud)
  MYSQL_HOST: process.env.MYSQL_HOST || 'localhost',
  MYSQL_USER: process.env.MYSQL_USER || 'hrapp',
  MYSQL_PASSWORD: process.env.MYSQL_PASSWORD || 'your_password_here',
  MYSQL_DATABASE: process.env.MYSQL_DATABASE || 'hrdb',

  // Sheet names (maps to database tables)
  SHEETS: {
    EMPLOYEES: 'EMPLOYEE DETAILS',
    MASTERSALARY: 'MASTERSALARY',
    EMPLOYEE_LOANS: 'EmployeeLoans',
    LEAVE: 'LEAVE',
    PENDING_TIMESHEETS: 'PendingTimesheets',
    RAW_CLOCK_DATA: 'RAW_CLOCK_DATA',
    CLOCK_IN_IMPORTS: 'CLOCK_IN_IMPORTS',
    USER_CONFIG: 'UserConfig'
  }
};

// ============================================================
// GOOGLE SHEETS CLIENT
// ============================================================

async function getGoogleSheetsClient() {
  const auth = new google.auth.GoogleAuth({
    keyFile: CONFIG.SERVICE_ACCOUNT_FILE,
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });

  const authClient = await auth.getClient();
  return google.sheets({ version: 'v4', auth: authClient });
}

// ============================================================
// MYSQL CONNECTION
// ============================================================

async function getMySQLConnection() {
  return await mysql.createConnection({
    host: CONFIG.MYSQL_HOST,
    user: CONFIG.MYSQL_USER,
    password: CONFIG.MYSQL_PASSWORD,
    database: CONFIG.MYSQL_DATABASE,
    multipleStatements: true
  });
}

// ============================================================
// UTILITY FUNCTIONS
// ============================================================

function formatDate(dateValue) {
  if (!dateValue) return null;

  // Google Sheets serial date to JS Date
  if (typeof dateValue === 'number') {
    const date = new Date((dateValue - 25569) * 86400 * 1000);
    return date.toISOString().split('T')[0];
  }

  // String date
  const parsed = new Date(dateValue);
  return isNaN(parsed) ? null : parsed.toISOString().split('T')[0];
}

function formatDateTime(dateValue) {
  if (!dateValue) return null;

  // Google Sheets serial date to JS Date
  if (typeof dateValue === 'number') {
    const date = new Date((dateValue - 25569) * 86400 * 1000);
    return date.toISOString().slice(0, 19).replace('T', ' ');
  }

  // String datetime
  const parsed = new Date(dateValue);
  return isNaN(parsed) ? null : parsed.toISOString().slice(0, 19).replace('T', ' ');
}

function parseNumber(value) {
  if (value === '' || value === null || value === undefined) return 0;
  const num = parseFloat(value);
  return isNaN(num) ? 0 : num;
}

function buildObjectFromRow(row, headers) {
  const obj = {};
  headers.forEach((header, index) => {
    obj[header] = row[index] || '';
  });
  return obj;
}

// ============================================================
// MIGRATION FUNCTIONS
// ============================================================

async function migrateEmployees(sheets, connection) {
  console.log('\n📋 Migrating EMPLOYEE DETAILS...');

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: CONFIG.SPREADSHEET_ID,
    range: `${CONFIG.SHEETS.EMPLOYEES}!A:AZ`,
  });

  const rows = response.data.values;
  if (!rows || rows.length === 0) {
    console.log('⚠️  No employee data found');
    return 0;
  }

  const headers = rows[0];
  const dataRows = rows.slice(1);

  let migrated = 0;

  for (const row of dataRows) {
    const emp = buildObjectFromRow(row, headers);

    // Skip empty rows
    if (!emp.id || emp.id === '') continue;

    try {
      await connection.execute(`
        INSERT INTO employees (
          id, refname, employee_name, surname, id_number, date_of_birth,
          contact_number, alternative_contact, alt_contact_name, address,
          employer, employment_status, hourly_rate, employment_date, termination_date,
          clock_in_ref, income_tax_number, overall_size, shoe_size,
          union_member, union_fee, retirement_member, retirement_account_number,
          notes, user, timestamp, modified_by, last_modified
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON DUPLICATE KEY UPDATE
          employee_name = VALUES(employee_name),
          surname = VALUES(surname),
          hourly_rate = VALUES(hourly_rate)
      `, [
        emp.id,
        emp.REFNAME || '',
        emp['EMPLOYEE NAME'] || '',
        emp.SURNAME || '',
        emp['ID NUMBER'] || '',
        formatDate(emp['DATE OF BIRTH']),
        emp['CONTACT NUMBER'] || '',
        emp['ALTERNATIVE CONTACT'] || '',
        emp['ALT CONTACT NAME'] || '',
        emp.ADDRESS || '',
        emp.EMPLOYER || 'SA Grinding Wheels',
        emp['EMPLOYMENT STATUS'] || 'Permanent',
        parseNumber(emp['HOURLY RATE']),
        formatDate(emp['EMPLOYMENT DATE']),
        formatDate(emp['TERMINATION DATE']),
        emp.ClockInRef || '',
        emp['INCOME TAX NUMBER'] || '',
        emp.OveralSize || '',
        emp.ShoeSize || '',
        emp.UNIONMEM || '',
        parseNumber(emp.UNIONFEE),
        emp.RETMEMBER || '',
        emp.RETFACCNUMBER || '',
        emp.NOTES || '',
        emp.USER || '',
        formatDateTime(emp.TIMESTAMP) || null,
        emp.MODIFIED_BY || '',
        formatDateTime(emp.LAST_MODIFIED) || null
      ]);

      migrated++;
    } catch (error) {
      console.error(`❌ Error migrating employee ${emp['EMPLOYEE NAME']}: ${error.message}`);
    }
  }

  console.log(`✅ Migrated ${migrated} employees`);
  return migrated;
}

async function migrateMasterSalary(sheets, connection) {
  console.log('\n📋 Migrating MASTERSALARY (this may take a while - 4,160+ records)...');

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: CONFIG.SPREADSHEET_ID,
    range: `${CONFIG.SHEETS.MASTERSALARY}!A:AZ`,
  });

  const rows = response.data.values;
  if (!rows || rows.length === 0) {
    console.log('⚠️  No salary data found');
    return 0;
  }

  const headers = rows[0];
  const dataRows = rows.slice(1);

  let migrated = 0;
  let batchSize = 100;
  let batch = [];

  for (const row of dataRows) {
    const payslip = buildObjectFromRow(row, headers);

    // Skip empty rows
    if (!payslip.RECORDNUMBER || payslip.RECORDNUMBER === '') continue;

    batch.push([
      parseInt(payslip.RECORDNUMBER) || null,
      payslip.id || '',
      payslip['EMPLOYEE NAME'] || '',
      payslip.EMPLOYER || 'SA Grinding Wheels',
      payslip['EMPLOYMENT STATUS'] || 'Permanent',
      formatDate(payslip.WEEKENDING),
      parseInt(payslip.HOURS) || 0,
      parseInt(payslip.MINUTES) || 0,
      parseInt(payslip.OVERTIMEHOURS) || 0,
      parseInt(payslip.OVERTIMEMINUTES) || 0,
      parseNumber(payslip.HOURLYRATE),
      parseNumber(payslip.STANDARDTIME),
      parseNumber(payslip.OVERTIME),
      parseNumber(payslip['LEAVE PAY']),
      parseNumber(payslip['BONUS PAY']),
      parseNumber(payslip.OTHERINCOME),
      payslip['OTHER INCOME TEXT'] || '',
      parseNumber(payslip.GROSSSALARY),
      parseNumber(payslip.UIF),
      parseNumber(payslip['OTHER DEDUCTIONS']),
      payslip['OTHER DEDUCTIONS TEXT'] || '',
      parseNumber(payslip.CurrentLoanBalance),
      parseNumber(payslip.LoanDeductionThisWeek),
      parseNumber(payslip.NewLoanThisWeek),
      payslip.LoanDisbursementType || '',
      parseNumber(payslip.UpdatedLoanBalance),
      payslip.LoanRepaymentLogged === 'TRUE' || payslip.LoanRepaymentLogged === true,
      parseNumber(payslip.TOTALDEDUCTIONS),
      parseNumber(payslip.NETTSALARY),
      parseNumber(payslip.PaidtoAccount),
      payslip.FILENAME || '',
      payslip.FILELINK || '',
      payslip.NOTES || '',
      payslip.USER || '',
      formatDateTime(payslip.TIMESTAMP) || null,
      payslip.MODIFIED_BY || '',
      formatDateTime(payslip.LAST_MODIFIED) || null
    ]);

    // Execute batch
    if (batch.length >= batchSize) {
      try {
        for (const values of batch) {
          await connection.execute(`
            INSERT INTO mastersalary (
              record_number, employee_id, employee_name, employer, employment_status,
              week_ending, hours, minutes, overtime_hours, overtime_minutes,
              hourly_rate, standard_time, overtime, leave_pay, bonus_pay,
              other_income, other_income_text, gross_salary, uif,
              other_deductions, other_deductions_text,
              current_loan_balance, loan_deduction_this_week, new_loan_this_week,
              loan_disbursement_type, updated_loan_balance, loan_repayment_logged,
              total_deductions, net_salary, paid_to_account,
              filename, filelink, notes, user, timestamp, modified_by, last_modified
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON DUPLICATE KEY UPDATE
              employee_id = VALUES(employee_id),
              week_ending = VALUES(week_ending)
          `, values);
        }
        migrated += batch.length;
        console.log(`   ✓ Migrated ${migrated} payslips...`);
      } catch (error) {
        console.error(`❌ Batch error: ${error.message}`);
      }
      batch = [];
    }
  }

  // Execute remaining batch
  if (batch.length > 0) {
    try {
      for (const values of batch) {
        await connection.execute(`
          INSERT INTO mastersalary (
            record_number, employee_id, employee_name, employer, employment_status,
            week_ending, hours, minutes, overtime_hours, overtime_minutes,
            hourly_rate, standard_time, overtime, leave_pay, bonus_pay,
            other_income, other_income_text, gross_salary, uif,
            other_deductions, other_deductions_text,
            current_loan_balance, loan_deduction_this_week, new_loan_this_week,
            loan_disbursement_type, updated_loan_balance, loan_repayment_logged,
            total_deductions, net_salary, paid_to_account,
            filename, filelink, notes, user, timestamp, modified_by, last_modified
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
          ON DUPLICATE KEY UPDATE
            employee_id = VALUES(employee_id),
            week_ending = VALUES(week_ending)
        `, values);
      }
      migrated += batch.length;
    } catch (error) {
      console.error(`❌ Final batch error: ${error.message}`);
    }
  }

  console.log(`✅ Migrated ${migrated} payslips`);
  return migrated;
}

async function migrateEmployeeLoans(sheets, connection) {
  console.log('\n📋 Migrating EmployeeLoans...');

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: CONFIG.SPREADSHEET_ID,
    range: `${CONFIG.SHEETS.EMPLOYEE_LOANS}!A:AZ`,
  });

  const rows = response.data.values;
  if (!rows || rows.length === 0) {
    console.log('⚠️  No loan data found');
    return 0;
  }

  const headers = rows[0];
  const dataRows = rows.slice(1);

  let migrated = 0;

  for (const row of dataRows) {
    const loan = buildObjectFromRow(row, headers);

    // Skip empty rows
    if (!loan['Employee ID'] || loan['Employee ID'] === '') continue;

    try {
      await connection.execute(`
        INSERT INTO employee_loans (
          employee_id, salary_link, transaction_date, loan_type,
          amount, balance_after, notes, timestamp, user
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
      `, [
        loan['Employee ID'],
        parseInt(loan.SalaryLink) || null,
        formatDate(loan.TransactionDate),
        loan.LoanType || 'Disbursement',
        parseNumber(loan.Amount),
        parseNumber(loan.BalanceAfter),
        loan.Notes || '',
        formatDateTime(loan.Timestamp) || null,
        loan.User || ''
      ]);

      migrated++;
    } catch (error) {
      console.error(`❌ Error migrating loan record: ${error.message}`);
    }
  }

  console.log(`✅ Migrated ${migrated} loan records`);
  return migrated;
}

async function migrateLeaveRecords(sheets, connection) {
  console.log('\n📋 Migrating LEAVE records...');

  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: CONFIG.SPREADSHEET_ID,
      range: `${CONFIG.SHEETS.LEAVE}!A:AZ`,
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) {
      console.log('⚠️  No leave data found');
      return 0;
    }

    const headers = rows[0];
    const dataRows = rows.slice(1);

    let migrated = 0;

    for (const row of dataRows) {
      const leave = buildObjectFromRow(row, headers);

      // Skip empty rows
      if (!leave['Employee ID'] || leave['Employee ID'] === '') continue;

      try {
        await connection.execute(`
          INSERT INTO leave_records (
            employee_id, employee_name, leave_date, leave_reason,
            notes, user, timestamp
          ) VALUES (?, ?, ?, ?, ?, ?, ?)
        `, [
          leave['Employee ID'] || '',
          leave['Employee Name'] || '',
          formatDate(leave['Leave Date']),
          leave['Leave Reason'] || 'AWOL',
          leave.Notes || '',
          leave.User || '',
          formatDateTime(leave.Timestamp) || null
        ]);

        migrated++;
      } catch (error) {
        console.error(`❌ Error migrating leave record: ${error.message}`);
      }
    }

    console.log(`✅ Migrated ${migrated} leave records`);
    return migrated;
  } catch (error) {
    console.log(`⚠️  Leave sheet not found or error: ${error.message}`);
    return 0;
  }
}

// ============================================================
// MAIN MIGRATION
// ============================================================

async function main() {
  console.log('════════════════════════════════════════════════════');
  console.log('   HR PAYROLL MIGRATION: Google Sheets → MySQL');
  console.log('════════════════════════════════════════════════════\n');

  let sheetsClient;
  let mysqlConnection;

  try {
    // Connect to Google Sheets
    console.log('🔌 Connecting to Google Sheets...');
    sheetsClient = await getGoogleSheetsClient();
    console.log('✅ Connected to Google Sheets\n');

    // Connect to MySQL
    console.log('🔌 Connecting to MySQL database...');
    mysqlConnection = await getMySQLConnection();
    console.log('✅ Connected to MySQL\n');

    // Disable foreign key checks temporarily for faster import
    await mysqlConnection.execute('SET FOREIGN_KEY_CHECKS = 0');

    // Run migrations
    const stats = {
      employees: 0,
      payslips: 0,
      loans: 0,
      leave: 0
    };

    stats.employees = await migrateEmployees(sheetsClient, mysqlConnection);
    stats.payslips = await migrateMasterSalary(sheetsClient, mysqlConnection);
    stats.loans = await migrateEmployeeLoans(sheetsClient, mysqlConnection);
    stats.leave = await migrateLeaveRecords(sheetsClient, mysqlConnection);

    // Re-enable foreign key checks
    await mysqlConnection.execute('SET FOREIGN_KEY_CHECKS = 1');

    // Summary
    console.log('\n════════════════════════════════════════════════════');
    console.log('   MIGRATION COMPLETE!');
    console.log('════════════════════════════════════════════════════');
    console.log(`✅ Employees:     ${stats.employees}`);
    console.log(`✅ Payslips:      ${stats.payslips}`);
    console.log(`✅ Loan records:  ${stats.loans}`);
    console.log(`✅ Leave records: ${stats.leave}`);
    console.log('════════════════════════════════════════════════════\n');

    console.log('🎉 Data migration successful!');
    console.log('📝 Next steps:');
    console.log('   1. Verify data in MySQL: mysql -u hrapp -p hrdb');
    console.log('   2. Run API server: cd api && npm start');
    console.log('   3. Update frontend to use new API endpoints\n');

  } catch (error) {
    console.error('\n❌ MIGRATION FAILED:', error.message);
    console.error(error.stack);
    process.exit(1);
  } finally {
    if (mysqlConnection) {
      await mysqlConnection.end();
      console.log('🔌 MySQL connection closed');
    }
  }
}

// Run migration
if (require.main === module) {
  main().catch(console.error);
}

module.exports = { main };
