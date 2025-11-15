/**
 * CONFIG.GS - System Configuration and Constants
 * HR Payroll System for SA Grinding Wheels & Scorpio Abrasives
 *
 * This file contains all constants, configuration values, and dropdown options
 * used throughout the system. Update these values to change system behavior.
 */

// ==================== EMPLOYER CONFIGURATION ====================

/**
 * List of valid employers
 * Employees can be assigned to either employer
 */
var EMPLOYER_LIST = [
  'SA Grinding Wheels',
  'Scorpio Abrasives'
];

/**
 * Company contact information for PDF generation
 */
var COMPANY_INFO = {
  address1: '18 DODGE STREET',
  address2: 'AUREUS',
  address3: 'RANDFONTEIN',
  phone: '(011) 693 4278 / 083 338 5609',
  email1: 'INFO@SAGRINDING.CO.ZA',
  email2: 'INFO@SCORPIOABRASIVES.CO.ZA'
};

// ==================== EMPLOYMENT STATUS ====================

/**
 * Valid employment status options
 * Affects UIF calculation (only permanent employees pay UIF)
 */
var EMPLOYMENT_STATUS_LIST = [
  'Permanent',
  'Temporary',
  'Contract'
];

/**
 * UIF rate (1% for permanent employees)
 */
var UIF_RATE = 0.01;

/**
 * Overtime multiplier (1.5x standard rate)
 */
var OVERTIME_MULTIPLIER = 1.5;

// ==================== LEAVE CONFIGURATION ====================

/**
 * Valid leave reasons
 */
var LEAVE_REASONS = [
  'AWOL',
  'Sick Leave',
  'Annual Leave',
  'Unpaid Leave'
];

// ==================== LOAN CONFIGURATION ====================

/**
 * Loan transaction types
 */
var LOAN_TYPES = [
  'Disbursement',
  'Repayment'
];

/**
 * Loan disbursement modes
 */
var DISBURSEMENT_MODES = [
  'With Salary',   // Loan amount added to Paid to Account
  'Separate',      // Loan disbursed separately, not added to Paid to Account
  'Manual Entry'   // Manual loan entry (not linked to payslip)
];

// ==================== TIMESHEET CONFIGURATION ====================

/**
 * Timesheet status options
 */
var TIMESHEET_STATUS = [
  'Pending',
  'Approved',
  'Rejected'
];

// ==================== FIELD DEFINITIONS ====================

/**
 * Required fields for employee creation
 * These fields must have values before saving
 */
var EMPLOYEE_REQUIRED_FIELDS = [
  'EMPLOYEE NAME',
  'SURNAME',
  'EMPLOYER',
  'HOURLY RATE',
  'ID NUMBER',
  'CONTACT NUMBER',
  'ADDRESS',
  'EMPLOYMENT STATUS',
  'ClockInRef'
];

/**
 * All employee fields (required + optional)
 */
var EMPLOYEE_ALL_FIELDS = [
  'id',                      // System-generated UUID
  'REFNAME',                 // Auto-generated: "FirstName Surname"
  'EMPLOYEE NAME',           // Required
  'SURNAME',                 // Required
  'EMPLOYER',                // Required
  'HOURLY RATE',             // Required
  'ID NUMBER',               // Required (unique)
  'CONTACT NUMBER',          // Required
  'ADDRESS',                 // Required
  'EMPLOYMENT STATUS',       // Required
  'ClockInRef',              // Required (unique)
  'DATE OF BIRTH',           // Optional
  'ALTERNATIVE CONTACT',     // Optional
  'ALT CONTACT NAME',        // Optional
  'INCOME TAX NUMBER',       // Optional
  'EMPLOYMENT DATE',         // Optional
  'TERMINATION DATE',        // Optional
  'NOTES',                   // Optional
  'OveralSize',              // Optional (uniform)
  'ShoeSize',                // Optional (uniform)
  'UNIONMEM',                // Legacy (keep but don't use)
  'UNIONFEE',                // Legacy (keep but don't use)
  'RETMEMBER',               // Legacy (keep but don't use)
  'RETFACCNUMBER',           // Legacy (keep but don't use)
  'USER',                    // System audit field
  'TIMESTAMP',               // System audit field
  'MODIFIED_BY',             // System audit field
  'LAST_MODIFIED'            // System audit field
];

/**
 * Payslip fields
 */
var PAYSLIP_FIELDS = [
  'RECORDNUMBER',            // Unique payslip number
  'id',                      // Employee ID (links to EMPLOYEE DETAILS)
  'TIMESTAMP',               // Creation timestamp
  'EMPLOYEE NAME',           // Reference to employee
  'EMPLOYER',                // Looked up from employee
  'EMPLOYMENT STATUS',       // Looked up from employee
  'WEEKENDING',              // Week ending date
  'HOURS',                   // Standard hours
  'MINUTES',                 // Standard minutes
  'OVERTIMEHOURS',           // Overtime hours
  'OVERTIMEMINUTES',         // Overtime minutes
  'HOURLYRATE',              // Looked up from employee
  'STANDARDTIME',            // Calculated
  'OVERTIME',                // Calculated
  'LEAVE PAY',               // Manual entry
  'BONUS PAY',               // Manual entry
  'OTHERINCOME',             // Manual entry
  'GROSSSALARY',             // Calculated
  'UIF',                     // Calculated
  'OTHER DEDUCTIONS',        // Manual entry
  'OTHER DEDUCTIONS TEXT',   // Description
  'CurrentLoanBalance',      // Looked up from loans
  'LoanDeductionThisWeek',   // Manual entry
  'NewLoanThisWeek',         // Manual entry
  'LoanDisbursementType',    // Manual entry
  'UpdatedLoanBalance',      // Calculated
  'LoanRepaymentLogged',     // System flag
  'TOTALDEDUCTIONS',         // Calculated
  'NETTSALARY',              // Calculated
  'PaidToAccount',           // Calculated (CRITICAL)
  'FILENAME',                // PDF filename
  'FILELINK',                // PDF URL
  'NOTES',                   // Free text
  'USER',                    // System audit
  'MODIFIED_BY',             // System audit
  'LAST_MODIFIED'            // System audit
];

// ==================== VALIDATION PATTERNS ====================

/**
 * South African ID Number format (13 digits)
 */
var SA_ID_NUMBER_PATTERN = /^\d{13}$/;

/**
 * South African phone number formats
 * Accepts: 0821234567, +27821234567, 082 123 4567, etc.
 */
var SA_PHONE_PATTERN = /^(\+27|0)[0-9]{9}$/;

/**
 * Email validation pattern
 */
var EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

// ==================== SYSTEM SETTINGS ====================

/**
 * Maximum number of records to display per page in lists
 */
var RECORDS_PER_PAGE = 50;

/**
 * Report generation timeout (milliseconds)
 */
var REPORT_TIMEOUT = 300000; // 5 minutes

/**
 * Auto-sync change detection window (milliseconds)
 * Only process changes within the last 5 minutes to avoid reprocessing old data
 */
var CHANGE_DETECTION_WINDOW = 300000; // 5 minutes

/**
 * Currency symbol
 */
var CURRENCY_SYMBOL = 'R';

/**
 * Date format for display
 */
var DATE_FORMAT = 'DD MMMM YYYY'; // e.g., "17 October 2025"

/**
 * Short date format for tables
 */
var SHORT_DATE_FORMAT = 'YYYY-MM-DD'; // e.g., "2025-10-17"

// ==================== GOOGLE DRIVE SETTINGS ====================

/**
 * Folder name for storing generated reports
 */
var REPORTS_FOLDER_NAME = 'Payroll Reports';

/**
 * Folder name for storing generated payslip PDFs
 */
var PAYSLIPS_FOLDER_NAME = 'Payslip PDFs';

/**
 * Report file naming patterns
 */
var REPORT_FILE_NAMES = {
  outstandingLoans: 'Outstanding Loans Report - {date}',
  individualStatement: 'Loan Statement - {employeeName} - {startDate} to {endDate}',
  weeklyPayroll: 'Weekly Payroll Summary - Week Ending {weekEnding}'
};

// ==================== SHEET NAMES ====================

/**
 * Expected sheet names (case-insensitive, whitespace-removed matching)
 * The system will look for these names in the spreadsheet
 */
var EXPECTED_SHEET_NAMES = {
  salary: ['MASTERSALARY', 'salaryrecords', 'payroll'],
  loans: ['EmployeeLoans', 'loantransactions', 'loans'],
  empdetails: ['EMPDETAILS', 'employeedetails', 'employees'],
  leave: ['LEAVE', 'leaverecords'],
  pending: ['PendingTimesheets', 'pendingtimesheets', 'timeapproval']
};

// ==================== ERROR MESSAGES ====================

/**
 * Standard error messages
 */
var ERROR_MESSAGES = {
  // General
  SHEET_NOT_FOUND: 'Required sheet not found: {sheetName}',
  INVALID_INPUT: 'Invalid input data provided',
  OPERATION_FAILED: 'Operation failed: {operation}',

  // Employee
  EMPLOYEE_NOT_FOUND: 'Employee not found: {employeeName}',
  DUPLICATE_ID_NUMBER: 'ID Number already exists: {idNumber}',
  DUPLICATE_CLOCK_REF: 'Clock Number already exists: {clockInRef}',
  INVALID_HOURLY_RATE: 'Hourly rate must be greater than 0',
  MISSING_REQUIRED_FIELDS: 'Missing required fields: {fields}',

  // Payroll
  DUPLICATE_PAYSLIP: 'Payslip already exists for {employeeName} on {weekEnding}',
  INVALID_TIME_VALUES: 'Time values cannot be negative',
  LOAN_EXCEEDS_BALANCE: 'Loan deduction exceeds current balance',

  // Leave
  INVALID_DATE_RANGE: 'Return date must be greater than or equal to start date',

  // Loans
  INVALID_LOAN_AMOUNT: 'Loan amount cannot be zero',
  DUPLICATE_SALARY_LINK: 'Loan record already exists for payslip: {recordNumber}',

  // Timesheets
  TIMESHEET_ALREADY_APPROVED: 'Timesheet already approved for {employeeName} on {weekEnding}',
  PAYSLIP_ALREADY_EXISTS: 'Payslip already exists, cannot approve timesheet'
};

// ==================== SUCCESS MESSAGES ====================

/**
 * Standard success messages
 */
var SUCCESS_MESSAGES = {
  EMPLOYEE_CREATED: 'Employee created successfully',
  EMPLOYEE_UPDATED: 'Employee updated successfully',
  EMPLOYEE_TERMINATED: 'Employee terminated successfully',

  LEAVE_RECORDED: 'Leave recorded successfully',

  LOAN_CREATED: 'Loan transaction recorded successfully',
  LOAN_BALANCE_UPDATED: 'Loan balance updated successfully',

  TIMESHEET_IMPORTED: 'Timesheet imported successfully',
  TIMESHEET_APPROVED: 'Timesheet approved and payslip created',
  TIMESHEET_REJECTED: 'Timesheet rejected',

  PAYSLIP_CREATED: 'Payslip created successfully',
  PAYSLIP_UPDATED: 'Payslip updated successfully',
  PDF_GENERATED: 'PDF generated successfully',

  REPORT_GENERATED: 'Report generated successfully',

  SYNC_COMPLETE: 'Loan sync completed successfully'
};

// ==================== HELPER FUNCTIONS ====================

/**
 * Get configuration value by key
 *
 * @param {string} key - Configuration key
 * @returns {*} Configuration value
 */
function getConfig(key) {
  const configs = {
    EMPLOYER_LIST: EMPLOYER_LIST,
    EMPLOYMENT_STATUS_LIST: EMPLOYMENT_STATUS_LIST,
    LEAVE_REASONS: LEAVE_REASONS,
    LOAN_TYPES: LOAN_TYPES,
    DISBURSEMENT_MODES: DISBURSEMENT_MODES,
    TIMESHEET_STATUS: TIMESHEET_STATUS,
    EMPLOYEE_REQUIRED_FIELDS: EMPLOYEE_REQUIRED_FIELDS,
    EMPLOYEE_ALL_FIELDS: EMPLOYEE_ALL_FIELDS,
    UIF_RATE: UIF_RATE,
    OVERTIME_MULTIPLIER: OVERTIME_MULTIPLIER,
    CURRENCY_SYMBOL: CURRENCY_SYMBOL,
    COMPANY_INFO: COMPANY_INFO
  };

  return configs[key];
}

/**
 * Validate if value is in allowed list
 *
 * @param {string} value - Value to validate
 * @param {Array} allowedValues - List of allowed values
 * @returns {boolean} True if valid
 */
function isValidEnum(value, allowedValues) {
  return allowedValues.includes(value);
}
