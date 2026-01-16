-- ============================================================
-- HR PAYROLL SYSTEM - MySQL Database Schema
-- Oracle Cloud Always Free Tier Migration
-- ============================================================
-- This schema mirrors the Google Sheets structure from the Apps Script system
-- Optimized for performance with proper indexes and foreign keys
-- ============================================================

-- Create database (if not exists)
CREATE DATABASE IF NOT EXISTS hrdb CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
USE hrdb;

-- ============================================================
-- TABLE: employees (EMPLOYEE DETAILS sheet)
-- ============================================================
CREATE TABLE IF NOT EXISTS employees (
    -- Primary Key
    id VARCHAR(50) PRIMARY KEY COMMENT 'System-generated UUID',

    -- Employee Information
    refname VARCHAR(100) COMMENT 'Auto-generated: FirstName Surname',
    employee_name VARCHAR(100) NOT NULL COMMENT 'First name',
    surname VARCHAR(100) NOT NULL COMMENT 'Last name',
    id_number VARCHAR(13) NOT NULL UNIQUE COMMENT 'SA ID Number (13 digits)',
    date_of_birth DATE COMMENT 'Date of birth',

    -- Contact Information
    contact_number VARCHAR(20) NOT NULL COMMENT 'Primary contact',
    alternative_contact VARCHAR(20) COMMENT 'Alternative contact number',
    alt_contact_name VARCHAR(100) COMMENT 'Alternative contact person name',
    address TEXT NOT NULL COMMENT 'Physical address',

    -- Employment Details
    employer ENUM('SA Grinding Wheels', 'Scorpio Abrasives') NOT NULL COMMENT 'Employer name',
    employment_status ENUM('Permanent', 'Temporary', 'Contract') NOT NULL COMMENT 'Employment type',
    hourly_rate DECIMAL(10,2) NOT NULL COMMENT 'Hourly pay rate',
    employment_date DATE COMMENT 'Date of employment',
    termination_date DATE COMMENT 'Date of termination (if applicable)',

    -- System References
    clock_in_ref VARCHAR(50) NOT NULL UNIQUE COMMENT 'Clock-in system reference',
    income_tax_number VARCHAR(20) COMMENT 'Tax number',

    -- Uniform Sizes
    overall_size VARCHAR(10) COMMENT 'Overall uniform size',
    shoe_size VARCHAR(10) COMMENT 'Shoe size',

    -- Legacy Fields (keep for compatibility)
    union_member VARCHAR(10) COMMENT 'Legacy - union membership',
    union_fee DECIMAL(10,2) COMMENT 'Legacy - union fee',
    retirement_member VARCHAR(10) COMMENT 'Legacy - retirement fund',
    retirement_account_number VARCHAR(50) COMMENT 'Legacy - retirement account',

    -- Notes
    notes TEXT COMMENT 'Additional notes',

    -- Audit Fields
    user VARCHAR(100) COMMENT 'Created by user',
    timestamp DATETIME DEFAULT CURRENT_TIMESTAMP COMMENT 'Creation timestamp',
    modified_by VARCHAR(100) COMMENT 'Last modified by',
    last_modified DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT 'Last modification time',

    -- Indexes for performance
    INDEX idx_employee_name (employee_name, surname),
    INDEX idx_employer (employer),
    INDEX idx_employment_status (employment_status),
    INDEX idx_clock_in_ref (clock_in_ref)
) ENGINE=InnoDB COMMENT='Employee master data';

-- ============================================================
-- TABLE: mastersalary (MASTERSALARY sheet)
-- ============================================================
CREATE TABLE IF NOT EXISTS mastersalary (
    -- Primary Key
    record_number INT AUTO_INCREMENT PRIMARY KEY COMMENT 'Unique payslip number',

    -- Employee Reference
    employee_id VARCHAR(50) NOT NULL COMMENT 'FK to employees.id',
    employee_name VARCHAR(100) NOT NULL COMMENT 'Employee name (denormalized)',
    employer ENUM('SA Grinding Wheels', 'Scorpio Abrasives') NOT NULL COMMENT 'Employer (denormalized)',
    employment_status ENUM('Permanent', 'Temporary', 'Contract') NOT NULL COMMENT 'Employment status (denormalized)',

    -- Payslip Period
    week_ending DATE NOT NULL COMMENT 'Week ending date (Saturday)',

    -- Time Worked
    hours INT DEFAULT 0 COMMENT 'Standard hours worked',
    minutes INT DEFAULT 0 COMMENT 'Standard minutes worked',
    overtime_hours INT DEFAULT 0 COMMENT 'Overtime hours worked',
    overtime_minutes INT DEFAULT 0 COMMENT 'Overtime minutes worked',

    -- Pay Rate
    hourly_rate DECIMAL(10,2) NOT NULL COMMENT 'Hourly rate (denormalized)',

    -- Calculated Earnings
    standard_time DECIMAL(10,2) DEFAULT 0.00 COMMENT 'Standard time pay',
    overtime DECIMAL(10,2) DEFAULT 0.00 COMMENT 'Overtime pay (1.5x)',
    leave_pay DECIMAL(10,2) DEFAULT 0.00 COMMENT 'Leave pay',
    bonus_pay DECIMAL(10,2) DEFAULT 0.00 COMMENT 'Bonus pay',
    other_income DECIMAL(10,2) DEFAULT 0.00 COMMENT 'Other income',
    other_income_text VARCHAR(255) COMMENT 'Description of other income',
    gross_salary DECIMAL(10,2) DEFAULT 0.00 COMMENT 'Total gross salary',

    -- Deductions
    uif DECIMAL(10,2) DEFAULT 0.00 COMMENT 'UIF (1% for permanent only)',
    other_deductions DECIMAL(10,2) DEFAULT 0.00 COMMENT 'Other deductions',
    other_deductions_text VARCHAR(255) COMMENT 'Description of other deductions',

    -- Loan Fields
    current_loan_balance DECIMAL(10,2) DEFAULT 0.00 COMMENT 'Loan balance before this payslip',
    loan_deduction_this_week DECIMAL(10,2) DEFAULT 0.00 COMMENT 'Loan repayment amount',
    new_loan_this_week DECIMAL(10,2) DEFAULT 0.00 COMMENT 'New loan disbursed',
    loan_disbursement_type ENUM('With Salary', 'Separate', 'Repayment', '') COMMENT 'How loan is paid',
    updated_loan_balance DECIMAL(10,2) DEFAULT 0.00 COMMENT 'Loan balance after this payslip',
    loan_repayment_logged BOOLEAN DEFAULT FALSE COMMENT 'Whether loan sync completed',

    -- Calculated Totals
    total_deductions DECIMAL(10,2) DEFAULT 0.00 COMMENT 'Total deductions',
    net_salary DECIMAL(10,2) DEFAULT 0.00 COMMENT 'Net salary (Gross - Deductions)',
    paid_to_account DECIMAL(10,2) DEFAULT 0.00 COMMENT 'CRITICAL: Amount paid to employee account',

    -- PDF Generation
    filename VARCHAR(255) COMMENT 'PDF filename',
    filelink TEXT COMMENT 'PDF URL/path',

    -- Notes
    notes TEXT COMMENT 'Additional notes',

    -- Audit Fields
    user VARCHAR(100) COMMENT 'Created by user',
    timestamp DATETIME DEFAULT CURRENT_TIMESTAMP COMMENT 'Creation timestamp',
    modified_by VARCHAR(100) COMMENT 'Last modified by',
    last_modified DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT 'Last modification time',

    -- Foreign Key
    FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE RESTRICT ON UPDATE CASCADE,

    -- Indexes for performance (CRITICAL for fixing the 2-minute problem!)
    INDEX idx_employee_id (employee_id),
    INDEX idx_week_ending (week_ending),
    INDEX idx_employee_week (employee_id, week_ending),
    INDEX idx_employer (employer),
    INDEX idx_timestamp (timestamp)
) ENGINE=InnoDB COMMENT='Payslip records (weekly pay)';

-- ============================================================
-- TABLE: employee_loans (EmployeeLoans sheet)
-- ============================================================
CREATE TABLE IF NOT EXISTS employee_loans (
    -- Primary Key
    id INT AUTO_INCREMENT PRIMARY KEY COMMENT 'Unique loan transaction ID',

    -- Employee Reference
    employee_id VARCHAR(50) NOT NULL COMMENT 'FK to employees.id',

    -- Link to Payslip (if applicable)
    salary_link INT COMMENT 'FK to mastersalary.record_number (NULL for manual entries)',

    -- Transaction Details
    transaction_date DATE NOT NULL COMMENT 'Date of transaction',
    loan_type ENUM('Disbursement', 'Repayment') NOT NULL COMMENT 'Transaction type',
    amount DECIMAL(10,2) NOT NULL COMMENT 'Transaction amount',
    balance_after DECIMAL(10,2) NOT NULL COMMENT 'Loan balance after this transaction',

    -- Notes
    notes TEXT COMMENT 'Transaction notes',

    -- Audit Fields
    timestamp DATETIME DEFAULT CURRENT_TIMESTAMP COMMENT 'Creation timestamp',
    user VARCHAR(100) COMMENT 'Created by user',

    -- Foreign Keys
    FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE RESTRICT ON UPDATE CASCADE,
    FOREIGN KEY (salary_link) REFERENCES mastersalary(record_number) ON DELETE SET NULL ON UPDATE CASCADE,

    -- Indexes for performance (CRITICAL for getCurrentLoanBalance speed!)
    INDEX idx_employee_date (employee_id, transaction_date, timestamp),
    INDEX idx_salary_link (salary_link)
) ENGINE=InnoDB COMMENT='Employee loan transaction ledger';

-- ============================================================
-- TABLE: leave_records (LEAVE sheet)
-- ============================================================
CREATE TABLE IF NOT EXISTS leave_records (
    -- Primary Key
    id INT AUTO_INCREMENT PRIMARY KEY COMMENT 'Unique leave record ID',

    -- Employee Reference
    employee_id VARCHAR(50) NOT NULL COMMENT 'FK to employees.id',
    employee_name VARCHAR(100) NOT NULL COMMENT 'Employee name (denormalized)',

    -- Leave Details
    leave_date DATE NOT NULL COMMENT 'Date of leave',
    leave_reason ENUM('AWOL', 'Sick Leave', 'Annual Leave', 'Unpaid Leave') NOT NULL COMMENT 'Reason for leave',

    -- Notes
    notes TEXT COMMENT 'Additional notes',

    -- Audit Fields
    user VARCHAR(100) COMMENT 'Created by user',
    timestamp DATETIME DEFAULT CURRENT_TIMESTAMP COMMENT 'Creation timestamp',

    -- Foreign Key
    FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE RESTRICT ON UPDATE CASCADE,

    -- Indexes
    INDEX idx_employee_id (employee_id),
    INDEX idx_leave_date (leave_date),
    INDEX idx_employee_date (employee_id, leave_date)
) ENGINE=InnoDB COMMENT='Leave tracking records';

-- ============================================================
-- TABLE: pending_timesheets (PendingTimesheets sheet)
-- ============================================================
CREATE TABLE IF NOT EXISTS pending_timesheets (
    -- Primary Key
    id INT AUTO_INCREMENT PRIMARY KEY COMMENT 'Unique timesheet ID',

    -- Employee Reference
    employee_id VARCHAR(50) NOT NULL COMMENT 'FK to employees.id',
    employee_name VARCHAR(100) NOT NULL COMMENT 'Employee name (denormalized)',
    clock_in_ref VARCHAR(50) NOT NULL COMMENT 'Clock-in reference',

    -- Timesheet Period
    week_ending DATE NOT NULL COMMENT 'Week ending date',

    -- Time Data
    total_hours INT DEFAULT 0 COMMENT 'Total hours worked',
    total_minutes INT DEFAULT 0 COMMENT 'Total minutes worked',
    overtime_hours INT DEFAULT 0 COMMENT 'Overtime hours',
    overtime_minutes INT DEFAULT 0 COMMENT 'Overtime minutes',

    -- Status
    status ENUM('Pending', 'Approved', 'Rejected') DEFAULT 'Pending' COMMENT 'Approval status',

    -- Audit Fields
    timestamp DATETIME DEFAULT CURRENT_TIMESTAMP COMMENT 'Creation timestamp',
    approved_by VARCHAR(100) COMMENT 'Approved/Rejected by user',
    approved_at DATETIME COMMENT 'Approval/Rejection timestamp',

    -- Foreign Key
    FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE RESTRICT ON UPDATE CASCADE,

    -- Indexes
    INDEX idx_employee_id (employee_id),
    INDEX idx_week_ending (week_ending),
    INDEX idx_status (status),
    INDEX idx_employee_week (employee_id, week_ending)
) ENGINE=InnoDB COMMENT='Pending timesheet approvals';

-- ============================================================
-- TABLE: raw_clock_data (RAW_CLOCK_DATA sheet)
-- ============================================================
CREATE TABLE IF NOT EXISTS raw_clock_data (
    -- Primary Key
    id INT AUTO_INCREMENT PRIMARY KEY COMMENT 'Unique clock-in record ID',

    -- Clock Data
    clock_in_ref VARCHAR(50) NOT NULL COMMENT 'Employee clock-in reference',
    clock_timestamp DATETIME NOT NULL COMMENT 'Clock-in/out timestamp',
    clock_type ENUM('In', 'Out') NOT NULL COMMENT 'Clock in or out',

    -- Import Tracking
    import_batch_id INT COMMENT 'FK to clock_in_imports.id',

    -- Indexes
    INDEX idx_clock_in_ref (clock_in_ref),
    INDEX idx_clock_timestamp (clock_timestamp),
    INDEX idx_import_batch (import_batch_id)
) ENGINE=InnoDB COMMENT='Raw clock-in/out data';

-- ============================================================
-- TABLE: clock_in_imports (CLOCK_IN_IMPORTS sheet)
-- ============================================================
CREATE TABLE IF NOT EXISTS clock_in_imports (
    -- Primary Key
    id INT AUTO_INCREMENT PRIMARY KEY COMMENT 'Unique import batch ID',

    -- Import Details
    import_date DATETIME DEFAULT CURRENT_TIMESTAMP COMMENT 'When data was imported',
    records_imported INT DEFAULT 0 COMMENT 'Number of records imported',
    imported_by VARCHAR(100) COMMENT 'User who imported',

    -- Notes
    notes TEXT COMMENT 'Import notes',

    -- Indexes
    INDEX idx_import_date (import_date)
) ENGINE=InnoDB COMMENT='Clock-in import batch tracking';

-- ============================================================
-- TABLE: user_config (UserConfig sheet)
-- ============================================================
CREATE TABLE IF NOT EXISTS user_config (
    -- Primary Key
    id INT AUTO_INCREMENT PRIMARY KEY COMMENT 'Unique user ID',

    -- User Details
    email VARCHAR(255) NOT NULL UNIQUE COMMENT 'User email (Google account)',
    full_name VARCHAR(100) COMMENT 'Full name',
    role ENUM('Admin', 'Manager', 'User') DEFAULT 'User' COMMENT 'User role',

    -- Authentication
    password_hash VARCHAR(255) COMMENT 'Password hash (for non-Google auth)',
    is_active BOOLEAN DEFAULT TRUE COMMENT 'Whether user account is active',

    -- Audit Fields
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP COMMENT 'Account creation',
    last_login DATETIME COMMENT 'Last login timestamp',

    -- Indexes
    INDEX idx_email (email),
    INDEX idx_is_active (is_active)
) ENGINE=InnoDB COMMENT='Authorized users and authentication';

-- ============================================================
-- VIEWS - For backward compatibility and common queries
-- ============================================================

-- View: Current loan balances for all employees
CREATE OR REPLACE VIEW v_current_loan_balances AS
SELECT
    e.id AS employee_id,
    e.employee_name,
    e.surname,
    COALESCE(
        (SELECT balance_after
         FROM employee_loans
         WHERE employee_id = e.id
         ORDER BY transaction_date DESC, timestamp DESC
         LIMIT 1),
        0
    ) AS current_balance
FROM employees e
WHERE e.termination_date IS NULL OR e.termination_date > CURDATE();

-- View: Recent payslips (last 100)
CREATE OR REPLACE VIEW v_recent_payslips AS
SELECT
    record_number,
    employee_id,
    employee_name,
    week_ending,
    gross_salary,
    net_salary,
    paid_to_account,
    timestamp
FROM mastersalary
ORDER BY record_number DESC
LIMIT 100;

-- View: Active employees
CREATE OR REPLACE VIEW v_active_employees AS
SELECT *
FROM employees
WHERE termination_date IS NULL OR termination_date > CURDATE();

-- ============================================================
-- INITIAL DATA - Insert default admin user
-- ============================================================

-- Default admin user (update email as needed)
INSERT INTO user_config (email, full_name, role, is_active)
VALUES ('admin@sagrinding.co.za', 'System Administrator', 'Admin', TRUE)
ON DUPLICATE KEY UPDATE email = email;

-- ============================================================
-- PERFORMANCE OPTIMIZATION
-- ============================================================

-- Enable query cache for better performance
SET GLOBAL query_cache_size = 16777216; -- 16MB
SET GLOBAL query_cache_type = 1;

-- ============================================================
-- SCHEMA COMPLETE
-- ============================================================
-- This schema is optimized for:
-- 1. Fast lookups (indexed foreign keys)
-- 2. Data integrity (foreign key constraints)
-- 3. Query performance (composite indexes on common queries)
-- 4. Audit trail (timestamp fields)
-- 5. Scalability (InnoDB engine with proper indexes)
--
-- Expected performance improvements:
-- - Edit payslip: 2 minutes → <1 second (120x faster!)
-- - Get loan balance: Full sheet load → Indexed query (<20ms)
-- - List payslips: Full sheet load → Filtered query (<50ms)
-- ============================================================
