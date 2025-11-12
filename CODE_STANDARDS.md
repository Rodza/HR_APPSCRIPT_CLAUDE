# Code Standards & Conventions

**Coding guidelines for HR System development**
**Version**: 1.0 | **Date**: October 28, 2025

---

## File Structure

### Apps Script Files (.gs)
```
Code.gs           - Main entry points (doGet, doPost, onChange)
Config.gs         - Constants and configuration
Utils.gs          - Shared utility functions
Employees.gs      - Employee management
Leave.gs          - Leave tracking
Loans.gs          - Loan management
Timesheets.gs     - Time approval workflow
Payroll.gs        - Payroll processing
Reports.gs        - Report generation
Triggers.gs       - Trigger management
```

### HTML Files (.html)
```
Dashboard.html           - Main container with navigation
EmployeeForm.html        - Add/Edit employee
EmployeeList.html        - Employee list view
LeaveForm.html          - Leave entry
LeaveList.html          - Leave history
LoanForm.html           - Loan transaction entry
LoanList.html           - Loan history
LoanBalanceWidget.html  - Reusable loan balance display
TimesheetImport.html    - Excel upload
TimesheetApproval.html  - Pending timesheet review
PayrollForm.html        - Payslip entry
PayrollList.html        - Payslip list
ReportsMenu.html        - Report generation menu
```

---

## Naming Conventions

### Functions
```javascript
// Use camelCase for function names
function getEmployeeById(id) { ... }
function createPayslip(data) { ... }
function validateEmployee(data) { ... }

// Prefix with verb (get, set, create, update, delete, validate, calculate)
function calculateGrossSalary(data) { ... }
function deleteOldReportFile(reportType) { ... }

// Boolean functions start with is/has/can
function isEmployeeActive(id) { ... }
function hasOutstandingLoans(employeeId) { ... }
function canApproveTimesheet(id) { ... }
```

### Variables
```javascript
// Use camelCase for variables
const employeeName = "John Smith";
const hourlyRate = 33.96;
const weekEnding = new Date();

// Constants in UPPER_SNAKE_CASE
const EMPLOYER_LIST = ["SA Grinding Wheels", "Scorpio Abrasives"];
const MAX_RETRY_COUNT = 3;
const DEFAULT_TIMEOUT = 30000;

// Arrays pluralized
const employees = [];
const payslips = [];
const loanTransactions = [];
```

### Sheet References
```javascript
// Use descriptive names matching purpose
const empSheet = sheets.empdetails;      // Employee details
const salarySheet = sheets.salary;       // Payroll records
const loanSheet = sheets.loans;          // Loan transactions
const leaveSheet = sheets.leave;         // Leave records
```

---

## Error Handling

### Standard Pattern
```javascript
function addEmployee(data) {
  try {
    // Validate input
    validateEmployee(data);
    
    // Get sheet reference
    const sheets = getSheets();
    const empSheet = sheets.empdetails;
    if (!empSheet) {
      throw new Error('Employee sheet not found');
    }
    
    // Perform operation
    const result = performOperation(empSheet, data);
    
    // Log success
    Logger.log('✅ Employee added successfully: ' + result.id);
    
    return { success: true, data: result };
    
  } catch (error) {
    // Log error with context
    Logger.log('❌ ERROR in addEmployee: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    
    // Return error response
    return { success: false, error: error.message };
  }
}
```

### Validation Pattern
```javascript
function validateEmployee(data) {
  const errors = [];
  
  // Check required fields
  if (!data.employeeName) errors.push('Employee Name is required');
  if (!data.surname) errors.push('Surname is required');
  if (!data.employer) errors.push('Employer is required');
  if (!data.hourlyRate || data.hourlyRate <= 0) {
    errors.push('Hourly Rate must be greater than 0');
  }
  
  // Check unique constraints
  if (data.idNumber && isIdNumberUsed(data.idNumber, data.id)) {
    errors.push('ID Number already exists');
  }
  
  // Throw if errors found
  if (errors.length > 0) {
    throw new Error(errors.join('; '));
  }
}
```

---

## Logging Standards

### Log Levels
```javascript
// ✅ Success (green checkmark)
Logger.log('✅ Operation completed successfully');

// ℹ️ Info (information)
Logger.log('ℹ️ Processing 15 records...');

// ⚠️ Warning (warning sign)
Logger.log('⚠️ Loan deduction exceeds balance');

// ❌ Error (red X)
Logger.log('❌ ERROR: Failed to create payslip');

// 🔍 Debug (magnifying glass)
Logger.log('🔍 DEBUG: Current balance = ' + balance);
```

### Log Format
```javascript
// Start of function
Logger.log('\n========== FUNCTION NAME ==========');
Logger.log('Input: ' + JSON.stringify(inputData));

// During processing
Logger.log('Step 1: Validating data...');
Logger.log('Step 2: Getting sheet reference...');
Logger.log('Step 3: Performing operation...');

// Success
Logger.log('✅ Result: ' + JSON.stringify(result));
Logger.log('========== FUNCTION COMPLETE ==========\n');

// Error
Logger.log('❌ ERROR: ' + error.message);
Logger.log('Stack: ' + error.stack);
Logger.log('========== FUNCTION FAILED ==========\n');
```

---

## Data Access Patterns

### Sheet Locator (Utils.gs)
```javascript
function getSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  const sheetMap = {};
  
  Logger.log('📋 Scanning spreadsheet for sheets...');
  
  allSheets.forEach(sheet => {
    const originalName = sheet.getName();
    const name = originalName.toLowerCase().replace(/\s+/g, '');
    
    if (name === 'mastersalary' || name === 'salaryrecords') {
      sheetMap.salary = sheet;
    } else if (name === 'employeeloans' || name === 'loantransactions') {
      sheetMap.loans = sheet;
    } else if (name === 'empdetails' || name === 'employeedetails') {
      sheetMap.empdetails = sheet;
    } else if (name === 'leave') {
      sheetMap.leave = sheet;
    } else if (name === 'pendingtimesheets') {
      sheetMap.pending = sheet;
    }
  });
  
  Logger.log('📊 Found sheets: ' + Object.keys(sheetMap).join(', '));
  
  return sheetMap;
}
```

### Reading Data
```javascript
// Get all data (avoid for large tables)
const data = sheet.getDataRange().getValues();

// Get specific range
const data = sheet.getRange(2, 1, lastRow - 1, numColumns).getValues();

// Get headers separately
const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

// Find column index
const nameCol = headers.indexOf('EMPLOYEE NAME');
```

### Writing Data
```javascript
// Append new row
sheet.appendRow([value1, value2, value3, ...]);

// Update specific cell
sheet.getRange(rowIndex, colIndex).setValue(newValue);

// Update multiple cells
sheet.getRange(rowIndex, 1, 1, numCols).setValues([[val1, val2, val3, ...]]);

// Batch updates (more efficient)
const updates = [];
updates.push([val1, val2, val3]);
updates.push([val4, val5, val6]);
sheet.getRange(startRow, 1, updates.length, numCols).setValues(updates);
SpreadsheetApp.flush();  // Force immediate write
```

---

## Calculation Functions

### Standard Pattern
```javascript
function calculatePayslip(data) {
  Logger.log('🔢 Calculating payslip for: ' + data.employeeName);
  
  // 1. Standard Time
  const standardTime = (data.hours * data.hourlyRate) + 
                       ((data.hourlyRate / 60) * data.minutes);
  Logger.log('  Standard Time: R' + standardTime.toFixed(2));
  
  // 2. Overtime (1.5x)
  const overtime = (data.overtimeHours * data.hourlyRate * 1.5) + 
                   ((data.hourlyRate / 60) * data.overtimeMinutes * 1.5);
  Logger.log('  Overtime: R' + overtime.toFixed(2));
  
  // 3. Gross Salary
  const grossSalary = standardTime + overtime + 
                     (data.leavePay || 0) + 
                     (data.bonusPay || 0) + 
                     (data.otherIncome || 0);
  Logger.log('  Gross Salary: R' + grossSalary.toFixed(2));
  
  // 4. UIF (1% if permanent)
  const uif = (data.employmentStatus === 'Permanent') ? 
              (grossSalary * 0.01) : 0;
  Logger.log('  UIF: R' + uif.toFixed(2));
  
  // 5. Total Deductions
  const totalDeductions = uif + (data.otherDeductions || 0) + 
                         (data.loanDeduction || 0);
  Logger.log('  Total Deductions: R' + totalDeductions.toFixed(2));
  
  // 6. Net Salary
  const netSalary = grossSalary - totalDeductions;
  Logger.log('  Net Salary: R' + netSalary.toFixed(2));
  
  // 7. Paid to Account
  const newLoanAmount = (data.loanDisbursementType === 'With Salary') ? 
                        (data.newLoan || 0) : 0;
  const paidToAccount = netSalary - (data.loanDeduction || 0) + newLoanAmount;
  Logger.log('  Paid to Account: R' + paidToAccount.toFixed(2));
  
  return {
    standardTime: parseFloat(standardTime.toFixed(2)),
    overtime: parseFloat(overtime.toFixed(2)),
    grossSalary: parseFloat(grossSalary.toFixed(2)),
    uif: parseFloat(uif.toFixed(2)),
    totalDeductions: parseFloat(totalDeductions.toFixed(2)),
    netSalary: parseFloat(netSalary.toFixed(2)),
    paidToAccount: parseFloat(paidToAccount.toFixed(2))
  };
}
```

### Rounding Rules
```javascript
// Always round currency to 2 decimal places
const amount = parseFloat(value.toFixed(2));

// For display
function formatCurrency(amount) {
  return 'R' + amount.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}
// Example: formatCurrency(1341.42) → "R1,341.42"
```

---

## Date Handling

### Standard Pattern
```javascript
// Parse date from various sources
function parseDate(dateValue) {
  if (dateValue instanceof Date) {
    return dateValue;
  }
  if (typeof dateValue === 'string') {
    return new Date(dateValue);
  }
  if (typeof dateValue === 'number') {
    return new Date(dateValue);
  }
  throw new Error('Invalid date format');
}

// Format for display
function formatDate(date) {
  const d = parseDate(date);
  const day = d.getDate();
  const month = d.toLocaleString('en-US', { month: 'long' });
  const year = d.getFullYear();
  return `${day} ${month} ${year}`;
}
// Example: formatDate(new Date('2025-10-17')) → "17 October 2025"

// Format for storage (ISO string)
function formatDateISO(date) {
  return parseDate(date).toISOString();
}
// Example: "2025-10-17T00:00:00.000Z"

// Week ending date (always use Friday)
function getWeekEnding(date) {
  const d = parseDate(date);
  const dayOfWeek = d.getDay();
  const daysUntilFriday = (5 - dayOfWeek + 7) % 7;
  const friday = new Date(d);
  friday.setDate(d.getDate() + daysUntilFriday);
  return friday;
}
```

---

## Audit Trail Pattern

### Standard Implementation
```javascript
function addAuditFields(data, isCreate = true) {
  const user = Session.getActiveUser().getEmail();
  const now = new Date();
  
  if (isCreate) {
    data.USER = user;
    data.TIMESTAMP = now;
  }
  
  data.MODIFIED_BY = user;
  data.LAST_MODIFIED = now;
  
  return data;
}

// Usage in create function
function createEmployee(data) {
  const enrichedData = addAuditFields(data, true);
  // ... save to sheet
}

// Usage in update function
function updateEmployee(id, data) {
  const enrichedData = addAuditFields(data, false);
  // ... save to sheet
}
```

---

## UI/Server Communication

### Google Apps Script Pattern
```javascript
// Server-side function (in .gs file)
function serverFunction(param1, param2) {
  try {
    const result = performOperation(param1, param2);
    return { success: true, data: result };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// Client-side call (in HTML file)
<script>
  function callServerFunction() {
    google.script.run
      .withSuccessHandler(onSuccess)
      .withFailureHandler(onFailure)
      .serverFunction(param1, param2);
  }
  
  function onSuccess(response) {
    if (response.success) {
      console.log('Success:', response.data);
      // Update UI
    } else {
      console.error('Error:', response.error);
      // Show error message
    }
  }
  
  function onFailure(error) {
    console.error('Failure:', error.message);
    // Show error message
  }
</script>
```

---

## Performance Best Practices

### Batch Operations
```javascript
// ❌ Bad: Multiple individual writes
for (let i = 0; i < employees.length; i++) {
  sheet.getRange(i + 2, 1).setValue(employees[i].name);
  sheet.getRange(i + 2, 2).setValue(employees[i].rate);
}

// ✅ Good: Single batch write
const rows = employees.map(emp => [emp.name, emp.rate]);
sheet.getRange(2, 1, rows.length, 2).setValues(rows);
SpreadsheetApp.flush();
```

### Caching
```javascript
// ❌ Bad: Multiple calls to getSheets()
function process() {
  const emp1 = getEmployeeById('e1');  // Calls getSheets()
  const emp2 = getEmployeeById('e2');  // Calls getSheets() again
}

// ✅ Good: Cache sheet references
function process() {
  const sheets = getSheets();  // Call once
  const emp1 = getEmployeeByIdWithSheets(sheets, 'e1');
  const emp2 = getEmployeeByIdWithSheets(sheets, 'e2');
}
```

### Large Dataset Handling
```javascript
// For processing large tables (MASTERSALARY)
function processLargeTable() {
  const sheet = getSheets().salary;
  const lastRow = sheet.getLastRow();
  const batchSize = 1000;
  
  for (let startRow = 2; startRow <= lastRow; startRow += batchSize) {
    const numRows = Math.min(batchSize, lastRow - startRow + 1);
    const data = sheet.getRange(startRow, 1, numRows, 10).getValues();
    
    // Process batch
    processBatch(data);
    
    // Prevent timeout
    Utilities.sleep(100);
  }
}
```

---

## Testing Conventions

### Unit Test Pattern
```javascript
// Test function naming: test_FunctionName_Scenario
function test_calculatePayslip_StandardTime() {
  const testData = {
    hours: 39,
    minutes: 30,
    hourlyRate: 33.96,
    overtimeHours: 0,
    overtimeMinutes: 0,
    leavePay: 0,
    bonusPay: 0,
    otherIncome: 0,
    employmentStatus: 'Permanent',
    otherDeductions: 0,
    loanDeduction: 0
  };
  
  const result = calculatePayslip(testData);
  
  // Expected: (39 × 33.96) + ((33.96/60) × 30) = 1324.44 + 16.98 = 1341.42
  const expected = 1341.42;
  const actual = result.standardTime;
  
  if (Math.abs(actual - expected) < 0.01) {
    Logger.log('✅ TEST PASSED: Standard time calculation correct');
  } else {
    Logger.log('❌ TEST FAILED: Expected ' + expected + ', got ' + actual);
  }
}

// Run all tests
function runAllTests() {
  test_calculatePayslip_StandardTime();
  test_calculatePayslip_Overtime();
  test_calculatePayslip_UIF();
  // ... more tests
}
```

---

## Documentation Standards

### Function Documentation
```javascript
/**
 * Creates a new employee record
 * 
 * @param {Object} data - Employee data object
 * @param {string} data.employeeName - First name (required)
 * @param {string} data.surname - Last name (required)
 * @param {string} data.employer - Employer name (required)
 * @param {number} data.hourlyRate - Hourly rate (required, must be > 0)
 * @param {string} data.idNumber - SA ID number (required, must be unique)
 * @param {string} data.contactNumber - Phone number (required)
 * @param {string} data.address - Full address (required)
 * @param {string} data.employmentStatus - Employment status (required)
 * @param {string} data.clockInRef - Clock number (required, must be unique)
 * 
 * @returns {Object} Result object with success flag and data/error
 * @returns {boolean} returns.success - Whether operation succeeded
 * @returns {Object} [returns.data] - Employee data with generated ID
 * @returns {string} [returns.error] - Error message if failed
 * 
 * @example
 * const result = addEmployee({
 *   employeeName: 'John',
 *   surname: 'Smith',
 *   employer: 'SA Grinding Wheels',
 *   hourlyRate: 40.00,
 *   idNumber: '9501015800081',
 *   contactNumber: '0821234567',
 *   address: '123 Main St',
 *   employmentStatus: 'Permanent',
 *   clockInRef: '002'
 * });
 */
function addEmployee(data) {
  // Implementation
}
```

### Code Comments
```javascript
// Single-line comment for simple explanation
const rate = employee.hourlyRate;

/* 
 * Multi-line comment for complex logic explanation
 * This calculation handles the edge case where...
 */
const adjusted = calculateAdjusted(rate);

// TODO: Implement validation for X
// FIXME: Handle edge case when Y
// NOTE: This assumes Z
// HACK: Temporary workaround until we fix the root cause
```

---

## Security Guidelines

### Input Sanitization
```javascript
function sanitizeInput(input) {
  if (typeof input === 'string') {
    // Remove HTML tags
    input = input.replace(/<[^>]*>/g, '');
    // Trim whitespace
    input = input.trim();
  }
  return input;
}

// Apply to all user inputs
data.employeeName = sanitizeInput(data.employeeName);
```

### Permission Checks
```javascript
function requireAdmin() {
  const user = Session.getActiveUser().getEmail();
  const adminEmails = ['owner@sagrinding.co.za', 'admin@sagrinding.co.za'];
  
  if (!adminEmails.includes(user)) {
    throw new Error('Unauthorized: Admin access required');
  }
}

// Use at start of sensitive functions
function deleteEmployee(id) {
  requireAdmin();
  // ... proceed with deletion
}
```

---

## Git Commit Messages (if using version control)

### Format
```
<type>(<scope>): <subject>

<body>

<footer>
```

### Types
- **feat**: New feature
- **fix**: Bug fix
- **refactor**: Code refactoring
- **docs**: Documentation changes
- **test**: Adding tests
- **style**: Code style changes (formatting)

### Examples
```
feat(employees): Add employee transfer between employers

Implemented functionality to change employee's employer field
with proper validation and audit trail.

Closes #12

---

fix(payroll): Correct UIF calculation for temporary employees

UIF was being applied to temporary employees when it should
only apply to permanent staff.

Bug reported by user on 2025-10-20
```

---

## Code Review Checklist

Before submitting code for review:

- [ ] All functions have proper error handling (try-catch)
- [ ] Logging statements follow standard format
- [ ] No hardcoded values (use Config.gs constants)
- [ ] Input validation on all user-supplied data
- [ ] Audit fields added (USER, TIMESTAMP, MODIFIED_BY, LAST_MODIFIED)
- [ ] Function documentation complete
- [ ] No console.log statements (use Logger.log)
- [ ] Currency values rounded to 2 decimals
- [ ] Date handling uses standard utility functions
- [ ] Batch operations used for multiple writes
- [ ] Test cases pass
- [ ] No duplicate code (refactor into utilities)

---

**Document End**
