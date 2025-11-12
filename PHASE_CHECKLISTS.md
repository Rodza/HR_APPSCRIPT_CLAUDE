# Phase Completion Checklists

**Clear criteria for phase sign-off**
**Version**: 1.0 | **Date**: October 28, 2025

---

## Phase 1: Employee Database ✓

**Goal**: Complete employee master data management
**Timeline**: Week 1-2 (3-5 days)
**Status**: 🔵 Starting

### Code Deliverables
- [ ] **Config.gs** created with all constants
  - [ ] EMPLOYER_LIST = ["SA Grinding Wheels", "Scorpio Abrasives"]
  - [ ] EMPLOYMENT_STATUS_LIST = ["Permanent", "Temporary", "Contract"]
  - [ ] Required fields list defined
  
- [ ] **Utils.gs** created with helper functions
  - [ ] getSheets() - Sheet locator
  - [ ] getCurrentUser() - Get active user email
  - [ ] formatDate(date) - Format date for display
  - [ ] formatCurrency(amount) - Format currency with R symbol
  - [ ] generateUUID() - Generate unique IDs
  - [ ] validateSAIdNumber(id) - Validate SA ID format
  - [ ] validatePhoneNumber(phone) - Validate phone format
  
- [ ] **Employees.gs** created with CRUD operations
  - [ ] addEmployee(data) - Create new employee
  - [ ] updateEmployee(id, data) - Update existing employee
  - [ ] getEmployeeById(id) - Get single employee
  - [ ] getEmployeeByName(refName) - Lookup by name
  - [ ] listEmployees(filters) - Get all/filtered employees
  - [ ] terminateEmployee(id, terminationDate) - Set termination
  - [ ] validateEmployee(data) - Validation logic
  - [ ] Helper functions: isIdNumberUsed(), isClockInRefUsed()

### UI Deliverables
- [ ] **Dashboard.html** - Basic shell with navigation
  - [ ] Header with app title
  - [ ] Navigation menu (Employees, Leave, Loans, Timesheets, Payroll, Reports)
  - [ ] Content area
  - [ ] Basic CSS styling (Tailwind or Bootstrap)
  
- [ ] **EmployeeForm.html** - Add/Edit employee form
  - [ ] All required fields with labels
  - [ ] Employer dropdown (2 options)
  - [ ] Employment status dropdown (3 options)
  - [ ] Field validation (client-side)
  - [ ] Submit button
  - [ ] Cancel button
  - [ ] Success/error messages
  
- [ ] **EmployeeList.html** - Employee list view
  - [ ] Table with key columns (Name, Employer, Rate, Status, Clock#)
  - [ ] Search box (filter by name)
  - [ ] Employer filter dropdown
  - [ ] Edit button per row
  - [ ] View button per row
  - [ ] Add New Employee button
  - [ ] Pagination (if > 50 records)

### Data Validation
- [ ] Required fields enforced (9 mandatory fields)
- [ ] Hourly rate must be > 0
- [ ] ID Number must be unique
- [ ] ClockInRef must be unique
- [ ] Phone number validates as SA format
- [ ] Employer must be one of 2 valid options
- [ ] Employment status must be one of 3 valid options

### Functional Tests
- [ ] Can add new employee with all required fields
- [ ] Cannot save with missing required fields (shows error)
- [ ] Cannot save with duplicate ID Number (shows error)
- [ ] Cannot save with duplicate Clock Number (shows error)
- [ ] Cannot save with hourly rate = 0 or negative (shows error)
- [ ] Can edit existing employee
- [ ] Changes are saved correctly
- [ ] Can view employee profile
- [ ] Can see employee list
- [ ] Can search by name (partial match)
- [ ] Can filter by employer (SA Grinding Wheels / Scorpio Abrasives)
- [ ] REFNAME auto-generates as "FirstName Surname"
- [ ] Audit trail records USER and TIMESTAMP on create
- [ ] Audit trail records MODIFIED_BY and LAST_MODIFIED on update

### Data Tests
- [ ] Add 3 test employees successfully
- [ ] Edit employee name - verify REFNAME updates
- [ ] Change employee's employer - verify update
- [ ] Try to add employee with duplicate ID - verify rejection
- [ ] Try to add employee with duplicate Clock# - verify rejection

### Performance Tests
- [ ] List loads in < 3 seconds with 18 employees
- [ ] Search responds in < 1 second
- [ ] Filter responds in < 1 second
- [ ] Form saves in < 2 seconds

### User Acceptance
- [ ] Owner can add employee without assistance
- [ ] Admin can edit employee without assistance
- [ ] Error messages are clear and helpful
- [ ] Form is intuitive (no training needed)

### Documentation
- [ ] Code comments on complex logic
- [ ] Function headers with JSDoc
- [ ] README with deployment instructions

### Sign-Off
- [ ] All checklist items completed
- [ ] Tested with real user
- [ ] No critical bugs
- [ ] Ready for Phase 2

**Phase 1 Complete**: _____ (Date) | **Signed**: _____ (Name)

---

## Phase 2: Leave & Loans ✓

**Goal**: Leave tracking + Loan ledger functional
**Timeline**: Week 3 (5-7 days)
**Status**: ⚪ Pending

### Code Deliverables
- [ ] **Leave.gs** created with functions
  - [ ] addLeave(data) - Record leave
  - [ ] getLeaveHistory(employeeId) - Get leave for employee
  - [ ] listLeave(filters) - Get all/filtered leave
  - [ ] validateLeave(data) - Validation logic
  
- [ ] **Loans.gs** created with functions
  - [ ] addLoanTransaction(data) - Record disbursement/repayment
  - [ ] getCurrentLoanBalance(employeeId) - Get latest balance
  - [ ] getLoanHistory(employeeId, startDate, endDate) - Get transactions
  - [ ] recalculateLoanBalances(employeeId) - Recalc balances chronologically
  - [ ] findLoanRecordBySalaryLink(recordNumber) - Find by payslip
  - [ ] validateLoan(data) - Validation logic

### UI Deliverables
- [ ] **LeaveForm.html** - Leave entry form
  - [ ] Employee dropdown/search
  - [ ] Start date picker
  - [ ] End date picker
  - [ ] Reason dropdown (4 options)
  - [ ] Notes field
  - [ ] Optional image upload
  - [ ] Total days auto-calculates
  - [ ] Submit/Cancel buttons
  
- [ ] **LeaveList.html** - Leave history view
  - [ ] Table with leave records
  - [ ] Filter by employee
  - [ ] Filter by date range
  - [ ] Filter by reason
  - [ ] View supporting documents
  
- [ ] **LoanForm.html** - Loan transaction form
  - [ ] Employee dropdown/search
  - [ ] Transaction type (Disbursement/Repayment)
  - [ ] Amount field
  - [ ] Disbursement mode dropdown (3 options)
  - [ ] Notes field
  - [ ] Current balance display (prominent)
  - [ ] Submit/Cancel buttons
  
- [ ] **LoanList.html** - Loan history view
  - [ ] Table with transactions
  - [ ] Filter by employee
  - [ ] Filter by date range
  - [ ] Filter by type
  - [ ] Shows running balance
  - [ ] Highlight current balance
  
- [ ] **LoanBalanceWidget.html** - Reusable balance display
  - [ ] Shows current balance for employee
  - [ ] Red if balance > 0
  - [ ] Green if balance = 0
  - [ ] Can be embedded in other views

### Data Validation
- [ ] Leave: End date >= Start date
- [ ] Leave: Total days calculated correctly
- [ ] Leave: Reason must be one of 4 options
- [ ] Loan: Amount must be > 0
- [ ] Loan: Type must match amount sign
- [ ] Loan: Cannot create duplicate SalaryLink

### Functional Tests
- [ ] Can record leave for employee
- [ ] Cannot save leave with end < start (shows error)
- [ ] Total days calculates correctly
- [ ] Can view leave history for employee
- [ ] Can filter leave by reason
- [ ] Can record loan disbursement
- [ ] Can record loan repayment
- [ ] Current balance calculates correctly
- [ ] Cannot create two loans for same payslip (shows error)
- [ ] Loan balance recalculation works (chronological)
- [ ] Can view loan history for employee
- [ ] Transaction list shows running balance correctly

### Data Tests
- [ ] Add leave record - verify saved
- [ ] Add loan disbursement R500 - verify balance = 500
- [ ] Add repayment R150 - verify balance = 350
- [ ] Add another repayment R350 - verify balance = 0
- [ ] Get loan balance - verify correct current balance
- [ ] Manually recalculate balances - verify chronological order maintained
- [ ] Create loan with past date - verify sorted correctly

### Integration Tests
- [ ] Leave form integrates with employee list
- [ ] Loan form integrates with employee list
- [ ] Loan balance widget shows correct data
- [ ] Employee profile shows current loan balance

### Performance Tests
- [ ] Leave list loads in < 3 seconds
- [ ] Loan list loads in < 3 seconds
- [ ] Balance calculation < 1 second per employee
- [ ] Recalculation of all balances < 10 seconds

### User Acceptance
- [ ] User can record leave easily
- [ ] User can record loan disbursement easily
- [ ] User can record loan repayment easily
- [ ] Loan balance is clearly visible
- [ ] Transaction history is easy to understand

### Sign-Off
- [ ] All checklist items completed
- [ ] Tested with real user
- [ ] No critical bugs
- [ ] Ready for Phase 3

**Phase 2 Complete**: _____ (Date) | **Signed**: _____ (Name)

---

## Phase 3: Time Approval System ✓

**Goal**: Time approval workflow functional
**Timeline**: Week 4 (3-5 days)
**Status**: ⚪ Pending

### Code Deliverables
- [ ] **Timesheets.gs** created with functions
  - [ ] importTimesheetExcel(fileBlob) - Import Excel file
  - [ ] parseExcelData(fileBlob) - Parse Excel to data array
  - [ ] addPendingTimesheet(data) - Add to pending table
  - [ ] updatePendingTimesheet(id, data) - Edit pending record
  - [ ] approveTimesheet(id) - Create payslip from pending
  - [ ] rejectTimesheet(id) - Mark as rejected
  - [ ] listPendingTimesheets(filters) - Get pending records
  - [ ] validateTimesheet(data) - Validation logic

### Table Setup
- [ ] **PendingTimesheets** table created in Google Sheets
  - [ ] All required columns
  - [ ] Proper headers
  - [ ] Data validation on columns

### UI Deliverables
- [ ] **TimesheetImport.html** - Excel upload interface
  - [ ] File upload button (accepts .xlsx, .xls, .csv)
  - [ ] Upload progress indicator
  - [ ] Success/error messages
  - [ ] Preview of imported data
  - [ ] Confirm import button
  
- [ ] **TimesheetApproval.html** - Pending timesheet review
  - [ ] Editable table showing all pending timesheets
  - [ ] Inline editing for hours/minutes
  - [ ] Employee name (not editable)
  - [ ] Week ending (not editable)
  - [ ] Notes field (editable)
  - [ ] Approve button per row
  - [ ] Reject button per row
  - [ ] Bulk approve selected
  - [ ] Filter by status (Pending/Approved/Rejected)
  - [ ] Sort by employee or date

### Excel Import
- [ ] Accepts standard format from HTML analyzer
- [ ] Parses columns correctly
- [ ] Handles missing values gracefully
- [ ] Validates employee names exist
- [ ] Validates dates are valid
- [ ] Validates time values are numeric and >= 0

### Data Validation
- [ ] Employee must exist in EMPLOYEE DETAILS
- [ ] Week ending must be valid date
- [ ] Cannot approve duplicate week for same employee
- [ ] All time values must be >= 0
- [ ] Status must be Pending/Approved/Rejected

### Functional Tests
- [ ] Can upload Excel file successfully
- [ ] Excel data parses correctly
- [ ] Data appears in pending table with Status = Pending
- [ ] Can edit hours in pending table
- [ ] Can edit minutes in pending table
- [ ] Can edit overtime in pending table
- [ ] Can edit notes in pending table
- [ ] Cannot edit employee name or week ending
- [ ] Click Approve - creates payslip in MASTERSALARY
- [ ] Approved record marked as Status = Approved
- [ ] Click Reject - marked as Status = Rejected
- [ ] Cannot approve if payslip already exists (shows error)
- [ ] Can filter pending by status
- [ ] Can bulk approve multiple records

### Data Tests
- [ ] Import Excel with 2 employees - verify both added to pending
- [ ] Edit hours for one record - verify saved
- [ ] Approve one record - verify payslip created
- [ ] Check approved record - verify Status = Approved
- [ ] Try to approve same week again - verify error
- [ ] Reject one record - verify Status = Rejected

### Integration Tests
- [ ] Approved timesheet creates correct payslip in MASTERSALARY
- [ ] Payslip has correct employee, week ending, hours
- [ ] Payslip calculations are correct
- [ ] If payslip has loan activity, auto-sync triggers (test in Phase 4)

### Performance Tests
- [ ] Import 18 records < 10 seconds
- [ ] Display pending table < 3 seconds
- [ ] Approve record < 5 seconds
- [ ] Bulk approve 18 records < 30 seconds

### User Acceptance
- [ ] User can upload Excel easily
- [ ] User can review and edit times easily
- [ ] Approve process is intuitive
- [ ] Error messages are clear
- [ ] Status tracking is visible

### Sign-Off
- [ ] All checklist items completed
- [ ] Tested with real user
- [ ] No critical bugs
- [ ] Ready for Phase 4

**Phase 3 Complete**: _____ (Date) | **Signed**: _____ (Name)

---

## Phase 4: Payroll Engine ✓ 🔴 CRITICAL

**Goal**: Complete payroll processing with auto-sync
**Timeline**: Week 5-6 (7-10 days)
**Status**: ⚪ Pending

### Code Deliverables
- [ ] **Payroll.gs** created with functions
  - [ ] createPayslip(data) - Create new payslip
  - [ ] updatePayslip(recordNumber, data) - Edit existing payslip
  - [ ] getPayslip(recordNumber) - Get single payslip
  - [ ] listPayslips(filters) - Get all/filtered payslips
  - [ ] calculatePayslip(data) - All calculation logic
  - [ ] generatePayslipPDF(recordNumber) - Generate PDF
  - [ ] validatePayslip(data) - Validation logic
  - [ ] checkDuplicatePayslip(employeeId, weekEnding) - Duplicate check
  
- [ ] **Triggers.gs** created with functions
  - [ ] installOnChangeTrigger() - Setup onChange trigger
  - [ ] onChange(e) - Auto-sync trigger handler
  - [ ] uninstallTriggers() - Cleanup triggers

### Calculation Functions
- [ ] calculateStandardTime(hours, minutes, rate) - Working correctly
- [ ] calculateOvertime(hours, minutes, rate) - 1.5x multiplier correct
- [ ] calculateGrossSalary(components) - Sum all earnings
- [ ] calculateUIF(gross, status) - 1% for permanent only
- [ ] calculateTotalDeductions(components) - Sum all deductions
- [ ] calculateNetSalary(gross, deductions) - Subtract deductions
- [ ] calculatePaidToAccount(net, loan, newLoan, type) - Final amount

### UI Deliverables
- [ ] **PayrollForm.html** - Payslip entry form
  - [ ] Employee dropdown/search
  - [ ] Week ending date picker
  - [ ] Hours and minutes fields (standard)
  - [ ] Overtime hours and minutes fields
  - [ ] Leave pay field
  - [ ] Bonus pay field
  - [ ] Other income field
  - [ ] Other deductions field
  - [ ] Other deductions text field
  - [ ] Loan section:
    - [ ] Current loan balance (prominent display)
    - [ ] Loan deduction this week field
    - [ ] New loan this week field
    - [ ] Loan disbursement type dropdown
    - [ ] Updated balance (auto-calculated)
  - [ ] Real-time calculation preview:
    - [ ] Standard time
    - [ ] Overtime
    - [ ] Gross salary
    - [ ] UIF
    - [ ] Total deductions
    - [ ] Net salary
    - [ ] **Paid to Account** (most prominent)
  - [ ] Notes field
  - [ ] Save button
  - [ ] Save & Generate PDF button
  - [ ] Cancel button
  
- [ ] **PayrollList.html** - Payslip list view
  - [ ] Table with key columns
  - [ ] Filter by employee
  - [ ] Filter by week ending
  - [ ] Filter by employer
  - [ ] Filter by date range
  - [ ] Edit button per row
  - [ ] View PDF button per row
  - [ ] Generate PDF button (if not yet generated)
  - [ ] Add New Payslip button

### PDF Generation
- [ ] PDF template matches sample #7916 format
- [ ] Header section with company details
- [ ] Earnings table (Normal Time, Overtime, Additional Pay)
- [ ] Deductions table (UIF, Other, Loans)
- [ ] Loan opening balance shown
- [ ] Loan repayment shown (as negative)
- [ ] Loan closing balance shown
- [ ] Notes section
- [ ] Summary section (Gross, Deductions, Net, Paid to Account)
- [ ] Currency formatted as "R#,###.##"
- [ ] PDF saved to Google Drive
- [ ] URL stored in FILELINK field
- [ ] Filename: {RECORDNUMBER}.pdf

### Auto-Sync Logic
- [ ] onChange trigger installed on MASTERSALARY sheet
- [ ] Trigger fires within 5 minutes of change
- [ ] Detects loan activity (LoanDeductionThisWeek > 0 OR NewLoanThisWeek > 0)
- [ ] Gets employee details (ID, Name)
- [ ] Gets current loan balance
- [ ] Checks if loan record exists (via SalaryLink)
- [ ] If exists: Updates existing loan record
- [ ] If not: Creates new loan record
- [ ] Recalculates balances chronologically
- [ ] Updates CurrentLoanBalance in MASTERSALARY
- [ ] Marks LoanRepaymentLogged = TRUE
- [ ] Handles errors gracefully (logs, doesn't crash)
- [ ] Script lock prevents concurrent execution

### Data Validation
- [ ] Cannot create duplicate payslip (same employee + week ending)
- [ ] Week ending must be valid date
- [ ] All time values must be >= 0
- [ ] All amounts must be >= 0
- [ ] Warning if loan deduction > current balance

### Functional Tests
- [ ] Can create payslip manually
- [ ] Can create payslip from approved timesheet
- [ ] All fields populate correctly
- [ ] Real-time calculations update as user types
- [ ] Standard time calculates correctly
- [ ] Overtime calculates correctly (1.5x rate)
- [ ] Gross salary correct
- [ ] UIF 1% applies to permanent employees only
- [ ] UIF does not apply to temporary/contract employees
- [ ] Total deductions correct
- [ ] Net salary correct
- [ ] **Paid to Account calculates correctly** (critical!)
- [ ] Cannot create duplicate payslip (shows error)
- [ ] Can edit existing payslip
- [ ] Changes save correctly
- [ ] Can view payslip details
- [ ] Can filter payslips by employee
- [ ] Can filter payslips by date
- [ ] Can filter payslips by employer

### PDF Tests
- [ ] Can generate PDF for payslip
- [ ] PDF matches sample #7916 format
- [ ] All values appear correctly
- [ ] Currency formatted correctly
- [ ] Company details correct
- [ ] PDF saves to Google Drive
- [ ] URL stored in FILELINK field
- [ ] Can view PDF from list
- [ ] Can regenerate PDF if needed

### Auto-Sync Tests (Critical)
- [ ] Create payslip with loan deduction - auto-sync triggers
- [ ] Loan repayment record created correctly
- [ ] Balance updated correctly
- [ ] Create payslip with new loan (With Salary) - auto-sync triggers
- [ ] Loan disbursement record created correctly
- [ ] Balance updated correctly
- [ ] Create payslip with new loan (Separate) - auto-sync triggers
- [ ] Loan disbursement record created correctly
- [ ] Balance updated correctly (not added to Paid to Account)
- [ ] Edit payslip loan amount - existing loan record updates (not duplicate)
- [ ] Edit payslip without loan activity - no loan record created
- [ ] Auto-sync doesn't create duplicates
- [ ] Balances recalculate chronologically
- [ ] CurrentLoanBalance in MASTERSALARY updates after sync
- [ ] Multiple rapid edits don't cause issues (script lock works)

### Calculation Validation (Against Historical Data)
- [ ] Test Case 1 (Sample #7916):
  - [ ] Input: 39h 30m @ R33.96, Permanent, Loan repayment R150
  - [ ] Standard Time: R1,341.42 ✓
  - [ ] Overtime: R0.00 ✓
  - [ ] Gross: R1,341.42 ✓
  - [ ] UIF: R13.41 ✓
  - [ ] Net: R1,328.01 ✓
  - [ ] Paid to Account: R1,178.01 ✓
- [ ] Test Case 2:
  - [ ] Input: 40h 0m @ R40.00, Permanent, New loan R500 (with salary)
  - [ ] Standard Time: R1,600.00 ✓
  - [ ] Overtime: R0.00 ✓
  - [ ] Gross: R1,600.00 ✓
  - [ ] UIF: R16.00 ✓
  - [ ] Net: R1,584.00 ✓
  - [ ] Paid to Account: R2,084.00 (Net + Loan) ✓
- [ ] Test Case 3:
  - [ ] Input: 40h 0m @ R35.00, Temporary, No loans
  - [ ] Standard Time: R1,400.00 ✓
  - [ ] UIF: R0.00 (temporary - no UIF) ✓
  - [ ] Net: R1,400.00 ✓
  - [ ] Paid to Account: R1,400.00 ✓
- [ ] Test Case 4:
  - [ ] Input: 35h 0m + 5h OT @ R30.00, Permanent, No loans
  - [ ] Standard Time: R1,050.00 ✓
  - [ ] Overtime: R225.00 (5 × 30 × 1.5) ✓
  - [ ] Gross: R1,275.00 ✓
  - [ ] UIF: R12.75 ✓
  - [ ] Net: R1,262.25 ✓
  - [ ] Paid to Account: R1,262.25 ✓

### Integration Tests
- [ ] Approved timesheet creates correct payslip
- [ ] Loan activity triggers auto-sync
- [ ] Auto-sync updates EmployeeLoans table
- [ ] Auto-sync updates CurrentLoanBalance in payslip
- [ ] Loan balance widget shows updated balance
- [ ] Employee loan history shows new transaction

### Performance Tests
- [ ] Create payslip < 5 seconds
- [ ] Calculate payslip < 1 second
- [ ] Generate PDF < 10 seconds
- [ ] Auto-sync < 5 seconds
- [ ] List payslips < 3 seconds (with 4,160 records)

### User Acceptance
- [ ] User can create payslip easily
- [ ] Calculations are clear and visible
- [ ] Loan balance is prominent
- [ ] Paid to Account is clearly shown
- [ ] PDF looks professional
- [ ] Auto-sync works invisibly (no user action needed)

### Sign-Off
- [ ] All checklist items completed
- [ ] All calculations validated
- [ ] Tested with real user
- [ ] Parallel run with AppSheet for 2 pay cycles
- [ ] No critical bugs
- [ ] Ready for Phase 5

**Phase 4 Complete**: _____ (Date) | **Signed**: _____ (Name)

---

## Phase 5: Reports ✓

**Goal**: All reports functional
**Timeline**: Week 7 (3-5 days)
**Status**: ⚪ Pending

### Code Deliverables
- [ ] **Reports.gs** created with functions
  - [ ] generateOutstandingLoansReport(asOfDate) - Outstanding loans
  - [ ] generateIndividualStatementReport(employeeId, startDate, endDate) - Loan statement
  - [ ] generateWeeklyPayrollSummaryReport(weekEnding) - Weekly summary
  - [ ] deleteOldReportFiles(reportType) - Cleanup old reports
  - [ ] getOrCreateReportsFolder() - Get/create folder
  - [ ] formatReportSheet(sheet, data, reportType) - Format report

### UI Deliverables
- [ ] **ReportsMenu.html** - Report generation interface
  - [ ] Report type selection
  - [ ] Parameter inputs:
    - [ ] Outstanding Loans: As-of date
    - [ ] Individual Statement: Employee, Start date, End date
    - [ ] Weekly Payroll: Week ending date
  - [ ] Generate button
  - [ ] Loading indicator
  - [ ] Success message with report URL
  - [ ] Error message if failure
  - [ ] Open report button

### Report 1: Outstanding Loans
- [ ] Accepts as-of date parameter
- [ ] Gets all employees with balance > 0 as of date
- [ ] Creates Google Sheet with:
  - [ ] Report title and date
  - [ ] Table with columns: Employee Name, Balance, Last Transaction
  - [ ] Total outstanding at bottom
- [ ] Formatted professionally
- [ ] Sorted by balance (descending)
- [ ] Saved to "Payroll Reports" folder
- [ ] Sharing set to "Anyone with link"
- [ ] Returns URL

### Report 2: Individual Statement
- [ ] Accepts employee ID, start date, end date
- [ ] Gets opening balance (as of start date)
- [ ] Gets all transactions in date range
- [ ] Creates Google Sheet with:
  - [ ] Employee name and details
  - [ ] Opening balance
  - [ ] Transaction table: Date, Type, Amount, Balance Before, Balance After, Notes
  - [ ] Closing balance
- [ ] Formatted professionally
- [ ] Running balance column
- [ ] Saved to "Payroll Reports" folder
- [ ] Sharing set to "Anyone with link"
- [ ] Returns URL

### Report 3: Weekly Payroll Summary
- [ ] Accepts week ending date
- [ ] Gets all payslips for that week
- [ ] Creates Google Sheet with 2 tabs:
  
  **Tab 1: Payroll Register**
  - [ ] Employee Name
  - [ ] Employer
  - [ ] Hours (Standard + OT)
  - [ ] Gross Pay
  - [ ] Deductions (UIF, Other, Loans)
  - [ ] Net Pay
  - [ ] Loan Transaction (if any)
  - [ ] Paid to Account
  
  **Tab 2: Summary**
  - [ ] Total Employees
  - [ ] Total Hours
  - [ ] Total Gross
  - [ ] Total Deductions
  - [ ] Total Net
  - [ ] Total Paid to Accounts
  - [ ] Breakdown by Employer
- [ ] Formatted professionally
- [ ] Totals calculated correctly
- [ ] Saved to "Payroll Reports" folder
- [ ] Sharing set to "Anyone with link"
- [ ] Returns URL

### Data Tests
- [ ] Generate Outstanding Loans report - verify correct balances
- [ ] Generate Individual Statement - verify transactions correct
- [ ] Generate Weekly Payroll Summary - verify totals correct
- [ ] Generate report with no data - handles gracefully
- [ ] Generate multiple reports - no file conflicts

### Functional Tests
- [ ] Can select report type
- [ ] Can enter parameters
- [ ] Can generate report
- [ ] Report URL returned
- [ ] Can open report in new tab
- [ ] Report formatted correctly
- [ ] Report calculations accurate
- [ ] Old report files deleted (optional)
- [ ] Reports saved to correct folder
- [ ] Reports are shareable

### Performance Tests
- [ ] Outstanding Loans report < 15 seconds
- [ ] Individual Statement report < 20 seconds
- [ ] Weekly Payroll Summary report < 30 seconds

### User Acceptance
- [ ] User can generate reports easily
- [ ] Reports are professional
- [ ] Reports are accurate
- [ ] Reports are easy to read
- [ ] Can access reports after generation

### Sign-Off
- [ ] All checklist items completed
- [ ] All report types tested
- [ ] Tested with real user
- [ ] No critical bugs
- [ ] Ready for Phase 6

**Phase 5 Complete**: _____ (Date) | **Signed**: _____ (Name)

---

## Phase 6: Testing & Parallel Run ✓

**Goal**: Production-ready system
**Timeline**: Week 8-10 (14-21 days)
**Status**: ⚪ Pending

### Week 8: Comprehensive Testing

#### System Testing
- [ ] All modules integrated correctly
- [ ] Navigation between modules works
- [ ] Data flows correctly between tables
- [ ] No broken links or errors
- [ ] All features accessible

#### Regression Testing
- [ ] Employee management still works
- [ ] Leave tracking still works
- [ ] Loan management still works
- [ ] Time approval still works
- [ ] Payroll processing still works
- [ ] Reports still work
- [ ] Auto-sync still works

#### Edge Case Testing
- [ ] Handle blank/empty fields gracefully
- [ ] Handle negative values gracefully
- [ ] Handle future dates appropriately
- [ ] Handle very large numbers
- [ ] Handle special characters in text
- [ ] Handle duplicate submissions
- [ ] Handle concurrent edits

#### Error Handling Testing
- [ ] All validations trigger correctly
- [ ] Error messages are clear
- [ ] System recovers from errors
- [ ] No data corruption on error
- [ ] Logging captures errors

#### Performance Testing
- [ ] Load 4,160 payslip records - system responsive
- [ ] Generate all report types - all < 30 seconds
- [ ] 3 concurrent users - no issues
- [ ] Large imports (50+ records) - no timeout
- [ ] Rapid operations - no race conditions

#### Security Testing
- [ ] Only authorized users can access
- [ ] Cannot delete data accidentally
- [ ] Audit trail captures all changes
- [ ] No SQL injection vulnerabilities
- [ ] No XSS vulnerabilities

#### Browser Compatibility Testing
- [ ] Works in Chrome
- [ ] Works in Edge
- [ ] Works in Firefox (if used)
- [ ] Desktop only (no mobile requirement)

### Week 9-10: Parallel Run

#### Setup
- [ ] Keep AppSheet running
- [ ] Start using new system alongside
- [ ] Process same data in both systems
- [ ] Compare results weekly

#### Week 9 Parallel Testing
- [ ] Week 1 payroll: Process in both systems
- [ ] Compare all calculations - must match 100%
- [ ] Compare loan balances - must match 100%
- [ ] Compare reports - must match
- [ ] Document any discrepancies
- [ ] Fix issues found
- [ ] User feedback collected

#### Week 10 Parallel Testing
- [ ] Week 2 payroll: Process in both systems
- [ ] Compare all calculations - must match 100%
- [ ] Compare loan balances - must match 100%
- [ ] Compare reports - must match
- [ ] No discrepancies allowed
- [ ] User comfortable with new system
- [ ] Decision: Ready for go-live?

### Final Validation
- [ ] All 4,160+ historical records intact
- [ ] All calculations accurate (100%)
- [ ] All loan balances correct
- [ ] All reports accurate
- [ ] No data loss
- [ ] No errors in logs
- [ ] Performance acceptable
- [ ] Users confident

### User Training (Minimal)
- [ ] Quick walkthrough of new system
- [ ] Show differences from AppSheet
- [ ] Demonstrate key workflows
- [ ] Answer questions
- [ ] Provide quick reference card

### Documentation
- [ ] User Guide created
- [ ] Admin Guide created
- [ ] Troubleshooting Guide created
- [ ] Deployment documentation
- [ ] Handover documentation

### Go-Live Preparation
- [ ] Freeze AppSheet data
- [ ] Final backup of all data
- [ ] Set read-only mode on AppSheet (if possible)
- [ ] Full switch to new system
- [ ] Monitor first payroll cycle closely
- [ ] Quick fixes if needed

### Post-Launch (Week 11)
- [ ] Monitor first week closely
- [ ] Address any issues immediately
- [ ] Collect user feedback
- [ ] Make minor adjustments
- [ ] Confirm system stable
- [ ] Close project

### Success Criteria (Final)
- [ ] ✅ All payrolls processed correctly (100% accuracy)
- [ ] ✅ Loan balances always accurate
- [ ] ✅ No duplicate records
- [ ] ✅ All reports generate successfully
- [ ] ✅ PDF matches required format
- [ ] ✅ Audit trail tracks all changes
- [ ] ✅ Payslip creation < 5 seconds
- [ ] ✅ Report generation < 30 seconds
- [ ] ✅ Page loads < 3 seconds
- [ ] ✅ Supports 3 concurrent users
- [ ] ✅ No data corruption
- [ ] ✅ Simple, intuitive interface
- [ ] ✅ Clear error messages
- [ ] ✅ Fewer clicks than AppSheet
- [ ] ✅ Users can operate without training
- [ ] ✅ Easy to see loan balances
- [ ] ✅ Easy to generate reports
- [ ] ✅ Cost savings vs AppSheet
- [ ] ✅ Time savings in payroll processing
- [ ] ✅ No disruption to weekly cycle
- [ ] ✅ Historical data preserved

### Sign-Off
- [ ] All checklist items completed
- [ ] Parallel run successful
- [ ] Users satisfied
- [ ] No critical bugs
- [ ] System production-ready
- [ ] **PROJECT COMPLETE**

**Phase 6 Complete**: _____ (Date) | **Signed**: _____ (Name)

---

## Overall Project Sign-Off

**Project**: HR System Migration (AppSheet to Apps Script)
**Duration**: _____ weeks
**Total Phases**: 6
**Status**: ✅ COMPLETE

### Final Approvals
- [ ] All phases completed
- [ ] All acceptance criteria met
- [ ] All tests passed
- [ ] Users trained
- [ ] Documentation complete
- [ ] AppSheet decommissioned
- [ ] Cost savings achieved
- [ ] System stable and reliable

**Project Manager**: _____ (Signature) | **Date**: _____

**Client Sign-Off**: _____ (Signature) | **Date**: _____

---

**Document End**
