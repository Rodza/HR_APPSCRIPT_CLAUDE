# HR System - FINAL CONFIRMED REQUIREMENTS

**Date**: October 28, 2025
**Status**: ALL REQUIREMENTS CONFIRMED - READY TO BUILD

---

## SYSTEM OVERVIEW

### Purpose
Weekly advance payment system for 15-18 employees. Accountant handles formal monthly payslips with statutory requirements. This system tracks:
- Weekly hours and payments
- Loan balances
- Leave records (tracking only)
- Basic reporting

### Users
- **Owner + Admin**: 2 users, desktop only, limited tech-savvy
- **No employee self-service needed**

### Architecture
- **Data**: Google Sheets (no database migration)
- **Code**: Google Apps Script (modular structure)
- **UI**: Single comprehensive dashboard (HTML/JavaScript)
- **Migration**: Parallel run with AppSheet until confident

### Scale
- **Employees**: 15-18 active
- **Historical Data**: October 2021 onwards (~4,160 payslip records)
- **Frequency**: Weekly payroll
- **Concurrent Users**: Maximum 3
- **Performance**: Report generation < 30 seconds acceptable

---

## DATA MODEL

### Tables to Keep

#### 1. EMPLOYEE DETAILS (Master Data)
**Purpose**: Employee registry

**Required Fields** (Must have value):
- EMPLOYEE NAME (First Name)
- SURNAME
- EMPLOYER (Dropdown: "SA Grinding Wheels" / "Scorpio Abrasives")
- HOURLY RATE (Decimal)
- ID NUMBER (Text)
- CONTACT NUMBER (Phone)
- ADDRESS (Address)
- EMPLOYMENT STATUS (Enum: Permanent/Temporary/Contract)
- ClockInRef (Clock Number - Text)

**Optional Fields** (Can be blank):
- DATE OF BIRTH (Date)
- ALTERNATIVE CONTACT (Phone)
- ALT CONTACT NAME (Name)
- INCOME TAX NUMBER (Text)
- EMPLOYMENT DATE (Date)
- TERMINATION DATE (Date)
- NOTES (LongText)
- OveralSize (Number) - for uniforms
- ShoeSize (Number) - for uniforms
- UNIONMEM (Yes/No) - legacy, keep for reference
- UNIONFEE (Price) - legacy, keep for reference
- RETMEMBER (Yes/No) - legacy, keep for reference
- RETFACCNUMBER (Text) - legacy, keep for reference

**System Fields**:
- USER (Text) - who created
- TIMESTAMP (DateTime) - when created
- MODIFIED_BY (Text) - who last modified
- LAST_MODIFIED (DateTime) - when last modified

**Notes**:
- Employees CAN move between employers (SA Grinding Wheels ↔ Scorpio Abrasives)
- REFNAME field will be auto-generated (e.g., "Archie Patrick")

---

#### 2. MASTERSALARY (Payroll Records)
**Purpose**: Individual weekly payslip records

**Key Fields**:
- RECORDNUMBER (Number) - Unique payslip number (like 7916 in sample)
- TIMESTAMP (DateTime)
- EMPLOYEE NAME (Reference to EMPLOYEE DETAILS)
- EMPLOYER (Looked up from employee)
- EMPLOYMENT STATUS (Looked up from employee)
- WEEKENDING (Date)

**Time Entry** (Manual entry by user):
- HOURS (Number) - Standard hours
- MINUTES (Number) - Standard minutes
- OVERTIMEHOURS (Number)
- OVERTIMEMINUTES (Number)

**Calculated Earnings**:
- HOURLYRATE (Looked up from employee)
- STANDARDTIME (Decimal) = (HOURS × HOURLYRATE) + ((HOURLYRATE/60) × MINUTES)
- OVERTIME (Decimal) = (OVERTIMEHOURS × HOURLYRATE × 1.5) + ((HOURLYRATE/60) × OVERTIMEMINUTES × 1.5)
- GROSSSALARY (Price) = STANDARDTIME + OVERTIME + LEAVE PAY + BONUS PAY + OTHERINCOME

**Additional Pay** (Manual entry):
- LEAVE PAY (Price)
- BONUS PAY (Price)
- OTHERINCOME (Price)

**Deductions**:
- UIF (Price) = GROSSSALARY × 0.01 (if EMPLOYMENT STATUS = "PERMANENT")
- OTHER DEDUCTIONS (Price) - Manual entry
- OTHER DEDUCTIONS TEXT (Text) - Description

**Loan Integration**:
- CurrentLoanBalance (Price) - Latest balance from EmployeeLoans
- LoanDeductionThisWeek (Price) - Manual entry
- NewLoanThisWeek (Price) - Manual entry
- LoanDisbursementType (Enum: "With Salary" / "Separate")
- UpdatedLoanBalance (Price) = CurrentLoanBalance - LoanDeductionThisWeek + NewLoanThisWeek
- LoanRepaymentLogged (Yes/No) - System flag

**Final Calculations**:
- TOTALDEDUCTIONS (Price) = UIF + OTHER DEDUCTIONS + LoanDeductionThisWeek
- NETTSALARY (Price) = GROSSSALARY - TOTALDEDUCTIONS
- PaidToAccount (Price) = NETTSALARY - LoanDeductionThisWeek + (NewLoanThisWeek if LoanDisbursementType = "With Salary", else 0)

**Document Generation**:
- FILENAME (Text) - PDF filename
- FILELINK (File) - Link to generated PDF

**System Fields**:
- USER (Text) - who created
- MODIFIED_BY (Text) - who last modified
- LAST_MODIFIED (DateTime) - when last modified
- NOTES (LongText)

---

#### 3. EmployeeLoans (Loan Ledger)
**Purpose**: Transaction log for all employee loans

**Fields**:
- LoanID (Text) - UUID
- Employee ID (Ref to EMPLOYEE DETAILS)
- Employee Name (Text) - Dereferenced from Employee ID
- Timestamp (DateTime) - System timestamp
- TransactionDate (Date) - Date of transaction (NEVER updated after creation)
- LoanAmount (Price) - Positive for disbursement, negative for repayment
- LoanType (Enum: "Disbursement" / "Repayment")
- DisbursementMode (Enum: "With Salary" / "Separate" / "Manual Entry")
- SalaryLink (Text) - RECORDNUMBER from MASTERSALARY (if linked)
- BalanceBefore (Price) - Balance before this transaction
- BalanceAfter (Price) - Balance after this transaction
- Notes (LongText)

**Critical Rules**:
- LoanAmount: Positive for new loan, negative for repayment
- TransactionDate: PRESERVED on edits (only Timestamp updates)
- SalaryLink: Prevents duplicate loan records (one per payslip)
- Balance recalculation: Always chronological by TransactionDate, then Timestamp

---

#### 4. LEAVE (Leave Tracking)
**Purpose**: Record keeping only (not linked to salary calculations)

**Fields**:
- TIMESTAMP (DateTime)
- EMPLOYEE NAME (Ref to EMPLOYEE DETAILS)
- STARTDATE.LEAVE (Date)
- RETURNDATE.LEAVE (Date)
- TOTALDAYS.LEAVE (Number)
- REASON (Enum: AWOL / Sick Leave / Annual Leave / Unpaid Leave)
- NOTES (Text)
- USER (Text) - who created
- IMAGE (Image) - Supporting documentation if needed

---

#### 5. PendingTimesheets (NEW - Time Approval)
**Purpose**: Staging table for time approval workflow

**Fields**:
- ID (Text) - Unique identifier
- EMPLOYEE NAME (Ref to EMPLOYEE DETAILS)
- WEEKENDING (Date)
- HOURS (Number) - Standard hours
- MINUTES (Number) - Standard minutes
- OVERTIMEHOURS (Number)
- OVERTIMEMINUTES (Number)
- NOTES (Text)
- STATUS (Enum: "Pending" / "Approved" / "Rejected")
- IMPORTED_BY (Text) - who imported
- IMPORTED_DATE (DateTime)
- REVIEWED_BY (Text) - who reviewed
- REVIEWED_DATE (DateTime)

**Workflow**:
1. Import from Excel (HTML analyzer output)
2. User reviews and edits
3. User clicks "Approve"
4. System creates MASTERSALARY record
5. Row marked as "Approved"

---

### Tables to Remove
- ❌ Process for PaySlip - 1 Process Table
- ❌ Process for Update Changes Process Table
- ❌ Process for Batch Print Process Table
- ❌ Process for Generate Outstanding Loans Report - 1 Process Table
- ❌ Process for Generate Individual Statement Process Table
- ❌ Process for Generate Weekly Payroll Summary - 1 Process Table
- ❌ New step Output (all variants)
- ❌ Delete Row Output
- ❌ Reports (report metadata tracking not needed)
- ❌ FilterReport (will be replaced with UI filters)
- ❌ DateSelection (will be replaced with UI date picker)

---

### Fields to Remove from EMPLOYEE DETAILS
- ❌ ATTENDANCE (legacy artifact)
- ❌ ETISK, RETADMIN, ETA, EMPLOYEECONT (retirement - not used)
- ❌ RISKDED, BROKERDED, RADED, EMPLCONTDED (retirement - not used)

### Fields to Remove from MASTERSALARY
- ❌ ATTENDANCECHECK, ATTENDANCE (legacy)
- ❌ UNIONDEDUCTION (not used)
- ❌ RETGROUPDED (not used)
- ❌ DEDUCTIONS (formula field - replaced by manual OTHER DEDUCTIONS)
- ❌ RISKDED, BROKERDED, RADED, EMPLCONTDED, TOTALRETFUND (retirement - not used)
- ❌ CHANGECOUNTER, INITIAL_EMPLOYEE, BOTRUN (AppSheet artifacts)
- ❌ MostRecent, Related Datas (AppSheet virtual columns)

---

## PAYSLIP FORMAT (PDF)

Based on sample payslip #7916:

### Header Section
```
WEEKLY PAY REMITTANCE

PAYSLIP NUMBER: [RECORDNUMBER]                    [EMPLOYER NAME]
                                                   18 DODGE STREET
EMPLOYEE: [First Name] [Surname]                   AUREUS
WEEKENDING: [MM/DD/YYYY]                          RANDFONTEIN
EMPLOYMENT STATUS: [Permanent/Temporary]           (011) 693 4278 / 083 338 5609
                                                   INFO@SAGRINDING.CO.ZA /
                                                   INFO@SCORPIOABRASIVES.CO.ZA
```

### Earnings & Deductions Table
```
+-------------------+-------+--------+--------+-----------+------------------------+
|     EARNINGS                                 |       DEDUCTIONS              |
+-------------------+-------+--------+--------+-----------+------------------------+
|                   | HOURS | MINUTES| RATE   | AMOUNT    | UIF            | R###  |
+-------------------+-------+--------+--------+-----------+------------------------+
| NORMAL TIME       | ##    | ##     | R##.## | R#,###.## | OTHER DEDUCTIONS      |
| OVER TIME         |       |        |        | R#.##     |                       |
+-------------------+-------+--------+--------+-----------+-----------------------+
| ADDITIONAL PAY    | ATTENDANCE     | 0      | LOANS OPENING    | R###.##      |
|                   |                |        | BALANCE          |              |
|                   | BONUS          |        | LOAN/REPAYMENT   | -R###.##     |
|                   | LEAVE          |        | LOANS CLOSING    | R#.##        |
|                   | OTHER          |        | BALANCE          |              |
+-------------------+----------------+--------+----------------------+-----------+
```

### Notes Section
```
+-----------------------------------------------------------------------------+
| NOTES                                                                       |
|                                                                             |
+-----------------------------------------------------------------------------+
```

### Summary Section
```
+----------------------------------+------------------------------------------+
| GROSS PAY                 R#,###.## | TOTAL DEDUCTIONS            R##.##   |
+----------------------------------+------------------------------------------+
| NETT PAY                                                        R#,###.##  |
+----------------------------------+------------------------------------------+
| AMOUNT PAID TO ACCOUNT                                          R#,###.##  |
+-----------------------------------------------------------------------------+
```

**Key Requirements**:
- Format currency as "R#,###.##" (South African Rand)
- Show loan opening balance, repayment (as negative), closing balance
- **AMOUNT PAID TO ACCOUNT** is the critical final number (what actually goes to bank)
- Notes section for any additional information

---

## FUNCTIONAL REQUIREMENTS

### Module 1: Employee Management

**Features**:
1. Add new employee (form with required field validation)
2. Edit employee details
3. View employee profile
4. List all employees (with search and employer filter)
5. Transfer employee between employers
6. Terminate employee (set TERMINATION DATE, change status)

**Validation Rules**:
- Required fields cannot be blank
- Hourly rate must be > 0
- ID NUMBER must be unique
- CONTACT NUMBER must be valid South African format
- ClockInRef must be unique

**UI Requirements**:
- Simple form layout
- Clear field labels
- Error messages for validation failures
- Success confirmation on save
- Employer filter dropdown on list view

---

### Module 2: Leave Tracking

**Features**:
1. Record leave (start date, end date, reason)
2. View leave history per employee
3. Search/filter leave records
4. Upload supporting documents (optional)

**Validation Rules**:
- End date must be >= Start date
- Total days calculated automatically
- Cannot record leave for future dates (or allow with warning)

**UI Requirements**:
- Calendar date picker
- Reason dropdown (AWOL, Sick, Annual, Unpaid)
- Notes field for details
- Image upload (optional)

---

### Module 3: Loan Management

**Features**:
1. Record new loan (disbursement)
2. Record repayment (manual entry)
3. View loan history per employee
4. View current balance per employee
5. Automatic sync with payroll deductions

**Validation Rules**:
- Loan amount must be > 0
- Cannot create duplicate SalaryLink
- Balance calculation must be chronological

**UI Requirements**:
- Simple form for disbursement (Amount, Type, Mode, Notes)
- Simple form for repayment (Amount, Notes)
- Clear display of current balance
- Transaction history table
- Highlight overdue balances (if any logic needed)

**Automatic Sync Process**:
1. User creates/edits payslip with LoanDeductionThisWeek or NewLoanThisWeek
2. onChange trigger fires within 5 minutes
3. System checks if loan record exists (via SalaryLink)
4. If exists: Update record
5. If not: Create new record
6. Recalculate all balances for employee chronologically
7. Update CurrentLoanBalance in MASTERSALARY

---

### Module 4: Time Approval Workflow

**Features**:
1. Import Excel from HTML analyzer
2. Display in editable table (PendingTimesheets)
3. User can adjust hours/minutes
4. Approve button creates payslip
5. Track status (Pending/Approved/Rejected)

**Excel Import Format** (from HTML analyzer):
```
Employee Name, Week Ending, Standard Hours, Standard Minutes, Overtime Hours, Overtime Minutes, Notes
Archie Patrick, 2025-10-17, 39, 30, 0, 0, On time all week
```

**Validation Rules**:
- Employee must exist in EMPLOYEE DETAILS
- Week ending must be a valid date
- Cannot approve duplicate week for same employee
- Hours/minutes must be >= 0

**UI Requirements**:
- Upload button for Excel file
- Editable table (inline editing)
- Approve/Reject buttons per row
- Filter by status
- Clear indication of what's pending

**Workflow**:
```
User uploads Excel → Parse and insert into PendingTimesheets (Status = Pending)
                                    ↓
User reviews table, makes adjustments → Click "Approve" button
                                    ↓
System validates → Creates MASTERSALARY record with approved hours
                                    ↓
Mark row as "Approved" → Trigger loan sync if applicable
```

---

### Module 5: Payroll Processing

**Features**:
1. Create new payslip (manual entry)
2. Create from approved timesheet (auto-populate)
3. Edit existing payslip
4. View payslip details
5. Generate PDF
6. List all payslips (with filters: employee, week ending, employer)

**Calculation Engine**:
```javascript
// 1. Standard Time
standardTime = (hours × hourlyRate) + ((hourlyRate / 60) × minutes)

// 2. Overtime
overtime = (overtimeHours × hourlyRate × 1.5) + ((hourlyRate / 60) × overtimeMinutes × 1.5)

// 3. Gross Salary
grossSalary = standardTime + overtime + leavePay + bonusPay + otherIncome

// 4. UIF (if permanent employee)
uif = (employmentStatus === "Permanent") ? (grossSalary × 0.01) : 0

// 5. Total Deductions
totalDeductions = uif + otherDeductions

// 6. Net Salary
netSalary = grossSalary - totalDeductions

// 7. Paid to Account
paidToAccount = netSalary - loanDeductionThisWeek + 
                ((loanDisbursementType === "With Salary") ? newLoanThisWeek : 0)
```

**Validation Rules**:
- WEEKENDING must be a valid date
- Cannot create duplicate payslip (same employee, same week ending)
- Hours/minutes must be >= 0
- All amounts must be >= 0
- If LoanDeductionThisWeek > CurrentLoanBalance, show warning

**UI Requirements**:
- Form with all fields clearly labeled
- Auto-lookup hourly rate from employee
- Real-time calculation preview
- **Prominent display of CurrentLoanBalance**
- Clear indication of PaidToAccount
- Generate PDF button
- View PDF button (if already generated)

**PDF Generation**:
- Use Google Docs template or HTML-to-PDF
- Match format from sample #7916
- Save to Google Drive
- Store URL in FILELINK field
- Filename: [RECORDNUMBER].pdf

---

### Module 6: Reporting

#### Report 1: Outstanding Loans Report
**Purpose**: Show all employees with current loan balances > 0

**Input**: As-of date (default: today)

**Output**: Google Sheets with:
- Employee Name
- Outstanding Balance
- Last Transaction Date
- Total Outstanding (sum)

**Process**:
1. For each employee, get latest BalanceAfter as of specified date
2. Filter to balance > 0
3. Sort by balance (descending)
4. Create formatted Google Sheet
5. Return URL

---

#### Report 2: Individual Statement
**Purpose**: Complete loan transaction history for one employee

**Input**: Employee ID, Start Date, End Date

**Output**: Google Sheets with:
- Employee name and details
- Opening balance (as of start date)
- All transactions in date range
- Transaction-by-transaction breakdown (Date, Type, Amount, Balance)
- Closing balance (as of end date)

**Process**:
1. Get opening balance (latest balance before start date)
2. Get all transactions in date range
3. Create formatted Google Sheet with running balance
4. Return URL

---

#### Report 3: Weekly Payroll Summary
**Purpose**: Detailed breakdown of all payslips for a specific week

**Input**: Week ending date

**Output**: Google Sheets with two tabs:

**Tab 1: Payroll Register**
- Employee Name
- Employer
- Hours (Standard + OT)
- Gross Pay
- Deductions (UIF, Other, Loans)
- Net Pay
- Loan Transaction (if any)
- Paid to Account

**Tab 2: Summary**
- Total Employees
- Total Hours
- Total Gross
- Total Deductions
- Total Net
- Total Paid to Accounts
- Breakdown by Employer

**Process**:
1. Get all MASTERSALARY records for specified week ending
2. Extract key fields
3. Create formatted Google Sheet
4. Return URL

---

**Report Generation Rules**:
- Delete old report files before generating new ones (optional)
- Save all reports to "Payroll Reports" folder in Google Drive
- Set sharing: "Anyone with link can view"
- Return URL to user

---

## TECHNICAL SPECIFICATIONS

### Google Apps Script Structure

```
Code.gs (main entry points)
├── doGet() - Web app handler (dashboard)
├── doPost() - Webhook handler (if needed for reports)
└── onChange() - Auto-trigger for loan sync

Config.gs
├── EMPLOYER_LIST = ["SA Grinding Wheels", "Scorpio Abrasives"]
├── EMPLOYMENT_STATUS_LIST = ["Permanent", "Temporary", "Contract"]
├── LEAVE_REASONS = ["AWOL", "Sick Leave", "Annual Leave", "Unpaid Leave"]
├── LOAN_TYPES = ["Disbursement", "Repayment"]
├── DISBURSEMENT_MODES = ["With Salary", "Separate", "Manual Entry"]
└── Required field lists

Utils.gs
├── getSheets() - Sheet locator
├── getCurrentUser()
├── formatDate()
├── formatCurrency()
├── generateUUID()
├── validateSAIdNumber()
├── validatePhoneNumber()
└── Logger helpers

Employees.gs
├── addEmployee(data)
├── updateEmployee(id, data)
├── getEmployeeById(id)
├── getEmployeeByName(name)
├── listEmployees(filters)
├── terminateEmployee(id, terminationDate)
├── validateEmployee(data)
└── employeeLookups (name, rate, status, etc.)

Leave.gs
├── addLeave(data)
├── getLeaveHistory(employeeId)
├── listLeave(filters)
└── validateLeave(data)

Loans.gs
├── addLoanTransaction(data)
├── getCurrentLoanBalance(employeeId)
├── getLoanHistory(employeeId, startDate, endDate)
├── recalculateLoanBalances(employeeId)
├── findLoanRecordBySalaryLink(recordNumber)
├── syncLoanForPayslip(recordNumber) - Auto-sync function
├── cleanupDuplicates()
└── validateLoan(data)

Timesheets.gs
├── importTimesheetExcel(fileBlob)
├── parseExcelData(fileBlob)
├── addPendingTimesheet(data)
├── updatePendingTimesheet(id, data)
├── approveTimesheet(id) - Creates payslip
├── rejectTimesheet(id)
├── listPendingTimesheets(filters)
└── validateTimesheet(data)

Payroll.gs
├── createPayslip(data)
├── updatePayslip(recordNumber, data)
├── getPayslip(recordNumber)
├── listPayslips(filters)
├── calculatePayslip(data) - All formulas
├── generatePayslipPDF(recordNumber)
├── validatePayslip(data)
└── checkDuplicatePayslip(employeeId, weekEnding)

Reports.gs
├── generateOutstandingLoansReport(asOfDate)
├── generateIndividualStatementReport(employeeId, startDate, endDate)
├── generateWeeklyPayrollSummaryReport(weekEnding)
├── deleteOldReportFiles(reportType)
├── getOrCreateReportsFolder()
└── formatReportSheet(sheet, data, reportType)

Triggers.gs
├── installOnChangeTrigger() - Setup onChange
├── onChange(e) - Main trigger handler
└── uninstallTriggers() - Cleanup
```

### HTML UI Structure

```
Dashboard.html (main container)
├── Navigation menu (Employees, Leave, Loans, Timesheets, Payroll, Reports)
├── CSS (Tailwind or Bootstrap)
└── JavaScript (calls Apps Script functions)

Components/
├── EmployeeForm.html
├── EmployeeList.html
├── LeaveForm.html
├── LeaveList.html
├── LoanForm.html
├── LoanList.html
├── TimesheetImport.html
├── TimesheetApproval.html
├── PayrollForm.html
├── PayrollList.html
├── ReportsMenu.html
└── LoanBalanceWidget.html (reusable component)
```

---

## AUDIT TRAIL SPECIFICATION

### For All Tables
**On Create**:
- USER = Session.getActiveUser().getEmail()
- TIMESTAMP = new Date()

**On Update**:
- MODIFIED_BY = Session.getActiveUser().getEmail()
- LAST_MODIFIED = new Date()

### Special Cases
**MASTERSALARY**:
- Track all changes to payslips
- Log when PDF generated
- Log when loan sync triggered

**EmployeeLoans**:
- TransactionDate NEVER changes after creation
- Timestamp updates on edit (for deduplication)
- Log all balance recalculations

---

## MIGRATION PLAN

### Phase 1: Employee Database (Week 1-2)
**Deliverable**: Employee management fully functional

**Tasks**:
1. Create Employees.gs with all CRUD functions
2. Create EmployeeForm.html
3. Create EmployeeList.html with employer filter
4. Add validation
5. Test with sample data
6. Deploy to test users

**Acceptance Criteria**:
- Can add employee with all required fields
- Can edit employee
- Can view employee list
- Can filter by employer
- Can search by name
- Validation prevents invalid data

---

### Phase 2: Leave & Loans (Week 3)
**Deliverable**: Leave tracking + Loan ledger working

**Tasks**:
1. Create Leave.gs
2. Create Loans.gs
3. Create LeaveForm.html and LeaveList.html
4. Create LoanForm.html and LoanList.html
5. Implement loan balance calculation
6. Test loan transaction history
7. Deploy to test users

**Acceptance Criteria**:
- Can record leave
- Can view leave history
- Can create loan disbursement
- Can create loan repayment
- Current balance calculates correctly
- Transaction history displays chronologically

---

### Phase 3: Time Approval System (Week 4)
**Deliverable**: Time approval workflow functional

**Tasks**:
1. Create Timesheets.gs
2. Create PendingTimesheets table
3. Create TimesheetImport.html (Excel upload)
4. Create TimesheetApproval.html (editable table)
5. Implement approve/reject logic
6. Test with sample Excel
7. Deploy to test users

**Acceptance Criteria**:
- Can upload Excel file
- Data imports correctly to PendingTimesheets
- Can edit hours in table
- Approve button creates payslip correctly
- Status tracking works

---

### Phase 4: Payroll Engine (Week 5-6)
**Deliverable**: Complete payroll processing

**Tasks**:
1. Create Payroll.gs with all calculation logic
2. Create PayrollForm.html with real-time calculations
3. Create PayrollList.html with filters
4. Integrate with approved timesheets
5. Implement PDF generation (matching sample #7916)
6. Implement auto-sync with loans (onChange trigger)
7. Add loan balance widget to form
8. Test all calculations against sample data
9. Test loan sync
10. Deploy to test users

**Acceptance Criteria**:
- Can create payslip manually
- Can create payslip from approved timesheet
- All calculations match expected values (validate against historical data)
- PDF generates correctly (matches sample format)
- Loan sync works automatically
- CurrentLoanBalance displays on form
- PaidToAccount calculates correctly

---

### Phase 5: Reports (Week 7)
**Deliverable**: All reports functional

**Tasks**:
1. Create Reports.gs
2. Implement Outstanding Loans Report
3. Implement Individual Statement Report
4. Implement Weekly Payroll Summary Report
5. Create ReportsMenu.html
6. Test report generation
7. Test report formatting
8. Deploy to test users

**Acceptance Criteria**:
- Outstanding Loans Report generates correctly
- Individual Statement Report shows correct history
- Weekly Payroll Summary Report has correct totals
- All reports properly formatted
- Reports save to correct folder
- URLs return correctly

---

### Phase 6: Testing & Parallel Run (Week 8-9)
**Deliverable**: Production-ready system

**Tasks**:
1. Create comprehensive test plan
2. Validate calculations against all historical data (sample set)
3. Run parallel with AppSheet for 2-4 pay cycles
4. Document any discrepancies
5. Fix issues
6. User feedback session
7. Final adjustments
8. Data freeze on AppSheet
9. Go-live

**Acceptance Criteria**:
- All calculations match AppSheet/manual calculations
- No data loss
- No errors in production use
- Users can operate system without assistance
- Performance acceptable (< 30 seconds for reports)
- All features working as expected

---

## SUCCESS CRITERIA

### Functional
- ✅ All payslips calculate correctly (100% accuracy)
- ✅ Loan balances always accurate
- ✅ No duplicate records
- ✅ All reports generate successfully
- ✅ PDF matches required format
- ✅ Audit trail tracks all changes

### Performance
- ✅ Payslip creation < 5 seconds
- ✅ Report generation < 30 seconds
- ✅ Page loads < 3 seconds
- ✅ Supports 3 concurrent users
- ✅ No data corruption

### User Experience
- ✅ Simple, intuitive interface
- ✅ Clear error messages
- ✅ Fewer clicks than AppSheet
- ✅ Owner and admin can use without training
- ✅ Easy to see loan balances
- ✅ Easy to generate reports

### Business
- ✅ Cost savings vs AppSheet
- ✅ Time savings in payroll processing
- ✅ No disruption to weekly payroll cycle
- ✅ Historical data preserved

---

## KNOWN CONSTRAINTS & LIMITATIONS

### Google Apps Script Quotas
- 6-minute execution time limit (triggers)
- 30-minute execution time limit (web apps)
- Daily quota limits (unlikely to hit with 15-18 employees)

### Google Sheets Performance
- Acceptable for 15-18 employees
- 4,160+ historical records manageable
- May need optimization if grows to 50+ employees

### Concurrency
- Google Sheets has basic locking
- Script locks prevent concurrent updates
- Not an issue with 2-3 users

### Browser Compatibility
- Designed for desktop only
- Chrome/Edge recommended
- May not be optimized for mobile

---

## APPENDIX

### Sample Data Structures

#### Employee Record (JSON)
```json
{
  "id": "e0b6115a",
  "REFNAME": "Archie Patrick",
  "EMPLOYEE NAME": "Archie",
  "SURNAME": "Patrick",
  "EMPLOYER": "SA Grinding Wheels",
  "HOURLY RATE": 33.96,
  "EMPLOYMENT STATUS": "Permanent",
  "ID NUMBER": "9401015800081",
  "CONTACT NUMBER": "0821234567",
  "ADDRESS": "123 Main St, Johannesburg",
  "ClockInRef": "001",
  "OveralSize": 38,
  "ShoeSize": 9,
  "USER": "owner@sagrinding.co.za",
  "TIMESTAMP": "2021-10-15T08:00:00Z"
}
```

#### Payslip Record (JSON)
```json
{
  "RECORDNUMBER": 7916,
  "TIMESTAMP": "2025-10-17T09:30:00Z",
  "EMPLOYEE NAME": "Archie Patrick",
  "EMPLOYER": "SA Grinding Wheels",
  "EMPLOYMENT STATUS": "Permanent",
  "WEEKENDING": "2025-10-17",
  "HOURS": 39,
  "MINUTES": 30,
  "OVERTIMEHOURS": 0,
  "OVERTIMEMINUTES": 0,
  "HOURLYRATE": 33.96,
  "STANDARDTIME": 1341.42,
  "OVERTIME": 0,
  "LEAVE PAY": 0,
  "BONUS PAY": 0,
  "OTHERINCOME": 0,
  "GROSSSALARY": 1341.42,
  "UIF": 13.41,
  "OTHER DEDUCTIONS": 0,
  "CurrentLoanBalance": 150.00,
  "LoanDeductionThisWeek": 150.00,
  "NewLoanThisWeek": 0,
  "UpdatedLoanBalance": 0,
  "TOTALDEDUCTIONS": 13.41,
  "NETTSALARY": 1328.01,
  "PaidToAccount": 1178.01,
  "USER": "admin@sagrinding.co.za",
  "LAST_MODIFIED": "2025-10-17T09:30:00Z"
}
```

#### Loan Transaction Record (JSON)
```json
{
  "LoanID": "550e8400-e29b-41d4-a716-446655440000",
  "Employee ID": "e0b6115a",
  "Employee Name": "Archie Patrick",
  "Timestamp": "2025-10-17T09:30:00Z",
  "TransactionDate": "2025-10-17",
  "LoanAmount": -150.00,
  "LoanType": "Repayment",
  "DisbursementMode": "With Salary",
  "SalaryLink": "7916",
  "BalanceBefore": 150.00,
  "BalanceAfter": 0.00,
  "Notes": "Repayment of R150.00 via payslip #7916"
}
```

---

## DOCUMENT STATUS

**Version**: 1.0 FINAL
**Date**: October 28, 2025
**Status**: ✅ APPROVED - READY FOR DEVELOPMENT
**Next Action**: Begin Phase 1 - Employee Database

---

**All requirements confirmed. No further clarifications needed. Ready to build.**
