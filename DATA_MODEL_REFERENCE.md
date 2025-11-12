# Data Model Reference

**Quick lookup for all table structures, field types, and relationships**
**Version**: 1.0 | **Date**: October 28, 2025

---

## Table Overview

| Table | Purpose | Records | Phase |
|-------|---------|---------|-------|
| EMPLOYEE DETAILS | Master employee registry | ~18 active | Phase 1 |
| LEAVE | Leave tracking (record keeping) | Ongoing | Phase 2 |
| EmployeeLoans | Loan transaction ledger | Ongoing | Phase 2 |
| PendingTimesheets | Time approval staging | Weekly batch | Phase 3 |
| MASTERSALARY | Weekly payslip records | ~4,160+ | Phase 4 |

---

## 1. EMPLOYEE DETAILS

**Purpose**: Central employee registry
**Primary Key**: id (auto-generated UUID)
**Unique Keys**: ID NUMBER, ClockInRef

### Fields

| Field Name | Type | Required | Notes |
|------------|------|----------|-------|
| **id** | Text | System | UUID, auto-generated |
| **REFNAME** | Text | System | Auto: "FirstName Surname" |
| **EMPLOYEE NAME** | Text | ✅ Yes | First name only |
| **SURNAME** | Text | ✅ Yes | Last name |
| **EMPLOYER** | Enum | ✅ Yes | "SA Grinding Wheels" / "Scorpio Abrasives" |
| **HOURLY RATE** | Decimal | ✅ Yes | Must be > 0 |
| **ID NUMBER** | Text | ✅ Yes | SA ID format, must be unique |
| **CONTACT NUMBER** | Phone | ✅ Yes | SA phone format |
| **ADDRESS** | Address | ✅ Yes | Full address |
| **EMPLOYMENT STATUS** | Enum | ✅ Yes | "Permanent" / "Temporary" / "Contract" |
| **ClockInRef** | Text | ✅ Yes | Clock number, must be unique |
| DATE OF BIRTH | Date | No | Optional |
| ALTERNATIVE CONTACT | Phone | No | Optional |
| ALT CONTACT NAME | Text | No | Optional |
| INCOME TAX NUMBER | Text | No | Optional |
| EMPLOYMENT DATE | Date | No | Optional |
| TERMINATION DATE | Date | No | Set when employee terminated |
| NOTES | LongText | No | Free text notes |
| OveralSize | Number | No | For uniform ordering |
| ShoeSize | Number | No | For uniform ordering |
| UNIONMEM | Yes/No | No | Legacy, keep but don't use |
| UNIONFEE | Price | No | Legacy, keep but don't use |
| RETMEMBER | Yes/No | No | Legacy, keep but don't use |
| RETFACCNUMBER | Text | No | Legacy, keep but don't use |
| **USER** | Text | System | Email of creator |
| **TIMESTAMP** | DateTime | System | Creation timestamp |
| **MODIFIED_BY** | Text | System | Email of last modifier |
| **LAST_MODIFIED** | DateTime | System | Last modification timestamp |

### Validation Rules
```javascript
// Required fields cannot be blank
if (!employeeName || !surname || !employer || !hourlyRate || !idNumber || 
    !contactNumber || !address || !employmentStatus || !clockInRef) {
  throw new Error("Missing required fields");
}

// Hourly rate must be positive
if (hourlyRate <= 0) {
  throw new Error("Hourly rate must be greater than 0");
}

// ID Number must be unique
if (idNumberExists(idNumber, currentEmployeeId)) {
  throw new Error("ID Number already exists");
}

// ClockInRef must be unique
if (clockInRefExists(clockInRef, currentEmployeeId)) {
  throw new Error("Clock Number already exists");
}
```

### Business Rules
- Employees CAN move between employers (update EMPLOYER field)
- REFNAME is auto-generated: `${EMPLOYEE_NAME} ${SURNAME}`
- Termination: Set TERMINATION DATE, optionally change status

---

## 2. MASTERSALARY

**Purpose**: Individual weekly payslip records
**Primary Key**: RECORDNUMBER (sequential number)
**Unique Constraint**: EMPLOYEE NAME + WEEKENDING (no duplicates)

### Fields

| Field Name | Type | Source | Formula/Notes |
|------------|------|--------|---------------|
| **RECORDNUMBER** | Number | System | Sequential, unique |
| **TIMESTAMP** | DateTime | System | Creation time |
| **EMPLOYEE NAME** | Reference | User | Links to EMPLOYEE DETAILS.REFNAME |
| **EMPLOYER** | Enum | Lookup | From EMPLOYEE DETAILS |
| **EMPLOYMENT STATUS** | Text | Lookup | From EMPLOYEE DETAILS |
| **WEEKENDING** | Date | User | Week ending date (e.g., Friday) |
| **HOURS** | Number | User | Standard hours worked |
| **MINUTES** | Number | User | Standard minutes worked |
| **OVERTIMEHOURS** | Number | User | Overtime hours |
| **OVERTIMEMINUTES** | Number | User | Overtime minutes |
| **HOURLYRATE** | Decimal | Lookup | From EMPLOYEE DETAILS |
| **STANDARDTIME** | Decimal | Calculated | `(HOURS × HOURLYRATE) + ((HOURLYRATE/60) × MINUTES)` |
| **OVERTIME** | Decimal | Calculated | `(OVERTIMEHOURS × HOURLYRATE × 1.5) + ((HOURLYRATE/60) × OVERTIMEMINUTES × 1.5)` |
| **LEAVE PAY** | Price | User | Manual entry |
| **BONUS PAY** | Price | User | Manual entry |
| **OTHERINCOME** | Price | User | Manual entry |
| **GROSSSALARY** | Price | Calculated | `STANDARDTIME + OVERTIME + LEAVE PAY + BONUS PAY + OTHERINCOME` |
| **UIF** | Price | Calculated | `GROSSSALARY × 0.01` if EMPLOYMENT STATUS = "PERMANENT", else 0 |
| **OTHER DEDUCTIONS** | Price | User | Manual entry |
| **OTHER DEDUCTIONS TEXT** | Text | User | Description of other deductions |
| **CurrentLoanBalance** | Price | Lookup | Latest balance from EmployeeLoans |
| **LoanDeductionThisWeek** | Price | User | Manual entry |
| **NewLoanThisWeek** | Price | User | Manual entry |
| **LoanDisbursementType** | Enum | User | "With Salary" / "Separate" |
| **UpdatedLoanBalance** | Price | Calculated | `CurrentLoanBalance - LoanDeductionThisWeek + NewLoanThisWeek` |
| **LoanRepaymentLogged** | Yes/No | System | Flag for auto-sync status |
| **TOTALDEDUCTIONS** | Price | Calculated | `UIF + OTHER DEDUCTIONS + LoanDeductionThisWeek` |
| **NETTSALARY** | Price | Calculated | `GROSSSALARY - TOTALDEDUCTIONS` |
| **PaidToAccount** | Price | Calculated | `NETTSALARY - LoanDeductionThisWeek + (NewLoanThisWeek if LoanDisbursementType = "With Salary", else 0)` |
| **FILENAME** | Text | System | PDF filename: `{RECORDNUMBER}.pdf` |
| **FILELINK** | File | System | URL to generated PDF |
| **NOTES** | LongText | User | Free text notes |
| **USER** | Text | System | Email of creator |
| **MODIFIED_BY** | Text | System | Email of last modifier |
| **LAST_MODIFIED** | DateTime | System | Last modification timestamp |

### Critical Calculations

```javascript
// 1. Standard Time
const standardTime = (hours * hourlyRate) + ((hourlyRate / 60) * minutes);

// 2. Overtime (1.5x multiplier)
const overtime = (overtimeHours * hourlyRate * 1.5) + 
                 ((hourlyRate / 60) * overtimeMinutes * 1.5);

// 3. Gross Salary
const grossSalary = standardTime + overtime + leavePay + bonusPay + otherIncome;

// 4. UIF (1% for permanent employees only)
const uif = (employmentStatus === "Permanent") ? (grossSalary * 0.01) : 0;

// 5. Total Deductions
const totalDeductions = uif + otherDeductions + loanDeductionThisWeek;

// 6. Net Salary
const netSalary = grossSalary - totalDeductions;

// 7. Paid to Account (CRITICAL)
const paidToAccount = netSalary - loanDeductionThisWeek + 
                      ((loanDisbursementType === "With Salary") ? newLoanThisWeek : 0);
```

### Validation Rules
```javascript
// Cannot create duplicate payslip
if (payslipExists(employeeName, weekEnding)) {
  throw new Error("Payslip already exists for this employee and week");
}

// All time values must be >= 0
if (hours < 0 || minutes < 0 || overtimeHours < 0 || overtimeMinutes < 0) {
  throw new Error("Time values cannot be negative");
}

// All amounts must be >= 0
if (leavePay < 0 || bonusPay < 0 || otherIncome < 0 || otherDeductions < 0) {
  throw new Error("Amounts cannot be negative");
}

// Warning if loan deduction > current balance
if (loanDeductionThisWeek > currentLoanBalance) {
  console.warn("Loan deduction exceeds current balance");
}
```

### Business Rules
- Week ending typically Friday
- Loan sync triggers automatically via onChange
- PDF generated on demand (not automatic)

---

## 3. EmployeeLoans

**Purpose**: Transaction log for all employee loans
**Primary Key**: LoanID (UUID)
**Unique Constraint**: SalaryLink (one loan record per payslip)

### Fields

| Field Name | Type | Source | Notes |
|------------|------|--------|-------|
| **LoanID** | Text | System | UUID, auto-generated |
| **Employee ID** | Reference | User/System | Links to EMPLOYEE DETAILS.id |
| **Employee Name** | Text | Dereferenced | Auto-filled from Employee ID |
| **Timestamp** | DateTime | System | Record timestamp (updates on edit) |
| **TransactionDate** | Date | User/System | NEVER changes after creation |
| **LoanAmount** | Price | User/System | Positive = disbursement, Negative = repayment |
| **LoanType** | Enum | System | "Disbursement" / "Repayment" |
| **DisbursementMode** | Enum | User | "With Salary" / "Separate" / "Manual Entry" |
| **SalaryLink** | Text | System | RECORDNUMBER from MASTERSALARY (if linked) |
| **BalanceBefore** | Price | Calculated | Balance before this transaction |
| **BalanceAfter** | Price | Calculated | Balance after this transaction |
| **Notes** | LongText | User/System | Transaction notes |

### Critical Rules

```javascript
// Positive amounts = Disbursement (new loan)
// Negative amounts = Repayment

// Balance calculation (chronological)
balanceAfter = balanceBefore + loanAmount;

// For disbursement:
loanAmount = 500;  // Positive
balanceAfter = balanceBefore + 500;

// For repayment:
loanAmount = -150;  // Negative
balanceAfter = balanceBefore - 150;
```

### Validation Rules
```javascript
// Cannot create duplicate SalaryLink
if (salaryLink && salaryLinkExists(salaryLink, currentLoanId)) {
  throw new Error("Loan record already exists for this payslip");
}

// Loan amount must not be zero
if (loanAmount === 0) {
  throw new Error("Loan amount cannot be zero");
}

// Type must match amount sign
if (loanType === "Disbursement" && loanAmount <= 0) {
  throw new Error("Disbursement must have positive amount");
}
if (loanType === "Repayment" && loanAmount >= 0) {
  throw new Error("Repayment must have negative amount");
}
```

### Business Rules
- **TransactionDate**: PRESERVED on edits (critical for chronological sorting)
- **Timestamp**: Updates on edits (for deduplication)
- Balance recalculation: ALWAYS chronological by TransactionDate, then Timestamp
- Auto-sync: Triggered by onChange when payslip has loan activity

---

## 4. LEAVE

**Purpose**: Leave record keeping (NOT linked to salary calculations)
**Primary Key**: Auto-increment row number
**Unique Key**: None (multiple leave records per employee allowed)

### Fields

| Field Name | Type | Source | Notes |
|------------|------|--------|-------|
| **TIMESTAMP** | DateTime | System | Record creation time |
| **EMPLOYEE NAME** | Reference | User | Links to EMPLOYEE DETAILS |
| **STARTDATE.LEAVE** | Date | User | Leave start date |
| **RETURNDATE.LEAVE** | Date | User | Expected return date |
| **TOTALDAYS.LEAVE** | Number | Calculated | Days between start and return |
| **REASON** | Enum | User | "AWOL" / "Sick Leave" / "Annual Leave" / "Unpaid Leave" |
| **NOTES** | Text | User | Additional details |
| **USER** | Text | System | Email of creator |
| **IMAGE** | Image | User | Optional supporting document |

### Validation Rules
```javascript
// Return date must be >= start date
if (returnDate < startDate) {
  throw new Error("Return date cannot be before start date");
}

// Total days calculation
totalDays = Math.floor((returnDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
```

### Business Rules
- Leave table is TRACKING ONLY
- Does NOT affect payroll calculations
- Leave Pay is entered manually in MASTERSALARY if applicable

---

## 5. PendingTimesheets (NEW)

**Purpose**: Staging table for time approval workflow
**Primary Key**: ID (UUID)
**Unique Constraint**: EMPLOYEE NAME + WEEKENDING (per status)

### Fields

| Field Name | Type | Source | Notes |
|------------|------|--------|-------|
| **ID** | Text | System | UUID, auto-generated |
| **EMPLOYEE NAME** | Reference | Import | Links to EMPLOYEE DETAILS |
| **WEEKENDING** | Date | Import | Week ending date |
| **HOURS** | Number | Import/Edit | Standard hours |
| **MINUTES** | Number | Import/Edit | Standard minutes |
| **OVERTIMEHOURS** | Number | Import/Edit | Overtime hours |
| **OVERTIMEMINUTES** | Number | Import/Edit | Overtime minutes |
| **NOTES** | Text | Import/Edit | Comments from time clock or user |
| **STATUS** | Enum | System | "Pending" / "Approved" / "Rejected" |
| **IMPORTED_BY** | Text | System | Email of importer |
| **IMPORTED_DATE** | DateTime | System | Import timestamp |
| **REVIEWED_BY** | Text | System | Email of reviewer |
| **REVIEWED_DATE** | DateTime | System | Review timestamp |

### Excel Import Format
```
Employee Name, Week Ending, Standard Hours, Standard Minutes, Overtime Hours, Overtime Minutes, Notes
Archie Patrick, 2025-10-17, 39, 30, 0, 0, On time all week
John Smith, 2025-10-17, 40, 0, 2, 0, Worked late Monday
```

### Validation Rules
```javascript
// Employee must exist
if (!employeeExists(employeeName)) {
  throw new Error("Employee not found: " + employeeName);
}

// Cannot approve duplicate week
if (payslipExists(employeeName, weekEnding)) {
  throw new Error("Payslip already exists for this week");
}

// All time values must be >= 0
if (hours < 0 || minutes < 0 || overtimeHours < 0 || overtimeMinutes < 0) {
  throw new Error("Time values cannot be negative");
}
```

### Workflow States
- **Pending**: Imported, awaiting review
- **Approved**: Creates MASTERSALARY record, triggers loan sync
- **Rejected**: Marked for deletion or correction

---

## Relationships

### EMPLOYEE DETAILS → Other Tables
```
EMPLOYEE DETAILS (1) ──→ (∞) MASTERSALARY
                    ├──→ (∞) EmployeeLoans
                    ├──→ (∞) LEAVE
                    └──→ (∞) PendingTimesheets
```

### MASTERSALARY ↔ EmployeeLoans
```
MASTERSALARY (1) ←──→ (0..1) EmployeeLoans
     (via RECORDNUMBER = SalaryLink)
```

### Lookup Chains
```
MASTERSALARY.EMPLOYEE NAME → EMPLOYEE DETAILS.REFNAME
    ↓ Lookup
EMPLOYEE DETAILS.EMPLOYER
EMPLOYEE DETAILS.EMPLOYMENT STATUS
EMPLOYEE DETAILS.HOURLY RATE
```

---

## Enums / Dropdown Values

### EMPLOYER
- "SA Grinding Wheels"
- "Scorpio Abrasives"

### EMPLOYMENT STATUS
- "Permanent"
- "Temporary"
- "Contract"

### LEAVE REASONS
- "AWOL"
- "Sick Leave"
- "Annual Leave"
- "Unpaid Leave"

### LOAN TYPE
- "Disbursement"
- "Repayment"

### DISBURSEMENT MODE
- "With Salary"
- "Separate"
- "Manual Entry"

### TIMESHEET STATUS
- "Pending"
- "Approved"
- "Rejected"

---

## Common Query Patterns

### Get Employee by ID
```javascript
const employee = employeeSheet.getDataRange().getValues()
  .find(row => row[idCol] === employeeId);
```

### Get Current Loan Balance
```javascript
const loanRecords = loanSheet.getDataRange().getValues()
  .filter(row => row[empIdCol] === employeeId)
  .sort((a, b) => new Date(b[timestampCol]) - new Date(a[timestampCol]));
const currentBalance = loanRecords.length > 0 ? loanRecords[0][balanceAfterCol] : 0;
```

### Check Duplicate Payslip
```javascript
const exists = salarySheet.getDataRange().getValues()
  .some(row => row[empNameCol] === employeeName && 
               row[weekEndingCol].getTime() === weekEnding.getTime());
```

### Get Latest Payslip
```javascript
const payslips = salarySheet.getDataRange().getValues()
  .filter(row => row[empNameCol] === employeeName)
  .sort((a, b) => b[recordNumberCol] - a[recordNumberCol]);
const latestPayslip = payslips[0];
```

---

## Data Integrity Rules

### Critical Constraints
1. **No Duplicate Payslips**: Same employee + same week ending
2. **No Duplicate SalaryLinks**: One loan record per payslip number
3. **Unique ID Numbers**: Across all employees
4. **Unique Clock Numbers**: Across all employees
5. **TransactionDate Immutability**: Never update after creation

### Referential Integrity
- All EMPLOYEE NAME references must exist in EMPLOYEE DETAILS
- All Employee ID references must exist in EMPLOYEE DETAILS
- SalaryLink should reference valid RECORDNUMBER (but not enforced)

### Calculation Integrity
- Loan balances must be chronologically correct
- Payslip calculations must be accurate (validate against test cases)
- PaidToAccount must equal actual bank transfer amount

---

## Performance Considerations

### Indexes (Manual Optimization)
- Sort EMPLOYEE DETAILS by REFNAME for faster lookups
- Sort EmployeeLoans by Employee ID, then TransactionDate
- Sort MASTERSALARY by RECORDNUMBER (descending) for latest records

### Large Dataset Handling
- 4,160+ payslip records: Filter by date range before processing
- Loan recalculation: Process one employee at a time
- Report generation: Use batch operations, flush() periodically

---

**Document End**
