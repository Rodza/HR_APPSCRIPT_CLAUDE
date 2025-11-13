# HTML Templates Guide

**Status**: Complete backend implementation (10 .gs files + Dashboard.html)
**Remaining**: 12 HTML files to complete the UI

All HTML files follow Bootstrap 5 + Font Awesome patterns from base HRIS.

---

## Completed Files ✅

1. **Dashboard.html** - Main container with sidebar navigation
2. **Code.gs** - Entry points (doGet, include functions)
3. **Config.gs** - Constants and configuration
4. **Utils.gs** - Shared utility functions
5. **Employees.gs** - Employee CRUD operations
6. **Leave.gs** - Leave tracking
7. **Loans.gs** - Loan management with auto-sync
8. **Timesheets.gs** - Time approval workflow
9. **Payroll.gs** - Payroll calculations (100% validated)
10. **Reports.gs** - 3 report types
11. **Triggers.gs** - onChange auto-sync

---

## Remaining HTML Files (12 files)

### Pattern to Follow

All HTML files should:
- Use Bootstrap 5 classes
- Include Font Awesome icons
- Have google.script.run calls to backend
- Include standard helper functions (showLoading, showSuccess, showError)
- Follow responsive design principles

---

## 1. EmployeeList.html

**Purpose**: Display all employees with search/filter

```html
<div class="card">
    <div class="card-header d-flex justify-content-between align-items-center">
        <h5><i class="fas fa-users"></i> Employees</h5>
        <button class="btn btn-primary btn-sm" onclick="showEmployeeForm()">
            <i class="fas fa-plus"></i> Add Employee
        </button>
    </div>
    <div class="card-body">
        <!-- Filter Section -->
        <div class="row mb-3">
            <div class="col-md-4">
                <input type="text" class="form-control" id="searchEmployee" placeholder="Search by name..." onkeyup="filterEmployees()">
            </div>
            <div class="col-md-3">
                <select class="form-select" id="filterEmployer" onchange="filterEmployees()">
                    <option value="">All Employers</option>
                    <option value="SA Grinding Wheels">SA Grinding Wheels</option>
                    <option value="Scorpio Abrasives">Scorpio Abrasives</option>
                </select>
            </div>
        </div>

        <!-- Table -->
        <div class="table-responsive">
            <table class="table table-hover">
                <thead>
                    <tr>
                        <th>Name</th>
                        <th>Employer</th>
                        <th>Clock#</th>
                        <th>Status</th>
                        <th>Hourly Rate</th>
                        <th>Actions</th>
                    </tr>
                </thead>
                <tbody id="employeesTableBody">
                    <tr><td colspan="6" class="text-center">Loading...</td></tr>
                </tbody>
            </table>
        </div>
    </div>
</div>

<script>
function loadEmployees() {
    google.script.run
        .withSuccessHandler(function(result) {
            if (result.success) {
                displayEmployees(result.data);
            } else {
                showError(result.error);
            }
        })
        .withFailureHandler(handleError)
        .listEmployees({});
}

function displayEmployees(employees) {
    const tbody = document.getElementById('employeesTableBody');

    if (employees.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="text-center">No employees found</td></tr>';
        return;
    }

    tbody.innerHTML = employees.map(emp => `
        <tr>
            <td>${emp.REFNAME}</td>
            <td>${emp.EMPLOYER}</td>
            <td>${emp.ClockInRef}</td>
            <td>${emp['EMPLOYMENT STATUS']}</td>
            <td>${formatCurrency(emp['HOURLY RATE'])}</td>
            <td>
                <button class="btn btn-sm btn-info" onclick="editEmployee('${emp.id}')">
                    <i class="fas fa-edit"></i>
                </button>
            </td>
        </tr>
    `).join('');
}

function showEmployeeForm(employeeId) {
    // Load EmployeeForm.html in modal or new view
    loadModule('employeeForm');
}

// Load employees on page load
loadEmployees();
</script>
```

**Backend calls:**
- `listEmployees(filters)` - Get all employees
- `getEmployeeById(id)` - For edit

---

## 2. EmployeeForm.html

**Purpose**: Add/edit employee with validation

```html
<div class="card">
    <div class="card-header">
        <h5><i class="fas fa-user-edit"></i> <span id="formTitle">Add Employee</span></h5>
    </div>
    <div class="card-body">
        <form id="employeeForm" onsubmit="saveEmployee(event)">
            <input type="hidden" id="employeeId">

            <div class="row">
                <div class="col-md-6">
                    <div class="mb-3">
                        <label for="employeeName" class="form-label">First Name *</label>
                        <input type="text" class="form-control" id="employeeName" required>
                    </div>
                </div>
                <div class="col-md-6">
                    <div class="mb-3">
                        <label for="surname" class="form-label">Surname *</label>
                        <input type="text" class="form-control" id="surname" required>
                    </div>
                </div>
            </div>

            <div class="row">
                <div class="col-md-6">
                    <div class="mb-3">
                        <label for="employer" class="form-label">Employer *</label>
                        <select class="form-select" id="employer" required>
                            <option value="">Select Employer</option>
                            <option value="SA Grinding Wheels">SA Grinding Wheels</option>
                            <option value="Scorpio Abrasives">Scorpio Abrasives</option>
                        </select>
                    </div>
                </div>
                <div class="col-md-6">
                    <div class="mb-3">
                        <label for="employmentStatus" class="form-label">Employment Status *</label>
                        <select class="form-select" id="employmentStatus" required>
                            <option value="">Select Status</option>
                            <option value="Permanent">Permanent</option>
                            <option value="Temporary">Temporary</option>
                            <option value="Contract">Contract</option>
                        </select>
                    </div>
                </div>
            </div>

            <div class="row">
                <div class="col-md-6">
                    <div class="mb-3">
                        <label for="hourlyRate" class="form-label">Hourly Rate *</label>
                        <input type="number" class="form-control" id="hourlyRate" step="0.01" min="0" required>
                    </div>
                </div>
                <div class="col-md-6">
                    <div class="mb-3">
                        <label for="clockInRef" class="form-label">Clock Number *</label>
                        <input type="text" class="form-control" id="clockInRef" required>
                    </div>
                </div>
            </div>

            <div class="row">
                <div class="col-md-6">
                    <div class="mb-3">
                        <label for="idNumber" class="form-label">ID Number *</label>
                        <input type="text" class="form-control" id="idNumber" maxlength="13" required>
                    </div>
                </div>
                <div class="col-md-6">
                    <div class="mb-3">
                        <label for="contactNumber" class="form-label">Contact Number *</label>
                        <input type="tel" class="form-control" id="contactNumber" required>
                    </div>
                </div>
            </div>

            <div class="mb-3">
                <label for="address" class="form-label">Address *</label>
                <textarea class="form-control" id="address" rows="2" required></textarea>
            </div>

            <div class="d-flex gap-2">
                <button type="submit" class="btn btn-primary">
                    <i class="fas fa-save"></i> Save
                </button>
                <button type="button" class="btn btn-secondary" onclick="cancelEmployeeForm()">
                    <i class="fas fa-times"></i> Cancel
                </button>
            </div>
        </form>
    </div>
</div>

<script>
function saveEmployee(event) {
    event.preventDefault();
    showLoading();

    const employeeId = document.getElementById('employeeId').value;
    const data = {
        employeeName: document.getElementById('employeeName').value,
        surname: document.getElementById('surname').value,
        employer: document.getElementById('employer').value,
        employmentStatus: document.getElementById('employmentStatus').value,
        hourlyRate: parseFloat(document.getElementById('hourlyRate').value),
        clockInRef: document.getElementById('clockInRef').value,
        idNumber: document.getElementById('idNumber').value,
        contactNumber: document.getElementById('contactNumber').value,
        address: document.getElementById('address').value
    };

    const method = employeeId ? 'updateEmployee' : 'addEmployee';
    const params = employeeId ? [employeeId, data] : [data];

    google.script.run
        .withSuccessHandler(function(result) {
            hideLoading();
            if (result.success) {
                showSuccess('Employee saved successfully');
                loadModule('employees');
            } else {
                showError(result.error);
            }
        })
        .withFailureHandler(handleError)
        [method](...params);
}
</script>
```

**Backend calls:**
- `addEmployee(data)` - Create new
- `updateEmployee(id, data)` - Update existing
- `getEmployeeById(id)` - Load for edit

---

## 3-4. Leave HTML Files

### LeaveList.html
- Display leave records with filters (employee, reason, date range)
- Button to add new leave
- Table showing: Employee, Start Date, Return Date, Total Days, Reason

### LeaveForm.html
- Employee dropdown (from listEmployees)
- Start date picker
- Return date picker
- Reason dropdown (AWOL, Sick Leave, Annual Leave, Unpaid Leave)
- Notes textarea
- Auto-calculate total days

**Backend calls:**
- `listLeave(filters)`
- `addLeave(data)`

---

## 5-7. Loan HTML Files

### LoanList.html
- Display loan transactions with filters
- Show running balance
- Color code: green for repayments, red for disbursements

### LoanForm.html
- Employee dropdown
- Transaction date
- Loan type (Disbursement/Repayment)
- Amount (positive or negative based on type)
- Disbursement mode dropdown
- Notes
- Display current balance prominently

### LoanBalanceWidget.html
- Reusable component
- Shows current loan balance for selected employee
- Can be embedded in PayrollForm

**Backend calls:**
- `getLoanHistory(employeeId, startDate, endDate)`
- `getCurrentLoanBalance(employeeId)`
- `addLoanTransaction(data)`

---

## 8-9. Timesheet HTML Files

### TimesheetImport.html
- File upload input (Excel)
- Preview table after upload
- Import button
- Progress indicator

### TimesheetApproval.html
- Table of pending timesheets
- Inline editing for hours/minutes
- Approve/Reject buttons per row
- Status badges (Pending/Approved/Rejected)
- Filter by employee and week

**Backend calls:**
- `importTimesheetExcel(fileBlob)`
- `listPendingTimesheets(filters)`
- `updatePendingTimesheet(id, data)`
- `approveTimesheet(id)`
- `rejectTimesheet(id, reason)`

---

## 10-11. Payroll HTML Files

### PayrollList.html
- Display all payslips with filters (employee, employer, week ending, date range)
- Table columns: Record#, Employee, Week Ending, Gross, Deductions, Net, Paid to Account
- Actions: View, Edit, Generate PDF

### PayrollForm.html (MOST COMPLEX)
**Features:**
- Employee dropdown (triggers hourly rate lookup)
- Week ending date picker
- Time inputs: Hours, Minutes, Overtime Hours, Overtime Minutes
- Additional pay: Leave Pay, Bonus Pay, Other Income
- Deductions: Other Deductions (with description)
- **Loan section** (critical):
  - Display Current Loan Balance (widget)
  - Loan Deduction This Week input
  - New Loan This Week input
  - Loan Disbursement Type dropdown (With Salary/Separate)
- **Real-time calculation preview:**
  - Standard Time
  - Overtime
  - Gross Salary
  - UIF (if permanent)
  - Total Deductions
  - Net Salary
  - **PAID TO ACCOUNT** (highlighted)
- Save button
- Generate PDF button (after save)

**JavaScript:**
- `calculatePreview()` function that calls backend calculatePayslip
- Auto-update on any field change
- Validation before save

**Backend calls:**
- `listEmployees()` - For dropdown
- `getEmployeeByName(name)` - Get hourly rate
- `getCurrentLoanBalance(employeeId)` - Show current balance
- `calculatePayslip(data)` - Real-time preview
- `createPayslip(data)` - Save
- `updatePayslip(recordNumber, data)` - Update
- `generatePayslipPDF(recordNumber)` - Generate PDF

---

## 12. ReportsMenu.html

**Purpose**: Report selection interface

```html
<div class="card">
    <div class="card-header">
        <h5><i class="fas fa-chart-bar"></i> Reports</h5>
    </div>
    <div class="card-body">
        <div class="row">
            <!-- Report 1: Outstanding Loans -->
            <div class="col-md-4">
                <div class="card">
                    <div class="card-body text-center">
                        <i class="fas fa-money-bill-wave fa-3x text-warning mb-3"></i>
                        <h5>Outstanding Loans Report</h5>
                        <p>Shows all employees with current loan balances</p>
                        <div class="mb-3">
                            <label>As of Date:</label>
                            <input type="date" class="form-control" id="loansReportDate">
                        </div>
                        <button class="btn btn-primary" onclick="generateLoansReport()">
                            <i class="fas fa-file-excel"></i> Generate
                        </button>
                    </div>
                </div>
            </div>

            <!-- Report 2: Individual Statement -->
            <div class="col-md-4">
                <div class="card">
                    <div class="card-body text-center">
                        <i class="fas fa-user-circle fa-3x text-info mb-3"></i>
                        <h5>Individual Statement</h5>
                        <p>Complete loan history for one employee</p>
                        <div class="mb-3">
                            <label>Employee:</label>
                            <select class="form-select" id="statementEmployee"></select>
                        </div>
                        <div class="mb-3">
                            <label>Start Date:</label>
                            <input type="date" class="form-control" id="statementStartDate">
                        </div>
                        <div class="mb-3">
                            <label>End Date:</label>
                            <input type="date" class="form-control" id="statementEndDate">
                        </div>
                        <button class="btn btn-primary" onclick="generateStatementReport()">
                            <i class="fas fa-file-excel"></i> Generate
                        </button>
                    </div>
                </div>
            </div>

            <!-- Report 3: Weekly Payroll Summary -->
            <div class="col-md-4">
                <div class="card">
                    <div class="card-body text-center">
                        <i class="fas fa-calendar-week fa-3x text-success mb-3"></i>
                        <h5>Weekly Payroll Summary</h5>
                        <p>Detailed breakdown for a specific week</p>
                        <div class="mb-3">
                            <label>Week Ending:</label>
                            <input type="date" class="form-control" id="weeklyReportDate">
                        </div>
                        <button class="btn btn-primary" onclick="generateWeeklyReport()">
                            <i class="fas fa-file-excel"></i> Generate
                        </button>
                    </div>
                </div>
            </div>
        </div>

        <!-- Report Results -->
        <div id="reportResults" class="mt-4" style="display: none;">
            <div class="alert alert-success">
                <h5><i class="fas fa-check-circle"></i> Report Generated!</h5>
                <p id="reportMessage"></p>
                <a id="reportLink" href="#" target="_blank" class="btn btn-success">
                    <i class="fas fa-external-link-alt"></i> Open Report
                </a>
            </div>
        </div>
    </div>
</div>

<script>
function loadEmployeeDropdown() {
    google.script.run
        .withSuccessHandler(function(result) {
            if (result.success) {
                const select = document.getElementById('statementEmployee');
                select.innerHTML = '<option value="">Select Employee</option>' +
                    result.data.map(emp => `<option value="${emp.id}">${emp.REFNAME}</option>`).join('');
            }
        })
        .listEmployees({ activeOnly: true });
}

function generateLoansReport() {
    showLoading();
    const asOfDate = document.getElementById('loansReportDate').value || new Date();

    google.script.run
        .withSuccessHandler(function(result) {
            hideLoading();
            if (result.success) {
                showReportResult('Outstanding Loans Report', result.data.url);
            } else {
                showError(result.error);
            }
        })
        .withFailureHandler(handleError)
        .generateOutstandingLoansReport(asOfDate);
}

function generateStatementReport() {
    showLoading();
    const employeeId = document.getElementById('statementEmployee').value;
    const startDate = document.getElementById('statementStartDate').value;
    const endDate = document.getElementById('statementEndDate').value;

    if (!employeeId || !startDate || !endDate) {
        hideLoading();
        showError('Please fill in all fields');
        return;
    }

    google.script.run
        .withSuccessHandler(function(result) {
            hideLoading();
            if (result.success) {
                showReportResult('Individual Statement', result.data.url);
            } else {
                showError(result.error);
            }
        })
        .withFailureHandler(handleError)
        .generateIndividualStatementReport(employeeId, startDate, endDate);
}

function generateWeeklyReport() {
    showLoading();
    const weekEnding = document.getElementById('weeklyReportDate').value;

    if (!weekEnding) {
        hideLoading();
        showError('Please select week ending date');
        return;
    }

    google.script.run
        .withSuccessHandler(function(result) {
            hideLoading();
            if (result.success) {
                showReportResult('Weekly Payroll Summary', result.data.url);
            } else {
                showError(result.error);
            }
        })
        .withFailureHandler(handleError)
        .generateWeeklyPayrollSummaryReport(weekEnding);
}

function showReportResult(reportName, url) {
    document.getElementById('reportMessage').textContent = reportName + ' generated successfully!';
    document.getElementById('reportLink').href = url;
    document.getElementById('reportResults').style.display = 'block';

    showSuccess(reportName + ' generated successfully!');
}

// Load employees on page load
loadEmployeeDropdown();
</script>
```

**Backend calls:**
- `generateOutstandingLoansReport(asOfDate)`
- `generateIndividualStatementReport(employeeId, startDate, endDate)`
- `generateWeeklyPayrollSummaryReport(weekEnding)`

---

## Implementation Checklist

### For Each HTML File:
1. ✅ Include Bootstrap 5 classes
2. ✅ Add Font Awesome icons
3. ✅ Implement google.script.run calls
4. ✅ Add loading/success/error handlers
5. ✅ Implement form validation
6. ✅ Format currency and dates
7. ✅ Test with real data
8. ✅ Mobile responsive

### Testing Steps:
1. Create each HTML file in Apps Script editor
2. Test in isolation via direct URL
3. Test via Dashboard navigation
4. Verify all backend calls work
5. Check error handling
6. Test edge cases

---

## Helper Functions (Available Globally from Dashboard.html)

```javascript
showLoading()            // Show loading overlay
hideLoading()            // Hide loading overlay
showSuccess(message)     // Green toast notification
showError(message)       // Red toast notification
showWarning(message)     // Yellow toast notification
showInfo(message)        // Blue toast notification
formatCurrency(amount)   // R1,234.56
formatDate(date)         // 17 October 2025
handleError(error)       // Standard error handler
```

---

## Deployment Order

1. Upload all .gs files (already complete ✅)
2. Upload Dashboard.html (already complete ✅)
3. Create remaining HTML files in this order:
   - EmployeeList.html, EmployeeForm.html
   - LeaveList.html, LeaveForm.html
   - LoanList.html, LoanForm.html, LoanBalanceWidget.html
   - TimesheetImport.html, TimesheetApproval.html
   - PayrollList.html, PayrollForm.html
   - ReportsMenu.html

4. Test each module as you create it
5. Fix any issues
6. Deploy web app
7. Share URL with users

---

**Document End**
