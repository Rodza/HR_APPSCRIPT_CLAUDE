# Deploy Monthly Report Feature to Apps Script

The monthly report feature has been implemented in the local files and committed to git, but needs to be deployed to your Google Apps Script project.

## Quick Deployment Steps

### Option 1: Using Clasp (Automated - Recommended)

If you have clasp configured on your local machine:

```bash
cd /path/to/HR_APPSCRIPT_CLAUDE
clasp login  # Only needed first time
clasp push
```

### Option 2: Manual Deployment (Copy/Paste)

Follow these steps to manually update the three modified files:

---

## 1. Update Utils.gs

**Open:** Google Apps Script Editor > Utils.gs

**Add these two functions** (scroll to around line 898, after the `isFriday()` function):

```javascript
/**
 * Get first Friday of a month
 *
 * @param {Date|string} date - Any date in the month
 * @returns {Date} First Friday of that month
 */
function getFirstFridayOfMonth(date) {
  var d = parseDate(date);

  // Set to first day of month
  var firstDay = new Date(d.getFullYear(), d.getMonth(), 1);

  // Get day of week for first day (0 = Sunday, 5 = Friday)
  var dayOfWeek = firstDay.getDay();

  // Calculate days until first Friday
  var daysUntilFriday;
  if (dayOfWeek <= 5) {
    // Friday is in the same week
    daysUntilFriday = 5 - dayOfWeek;
  } else {
    // Friday is in the next week (Saturday = 6, need to go to next Friday)
    daysUntilFriday = 6; // From Saturday to Friday
  }

  // Add days to get to first Friday
  var firstFriday = new Date(firstDay);
  firstFriday.setDate(firstDay.getDate() + daysUntilFriday);

  return firstFriday;
}

/**
 * Get last Friday of a month
 *
 * @param {Date|string} date - Any date in the month
 * @returns {Date} Last Friday of that month
 */
function getLastFridayOfMonth(date) {
  var d = parseDate(date);

  // Set to last day of month
  var lastDay = new Date(d.getFullYear(), d.getMonth() + 1, 0);

  // Get day of week for last day (0 = Sunday, 5 = Friday)
  var dayOfWeek = lastDay.getDay();

  // Calculate days back to last Friday
  var daysBackToFriday;
  if (dayOfWeek >= 5) {
    // Friday is in the same week
    daysBackToFriday = dayOfWeek - 5;
  } else {
    // Friday is in the previous week
    daysBackToFriday = dayOfWeek + 2; // e.g., if Thursday (4), go back 6 days
  }

  // Subtract days to get to last Friday
  var lastFriday = new Date(lastDay);
  lastFriday.setDate(lastDay.getDate() - daysBackToFriday);

  return lastFriday;
}
```

---

## 2. Update Reports.gs

**Open:** Google Apps Script Editor > Reports.gs

**Add this function** (after the `generateWeeklyPayrollSummaryReport()` function, around line 648):

Copy the entire `generateMonthlyPayrollSummaryReport()` function from:
`apps-script/Reports.gs` lines 650-963

**Also add the test function** (after the `test_weeklyPayrollReport()` function, around line 1057):

Copy the `test_monthlyPayrollReport()` function from:
`apps-script/Reports.gs` lines 1059-1082

---

## 3. Update ReportsMenu.html

**Open:** Google Apps Script Editor > ReportsMenu.html

### 3a. Add the Monthly Report Card

Find the `<!-- Report 3: Weekly Payroll Summary -->` section (around line 91).

After the closing `</div>` for Report 3 (around line 121), add:

```html
                <!-- Report 4: Monthly Payroll Summary -->
                <div class="col-md-4 mb-4">
                    <div class="report-card">
                        <div class="report-icon">
                            <i class="fas fa-calendar-alt text-primary"></i>
                        </div>
                        <h5>Monthly Payroll Summary</h5>
                        <p class="text-muted">
                            Monthly salary totals per employee (first Friday to last Friday)
                        </p>

                        <div class="mb-3">
                            <label class="form-label">Month:</label>
                            <input type="month" class="form-control" id="monthlyReportMonth">
                        </div>

                        <div class="mb-3" style="visibility: hidden;">
                            <label class="form-label">&nbsp;</label>
                            <input type="text" class="form-control">
                        </div>

                        <div class="mb-3" style="visibility: hidden;">
                            <label class="form-label">&nbsp;</label>
                            <input type="text" class="form-control">
                        </div>

                        <button class="btn btn-primary w-100" onclick="generateMonthlyReport()">
                            <i class="fas fa-file-excel"></i> Generate Report
                        </button>
                    </div>
                </div>
```

### 3b. Update Help Text

Find the `<!-- Help Text -->` section (around line 173).

Replace the `<ul>` list with:

```html
                <ul class="mb-0">
                    <li><strong>Outstanding Loans:</strong> Lists all employees with loan balances as of the specified date</li>
                    <li><strong>Individual Statement:</strong> Shows opening balance, all transactions, and closing balance for one employee</li>
                    <li><strong>Weekly Payroll Summary:</strong> Complete payroll register with summary totals by employer</li>
                    <li><strong>Monthly Payroll Summary:</strong> Monthly salary totals per employee from first Friday to last Friday of the month</li>
                </ul>
```

### 3c. Add JavaScript for Monthly Report

Find the `<script>` section (around line 188).

After the line that sets `weeklyReportDate.value`, add:

```javascript
        // Set current month as default for monthly report
        const currentMonth = new Date().toISOString().substring(0, 7);
        document.getElementById('monthlyReportMonth').value = currentMonth;
```

Then, after the `generateWeeklyReport()` function (around line 321), add:

```javascript
        function generateMonthlyReport() {
            const monthInput = document.getElementById('monthlyReportMonth').value;

            if (!monthInput) {
                showError('Please select a month');
                return;
            }

            // Convert YYYY-MM to a date (use 15th of month as mid-point)
            const monthDate = monthInput + '-15';

            showLoading();
            showInfo('Generating Monthly Payroll Summary... This may take a moment.');

            google.script.run
                .withSuccessHandler(function(result) {
                    hideLoading();
                    if (result.success) {
                        showReportResult(
                            'Monthly Payroll Summary',
                            result.data.url,
                            `Month: ${result.data.month} | Period: ${formatDate(result.data.periodStart)} to ${formatDate(result.data.periodEnd)} | Employees: ${result.data.totalEmployees} | Total Paid to Accounts: ${formatCurrency(result.data.totalPaidToAccounts)}`
                        );
                    } else {
                        showError('Failed to generate report: ' + result.error);
                    }
                })
                .withFailureHandler(function(error) {
                    hideLoading();
                    showError('Error: ' + error.message);
                })
                .generateMonthlyPayrollSummaryReport(monthDate);
        }
```

---

## 4. Test the Implementation

After deploying all changes:

1. In Apps Script Editor, select function: `test_monthlyPayrollReport`
2. Click "Run"
3. Check the execution log for success message
4. Open the generated report URL

---

## Alternative: Copy Full Files

If you prefer, you can copy the entire content of each modified file:

1. **Utils.gs**: Copy from `apps-script/Utils.gs`
2. **Reports.gs**: Copy from `apps-script/Reports.gs`
3. **ReportsMenu.html**: Copy from `apps-script/ReportsMenu.html`

---

## Verification

Once deployed, the monthly report feature should be available:

✅ New "Monthly Payroll Summary" card appears in Reports section
✅ Month picker allows selection of any month
✅ Report calculates from first Friday to last Friday of the month
✅ Groups all payslips by employee and sums totals
✅ Generates formatted Excel report with Register and Summary tabs

---

## Need Help?

If you encounter issues:
1. Check the execution log in Apps Script for specific errors
2. Verify all three files were updated correctly
3. Ensure you saved all changes in the Apps Script editor
4. Try running the test function: `test_monthlyPayrollReport()`
