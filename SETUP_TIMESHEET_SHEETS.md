# Timesheet Sheets Setup Guide

## Required Sheets

The Timesheet Import system requires 3 sheets with specific column headers:

1. **RAW_CLOCK_DATA** (21 columns)
2. **CLOCK_IN_IMPORTS** (14 columns)
3. **PENDING_TIMESHEETS** (25 columns)

---

## Quick Setup (Recommended)

### Option 1: From Google Sheets Menu

1. Open your HR spreadsheet
2. Click **HR System** menu
3. Click **Timesheets** → **⚙️ Setup Sheets**
4. Click **YES** to confirm
5. Done! All sheets will be created with correct headers

### Option 2: From Apps Script Editor

1. Open Apps Script Editor (Extensions → Apps Script)
2. Select function: `setupAllTimesheetSheets`
3. Click **Run**
4. Check execution log for results

### Option 3: Setup Individual Sheets

Run these functions separately if you only need specific sheets:

- `setupRawClockDataSheet()` - Creates RAW_CLOCK_DATA
- `setupClockInImportsSheet()` - Creates CLOCK_IN_IMPORTS
- `setupPendingTimesheetsSheet()` - Creates PENDING_TIMESHEETS

---

## Sheet Specifications

### 1. RAW_CLOCK_DATA (21 columns)

Stores permanent record of all clock-in transactions.

**Columns:**
```
1.  ID - UUID, auto-generated unique identifier
2.  IMPORT_ID - Links to CLOCK_IN_IMPORTS.IMPORT_ID
3.  EMPLOYEE_CLOCK_REF - Clock-in card number from device
4.  EMPLOYEE_NAME - Resolved from EMPLOYEE DETAILS
5.  EMPLOYEE_ID - Links to EMPLOYEE DETAILS.id
6.  WEEK_ENDING - Week ending date
7.  PUNCH_DATE - Date of clock punch
8.  DEVICE_NAME - Clock In/Bathroom/Exit
9.  PUNCH_TIME - Actual punch timestamp
10. DEPARTMENT - From clock-in data
11. DEVICE_SN - Device serial number
12. TYPE - Transaction type (Normal, etc.)
13. SOURCE - Transaction source (Access Device, etc.)
14. TIME_ZONE - Timezone from device
15. VERIFICATION_MODE - Verification method (Face, Fingerprint, etc.)
16. MOBILE_PUNCHCARD - Mobile punchcard number if applicable
17. UPLOAD_TIME - When data was uploaded from device
18. STATUS - Draft/Reviewed/Approved/Locked
19. CREATED_DATE - When record was imported
20. LOCKED_DATE - When record was locked
21. LOCKED_BY - User who locked record
```

**Purpose:**
- Permanent audit trail
- Raw transaction data
- Can be reprocessed if needed

---

### 2. CLOCK_IN_IMPORTS (14 columns)

Tracks import batches and prevents duplicates.

**Columns:**
```
1.  IMPORT_ID - UUID, unique import batch identifier
2.  IMPORT_DATE - When import was performed
3.  WEEK_ENDING - Week ending date for this batch
4.  FILENAME - Original uploaded file name
5.  FILE_HASH - Hash of file content (duplicate detection)
6.  TOTAL_RECORDS - Count of records in import
7.  MATCHED_EMPLOYEES - Successfully matched employees
8.  UNMATCHED_REFS - JSON array of unmatched ClockInRefs
9.  STATUS - Active/Replaced/Archived
10. IMPORTED_BY - User email who imported
11. OVERRIDE_APPLIED - If user overrode warnings
12. REPLACED_BY_IMPORT_ID - If replaced, links to new import
13. REPLACED_DATE - When this import was replaced
14. NOTES - Import notes or override reason
```

**Purpose:**
- Prevent duplicate imports
- Track import history
- Manage import replacements
- Audit trail of who imported what

---

### 3. PENDING_TIMESHEETS (25 columns)

Staging area for timesheet approval workflow.

**Columns:**
```
1.  ID - UUID, unique timesheet identifier
2.  RAW_DATA_IMPORT_ID - Links to CLOCK_IN_IMPORTS.IMPORT_ID
3.  EMPLOYEE_ID - Links to EMPLOYEE DETAILS.id
4.  EMPLOYEE NAME - From EMPLOYEE DETAILS (keep original format)
5.  EMPLOYEE_CLOCK_REF - Clock-in reference number
6.  WEEKENDING - Week ending date (keep original format)
7.  CALCULATED_TOTAL_HOURS - System calculated total hours
8.  CALCULATED_TOTAL_MINUTES - System calculated total minutes
9.  HOURS - Standard hours (manual entry)
10. MINUTES - Standard minutes (manual entry)
11. OVERTIMEHOURS - Overtime hours (manual entry)
12. OVERTIMEMINUTES - Overtime minutes (manual entry)
13. LUNCH_DEDUCTION_MIN - Calculated lunch deduction
14. BATHROOM_TIME_MIN - Total bathroom time
15. RECON_DETAILS - JSON with full reconciliation data
16. WARNINGS - JSON array of any warnings
17. NOTES - Manual notes/adjustments reason
18. STATUS - Pending/Approved/Rejected
19. IMPORTED_BY - User who created
20. IMPORTED_DATE - When timesheet was created
21. REVIEWED_BY - User who reviewed
22. REVIEWED_DATE - When approved/rejected
23. PAYSLIP_ID - Links to MASTERSALARY.RECORDNUMBER
24. IS_LOCKED - True once payslip generated
25. LOCKED_DATE - When record was locked
```

**Purpose:**
- Review imported timesheets before payroll
- Edit/adjust hours if needed
- Track approval workflow
- Link to payslips

---

## Manual Setup (If Needed)

If you prefer to create sheets manually:

1. **Create Sheet:**
   - Right-click sheet tab
   - Insert new sheet
   - Rename to exact name (e.g., `RAW_CLOCK_DATA`)

2. **Add Headers:**
   - Copy column names from specification above
   - Paste into Row 1
   - Format as bold
   - Set background color to gray (#f3f3f3)
   - Freeze row 1

3. **Verify:**
   - Column count must match exactly
   - Column names must match exactly (case-sensitive!)
   - No extra spaces in names

**⚠️ WARNING:** Manual setup is error-prone. Use automated setup functions instead!

---

## Troubleshooting

### Sheet Not Found Error

If you get "sheet not found" errors:

1. Check sheet names match **exactly** (case-sensitive):
   - ✅ `RAW_CLOCK_DATA`
   - ❌ `raw_clock_data`
   - ❌ `Raw Clock Data`

2. Run setup function:
   ```javascript
   setupAllTimesheetSheets()
   ```

### Column Order Wrong

If data appears in wrong columns:

1. **DO NOT** manually rearrange columns
2. Delete the sheet entirely
3. Re-run setup function
4. Re-import data

### Existing Data Migration

If you have existing data in wrong column order:

1. Use migration script: `migratePendingTimesheetsData()`
2. Verify with: `verifyMigration()`
3. See `MigratePendingTimesheets.gs` for details

---

## Verification

After setup, verify sheets are correct:

### Visual Check
1. Open each sheet
2. Verify header row matches specification
3. Check column count

### Code Check
Run this in Apps Script:
```javascript
function verifySetup() {
  var sheets = getSheets();

  Logger.log('RAW_CLOCK_DATA: ' + (sheets.rawClockData ? '✅ Found' : '❌ Missing'));
  Logger.log('CLOCK_IN_IMPORTS: ' + (sheets.clockImports ? '✅ Found' : '❌ Missing'));
  Logger.log('PENDING_TIMESHEETS: ' + (sheets.pending ? '✅ Found' : '❌ Missing'));
}
```

---

## Next Steps

After setup is complete:

1. ✅ Sheets are created with correct headers
2. ✅ Headers are formatted (bold, gray background)
3. ✅ Row 1 is frozen
4. ✅ Ready to import clock data

Now you can:
- **Import Clock Data**: HR System → Timesheets → Import Clock Data
- **Upload Excel file** from your clock-in system
- **Review timesheets** in Pending Approval
- **Approve/reject** and create payslips

---

## Support

If you encounter issues:

1. Check execution logs (View → Logs in Apps Script)
2. Verify sheet names match exactly
3. Re-run setup function
4. Check `SHEET_COLUMN_VERIFICATION.md` for column mappings

All setup functions are in: `apps-script/Timesheets.gs`
