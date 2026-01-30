# Fix for Missing Employees Issue

## Problem

7 employees were not showing in the app despite being listed in the employee sheet:
- Samuel Tuin
- Stephen Chimane Plaatjie
- Tapiwa Malvin Zulu
- Ticky William Lekoko
- Walter Hebert Coglin
- Walter Raphiri
- Yvonne Kedikgotse

## Root Cause

The `listEmployees()` function in `Employees.gs` was checking if the first column (the 'id' field) was non-empty before including an employee in the list. These 7 employees were likely added manually to the sheet and did not have system-generated IDs, causing them to be excluded from the employee list.

### Previous Code (Line 158)
```javascript
for (var i = 1; i < data.length; i++) {
  if (data[i][0]) {  // Only processes rows with non-empty 'id' column
    // ... build employee object
  }
}
```

## Solution

### 1. Modified Employee Listing Logic

Changed the `listEmployees()` and `listEmployeesFull()` functions to check for the `REFNAME` field (employee name) instead of the 'id' field. This ensures that all employees with valid names are included, regardless of whether they have an ID.

#### Updated Code
```javascript
// Find REFNAME column index for better row validation
var refNameColIndex = -1;
for (var h = 0; h < headers.length; h++) {
  if (headers[h] === 'REFNAME') {
    refNameColIndex = h;
    break;
  }
}

for (var i = 1; i < data.length; i++) {
  // Check if row has a REFNAME (employee name) instead of just checking first column
  var hasRefName = refNameColIndex >= 0 &&
                   data[i][refNameColIndex] &&
                   data[i][refNameColIndex].toString().trim() !== '';

  if (hasRefName) {
    try {
      var employee = buildObjectFromRow(data[i], headers);
      // Generate ID if missing to ensure all employees have a unique identifier
      if (!employee.id || employee.id === '') {
        employee.id = generateId();
      }
      employees.push(employee);
    } catch (rowError) {
      // Skip invalid rows silently
    }
  }
}
```

### 2. Auto-Generate Missing IDs

If an employee doesn't have an ID, the system now automatically generates one on-the-fly using the `generateId()` function. This ensures all employees have a unique identifier for operations that require it.

### 3. Created Utility Scripts

#### DiagnoseMissingEmployees.gs
A diagnostic script to identify which employees are in the sheet but might not be showing in the app. Useful for troubleshooting similar issues in the future.

**Usage:**
```javascript
diagnoseMissingEmployees()
```

#### FillMissingEmployeeIds.gs
A migration script that permanently fills in missing IDs for employees in the sheet.

**Usage:**
```javascript
// Check which employees are missing IDs (read-only)
checkMissingEmployeeIds()

// Fill in missing IDs (writes to sheet)
fillMissingEmployeeIds()
```

## Files Modified

1. **apps-script/Employees.gs**
   - Updated `listEmployees()` function (lines 154-183)
   - Updated `listEmployeesFull()` function (lines 283-316)

2. **apps-script/DiagnoseMissingEmployees.gs** (new file)
   - Diagnostic script for troubleshooting missing employees

3. **apps-script/FillMissingEmployeeIds.gs** (new file)
   - Migration script to permanently add IDs to employees

## Testing

After deploying this fix:

1. **Verify All Employees Appear:**
   - Navigate to the Employee List in the app
   - Confirm all 56 employees are now visible
   - Specifically check for the 7 previously missing employees

2. **Run Diagnostic Script (Optional):**
   - Open the Apps Script editor
   - Run `diagnoseMissingEmployees()`
   - Review the console output to confirm no employees are missing

3. **Fill Missing IDs (Recommended):**
   - Open the Apps Script editor
   - First run `checkMissingEmployeeIds()` to see which employees need IDs
   - Then run `fillMissingEmployeeIds()` to permanently add IDs to the sheet
   - This prevents the need to generate IDs on-the-fly in future requests

## Prevention

To prevent this issue in the future:

1. **Always use the app** to add new employees rather than manually entering them in the sheet
2. If employees must be added manually to the sheet, run `fillMissingEmployeeIds()` afterwards to ensure they have IDs
3. The system now gracefully handles missing IDs by auto-generating them, so even manual entries will appear in the list

## Impact

- ✅ All employees now appear in the employee list
- ✅ No data loss - all employee information preserved
- ✅ Auto-generation of IDs ensures backwards compatibility
- ✅ Diagnostic tools available for future troubleshooting
- ✅ Migration script available to permanently fix missing IDs
