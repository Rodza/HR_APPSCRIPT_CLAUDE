/**
 * Test that the sheet mapping fix worked
 * This should now correctly map EMPLOYEE DETAILS to empdetails
 * and NOT let EmployeeLoans overwrite it
 */
function testSheetMappingFix() {
  console.log('=== TESTING SHEET MAPPING FIX ===\n');

  try {
    // Get sheets using the fixed function
    var sheets = getSheets();

    console.log('Mapped sheets:', Object.keys(sheets).join(', '));

    // Test 1: empdetails should exist
    if (!sheets.empdetails) {
      console.error('✗ FAIL: empdetails sheet not mapped!');
      return false;
    }
    console.log('✓ empdetails sheet mapped to:', sheets.empdetails.getName());

    // Test 2: Should be "EMPLOYEE DETAILS" not "EmployeeLoans"
    if (sheets.empdetails.getName() === 'EMPLOYEE DETAILS') {
      console.log('✓ CORRECT: empdetails = EMPLOYEE DETAILS');
    } else {
      console.error('✗ WRONG: empdetails =', sheets.empdetails.getName());
      console.error('  (Expected: EMPLOYEE DETAILS)');
      return false;
    }

    // Test 3: loans should exist and be separate
    if (!sheets.loans) {
      console.error('✗ FAIL: loans sheet not mapped!');
      return false;
    }
    console.log('✓ loans sheet mapped to:', sheets.loans.getName());

    // Test 4: loans should be "EmployeeLoans"
    if (sheets.loans.getName() === 'EmployeeLoans') {
      console.log('✓ CORRECT: loans = EmployeeLoans');
    } else {
      console.error('✗ WRONG: loans =', sheets.loans.getName());
      return false;
    }

    // Test 5: They should be DIFFERENT sheets
    if (sheets.empdetails.getSheetId() !== sheets.loans.getSheetId()) {
      console.log('✓ CORRECT: empdetails and loans are separate sheets');
    } else {
      console.error('✗ FAIL: empdetails and loans are the SAME sheet!');
      return false;
    }

    // Test 6: Get employee headers to verify correct sheet
    console.log('\n--- Employee Sheet Headers ---');
    var empHeaders = sheets.empdetails.getRange(1, 1, 1, sheets.empdetails.getLastColumn()).getValues()[0];
    console.log('Headers:', empHeaders.slice(0, 10).join(', '));

    if (empHeaders[0] === 'id' && empHeaders[2] === 'EMPLOYEE NAME') {
      console.log('✓ CORRECT: Headers match employee sheet format');
    } else {
      console.error('✗ WRONG: Headers don\'t match employee format!');
      console.error('  First header:', empHeaders[0]);
      return false;
    }

    // Test 7: Try to list employees
    console.log('\n--- Testing listEmployees() ---');
    var result = listEmployees(null);

    console.log('Response format:', {
      success: result.success,
      dataType: Array.isArray(result.data) ? 'array' : typeof result.data,
      dataLength: result.data ? result.data.length : 0,
      error: result.error
    });

    if (result.success && Array.isArray(result.data)) {
      console.log('✓ listEmployees() returned proper format');
      console.log('  Found', result.data.length, 'employee(s)');

      if (result.data.length > 0) {
        var emp = result.data[0];
        console.log('\n--- First Employee ---');
        console.log('  REFNAME:', emp.REFNAME);
        console.log('  EMPLOYER:', emp.EMPLOYER);
        console.log('  Has LoanID?:', emp.hasOwnProperty('LoanID') ? 'YES (BAD!)' : 'NO (GOOD!)');

        if (emp.hasOwnProperty('LoanID')) {
          console.error('✗ FAIL: Employee data contains loan fields!');
          console.error('  This means the wrong sheet is being used.');
          return false;
        } else {
          console.log('✓ CORRECT: Employee data is clean (no loan fields)');
        }
      }
    } else {
      console.error('✗ FAIL: listEmployees() failed');
      console.error('  Error:', result.error);
      return false;
    }

    console.log('\n=== ALL TESTS PASSED! ===');
    console.log('Sheet mapping is now correct!');
    return true;

  } catch (error) {
    console.error('\n=== TEST FAILED ===');
    console.error('Error:', error.toString());
    console.error('Stack:', error.stack);
    return false;
  }
}
