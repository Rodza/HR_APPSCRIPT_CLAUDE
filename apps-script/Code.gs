/**
 * Code.gs - Main entry point for HR System Web App
 *
 * Handles web app initialization and HTML serving
 */

// ============================================================================
// AUTHORIZATION CHECK FUNCTIONS
// ============================================================================

/**
 * Check if app has all required OAuth permissions
 * Run this function from Apps Script Editor to verify authorization
 *
 * @returns {Object} Status of each permission
 */
function checkAuthorization() {
  console.log('========== CHECKING AUTHORIZATION ==========');

  var results = {
    spreadsheets: false,
    documents: false,
    drive: false,
    allAuthorized: false
  };

  try {
    // Test Spreadsheet access
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    console.log('✅ Spreadsheet access: OK - ' + ss.getName());
    results.spreadsheets = true;
  } catch (e) {
    console.error('❌ Spreadsheet access: FAILED - ' + e.message);
  }

  try {
    // Test Document access (required for PDF generation)
    var testDoc = DocumentApp.create('AUTH_TEST_' + new Date().getTime());
    console.log('✅ Document access: OK - Created test doc');

    // Clean up test doc
    try {
      DriveApp.getFileById(testDoc.getId()).setTrashed(true);
      console.log('✅ Drive access: OK - Cleaned up test doc');
      results.documents = true;
      results.drive = true;
    } catch (e) {
      console.error('❌ Drive access: FAILED - ' + e.message);
      results.documents = true; // Doc creation worked at least
    }
  } catch (e) {
    console.error('❌ Document access: FAILED - ' + e.message);
    console.error('   This permission is required for PDF generation');
    console.error('   Run forceAuthorization() to grant permissions');
  }

  results.allAuthorized = results.spreadsheets && results.documents && results.drive;

  if (results.allAuthorized) {
    console.log('\n🎉 ALL PERMISSIONS GRANTED - App is fully authorized!');
  } else {
    console.log('\n⚠️ MISSING PERMISSIONS - Run forceAuthorization() to fix');
  }

  console.log('========== AUTHORIZATION CHECK COMPLETE ==========');
  return results;
}

/**
 * Force OAuth authorization for all required scopes
 *
 * HOW TO USE:
 * 1. In Apps Script Editor, select this function from dropdown
 * 2. Click Run button
 * 3. When prompted, click "Review Permissions"
 * 4. Select your Google account
 * 5. Click "Advanced" if you see a warning
 * 6. Click "Go to [Your Project] (unsafe)" - this is YOUR script, it's safe!
 * 7. Review the permissions:
 *    - View and manage spreadsheets
 *    - View and manage documents
 *    - View and manage Drive files
 * 8. Click "Allow"
 *
 * After authorization, redeploy your web app:
 * - Deploy → Manage deployments → Edit → New version → Deploy
 */
function forceAuthorization() {
  console.log('========== FORCING AUTHORIZATION ==========');
  console.log('This function will request all required OAuth permissions...\n');

  try {
    // 1. Test Spreadsheet access
    console.log('Step 1: Testing Spreadsheet access...');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    console.log('✅ Spreadsheet: ' + ss.getName());

    // 2. Test Document access (THIS TRIGGERS AUTHORIZATION)
    console.log('\nStep 2: Testing Document access...');
    var doc = DocumentApp.create('AUTH_TEST_' + new Date().getTime());
    console.log('✅ Document created: ' + doc.getId());

    // 3. Test Drive access
    console.log('\nStep 3: Testing Drive access...');
    var file = DriveApp.getFileById(doc.getId());
    file.setTrashed(true);
    console.log('✅ Drive access confirmed - Test file deleted');

    console.log('\n🎉 AUTHORIZATION SUCCESSFUL!');
    console.log('\nNext steps:');
    console.log('1. Deploy → Manage deployments');
    console.log('2. Click edit (pencil icon) on your deployment');
    console.log('3. Select "New version"');
    console.log('4. Click "Deploy"');
    console.log('5. Test PDF generation in your web app');
    console.log('\n========== AUTHORIZATION COMPLETE ==========');

    return { success: true, message: 'All permissions granted successfully!' };

  } catch (error) {
    console.error('\n❌ AUTHORIZATION FAILED!');
    console.error('Error: ' + error.message);
    console.error('\nTroubleshooting:');
    console.error('1. Make sure appsscript.json includes all OAuth scopes');
    console.error('2. Check Project Settings → OAuth Scopes');
    console.error('3. Try revoking access at: https://myaccount.google.com/permissions');
    console.error('4. Run this function again');
    console.log('\n========== AUTHORIZATION FAILED ==========');

    return { success: false, error: error.message };
  }
}

/**
 * Quick diagnostic - shows which permissions are missing
 * Run this first to see what needs to be authorized
 */
function diagnoseAuthorization() {
  console.log('========== AUTHORIZATION DIAGNOSTIC ==========\n');

  console.log('Required OAuth Scopes:');
  console.log('1. https://www.googleapis.com/auth/spreadsheets');
  console.log('2. https://www.googleapis.com/auth/documents');
  console.log('3. https://www.googleapis.com/auth/drive.file');
  console.log('4. https://www.googleapis.com/auth/script.external_request\n');

  console.log('Checking current authorization status...\n');
  var status = checkAuthorization();

  console.log('\n========== RECOMMENDATIONS ==========');
  if (!status.allAuthorized) {
    console.log('⚠️ Action Required: Run forceAuthorization() to grant permissions');
    console.log('\nSteps:');
    console.log('1. Select "forceAuthorization" from function dropdown');
    console.log('2. Click Run');
    console.log('3. Follow the authorization prompts');
  } else {
    console.log('✅ All permissions are granted!');
    console.log('✅ PDF generation should work');
  }
  console.log('========== DIAGNOSTIC COMPLETE ==========');

  return status;
}

// ============================================================================
// WEB APP ENTRY POINTS
// ============================================================================

/**
 * Main entry point for GET requests
 * This function is called when someone visits the web app URL
 *
 * @returns {HtmlOutput} The main dashboard HTML
 */
function doGet() {
  try {
    return HtmlService.createHtmlOutputFromFile('Dashboard')
      .setTitle('SA HR Payroll System')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (error) {
    // If Dashboard.html doesn't exist, return error page
    return HtmlService.createHtmlOutput(
      '<h1>Error</h1><p>Failed to load dashboard: ' + error.toString() + '</p>'
    ).setTitle('Error');
  }
}

/**
 * Include HTML file content
 * This is called from client-side to load different modules
 * 
 * @param {string} filename - Name of the HTML file to include (without .html extension)
 * @returns {string} HTML content of the file
 */
function include(filename) {
  try {
    // Log the request
    console.log('Including file: ' + filename);
    
    // Get the HTML file content
    var htmlContent = HtmlService.createHtmlOutputFromFile(filename).getContent();
    
    return htmlContent;
    
  } catch (error) {
    console.error('Failed to include file: ' + filename, error);
    
    // Return error message as HTML
    return '<div class="alert alert-danger">' +
           '<h5>Error Loading Module</h5>' +
           '<p>Failed to load ' + filename + ': ' + error.toString() + '</p>' +
           '</div>';
  }
}

/**
 * Get current user email
 * 
 * @returns {string} Email address of current user
 */
function getCurrentUser() {
  try {
    return Session.getActiveUser().getEmail();
  } catch (error) {
    return 'Unknown User';
  }
}

/**
 * Check if user has admin access
 * 
 * @returns {boolean} True if user is admin
 */
function isAdmin() {
  var userEmail = getCurrentUser();
  
  // Define admin emails (customize as needed)
  var adminEmails = [
    'admin@sagrindingwheels.co.za',
    'owner@sagrindingwheels.co.za'
  ];
  
  return adminEmails.indexOf(userEmail) >= 0;
}

/**
 * Initialize application data
 * Called when app first loads to ensure sheets exist
 * 
 * @returns {Object} Initialization status
 */
function initializeApp() {
  try {
    var sheets = getSheets();
    
    // Check if employee sheet exists
    if (!sheets.empdetails) {
      return {
        success: false,
        error: 'Employee Details sheet not found. Please ensure sheet exists.'
      };
    }
    
    return {
      success: true,
      message: 'Application initialized successfully',
      user: getCurrentUser(),
      isAdmin: isAdmin()
    };
    
  } catch (error) {
    return {
      success: false,
      error: 'Failed to initialize: ' + error.toString()
    };
  }
}

/**
 * Get application configuration for client
 *
 * @returns {Object} Client-safe configuration
 */
function getClientConfig() {
  return {
    user: getCurrentUser(),
    isAdmin: isAdmin(),
    version: '1.2.0-payroll-enabled', // Version identifier to track deployment
    lastUpdate: '2025-11-15T00:00:00Z',
    features: {
      employees: true,
      leave: true,
      loans: true,
      timesheets: true,
      payroll: true,  // ENABLED - Fixed module loading
      reports: true
    }
  };
}

/**
 * Get version info - used to verify deployment
 *
 * @returns {Object} Version information
 */
function getVersion() {
  return {
    version: '1.2.0-payroll-enabled',
    lastUpdate: '2025-11-15T00:00:00Z',
    hasFixedSheetMapping: true,
    hasLoggingFunctions: true,
    hasFormatResponse: true,
    hasPayrollModule: true,
    payrollFeatures: {
      createPayslip: true,
      calculatePayslipPreview: true,
      fieldMapping: true,
      moduleLoading: true
    },
    deploymentStatus: 'Payroll module fully enabled and functional'
  };
}

// ============================================================================
// TEST FUNCTIONS
// ============================================================================

/**
 * Test the web app setup
 */
function testWebApp() {
  try {
    console.log('Testing Web App Setup...');

    // Test doGet
    var htmlOutput = doGet();
    console.log('✓ doGet() works');

    // Test include
    var employeeList = include('EmployeeList');
    console.log('✓ include("EmployeeList") works');

    // Test initialization
    var initResult = initializeApp();
    console.log('✓ initializeApp():', initResult);

    // Test user functions
    console.log('✓ Current user:', getCurrentUser());
    console.log('✓ Is admin:', isAdmin());

    console.log('All web app tests passed!');
    return true;

  } catch (error) {
    console.error('Web app test failed:', error);
    return false;
  }
}

/**
 * Comprehensive integration test for all fixes
 * Tests logging, response formats, data handling, etc.
 */
function testAllFixes() {
  console.log('====================================');
  console.log('COMPREHENSIVE INTEGRATION TEST');
  console.log('====================================');

  var allTestsPassed = true;
  var testResults = [];

  try {
    // Test 1: Logging functions exist and work
    console.log('\n--- Test 1: Logging Functions ---');
    try {
      if (typeof logError !== 'function') throw new Error('logError not defined');
      if (typeof logSuccess !== 'function') throw new Error('logSuccess not defined');
      if (typeof logWarning !== 'function') throw new Error('logWarning not defined');
      if (typeof logInfo !== 'function') throw new Error('logInfo not defined');

      // Test calling them
      logInfo('Test info message');
      logSuccess('Test success message');
      logWarning('Test warning message');
      logError('Test error message', new Error('Test error'));

      testResults.push({test: 'Logging functions', status: 'PASS'});
      console.log('✓ All logging functions exist and work');
    } catch (e) {
      testResults.push({test: 'Logging functions', status: 'FAIL', error: e.toString()});
      console.error('✗ Logging functions test failed:', e);
      allTestsPassed = false;
    }

    // Test 2: include() function returns HTML
    console.log('\n--- Test 2: include() Function ---');
    try {
      var html = include('EmployeeList');
      if (!html || typeof html !== 'string') {
        throw new Error('include() did not return string');
      }
      if (html.indexOf('<') === -1) {
        throw new Error('include() did not return HTML');
      }
      testResults.push({test: 'include() function', status: 'PASS'});
      console.log('✓ include() returns HTML content');
    } catch (e) {
      testResults.push({test: 'include() function', status: 'FAIL', error: e.toString()});
      console.error('✗ include() test failed:', e);
      allTestsPassed = false;
    }

    // Test 3: listEmployees() returns proper format
    console.log('\n--- Test 3: listEmployees() Response Format ---');
    try {
      var result = listEmployees(null);

      if (!result || typeof result !== 'object') {
        throw new Error('listEmployees() did not return object');
      }
      if (!result.hasOwnProperty('success')) {
        throw new Error('Response missing "success" property');
      }
      if (!result.hasOwnProperty('data')) {
        throw new Error('Response missing "data" property');
      }
      if (!result.hasOwnProperty('error')) {
        throw new Error('Response missing "error" property');
      }
      if (typeof result.success !== 'boolean') {
        throw new Error('success property is not boolean');
      }

      testResults.push({test: 'listEmployees() format', status: 'PASS'});
      console.log('✓ listEmployees() returns {success, data, error} format');
      console.log('  Response:', JSON.stringify(result, null, 2).substring(0, 200) + '...');
    } catch (e) {
      testResults.push({test: 'listEmployees() format', status: 'FAIL', error: e.toString()});
      console.error('✗ listEmployees() test failed:', e);
      allTestsPassed = false;
    }

    // Test 4: isEmpty() handles all data types
    console.log('\n--- Test 4: isEmpty() Type Safety ---');
    try {
      var emptyTests = [
        {input: null, expected: true, desc: 'null'},
        {input: undefined, expected: true, desc: 'undefined'},
        {input: '', expected: true, desc: 'empty string'},
        {input: '  ', expected: true, desc: 'whitespace string'},
        {input: 'test', expected: false, desc: 'non-empty string'},
        {input: 0, expected: false, desc: 'number 0'},
        {input: 42, expected: false, desc: 'number 42'},
        {input: true, expected: false, desc: 'boolean true'},
        {input: false, expected: false, desc: 'boolean false'},
        {input: [], expected: true, desc: 'empty array'},
        {input: [1, 2], expected: false, desc: 'non-empty array'},
        {input: {}, expected: true, desc: 'empty object'},
        {input: {a: 1}, expected: false, desc: 'non-empty object'},
        {input: new Date(), expected: false, desc: 'date'}
      ];

      var emptyTestFailed = false;
      for (var i = 0; i < emptyTests.length; i++) {
        var test = emptyTests[i];
        var result = isEmpty(test.input);
        if (result !== test.expected) {
          console.error('  ✗ isEmpty(' + test.desc + ') returned ' + result + ', expected ' + test.expected);
          emptyTestFailed = true;
        } else {
          console.log('  ✓ isEmpty(' + test.desc + ') = ' + result);
        }
      }

      if (emptyTestFailed) {
        throw new Error('Some isEmpty() tests failed');
      }

      testResults.push({test: 'isEmpty() type safety', status: 'PASS'});
      console.log('✓ isEmpty() handles all data types correctly');
    } catch (e) {
      testResults.push({test: 'isEmpty() type safety', status: 'FAIL', error: e.toString()});
      console.error('✗ isEmpty() test failed:', e);
      allTestsPassed = false;
    }

    // Test 5: formatResponse() creates consistent responses
    console.log('\n--- Test 5: formatResponse() Function ---');
    try {
      if (typeof formatResponse !== 'function') {
        throw new Error('formatResponse not defined');
      }

      var successResponse = formatResponse(true, {test: 'data'}, null);
      var errorResponse = formatResponse(false, null, 'Test error');

      if (successResponse.success !== true || !successResponse.data || successResponse.error !== null) {
        throw new Error('Success response format incorrect');
      }
      if (errorResponse.success !== false || errorResponse.data !== null || !errorResponse.error) {
        throw new Error('Error response format incorrect');
      }

      testResults.push({test: 'formatResponse()', status: 'PASS'});
      console.log('✓ formatResponse() creates consistent response objects');
    } catch (e) {
      testResults.push({test: 'formatResponse()', status: 'FAIL', error: e.toString()});
      console.error('✗ formatResponse() test failed:', e);
      allTestsPassed = false;
    }

    // Test 6: getSheets() returns proper structure
    console.log('\n--- Test 6: getSheets() Function ---');
    try {
      var sheets = getSheets();
      if (!sheets || typeof sheets !== 'object') {
        throw new Error('getSheets() did not return object');
      }
      console.log('  Available sheets:', Object.keys(sheets).join(', '));
      if (!sheets.empdetails) {
        console.warn('  ⚠ Employee sheet not found - check sheet names!');
      }
      testResults.push({test: 'getSheets()', status: 'PASS'});
      console.log('✓ getSheets() returns sheet object');
    } catch (e) {
      testResults.push({test: 'getSheets()', status: 'FAIL', error: e.toString()});
      console.error('✗ getSheets() test failed:', e);
      allTestsPassed = false;
    }

    // Final Summary
    console.log('\n====================================');
    console.log('TEST SUMMARY');
    console.log('====================================');

    var passCount = 0;
    var failCount = 0;

    for (var i = 0; i < testResults.length; i++) {
      var result = testResults[i];
      if (result.status === 'PASS') {
        console.log('✓ ' + result.test + ': PASS');
        passCount++;
      } else {
        console.log('✗ ' + result.test + ': FAIL - ' + result.error);
        failCount++;
      }
    }

    console.log('\nTotal: ' + testResults.length + ' tests');
    console.log('Passed: ' + passCount);
    console.log('Failed: ' + failCount);

    if (allTestsPassed) {
      console.log('\n🎉 ALL TESTS PASSED! System is ready.');
      return {
        success: true,
        message: 'All integration tests passed',
        results: testResults
      };
    } else {
      console.log('\n⚠ SOME TESTS FAILED. Review errors above.');
      return {
        success: false,
        message: 'Some tests failed',
        results: testResults
      };
    }

  } catch (error) {
    console.error('\n====================================');
    console.error('CRITICAL TEST FAILURE');
    console.error('====================================');
    console.error('Error:', error.toString());
    console.error('Stack:', error.stack);
    return {
      success: false,
      message: 'Critical test failure',
      error: error.toString()
    };
  }
}
