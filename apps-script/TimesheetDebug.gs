/**
 * TIMESHEETDEBUG.GS - Diagnostic and Debugging Tools
 * HR Payroll System for SA Grinding Wheels & Scorpio Abrasives
 *
 * This module provides comprehensive debugging tools to diagnose
 * authorization and server response issues in timesheet modules.
 */

// ==================== DIAGNOSTIC FUNCTIONS ====================

/**
 * Run comprehensive diagnostics on all timesheet functions
 * This will test each function and log the results
 *
 * @returns {Object} Diagnostic results
 */
function runTimesheetDiagnostics() {
  var results = {
    timestamp: new Date().toISOString(),
    tests: [],
    summary: {
      total: 0,
      passed: 0,
      failed: 0,
      warnings: 0
    }
  };

  Logger.log('\n' + '='.repeat(60));
  Logger.log('TIMESHEET MODULE DIAGNOSTICS');
  Logger.log('Started: ' + results.timestamp);
  Logger.log('='.repeat(60) + '\n');

  // Test 1: getTimeConfig
  results.tests.push(testGetTimeConfig());

  // Test 2: updateTimeConfig
  results.tests.push(testUpdateTimeConfig());

  // Test 3: resetTimeConfig
  results.tests.push(testResetTimeConfig());

  // Test 4: exportTimeConfig
  results.tests.push(testExportTimeConfig());

  // Test 5: listPendingTimesheets
  results.tests.push(testListPendingTimesheets());

  // Test 6: Sheet access
  results.tests.push(testSheetAccess());

  // Test 7: PropertiesService access
  results.tests.push(testPropertiesAccess());

  // Test 8: Session and authorization
  results.tests.push(testAuthorizationStatus());

  // Calculate summary
  results.tests.forEach(function(test) {
    results.summary.total++;
    if (test.status === 'PASSED') {
      results.summary.passed++;
    } else if (test.status === 'FAILED') {
      results.summary.failed++;
    } else if (test.status === 'WARNING') {
      results.summary.warnings++;
    }
  });

  // Print summary
  Logger.log('\n' + '='.repeat(60));
  Logger.log('DIAGNOSTIC SUMMARY');
  Logger.log('='.repeat(60));
  Logger.log('Total Tests: ' + results.summary.total);
  Logger.log('✅ Passed: ' + results.summary.passed);
  Logger.log('❌ Failed: ' + results.summary.failed);
  Logger.log('⚠️  Warnings: ' + results.summary.warnings);
  Logger.log('='.repeat(60) + '\n');

  return {
    success: true,
    data: results
  };
}

// ==================== INDIVIDUAL TEST FUNCTIONS ====================

/**
 * Test getTimeConfig function
 */
function testGetTimeConfig() {
  var testName = 'getTimeConfig()';
  Logger.log('\n--- Testing: ' + testName + ' ---');

  try {
    var result = getTimeConfig();

    // Check return format
    if (!result) {
      Logger.log('❌ FAILED: Function returned null/undefined');
      return {
        test: testName,
        status: 'FAILED',
        error: 'Function returned null/undefined',
        result: result
      };
    }

    if (typeof result !== 'object') {
      Logger.log('❌ FAILED: Function did not return an object');
      return {
        test: testName,
        status: 'FAILED',
        error: 'Expected object, got ' + typeof result,
        result: result
      };
    }

    if (!result.hasOwnProperty('success')) {
      Logger.log('❌ FAILED: Result missing "success" property');
      return {
        test: testName,
        status: 'FAILED',
        error: 'Missing "success" property',
        result: result
      };
    }

    if (!result.hasOwnProperty('data')) {
      Logger.log('❌ FAILED: Result missing "data" property');
      return {
        test: testName,
        status: 'FAILED',
        error: 'Missing "data" property',
        result: result
      };
    }

    if (result.success !== true) {
      Logger.log('⚠️  WARNING: success is not true');
      return {
        test: testName,
        status: 'WARNING',
        warning: 'success property is ' + result.success,
        result: result
      };
    }

    // Check data structure
    var config = result.data;
    var requiredKeys = ['standardStartTime', 'standardEndTime', 'graceMinutes'];
    var missingKeys = [];

    requiredKeys.forEach(function(key) {
      if (!config.hasOwnProperty(key)) {
        missingKeys.push(key);
      }
    });

    if (missingKeys.length > 0) {
      Logger.log('⚠️  WARNING: Config missing keys: ' + missingKeys.join(', '));
      return {
        test: testName,
        status: 'WARNING',
        warning: 'Missing config keys: ' + missingKeys.join(', '),
        result: result
      };
    }

    Logger.log('✅ PASSED: Function returned correct format');
    Logger.log('   Return format: {success: true, data: {...}}');
    Logger.log('   Config keys: ' + Object.keys(config).length);

    return {
      test: testName,
      status: 'PASSED',
      result: result
    };

  } catch (error) {
    Logger.log('❌ FAILED: Exception thrown: ' + error.message);
    Logger.log('   Stack: ' + error.stack);
    return {
      test: testName,
      status: 'FAILED',
      error: error.message,
      stack: error.stack
    };
  }
}

/**
 * Test updateTimeConfig function
 */
function testUpdateTimeConfig() {
  var testName = 'updateTimeConfig()';
  Logger.log('\n--- Testing: ' + testName + ' ---');

  try {
    var testConfig = {
      standardStartTime: '07:30',
      graceMinutes: 5
    };

    var result = updateTimeConfig(testConfig);

    if (!result) {
      Logger.log('❌ FAILED: Function returned null/undefined');
      return {
        test: testName,
        status: 'FAILED',
        error: 'Function returned null/undefined'
      };
    }

    if (!result.hasOwnProperty('success')) {
      Logger.log('❌ FAILED: Result missing "success" property');
      return {
        test: testName,
        status: 'FAILED',
        error: 'Missing "success" property',
        result: result
      };
    }

    Logger.log('✅ PASSED: Function returned result with success property');
    Logger.log('   Success: ' + result.success);

    return {
      test: testName,
      status: 'PASSED',
      result: result
    };

  } catch (error) {
    Logger.log('❌ FAILED: Exception thrown: ' + error.message);
    return {
      test: testName,
      status: 'FAILED',
      error: error.message
    };
  }
}

/**
 * Test resetTimeConfig function
 */
function testResetTimeConfig() {
  var testName = 'resetTimeConfig()';
  Logger.log('\n--- Testing: ' + testName + ' ---');

  try {
    var result = resetTimeConfig();

    if (!result) {
      Logger.log('❌ FAILED: Function returned null/undefined');
      return {
        test: testName,
        status: 'FAILED',
        error: 'Function returned null/undefined'
      };
    }

    if (!result.hasOwnProperty('success')) {
      Logger.log('❌ FAILED: Result missing "success" property');
      return {
        test: testName,
        status: 'FAILED',
        error: 'Missing "success" property'
      };
    }

    Logger.log('✅ PASSED: Function returned result with success property');

    return {
      test: testName,
      status: 'PASSED',
      result: result
    };

  } catch (error) {
    Logger.log('❌ FAILED: Exception thrown: ' + error.message);
    return {
      test: testName,
      status: 'FAILED',
      error: error.message
    };
  }
}

/**
 * Test exportTimeConfig function
 */
function testExportTimeConfig() {
  var testName = 'exportTimeConfig()';
  Logger.log('\n--- Testing: ' + testName + ' ---');

  try {
    var result = exportTimeConfig();

    if (!result) {
      Logger.log('❌ FAILED: Function returned null/undefined');
      return {
        test: testName,
        status: 'FAILED',
        error: 'Function returned null/undefined'
      };
    }

    if (!result.hasOwnProperty('success')) {
      Logger.log('❌ FAILED: Result missing "success" property');
      return {
        test: testName,
        status: 'FAILED',
        error: 'Missing "success" property'
      };
    }

    Logger.log('✅ PASSED: Function returned result with success property');

    return {
      test: testName,
      status: 'PASSED',
      result: result
    };

  } catch (error) {
    Logger.log('❌ FAILED: Exception thrown: ' + error.message);
    return {
      test: testName,
      status: 'FAILED',
      error: error.message
    };
  }
}

/**
 * Test listPendingTimesheets function
 */
function testListPendingTimesheets() {
  var testName = 'listPendingTimesheets()';
  Logger.log('\n--- Testing: ' + testName + ' ---');

  try {
    var result = listPendingTimesheets({});

    if (!result) {
      Logger.log('❌ FAILED: Function returned null/undefined');
      return {
        test: testName,
        status: 'FAILED',
        error: 'Function returned null/undefined'
      };
    }

    if (!result.hasOwnProperty('success')) {
      Logger.log('❌ FAILED: Result missing "success" property');
      return {
        test: testName,
        status: 'FAILED',
        error: 'Missing "success" property'
      };
    }

    Logger.log('✅ PASSED: Function returned result with success property');
    if (result.success) {
      Logger.log('   Records returned: ' + (result.data ? result.data.length : 0));
    }

    return {
      test: testName,
      status: 'PASSED',
      result: result
    };

  } catch (error) {
    Logger.log('❌ FAILED: Exception thrown: ' + error.message);
    return {
      test: testName,
      status: 'FAILED',
      error: error.message
    };
  }
}

/**
 * Test sheet access
 */
function testSheetAccess() {
  var testName = 'Sheet Access';
  Logger.log('\n--- Testing: ' + testName + ' ---');

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (!ss) {
      Logger.log('❌ FAILED: Cannot access active spreadsheet');
      return {
        test: testName,
        status: 'FAILED',
        error: 'Cannot access active spreadsheet'
      };
    }

    Logger.log('✅ PASSED: Can access spreadsheet: ' + ss.getName());

    // Check for PENDING_TIMESHEETS sheet
    var pendingSheet = ss.getSheetByName('PENDING_TIMESHEETS');
    if (pendingSheet) {
      Logger.log('   PENDING_TIMESHEETS sheet exists');
    } else {
      Logger.log('   ⚠️  PENDING_TIMESHEETS sheet not found');
    }

    return {
      test: testName,
      status: 'PASSED',
      spreadsheetName: ss.getName(),
      hasPendingSheet: !!pendingSheet
    };

  } catch (error) {
    Logger.log('❌ FAILED: Exception thrown: ' + error.message);
    return {
      test: testName,
      status: 'FAILED',
      error: error.message
    };
  }
}

/**
 * Test PropertiesService access
 */
function testPropertiesAccess() {
  var testName = 'PropertiesService Access';
  Logger.log('\n--- Testing: ' + testName + ' ---');

  try {
    var props = PropertiesService.getScriptProperties();

    if (!props) {
      Logger.log('❌ FAILED: Cannot access PropertiesService');
      return {
        test: testName,
        status: 'FAILED',
        error: 'Cannot access PropertiesService'
      };
    }

    // Try to read and write a test property
    props.setProperty('DIAGNOSTIC_TEST', 'test_value');
    var testValue = props.getProperty('DIAGNOSTIC_TEST');

    if (testValue === 'test_value') {
      Logger.log('✅ PASSED: PropertiesService read/write working');
      props.deleteProperty('DIAGNOSTIC_TEST');
    } else {
      Logger.log('⚠️  WARNING: PropertiesService write/read mismatch');
      return {
        test: testName,
        status: 'WARNING',
        warning: 'Could write but not read property'
      };
    }

    return {
      test: testName,
      status: 'PASSED'
    };

  } catch (error) {
    Logger.log('❌ FAILED: Exception thrown: ' + error.message);
    return {
      test: testName,
      status: 'FAILED',
      error: error.message
    };
  }
}

/**
 * Test authorization status
 */
function testAuthorizationStatus() {
  var testName = 'Authorization Status';
  Logger.log('\n--- Testing: ' + testName + ' ---');

  try {
    var user = Session.getActiveUser().getEmail();

    if (!user) {
      Logger.log('⚠️  WARNING: Cannot get active user email');
      return {
        test: testName,
        status: 'WARNING',
        warning: 'Cannot get active user email'
      };
    }

    Logger.log('✅ PASSED: User authorized: ' + user);

    // IMPORTANT: Also test Drive API access (needed for Excel import)
    Logger.log('Testing Drive API access...');
    try {
      // Try to use Drive Advanced Service
      var testFile = DriveApp.createFile('test', 'test');
      var testFileId = testFile.getId();

      // Try Drive API v3 (Advanced Service)
      var driveFile = Drive.Files.get(testFileId);
      Logger.log('✅ Drive API v3 access: OK');

      // Clean up
      DriveApp.getFileById(testFileId).setTrashed(true);
    } catch (driveError) {
      Logger.log('❌ CRITICAL: Drive API access FAILED');
      Logger.log('   Error: ' + driveError.message);
      Logger.log('   This is why Excel import fails!');
      return {
        test: testName,
        status: 'FAILED',
        error: 'Drive API not authorized: ' + driveError.message,
        user: user,
        critical: true,
        solution: 'Enable Drive API in Advanced Services'
      };
    }

    return {
      test: testName,
      status: 'PASSED',
      user: user,
      driveApiEnabled: true
    };

  } catch (error) {
    Logger.log('❌ FAILED: Exception thrown: ' + error.message);
    return {
      test: testName,
      status: 'FAILED',
      error: error.message
    };
  }
}

// ==================== FUNCTION CALL LOGGER ====================

/**
 * Wrapper function to log all calls to getTimeConfig
 * Use this temporarily to track when and how the function is called
 */
function getTimeConfig_DEBUG() {
  Logger.log('\n🔍 DEBUG: getTimeConfig() called');
  Logger.log('   Time: ' + new Date().toISOString());
  Logger.log('   Caller: ' + new Error().stack.split('\n')[2]);

  var result = getTimeConfig();

  Logger.log('   Result type: ' + typeof result);
  Logger.log('   Result is null: ' + (result === null));
  Logger.log('   Result is undefined: ' + (result === undefined));

  if (result) {
    Logger.log('   Result.success: ' + result.success);
    Logger.log('   Result.data exists: ' + (!!result.data));
    Logger.log('   Result keys: ' + Object.keys(result).join(', '));
  }

  return result;
}

/**
 * Test function that returns various response formats
 * to help debug client-side handling
 */
function testVariousResponseFormats() {
  return {
    success: true,
    data: {
      test1_null: null,
      test2_undefined: undefined,
      test3_emptyObject: {},
      test4_emptyArray: [],
      test5_string: 'test',
      test6_number: 123,
      test7_boolean: true,
      test8_nestedObject: {
        nested: {
          value: 'deep'
        }
      }
    }
  };
}

/**
 * Function to test if server-side returns are being received by client
 */
function pingTest() {
  Logger.log('🏓 PING received at: ' + new Date().toISOString());

  return {
    success: true,
    data: {
      message: 'PONG',
      timestamp: new Date().toISOString(),
      randomValue: Math.random()
    }
  };
}
