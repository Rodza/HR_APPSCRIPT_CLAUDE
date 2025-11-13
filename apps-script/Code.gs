/**
 * CODE.GS - Main Entry Points
 * HR Payroll System for SA Grinding Wheels & Scorpio Abrasives
 *
 * This file contains the main entry points for the web application:
 * - doGet() - Serves the web app UI
 * - doPost() - Handles POST requests (for future webhook integrations)
 * - Test functions for quick verification
 */

/**
 * Serves the web application
 * This function is called when users access the web app URL
 *
 * @param {Object} e - Event object containing request parameters
 * @returns {HtmlOutput} The HTML page to display
 */
function doGet(e) {
  try {
    Logger.log('\n========== DOGET CALLED ==========');
    Logger.log('ℹ️ Request parameters: ' + JSON.stringify(e.parameter));

    // Create HTML output from Dashboard.html
    const template = HtmlService.createTemplateFromFile('Dashboard');

    // You can pass server-side variables to the template if needed
    template.userEmail = Session.getActiveUser().getEmail();
    template.today = new Date().toISOString();

    const htmlOutput = template.evaluate()
      .setTitle('HR Payroll System')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');

    Logger.log('✅ Dashboard loaded successfully');
    Logger.log('========== DOGET COMPLETE ==========\n');

    return htmlOutput;

  } catch (error) {
    Logger.log('❌ ERROR in doGet: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== DOGET FAILED ==========\n');

    // Return error page
    return HtmlService.createHtmlOutput(
      '<h1>Error Loading Application</h1>' +
      '<p>' + error.message + '</p>' +
      '<p>Please contact your system administrator.</p>'
    );
  }
}

/**
 * Handles POST requests to the web app
 * Can be used for webhooks or API integrations
 *
 * @param {Object} e - Event object containing POST data
 * @returns {ContentService.TextOutput} JSON response
 */
function doPost(e) {
  try {
    Logger.log('\n========== DOPOST CALLED ==========');
    Logger.log('ℹ️ POST data: ' + JSON.stringify(e.postData));

    const params = JSON.parse(e.postData.contents);

    // Handle different POST actions
    let result;
    switch (params.action) {
      case 'generateReport':
        // Future: Handle report generation via webhook
        result = { success: true, message: 'Report generation initiated' };
        break;

      default:
        result = { success: false, error: 'Unknown action: ' + params.action };
    }

    Logger.log('✅ POST request handled');
    Logger.log('========== DOPOST COMPLETE ==========\n');

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log('❌ ERROR in doPost: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== DOPOST FAILED ==========\n');

    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Include HTML files in other HTML files
 * This is a special function used by HtmlService for modular HTML
 *
 * @param {string} filename - Name of the file to include
 * @returns {string} The file contents
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Get the current user's email
 * Used by UI to display logged-in user
 *
 * @returns {Object} Result with user email
 */
function getCurrentUserInfo() {
  try {
    const email = Session.getActiveUser().getEmail();
    return { success: true, data: { email: email } };
  } catch (error) {
    Logger.log('❌ ERROR getting user info: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Test function to verify system setup
 * Call this from Apps Script editor to verify everything is working
 *
 * @returns {Object} Test results
 */
function runSystemTests() {
  Logger.log('\n========== RUNNING SYSTEM TESTS ==========');

  const results = {
    sheetsAccess: false,
    configLoaded: false,
    utilsFunctional: false,
    allTests: false
  };

  try {
    // Test 1: Can we access sheets?
    Logger.log('🔍 Test 1: Checking sheet access...');
    const sheets = getSheets();
    if (sheets && Object.keys(sheets).length > 0) {
      results.sheetsAccess = true;
      Logger.log('✅ Sheet access OK. Found: ' + Object.keys(sheets).join(', '));
    } else {
      Logger.log('❌ No sheets found');
    }

    // Test 2: Config loaded?
    Logger.log('🔍 Test 2: Checking config...');
    if (typeof EMPLOYER_LIST !== 'undefined' && EMPLOYER_LIST.length > 0) {
      results.configLoaded = true;
      Logger.log('✅ Config loaded OK. Employers: ' + EMPLOYER_LIST.join(', '));
    } else {
      Logger.log('❌ Config not loaded');
    }

    // Test 3: Utils working?
    Logger.log('🔍 Test 3: Checking utils...');
    const testUuid = generateUUID();
    const testDate = formatDate(new Date());
    const testCurrency = formatCurrency(1234.56);
    if (testUuid && testDate && testCurrency) {
      results.utilsFunctional = true;
      Logger.log('✅ Utils working OK');
      Logger.log('  Sample UUID: ' + testUuid);
      Logger.log('  Sample Date: ' + testDate);
      Logger.log('  Sample Currency: ' + testCurrency);
    } else {
      Logger.log('❌ Utils not working');
    }

    // Overall result
    results.allTests = results.sheetsAccess && results.configLoaded && results.utilsFunctional;

    if (results.allTests) {
      Logger.log('✅ ALL TESTS PASSED');
    } else {
      Logger.log('⚠️ SOME TESTS FAILED');
    }

  } catch (error) {
    Logger.log('❌ ERROR during testing: ' + error.message);
    Logger.log('Stack: ' + error.stack);
  }

  Logger.log('========== SYSTEM TESTS COMPLETE ==========\n');

  return results;
}

/**
 * Quick test to verify the system is responsive
 *
 * @returns {string} Success message
 */
function ping() {
  return 'pong';
}

/**
 * Initialize the system (run once after deployment)
 * This function sets up triggers and performs initial checks
 *
 * @returns {Object} Initialization result
 */
function initializeSystem() {
  try {
    Logger.log('\n========== INITIALIZING SYSTEM ==========');

    // Run system tests first
    const testResults = runSystemTests();

    if (!testResults.allTests) {
      throw new Error('System tests failed. Please check configuration.');
    }

    // Install onChange trigger for auto-sync
    Logger.log('ℹ️ Installing onChange trigger...');
    installOnChangeTrigger();
    Logger.log('✅ Trigger installed');

    Logger.log('✅ System initialization complete');
    Logger.log('========== INITIALIZATION COMPLETE ==========\n');

    return {
      success: true,
      message: 'System initialized successfully',
      testResults: testResults
    };

  } catch (error) {
    Logger.log('❌ ERROR during initialization: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== INITIALIZATION FAILED ==========\n');

    return { success: false, error: error.message };
  }
}
