/**
 * Code.gs - Main entry point for HR System Web App
 * 
 * Handles web app initialization and HTML serving
 */

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
    features: {
      employees: true,
      leave: false,  // Phase 2
      loans: false,  // Phase 2
      timesheets: false,  // Phase 3
      payroll: false,  // Phase 4
      reports: false  // Phase 5
    }
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
