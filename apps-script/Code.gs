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
 * @param {Object} e - Event parameter with query string
 * @returns {HtmlOutput} Login page or dashboard
 */
function doGet(e) {
  try {
    var timestamp = new Date().toISOString();
    logInfo('========================================');
    logInfo('doGet called at: ' + timestamp);
    logInfo('Deployment URL: ' + ScriptApp.getService().getUrl());

    // Get page parameter for routing
    var page = (e && e.parameter) ? e.parameter.page : null;
    logInfo('Page requested: ' + (page || 'default (login/dashboard)'));

    // Handle password reset pages (no authentication required)
    if (page === 'forgot') {
      logInfo('Showing forgot password page');
      return createForgotPasswordPage();
    }

    if (page === 'reset-sent') {
      logInfo('Showing reset email sent page');
      return createResetEmailSentPage();
    }

    if (page === 'reset') {
      var token = (e && e.parameter) ? e.parameter.token : null;
      if (!token) {
        logWarning('Reset page requested without token');
        return createLoginPage();
      }
      logInfo('Showing password reset page with token');
      return createResetPasswordPage(token);
    }

    if (page === 'reset-success') {
      logInfo('Showing password reset success page');
      return createResetSuccessPage();
    }

    // Check if user has a valid session (handle undefined e parameter)
    var sessionToken = (e && e.parameter) ? e.parameter.session : null;
    logInfo('Session token from URL: ' + (sessionToken ? 'Present (' + sessionToken.substring(0, 8) + '...)' : 'Missing'));

    var userEmail = null;
    if (sessionToken) {
      try {
        userEmail = getUserFromSession(sessionToken);
        logInfo('User from session: ' + (userEmail || 'None (session expired or invalid)'));
      } catch (sessionError) {
        logError('Error getting user from session', sessionError);
        // Clear session and show login
        sessionToken = null;
        userEmail = null;
      }
    }

    if (userEmail) {
      // Valid session - show dashboard
      logInfo('✅ Valid session found - Showing dashboard for: ' + userEmail);
      logInfo('========================================');
      try {
        return createDashboardPage(userEmail, sessionToken);
      } catch (dashError) {
        logError('Error creating dashboard', dashError);
        throw dashError;
      }
    }

    // No valid session - show login page
    logInfo('ℹ️  No valid session - Showing login page');
    logInfo('========================================');
    try {
      return createLoginPage();
    } catch (loginError) {
      logError('Error creating login page', loginError);
      throw loginError;
    }

  } catch (error) {
    logError('❌ doGet error', error);
    logError('Stack trace: ' + error.stack);
    return HtmlService.createHtmlOutput(
      '<h1>Error Loading Application</h1>' +
      '<p><strong>Error:</strong> ' + error.toString() + '</p>' +
      '<p><strong>Stack:</strong> ' + error.stack + '</p>' +
      '<p><a href="' + ScriptApp.getService().getUrl() + '">Try Again</a></p>'
    ).setTitle('Error');
  }
}

/**
 * Handle login form submission
 *
 * @param {string} email - User email
 * @param {string} password - User password
 * @returns {Object} Result with session token if successful
 */
function handleLogin(email, password) {
  var result = authenticateUser(email, password);

  if (result.success) {
    var sessionToken = createUserSession(result.user.email);

    // Return the dashboard HTML directly instead of just the session token
    // This avoids the redirect issue with Apps Script iframe URLs
    try {
      var template = HtmlService.createTemplateFromFile('Dashboard');
      template.userEmail = result.user.email;
      template.sessionToken = sessionToken;

      var dashboardHtml = template.evaluate().getContent();

      return {
        success: true,
        sessionToken: sessionToken,
        user: result.user,
        dashboardHtml: dashboardHtml
      };
    } catch (error) {
      logError('Error generating dashboard HTML in handleLogin', error);
      return {
        success: false,
        error: 'Login succeeded but dashboard generation failed: ' + error.toString()
      };
    }
  }

  return result;
}

/**
 * Handle logout
 *
 * @param {string} sessionToken - Session token to destroy
 * @returns {Object} Result
 */
function handleLogout(sessionToken) {
  destroySession(sessionToken);
  return {
    success: true,
    message: 'Logged out successfully'
  };
}

/**
 * Handle password reset request
 * Sends reset email to user if email exists
 *
 * @param {string} email - User's email address
 * @returns {Object} Result
 */
function handlePasswordResetRequest(email) {
  try {
    if (!email) {
      return { success: false, error: 'Email is required' };
    }

    // Get the current web app URL (without parameters)
    var webAppUrl = ScriptApp.getService().getUrl();

    // Send reset email
    var result = sendPasswordResetEmail(email, webAppUrl);

    return result;

  } catch (error) {
    logError('handlePasswordResetRequest error', error);
    return {
      success: false,
      error: 'Failed to process reset request: ' + error.toString()
    };
  }
}

/**
 * Handle password reset with token
 * Validates token and updates password
 *
 * @param {string} token - Reset token
 * @param {string} newPassword - New password
 * @returns {Object} Result
 */
function handlePasswordReset(token, newPassword) {
  try {
    if (!token || !newPassword) {
      return { success: false, error: 'Token and new password are required' };
    }

    // Validate password strength (optional but recommended)
    if (newPassword.length < 6) {
      return { success: false, error: 'Password must be at least 6 characters long' };
    }

    // Reset password
    var result = resetPasswordWithToken(token, newPassword);

    return result;

  } catch (error) {
    logError('handlePasswordReset error', error);
    return {
      success: false,
      error: 'Failed to reset password: ' + error.toString()
    };
  }
}

/**
 * Validate password reset token
 * Returns whether token is valid without consuming it
 *
 * @param {string} token - Reset token to validate
 * @returns {Object} Result with validation status
 */
function handleValidateResetToken(token) {
  try {
    var result = validatePasswordResetToken(token);
    return result;
  } catch (error) {
    logError('handleValidateResetToken error', error);
    return {
      success: false,
      error: 'Failed to validate token: ' + error.toString()
    };
  }
}

/**
 * Get current logged-in user email from session
 *
 * @param {string} sessionToken - Session token
 * @returns {string|null} User email or null
 */
function getSessionUser(sessionToken) {
  return getUserFromSession(sessionToken);
}

/**
 * Create login page HTML
 * @returns {HtmlOutput} Login page
 */
function createLoginPage() {
  try {
    logInfo('Creating login page...');

    var loginHtml = HtmlService.createHtmlOutputFromFile('Login')
      .setTitle('Login - SA HR Payroll System');

    logInfo('Login page created successfully');
    return loginHtml;
  } catch (error) {
    logError('createLoginPage error', error);
    // Return error page instead of throwing
    return HtmlService.createHtmlOutput(
      '<h1>Login Page Error</h1>' +
      '<p>Error: ' + error.toString() + '</p>' +
      '<pre>' + error.stack + '</pre>'
    );
  }
}

/**
 * Create dashboard page with session
 * @param {string} userEmail - Logged in user's email
 * @param {string} sessionToken - Session token
 * @returns {HtmlOutput} Dashboard page
 */
function createDashboardPage(userEmail, sessionToken) {
  try {
    logInfo('Creating dashboard page for: ' + userEmail);
    logInfo('Session token: ' + (sessionToken ? 'Present' : 'Missing'));

    var template = HtmlService.createTemplateFromFile('Dashboard');
    template.userEmail = userEmail;
    template.sessionToken = sessionToken;

    logInfo('Template created, evaluating...');
    var output = template.evaluate()
      .setTitle('SA HR Payroll System');

    logInfo('Dashboard page created successfully');
    return output;

  } catch (error) {
    logError('createDashboardPage error', error);
    // Return error page instead of throwing
    return HtmlService.createHtmlOutput(
      '<h1>Dashboard Error</h1>' +
      '<p>Error creating dashboard for: ' + userEmail + '</p>' +
      '<p>Error: ' + error.toString() + '</p>' +
      '<pre>' + error.stack + '</pre>'
    );
  }
}

/**
 * Create access denied page HTML
 * @param {string} userEmail - User's email address
 * @returns {HtmlOutput} Access denied page
 */
function createAccessDeniedPage(userEmail) {
  var deniedHtml =
    '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '  <meta charset="utf-8">' +
    '  <meta name="viewport" content="width=device-width, initial-scale=1">' +
    '  <title>Access Denied - SA HR Payroll System</title>' +
    '  <style>' +
    '    body {' +
    '      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;' +
    '      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);' +
    '      display: flex;' +
    '      align-items: center;' +
    '      justify-content: center;' +
    '      min-height: 100vh;' +
    '      margin: 0;' +
    '      padding: 20px;' +
    '    }' +
    '    .container {' +
    '      background: white;' +
    '      border-radius: 12px;' +
    '      box-shadow: 0 20px 60px rgba(0,0,0,0.3);' +
    '      padding: 40px;' +
    '      max-width: 500px;' +
    '      text-align: center;' +
    '    }' +
    '    .icon {' +
    '      font-size: 64px;' +
    '      margin-bottom: 20px;' +
    '    }' +
    '    h1 {' +
    '      color: #dc3545;' +
    '      margin: 0 0 20px 0;' +
    '      font-size: 28px;' +
    '    }' +
    '    p {' +
    '      color: #666;' +
    '      line-height: 1.6;' +
    '      margin: 10px 0;' +
    '    }' +
    '    .user-info {' +
    '      background: #f8f9fa;' +
    '      border-radius: 6px;' +
    '      padding: 15px;' +
    '      margin: 20px 0;' +
    '      font-family: monospace;' +
    '      color: #495057;' +
    '    }' +
    '    .help-text {' +
    '      font-size: 14px;' +
    '      color: #6c757d;' +
    '      margin-top: 30px;' +
    '    }' +
    '  </style>' +
    '</head>' +
    '<body>' +
    '  <div class="container">' +
    '    <div class="icon">🔒</div>' +
    '    <h1>Access Denied</h1>' +
    '    <p>You are not authorized to access this application.</p>' +
    '    <div class="user-info">Your email: ' + userEmail + '</div>' +
    '    <p class="help-text">' +
    '      If you believe you should have access, please contact your administrator ' +
    '      and provide them with your email address shown above.' +
    '    </p>' +
    '  </div>' +
    '</body>' +
    '</html>';

  return HtmlService.createHtmlOutput(deniedHtml)
    .setTitle('Access Denied');
}

/**
 * Create forgot password page
 * @returns {HtmlOutput} Forgot password page
 */
function createForgotPasswordPage() {
  try {
    logInfo('Creating forgot password page...');

    var template = HtmlService.createTemplateFromFile('ForgotPassword');
    var forgotHtml = template.evaluate()
      .setTitle('Forgot Password - SA HR Payroll System');

    logInfo('Forgot password page created successfully');
    return forgotHtml;
  } catch (error) {
    logError('createForgotPasswordPage error', error);
    return HtmlService.createHtmlOutput(
      '<h1>Page Error</h1>' +
      '<p>Error: ' + error.toString() + '</p>' +
      '<pre>' + error.stack + '</pre>'
    );
  }
}

/**
 * Create reset email sent confirmation page
 * @returns {HtmlOutput} Reset email sent page
 */
function createResetEmailSentPage() {
  try {
    logInfo('Creating reset email sent page...');

    var template = HtmlService.createTemplateFromFile('ResetEmailSent');
    var sentHtml = template.evaluate()
      .setTitle('Password Reset Email Sent - SA HR Payroll System');

    logInfo('Reset email sent page created successfully');
    return sentHtml;
  } catch (error) {
    logError('createResetEmailSentPage error', error);
    return HtmlService.createHtmlOutput(
      '<h1>Page Error</h1>' +
      '<p>Error: ' + error.toString() + '</p>' +
      '<pre>' + error.stack + '</pre>'
    );
  }
}

/**
 * Create password reset form page
 * @param {string} token - Reset token
 * @returns {HtmlOutput} Password reset page
 */
function createResetPasswordPage(token) {
  try {
    logInfo('Creating reset password page...');

    var template = HtmlService.createTemplateFromFile('ResetPassword');
    template.token = token;

    var resetHtml = template.evaluate()
      .setTitle('Reset Password - SA HR Payroll System');

    logInfo('Reset password page created successfully');
    return resetHtml;
  } catch (error) {
    logError('createResetPasswordPage error', error);
    return HtmlService.createHtmlOutput(
      '<h1>Page Error</h1>' +
      '<p>Error: ' + error.toString() + '</p>' +
      '<pre>' + error.stack + '</pre>'
    );
  }
}

/**
 * Create reset success page
 * @returns {HtmlOutput} Reset success page
 */
function createResetSuccessPage() {
  try {
    logInfo('Creating reset success page...');

    var template = HtmlService.createTemplateFromFile('ResetSuccess');
    var successHtml = template.evaluate()
      .setTitle('Password Reset Successful - SA HR Payroll System');

    logInfo('Reset success page created successfully');
    return successHtml;
  } catch (error) {
    logError('createResetSuccessPage error', error);
    return HtmlService.createHtmlOutput(
      '<h1>Page Error</h1>' +
      '<p>Error: ' + error.toString() + '</p>' +
      '<pre>' + error.stack + '</pre>'
    );
  }
}

/**
 * Create configuration error page HTML
 * @param {string} errorDetails - Optional error details
 * @returns {HtmlOutput} Configuration error page
 */
function createConfigurationErrorPage(errorDetails) {
  var configErrorHtml =
    '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '  <meta charset="utf-8">' +
    '  <meta name="viewport" content="width=device-width, initial-scale=1">' +
    '  <title>Configuration Error - SA HR Payroll System</title>' +
    '  <style>' +
    '    body {' +
    '      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;' +
    '      background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);' +
    '      display: flex;' +
    '      align-items: center;' +
    '      justify-content: center;' +
    '      min-height: 100vh;' +
    '      margin: 0;' +
    '      padding: 20px;' +
    '    }' +
    '    .container {' +
    '      background: white;' +
    '      border-radius: 12px;' +
    '      box-shadow: 0 20px 60px rgba(0,0,0,0.3);' +
    '      padding: 40px;' +
    '      max-width: 600px;' +
    '    }' +
    '    .icon {' +
    '      font-size: 64px;' +
    '      margin-bottom: 20px;' +
    '      text-align: center;' +
    '    }' +
    '    h1 {' +
    '      color: #dc3545;' +
    '      margin: 0 0 20px 0;' +
    '      font-size: 28px;' +
    '      text-align: center;' +
    '    }' +
    '    h2 {' +
    '      color: #495057;' +
    '      font-size: 18px;' +
    '      margin: 30px 0 15px 0;' +
    '    }' +
    '    p, li {' +
    '      color: #666;' +
    '      line-height: 1.6;' +
    '      margin: 10px 0;' +
    '    }' +
    '    .error-box {' +
    '      background: #fff3cd;' +
    '      border: 1px solid #ffc107;' +
    '      border-radius: 6px;' +
    '      padding: 15px;' +
    '      margin: 20px 0;' +
    '      font-size: 14px;' +
    '      color: #856404;' +
    '    }' +
    '    .steps {' +
    '      background: #f8f9fa;' +
    '      border-radius: 6px;' +
    '      padding: 20px;' +
    '      margin: 20px 0;' +
    '    }' +
    '    .steps ol {' +
    '      margin: 10px 0;' +
    '      padding-left: 25px;' +
    '    }' +
    '    .steps li {' +
    '      margin: 8px 0;' +
    '    }' +
    '    code {' +
    '      background: #e9ecef;' +
    '      padding: 2px 6px;' +
    '      border-radius: 3px;' +
    '      font-family: monospace;' +
    '      font-size: 13px;' +
    '    }' +
    '    .highlight {' +
    '      background: #fff3cd;' +
    '      padding: 2px 6px;' +
    '      border-radius: 3px;' +
    '      font-weight: bold;' +
    '    }' +
    '  </style>' +
    '</head>' +
    '<body>' +
    '  <div class="container">' +
    '    <div class="icon">⚙️</div>' +
    '    <h1>Web App Configuration Required</h1>' +
    '    <p>The web app deployment settings need to be configured to allow user identification.</p>' +
    (errorDetails ? '<div class="error-box"><strong>Error:</strong> ' + errorDetails + '</div>' : '') +
    '    <h2>Administrator: Fix This Issue</h2>' +
    '    <div class="steps">' +
    '      <p><strong>Follow these steps to fix the configuration:</strong></p>' +
    '      <ol>' +
    '        <li>Open the spreadsheet in Google Sheets</li>' +
    '        <li>Go to <code>Extensions → Apps Script</code></li>' +
    '        <li>Click <code>Deploy → Manage deployments</code></li>' +
    '        <li>Click the <strong>Edit</strong> icon (pencil) on the active deployment</li>' +
    '        <li>Configure these settings:' +
    '          <ul>' +
    '            <li><span class="highlight">Execute as:</span> Me (owner email)</li>' +
    '            <li><span class="highlight">Who has access:</span> Anyone with Google account</li>' +
    '          </ul>' +
    '        </li>' +
    '        <li>Click <strong>New version</strong></li>' +
    '        <li>Click <strong>Deploy</strong></li>' +
    '        <li><strong>IMPORTANT:</strong> In the spreadsheet, go to <code>HR System → User Access → Sync to Script Properties</code></li>' +
    '        <li>This stores the whitelist without requiring spreadsheet access for users</li>' +
    '        <li>Copy the new web app URL and share it with users</li>' +
    '      </ol>' +
    '    </div>' +
    '    <h2>Alternative: Share Spreadsheet with Users</h2>' +
    '    <p>If the error persists, you may need to share the spreadsheet with your users:</p>' +
    '    <div class="steps">' +
    '      <ol>' +
    '        <li>Open the spreadsheet</li>' +
    '        <li>Click <strong>Share</strong> button (top right)</li>' +
    '        <li>Add user emails or share with "Anyone in organization"</li>' +
    '        <li>Set permission to <strong>Viewer</strong> (they don\'t need Edit access)</li>' +
    '        <li>Users can now access the web app</li>' +
    '      </ol>' +
    '    </div>' +
    '    <h2>Users: Authorization Required</h2>' +
    '    <p>When you first access the app after redeployment, you\'ll see an authorization screen:</p>' +
    '    <div class="steps">' +
    '      <ol>' +
    '        <li>Click <strong>Review Permissions</strong></li>' +
    '        <li>Select your Google account</li>' +
    '        <li>Click <strong>Advanced</strong> (if you see a warning)</li>' +
    '        <li>Click <strong>Go to [App Name] (unsafe)</strong> - This is safe, it\'s your organization\'s app</li>' +
    '        <li>Review and click <strong>Allow</strong></li>' +
    '      </ol>' +
    '    </div>' +
    '  </div>' +
    '</body>' +
    '</html>';

  return HtmlService.createHtmlOutput(configErrorHtml)
    .setTitle('Configuration Error');
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
 * For login system, this is called with sessionToken from client
 *
 * @param {string} sessionToken - Optional session token from logged-in user
 * @returns {string} Email address of current user
 */
function getCurrentUser(sessionToken) {
  try {
    // If session token provided, use that
    if (sessionToken) {
      var email = getUserFromSession(sessionToken);
      if (email) {
        return email;
      }
    }

    // Fallback: Try to get from Google Session (for spreadsheet menu functions)
    try {
      var effectiveEmail = Session.getEffectiveUser().getEmail();
      if (effectiveEmail && effectiveEmail !== '' && effectiveEmail !== 'system') {
        return effectiveEmail;
      }
    } catch (e) {
      // Ignore
    }

    return 'Unknown User';
  } catch (error) {
    Logger.log('Error getting user email: ' + error.toString());
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
 * Check if user is authorized to access the system
 * Checks against the UserConfig sheet for whitelisted users
 *
 * @returns {boolean} True if user is authorized
 */
function isAuthorizedUser(sessionToken) {
  var userEmail = getCurrentUser(sessionToken);

  // Unknown users are not authorized
  if (userEmail === 'Unknown User' || userEmail === '' || userEmail === 'system') {
    return false;
  }

  // SAFETY: If whitelist is empty, allow the script owner (for initial setup)
  var authorizedUsers = getAuthorizedUsers();
  if (authorizedUsers.length === 0) {
    // Allow the owner to access when whitelist is empty
    try {
      var effectiveEmail = Session.getEffectiveUser().getEmail();
      if (effectiveEmail && effectiveEmail !== '' && effectiveEmail === userEmail) {
        Logger.log('⚠️ Whitelist is empty - granting access to owner: ' + effectiveEmail);
        return true;
      }
    } catch (error) {
      Logger.log('Error checking owner access: ' + error.toString());
    }
  }

  // Check against whitelist
  return isUserAuthorized(userEmail);
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
    isAuthorized: isAuthorizedUser(),
    version: '1.4.0-user-whitelist', // Version identifier to track deployment
    lastUpdate: '2025-12-01T00:00:00Z',
    features: {
      employees: true,
      leave: true,
      loans: true,
      timesheets: true,
      timesheetIntegration: true,  // NEW - Clock-in integration
      payroll: true,  // ENABLED - Fixed module loading
      reports: true,
      userWhitelist: true  // NEW - User whitelist access control
    }
  };
}

// ============================================================================
// TIMESHEET MODULE FUNCTIONS
// ============================================================================

/**
 * Show timesheet import interface
 * Opens sidebar for importing clock-in data
 */
function showTimesheetImport() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('TimesheetImport')
      .setTitle('Import Clock Data')
      .setWidth(800);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Show timesheet approval dashboard
 * Opens sidebar for reviewing and approving timesheets
 */
function showTimesheetApproval() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('TimesheetApproval')
      .setTitle('Timesheet Approval')
      .setWidth(1000);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Show timesheet settings interface
 * Opens sidebar for configuring time processing rules
 */
function showTimesheetSettings() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('TimesheetSettings')
      .setTitle('Timesheet Settings')
      .setWidth(900);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Show Timesheet Debugger
 * Opens a comprehensive debugging tool to diagnose authorization and server issues
 */
function showTimesheetDebugger() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('TimesheetDebugger')
      .setTitle('🔍 Timesheet Debugger')
      .setWidth(1200);
    SpreadsheetApp.getUi().showModelessDialog(html, 'Timesheet Debugger - Diagnostic Tool');
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error loading debugger: ' + error.toString());
  }
}

/**
 * Run timesheet sheet setup - creates missing sheets with proper headers
 */
function runTimesheetSheetSetup() {
  try {
    var ui = SpreadsheetApp.getUi();

    // Confirm with user
    var response = ui.alert(
      'Setup Timesheet Sheets',
      'This will create the following sheets if they don\'t exist:\n\n' +
      '• RAW_CLOCK_DATA\n' +
      '• CLOCK_IN_IMPORTS\n' +
      '• PendingTimesheets\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    // Run setup
    var result = setupAllTimesheetSheets();

    if (result.success) {
      var message = 'Setup complete!\n\n';
      result.results.forEach(function(item) {
        var status = item.result.success ? '✅' : '❌';
        message += status + ' ' + item.sheet + '\n';
      });
      ui.alert('Success', message, ui.ButtonSet.OK);
    } else {
      ui.alert('Error', 'Setup failed: ' + result.error, ui.ButtonSet.OK);
    }

  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Fix RAW_CLOCK_DATA missing headers
 */
function runFixRawClockDataHeaders() {
  try {
    var ui = SpreadsheetApp.getUi();

    // Confirm with user
    var response = ui.alert(
      'Fix RAW_CLOCK_DATA Headers',
      'This will add missing column headers to RAW_CLOCK_DATA sheet.\n\n' +
      'Missing headers:\n' +
      '• STATUS (column 11)\n' +
      '• CREATED_DATE (column 12)\n' +
      '• LOCKED_DATE (column 13)\n' +
      '• LOCKED_BY (column 14)\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    // Run fix
    var result = fixRawClockDataHeaders();

    if (result.success) {
      var message = 'Headers fixed!\n\n';
      message += 'Before: ' + result.before + ' columns\n';
      message += 'After: ' + result.after + ' columns\n\n';
      if (result.added && result.added.length > 0) {
        message += 'Added headers:\n• ' + result.added.join('\n• ');
      }
      ui.alert('Success', message, ui.ButtonSet.OK);
    } else {
      ui.alert('Error', 'Fix failed: ' + result.error, ui.ButtonSet.OK);
    }

  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Fix PUNCH_TIME column format in RAW_CLOCK_DATA
 */
function runFixRawClockDataPunchTimeFormat() {
  try {
    var ui = SpreadsheetApp.getUi();

    // Confirm with user
    var response = ui.alert(
      'Fix PUNCH_TIME Format',
      'This will format the PUNCH_TIME column to show date AND time.\n\n' +
      'Current: 14/11/2025 (date only)\n' +
      'After: 2025-11-14 07:18:56 (date + time)\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    // Run fix
    var result = fixRawClockDataPunchTimeFormat();

    if (result.success) {
      var message = 'PUNCH_TIME column formatted!\n\n';
      message += 'Formatted ' + result.rowsFormatted + ' rows\n';
      message += 'Time component is now visible';
      ui.alert('Success', message, ui.ButtonSet.OK);
    } else {
      ui.alert('Error', 'Fix failed: ' + result.error, ui.ButtonSet.OK);
    }

  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Fix PendingTimesheets wrong headers
 */
function runFixPendingTimesheetsHeaders() {
  try {
    var ui = SpreadsheetApp.getUi();

    // Confirm with user
    var response = ui.alert(
      'Fix PendingTimesheets Headers',
      'This will correct wrong column header names in PendingTimesheets sheet.\n\n' +
      'Common issues:\n' +
      '• EMPLOYEE_NAME → EMPLOYEE NAME (with space)\n' +
      '• WEEK_ENDING → WEEKENDING (no underscore)\n' +
      '• STANDARD_HOURS → HOURS\n' +
      '• Column order mismatches\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    // Run fix
    var result = fixPendingTimesheetsHeaders();

    if (result.success) {
      var message = 'Headers fixed!\n\n';
      message += 'Corrected ' + result.corrected + ' header(s)\n\n';
      if (result.mismatches && result.mismatches.length > 0) {
        message += 'Fixed columns:\n';
        result.mismatches.slice(0, 5).forEach(function(m) {
          message += '• Col ' + m.position + ': "' + m.current + '" → "' + m.expected + '"\n';
        });
        if (result.mismatches.length > 5) {
          message += '... and ' + (result.mismatches.length - 5) + ' more\n';
        }
      }
      ui.alert('Success', message, ui.ButtonSet.OK);
    } else {
      ui.alert('Error', 'Fix failed: ' + result.error, ui.ButtonSet.OK);
    }

  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Menu handler for checking PendingTimesheets headers
 * Displays current vs expected headers
 */
function runCheckPendingTimesheetsHeaders() {
  try {
    var ui = SpreadsheetApp.getUi();

    // Run diagnostic
    var result = checkPendingTimesheetsHeaders();

    if (result.success) {
      var message = 'Check execution log for full details.\n\n';
      message += 'Current columns: ' + result.currentHeaders.length + '\n';
      message += 'Expected columns: ' + result.expectedHeaders.length + '\n\n';

      var mismatches = 0;
      for (var i = 0; i < result.currentHeaders.length; i++) {
        if (result.currentHeaders[i] !== result.expectedHeaders[i]) {
          mismatches++;
        }
      }

      if (mismatches > 0) {
        message += '⚠️ Found ' + mismatches + ' header mismatch(es)\n\n';
        message += 'Check execution log (View → Logs) for details.';
        ui.alert('Headers Need Fixing', message, ui.ButtonSet.OK);
      } else {
        message += '✅ All headers match!\n\n';
        ui.alert('Headers Correct', message, ui.ButtonSet.OK);
      }
    } else {
      ui.alert('Error', 'Check failed: ' + result.error, ui.ButtonSet.OK);
    }

  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Menu handler for setting up UserConfig sheet
 */
function runSetupUserConfigSheet() {
  try {
    var ui = SpreadsheetApp.getUi();

    // Confirm with user
    var response = ui.alert(
      'Setup UserConfig Sheet',
      'This will create the UserConfig sheet for managing authorized users.\n\n' +
      'The sheet will have two columns:\n' +
      '• Name - User\'s full name\n' +
      '• Email - User\'s email address (must match their Google account)\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    // Run setup
    var result = setupUserConfigSheet();

    if (result.success) {
      ui.alert(
        'Success',
        result.message + '\n\n' +
        'Sheet: ' + result.sheetName + '\n\n' +
        'Next steps:\n' +
        '1. Add authorized users to the UserConfig sheet\n' +
        '2. Enter their Name and Email (must match their Google account)\n' +
        '3. Redeploy the web app for changes to take effect',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('Error', 'Setup failed: ' + result.error, ui.ButtonSet.OK);
    }

  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Menu handler for viewing authorized users
 */
function runViewAuthorizedUsers() {
  try {
    var ui = SpreadsheetApp.getUi();

    // Get authorized users
    var authorizedUsers = getAuthorizedUsers();

    // Check sync status
    var scriptProperties = PropertiesService.getScriptProperties();
    var lastSync = scriptProperties.getProperty('USER_WHITELIST_LAST_SYNC');
    var syncStatus = lastSync ? '\n\nLast synced: ' + new Date(lastSync).toLocaleString() : '\n\n⚠️ Not synced to Script Properties yet';

    var message;
    if (authorizedUsers.length === 0) {
      message = 'No authorized users found.\n\n' +
                'The UserConfig sheet may not exist or may be empty.\n\n' +
                'Use "Setup UserConfig Sheet" to create it.';
    } else {
      message = 'Authorized Users (' + authorizedUsers.length + '):\n\n';
      for (var i = 0; i < Math.min(authorizedUsers.length, 10); i++) {
        message += (i + 1) + '. ' + authorizedUsers[i] + '\n';
      }
      if (authorizedUsers.length > 10) {
        message += '... and ' + (authorizedUsers.length - 10) + ' more\n';
      }
      message += syncStatus;
      message += '\n\nCurrent user: ' + getCurrentUser() + '\n' +
                 'Is authorized: ' + (isAuthorizedUser() ? 'YES ✓' : 'NO ✗');
    }

    ui.alert('Authorized Users', message, ui.ButtonSet.OK);

  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Menu handler for syncing UserConfig to Script Properties
 */
function runSyncUserConfig() {
  try {
    var ui = SpreadsheetApp.getUi();

    // Confirm with user
    var response = ui.alert(
      'Sync User Whitelist',
      'This will copy the email addresses from UserConfig sheet to Script Properties.\n\n' +
      'This allows users to access the web app without needing spreadsheet access.\n\n' +
      '⚠️ IMPORTANT: Run this after adding/removing users from UserConfig sheet.\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    // Run sync
    var result = syncUserConfigToProperties();

    if (result.success) {
      var userList = result.users.slice(0, 10).join('\n');
      if (result.users.length > 10) {
        userList += '\n... and ' + (result.users.length - 10) + ' more';
      }

      ui.alert(
        'Success',
        result.message + '\n\n' +
        'Synced users:\n' + userList + '\n\n' +
        '✓ Users can now access the web app without spreadsheet permissions\n' +
        '✓ Deploy as "Execute as: Me" for best security',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('Error', 'Sync failed: ' + result.error, ui.ButtonSet.OK);
    }

  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Create custom menu on spreadsheet open
 * Adds HR System menu with all modules
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  // Create main menu
  var menu = ui.createMenu('HR System');

  // Add Timesheets submenu
  menu.addSubMenu(ui.createMenu('Timesheets')
    .addItem('Import Clock Data', 'showTimesheetImport')
    .addItem('Pending Approval', 'showTimesheetApproval')
    .addItem('📊 Timesheet Breakdown', 'showTimesheetBreakdown')
    .addSeparator()
    .addItem('Settings', 'showTimesheetSettings')
    .addSeparator()
    .addItem('⚙️ Setup Sheets', 'runTimesheetSheetSetup')
    .addItem('🔧 Fix RAW_CLOCK_DATA Headers', 'runFixRawClockDataHeaders')
    .addItem('🔧 Fix PUNCH_TIME Format', 'runFixRawClockDataPunchTimeFormat')
    .addItem('🔧 Fix PendingTimesheets Headers', 'runFixPendingTimesheetsHeaders')
    .addSeparator()
    .addItem('📋 Check PendingTimesheets Headers', 'runCheckPendingTimesheetsHeaders'));

  // Add User Access submenu
  menu.addSubMenu(ui.createMenu('User Access')
    .addItem('⚙️ Setup UserConfig Sheet', 'runSetupUserConfigSheet')
    .addItem('👥 View Authorized Users', 'runViewAuthorizedUsers')
    .addSeparator()
    .addItem('🔄 Sync to Script Properties', 'runSyncUserConfig'));

  // Add other submenus (if you want to add more later)
  // menu.addSubMenu(ui.createMenu('Reports')...);

  menu.addToUi();
}

/**
 * Get version info - used to verify deployment
 *
 * @returns {Object} Version information
 */
function getVersion() {
  return {
    version: '1.4.0-user-whitelist',
    lastUpdate: '2025-12-01T00:00:00Z',
    hasFixedSheetMapping: true,
    hasLoggingFunctions: true,
    hasFormatResponse: true,
    hasPayrollModule: true,
    hasTimesheetIntegration: true,
    hasUserWhitelist: true,
    payrollFeatures: {
      createPayslip: true,
      calculatePayslipPreview: true,
      fieldMapping: true,
      moduleLoading: true,
      authorizationCheck: true
    },
    timesheetFeatures: {
      clockInImport: true,
      timesheetProcessor: true,
      configurableRules: true,
      duplicateDetection: true,
      employeeValidation: true,
      settingsInterface: true
    },
    securityFeatures: {
      userWhitelist: true,
      userConfigSheet: true,
      entryPointProtection: true,
      criticalOperationProtection: true
    },
    deploymentStatus: 'Payroll + Timesheet Integration + User Whitelist fully enabled'
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
