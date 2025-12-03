/**
 * DiagnosticLogin.gs
 *
 * Run these functions from the Apps Script Editor to diagnose and fix login issues
 */

/**
 * STEP 1: Run this to diagnose the login system
 *
 * This will check:
 * - UserConfig sheet exists
 * - Users are synced to properties
 * - At least one user exists
 * - Web app deployment settings
 */
function diagnoseLoginIssues() {
  console.log('========== LOGIN SYSTEM DIAGNOSTIC ==========\n');

  var issues = [];
  var warnings = [];
  var passed = [];

  // Check 1: UserConfig sheet exists
  console.log('📋 CHECK 1: UserConfig Sheet');
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('UserConfig');

    if (sheet) {
      var data = sheet.getDataRange().getValues();
      var rowCount = data.length - 1; // Exclude header

      passed.push('✅ UserConfig sheet exists with ' + rowCount + ' user(s)');
      console.log('   ✅ Sheet exists with ' + rowCount + ' user(s)\n');

      if (rowCount === 0) {
        warnings.push('⚠️  No users found in UserConfig sheet');
        console.log('   ⚠️  WARNING: No users found in sheet\n');
      } else {
        // Show first 3 users (email only)
        console.log('   Users found:');
        for (var i = 1; i < Math.min(4, data.length); i++) {
          console.log('   - ' + data[i][1]); // Email column
        }
        console.log('');
      }
    } else {
      issues.push('❌ UserConfig sheet not found');
      console.log('   ❌ FAILED: UserConfig sheet does not exist\n');
    }
  } catch (error) {
    issues.push('❌ Error checking UserConfig sheet: ' + error.toString());
    console.log('   ❌ ERROR: ' + error.toString() + '\n');
  }

  // Check 2: Script Properties
  console.log('🔧 CHECK 2: Script Properties');
  try {
    var props = PropertiesService.getScriptProperties();
    var userConfigData = props.getProperty('USER_CONFIG_DATA');

    if (userConfigData) {
      var users = JSON.parse(userConfigData);
      passed.push('✅ USER_CONFIG_DATA synced with ' + users.length + ' user(s)');
      console.log('   ✅ USER_CONFIG_DATA property exists');
      console.log('   ✅ Contains ' + users.length + ' user(s)');

      // Show emails
      console.log('   Synced users:');
      for (var i = 0; i < Math.min(3, users.length); i++) {
        console.log('   - ' + users[i].Email);
      }
      console.log('');

      var lastSync = props.getProperty('USER_CONFIG_LAST_SYNC');
      if (lastSync) {
        console.log('   Last sync: ' + lastSync + '\n');
      }
    } else {
      issues.push('❌ USER_CONFIG_DATA not synced to properties');
      console.log('   ❌ FAILED: USER_CONFIG_DATA property not found\n');
    }
  } catch (error) {
    issues.push('❌ Error checking properties: ' + error.toString());
    console.log('   ❌ ERROR: ' + error.toString() + '\n');
  }

  // Check 3: Web App Deployment
  console.log('🌐 CHECK 3: Web App Deployment');
  try {
    var url = ScriptApp.getService().getUrl();
    if (url) {
      passed.push('✅ Web app is deployed');
      console.log('   ✅ Web app is deployed');
      console.log('   📍 URL: ' + url + '\n');
    } else {
      warnings.push('⚠️  Web app may not be deployed');
      console.log('   ⚠️  WARNING: Could not get web app URL\n');
    }
  } catch (error) {
    warnings.push('⚠️  Error checking deployment: ' + error.toString());
    console.log('   ⚠️  WARNING: ' + error.toString() + '\n');
  }

  // Check 4: Test Login HTML
  console.log('📄 CHECK 4: Login HTML Template');
  try {
    var loginHtml = HtmlService.createHtmlOutputFromFile('Login');
    if (loginHtml) {
      passed.push('✅ Login.html template loads correctly');
      console.log('   ✅ Login.html can be loaded\n');
    }
  } catch (error) {
    issues.push('❌ Login.html failed to load: ' + error.toString());
    console.log('   ❌ FAILED: Login.html error - ' + error.toString() + '\n');
  }

  // Check 5: Test doGet()
  console.log('🚀 CHECK 5: doGet() Function');
  try {
    var testEvent = { parameter: {} };
    var output = doGet(testEvent);

    if (output && output.getContent) {
      var content = output.getContent();
      if (content && content.length > 0) {
        passed.push('✅ doGet() returns valid HTML output');
        console.log('   ✅ doGet() returns valid HTML (' + content.length + ' characters)\n');
      } else {
        issues.push('❌ doGet() returns empty content');
        console.log('   ❌ FAILED: doGet() returns empty content\n');
      }
    } else {
      issues.push('❌ doGet() does not return valid output');
      console.log('   ❌ FAILED: doGet() return value is invalid\n');
    }
  } catch (error) {
    issues.push('❌ doGet() throws error: ' + error.toString());
    console.log('   ❌ FAILED: doGet() error - ' + error.toString() + '\n');
  }

  // Summary
  console.log('========== DIAGNOSTIC SUMMARY ==========\n');

  if (passed.length > 0) {
    console.log('✅ PASSED (' + passed.length + '):');
    passed.forEach(function(msg) {
      console.log('   ' + msg);
    });
    console.log('');
  }

  if (warnings.length > 0) {
    console.log('⚠️  WARNINGS (' + warnings.length + '):');
    warnings.forEach(function(msg) {
      console.log('   ' + msg);
    });
    console.log('');
  }

  if (issues.length > 0) {
    console.log('❌ ISSUES FOUND (' + issues.length + '):');
    issues.forEach(function(msg) {
      console.log('   ' + msg);
    });
    console.log('');
    console.log('👉 Run fixLoginIssues() to automatically fix common issues\n');
  } else {
    console.log('🎉 No critical issues found!\n');
    console.log('If you\'re still seeing a blank screen:\n');
    console.log('1. Try accessing the web app in an incognito/private window');
    console.log('2. Clear your browser cache');
    console.log('3. Check browser console for JavaScript errors (F12)\n');
  }

  return {
    passed: passed,
    warnings: warnings,
    issues: issues,
    hasIssues: issues.length > 0
  };
}

/**
 * STEP 2: Run this to automatically fix common issues
 */
function fixLoginIssues() {
  console.log('========== FIXING LOGIN ISSUES ==========\n');

  var fixed = [];
  var errors = [];

  // Fix 1: Create UserConfig sheet if missing
  console.log('🔧 FIX 1: Setting up UserConfig sheet');
  try {
    var result = setupUserConfigSheet();
    if (result.success) {
      fixed.push('✅ UserConfig sheet created/verified');
      console.log('   ✅ UserConfig sheet is ready\n');
    } else {
      errors.push('❌ Failed to setup UserConfig: ' + result.error);
      console.log('   ❌ ERROR: ' + result.error + '\n');
    }
  } catch (error) {
    errors.push('❌ Error setting up UserConfig: ' + error.toString());
    console.log('   ❌ ERROR: ' + error.toString() + '\n');
  }

  // Fix 2: Sync users to properties
  console.log('🔄 FIX 2: Syncing users to Script Properties');
  try {
    var result = syncUserConfigToProperties();
    if (result.success) {
      fixed.push('✅ User data synced to properties');
      console.log('   ✅ Successfully synced ' + result.count + ' user(s)\n');
    } else {
      errors.push('❌ Failed to sync users: ' + result.error);
      console.log('   ❌ ERROR: ' + result.error + '\n');
    }
  } catch (error) {
    errors.push('❌ Error syncing users: ' + error.toString());
    console.log('   ❌ ERROR: ' + error.toString() + '\n');
  }

  // Summary
  console.log('========== FIX SUMMARY ==========\n');

  if (fixed.length > 0) {
    console.log('✅ FIXED (' + fixed.length + '):');
    fixed.forEach(function(msg) {
      console.log('   ' + msg);
    });
    console.log('');
  }

  if (errors.length > 0) {
    console.log('❌ ERRORS (' + errors.length + '):');
    errors.forEach(function(msg) {
      console.log('   ' + msg);
    });
    console.log('');
  }

  // Check if users exist
  var props = PropertiesService.getScriptProperties();
  var userConfigData = props.getProperty('USER_CONFIG_DATA');

  if (userConfigData) {
    var users = JSON.parse(userConfigData);
    if (users.length === 0) {
      console.log('⚠️  WARNING: No users exist yet!');
      console.log('👉 Run createTestUser() to create a user\n');
      return { success: false, needsUsers: true };
    } else {
      console.log('🎉 Login system is ready with ' + users.length + ' user(s)!\n');
      console.log('📍 Web App URL: ' + ScriptApp.getService().getUrl() + '\n');
      return { success: true, userCount: users.length };
    }
  } else {
    console.log('❌ User sync failed - please check the errors above\n');
    return { success: false };
  }
}

/**
 * STEP 3: Create a test user for testing login
 *
 * Change the default values below before running
 */
function createTestUser() {
  // ⚠️ CHANGE THESE VALUES BEFORE RUNNING ⚠️
  var userName = 'Test User';
  var userEmail = 'test@example.com';  // CHANGE THIS
  var userPassword = 'testpass123';    // CHANGE THIS

  console.log('========== CREATING TEST USER ==========\n');
  console.log('Name: ' + userName);
  console.log('Email: ' + userEmail);
  console.log('Password: ' + userPassword + '\n');

  var result = createNewUser(userName, userEmail, userPassword);

  if (result.success) {
    console.log('✅ SUCCESS! User created:\n');
    console.log('   Name: ' + result.user.name);
    console.log('   Email: ' + result.user.email);
    console.log('   Password: ' + userPassword + '\n');
    console.log('📍 Test the login at: ' + ScriptApp.getService().getUrl() + '\n');
    return result;
  } else {
    console.log('❌ FAILED: ' + result.error + '\n');
    return result;
  }
}

/**
 * STEP 4: Test the complete login flow
 */
function testLoginFlow() {
  console.log('========== TESTING LOGIN FLOW ==========\n');

  var testEmail = 'test@example.com';  // Change to your test user
  var testPassword = 'testpass123';    // Change to your test password

  console.log('Testing login with:');
  console.log('Email: ' + testEmail);
  console.log('Password: ' + testPassword + '\n');

  // Step 1: Test authentication
  console.log('1️⃣  Testing authenticateUser()...');
  var authResult = authenticateUser(testEmail, testPassword);

  if (authResult.success) {
    console.log('   ✅ Authentication successful');
    console.log('   User: ' + authResult.user.name + ' (' + authResult.user.email + ')\n');

    // Step 2: Test session creation
    console.log('2️⃣  Testing createUserSession()...');
    var sessionToken = createUserSession(authResult.user.email);
    console.log('   ✅ Session created');
    console.log('   Token: ' + sessionToken + '\n');

    // Step 3: Test session retrieval
    console.log('3️⃣  Testing getUserFromSession()...');
    var retrievedEmail = getUserFromSession(sessionToken);
    if (retrievedEmail === testEmail) {
      console.log('   ✅ Session retrieval successful');
      console.log('   Email: ' + retrievedEmail + '\n');
    } else {
      console.log('   ❌ Session retrieval failed');
      console.log('   Expected: ' + testEmail);
      console.log('   Got: ' + retrievedEmail + '\n');
      return { success: false, step: 'session_retrieval' };
    }

    // Step 4: Test handleLogin()
    console.log('4️⃣  Testing handleLogin()...');
    var loginResult = handleLogin(testEmail, testPassword);
    if (loginResult.success) {
      console.log('   ✅ handleLogin() successful');
      console.log('   Session Token: ' + loginResult.sessionToken + '\n');
    } else {
      console.log('   ❌ handleLogin() failed: ' + loginResult.error + '\n');
      return { success: false, step: 'handleLogin' };
    }

    // Step 5: Test session cleanup
    console.log('5️⃣  Testing destroySession()...');
    destroySession(sessionToken);
    var clearedEmail = getUserFromSession(sessionToken);
    if (!clearedEmail) {
      console.log('   ✅ Session destroyed successfully\n');
    } else {
      console.log('   ⚠️  Session may not have been destroyed\n');
    }

    console.log('🎉 ALL LOGIN TESTS PASSED!\n');
    console.log('📍 Try logging in at: ' + ScriptApp.getService().getUrl() + '\n');
    return { success: true };

  } else {
    console.log('   ❌ Authentication failed: ' + authResult.error + '\n');
    console.log('Possible reasons:');
    console.log('- User does not exist');
    console.log('- Wrong password');
    console.log('- Users not synced to properties\n');
    console.log('👉 Run fixLoginIssues() first, then createTestUser()\n');
    return { success: false, step: 'authentication', error: authResult.error };
  }
}

/**
 * QUICK FIX: Run this single function to set everything up
 */
function quickFixLogin() {
  console.log('========== QUICK FIX LOGIN ==========\n');
  console.log('This will:');
  console.log('1. Setup UserConfig sheet');
  console.log('2. Sync users to properties');
  console.log('3. Show you the next steps\n');

  var fixResult = fixLoginIssues();

  if (fixResult.needsUsers) {
    console.log('\n⚠️  NEXT STEP: Create a user');
    console.log('\nOption 1 - Create via script:');
    console.log('1. Edit createTestUser() function (change email/password)');
    console.log('2. Run createTestUser()\n');
    console.log('Option 2 - Create via sheet:');
    console.log('1. Open the UserConfig sheet in your spreadsheet');
    console.log('2. Add a row: Name | Email | (leave PasswordHash & Salt empty)');
    console.log('3. Run createNewUser("Name", "email@example.com", "password")\n');
    console.log('After creating a user, run testLoginFlow() to verify everything works.\n');
  }

  return fixResult;
}
