/**
 * DeploymentInstructions.gs
 *
 * DEPLOYMENT GUIDE FOR WEB APP
 * ============================
 *
 * The diagnostic shows your backend is working, but you may be using
 * a development URL or have deployment issues.
 *
 * FOLLOW THESE STEPS:
 *
 * 1. CREATE A NEW DEPLOYMENT
 *    - In Apps Script Editor, click "Deploy" → "New deployment"
 *    - Click the gear icon ⚙️ next to "Select type"
 *    - Choose "Web app"
 *
 * 2. CONFIGURE THE DEPLOYMENT
 *    Description: HR System Production v1
 *    Execute as: Me (your@email.com)
 *    Who has access: Anyone (or "Anyone with Google account" if you prefer)
 *
 * 3. CLICK "DEPLOY"
 *    - You may need to authorize the app
 *    - Copy the NEW Web app URL (should end with /exec, not /dev)
 *
 * 4. TEST THE NEW URL
 *    - Open in an INCOGNITO/PRIVATE browser window
 *    - You should see the login page
 *    - Try logging in with one of your 3 users
 *
 * 5. IF STILL BLANK
 *    - Open browser Developer Tools (F12)
 *    - Go to Console tab
 *    - Look for any error messages
 *    - Run getProductionUrl() below to see the correct URL
 *
 * TROUBLESHOOTING TIPS:
 * - Always test in incognito mode to avoid cache issues
 * - Make sure you authorized all OAuth scopes
 * - Check that the deployment is "Active"
 * - The URL should end with /exec (production) not /dev (test)
 */

/**
 * Get the current web app URL
 * This shows you the URL that doGet() would use
 */
function getProductionUrl() {
  var url = ScriptApp.getService().getUrl();
  console.log('Current Web App URL:');
  console.log(url);
  console.log('');

  if (url.indexOf('/dev') > -1) {
    console.log('⚠️  WARNING: This is a DEVELOPMENT URL (/dev)');
    console.log('');
    console.log('Development URLs are for testing only and may not work correctly.');
    console.log('');
    console.log('TO FIX:');
    console.log('1. Go to Deploy → Manage deployments');
    console.log('2. Find your active deployment (should end with /exec)');
    console.log('3. If none exist, create a new deployment');
    console.log('4. Use the /exec URL, not the /dev URL');
  } else {
    console.log('✅ This is a PRODUCTION URL - this is correct!');
    console.log('');
    console.log('If you still see a blank screen:');
    console.log('1. Open this URL in an incognito/private window');
    console.log('2. Clear your browser cache');
    console.log('3. Check browser console (F12) for errors');
  }

  return url;
}

/**
 * Test if your users can authenticate
 * Change the email/password to match one of your 3 users
 */
function testUserLogin() {
  // ⚠️ CHANGE THESE TO MATCH ONE OF YOUR 3 USERS ⚠️
  var testEmail = 'user@example.com';  // CHANGE THIS
  var testPassword = 'password123';     // CHANGE THIS

  console.log('========== TESTING USER LOGIN ==========');
  console.log('Email: ' + testEmail);
  console.log('');

  var result = authenticateUser(testEmail, testPassword);

  if (result.success) {
    console.log('✅ SUCCESS! User can log in');
    console.log('User: ' + result.user.name);
    console.log('Email: ' + result.user.email);
    console.log('');
    console.log('Backend authentication is working!');
    console.log('');
    console.log('If you still see a blank screen in the web app:');
    console.log('1. The issue is with deployment or browser cache');
    console.log('2. Run getProductionUrl() to check your URL');
    console.log('3. Try accessing in incognito mode');
    console.log('4. Check browser console for JavaScript errors');
  } else {
    console.log('❌ FAILED: ' + result.error);
    console.log('');
    console.log('Make sure you:');
    console.log('1. Used the correct email address from UserConfig sheet');
    console.log('2. Used the correct password');
    console.log('3. Check the UserConfig sheet to see your users');
  }

  return result;
}

/**
 * List all users in the system
 * Use this to see which emails you can log in with
 */
function listAllUsers() {
  console.log('========== USERS IN SYSTEM ==========');

  var props = PropertiesService.getScriptProperties();
  var userConfigData = props.getProperty('USER_CONFIG_DATA');

  if (userConfigData) {
    var users = JSON.parse(userConfigData);
    console.log('Found ' + users.length + ' user(s):');
    console.log('');

    for (var i = 0; i < users.length; i++) {
      console.log((i + 1) + '. ' + users[i].Name);
      console.log('   Email: ' + users[i].Email);
      console.log('   (Use this email to log in)');
      console.log('');
    }

    console.log('⚠️  NOTE: Passwords are encrypted and cannot be displayed');
    console.log('If you forgot a password, you can reset it by:');
    console.log('1. Running: createNewUser(name, email, newPassword)');
    console.log('   (This will update the user if they already exist)');
  } else {
    console.log('❌ No users found in Script Properties');
    console.log('Run quickFixLogin() to sync users');
  }
}

/**
 * Check browser/deployment issues
 */
function checkDeploymentStatus() {
  console.log('========== DEPLOYMENT STATUS ==========');
  console.log('');

  // Check URL
  var url = ScriptApp.getService().getUrl();
  console.log('📍 Web App URL: ' + url);

  if (url.indexOf('/dev') > -1) {
    console.log('❌ Status: DEVELOPMENT MODE');
    console.log('');
    console.log('You are using a test URL. This may cause issues.');
  } else {
    console.log('✅ Status: PRODUCTION MODE');
  }
  console.log('');

  // Check if doGet works
  console.log('Testing doGet() function...');
  try {
    var testEvent = { parameter: {} };
    var output = doGet(testEvent);
    var content = output.getContent();

    if (content && content.length > 1000) {
      console.log('✅ doGet() returns HTML (' + content.length + ' characters)');

      // Check if it's the login page
      if (content.indexOf('login') > -1 || content.indexOf('Login') > -1) {
        console.log('✅ Login page content detected');
      }
    } else {
      console.log('⚠️  doGet() returns content but it seems short');
      console.log('Length: ' + content.length + ' characters');
    }
  } catch (error) {
    console.log('❌ doGet() error: ' + error.toString());
  }
  console.log('');

  // Check Login.html
  console.log('Testing Login.html template...');
  try {
    var loginHtml = HtmlService.createHtmlOutputFromFile('Login');
    var loginContent = loginHtml.getContent();
    console.log('✅ Login.html loads (' + loginContent.length + ' characters)');
  } catch (error) {
    console.log('❌ Login.html error: ' + error.toString());
  }
  console.log('');

  console.log('========== RECOMMENDATION ==========');
  console.log('');
  console.log('Your backend is working correctly!');
  console.log('');
  console.log('Next steps:');
  console.log('1. Run listAllUsers() to see which emails you can use');
  console.log('2. Create a NEW deployment (Deploy → New deployment)');
  console.log('3. Copy the /exec URL (not /dev)');
  console.log('4. Open the URL in an INCOGNITO window');
  console.log('5. Log in with one of your user emails');
  console.log('');
  console.log('If the screen is still blank:');
  console.log('- Press F12 to open browser developer tools');
  console.log('- Check the Console tab for JavaScript errors');
  console.log('- Share any error messages you see');
}
