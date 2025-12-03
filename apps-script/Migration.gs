/**
 * Migration.gs - A file for one-time data migration scripts.
 */

/**
 * Migrates the UserConfig sheet from plain text passwords to hashed passwords.
 * This is a one-time operation.
 */
function migrateUsersToHashedPasswords() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('UserConfig');
    if (!sheet) {
      return { success: false, error: 'UserConfig sheet not found' };
    }

    // 1. Backup the sheet
    var backupSheet = sheet.copyTo(ss);
    backupSheet.setName('UserConfig_Backup_' + new Date().getTime());

    // 2. Read the data
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var emailIndex = indexOf(headers, 'Email');
    var passwordIndex = indexOf(headers, 'Password');

    if (passwordIndex === -1) {
      return { success: false, error: 'Password column not found. Already migrated?' };
    }

    // 3. Add new columns if they don't exist
    var passwordHashIndex = indexOf(headers, 'PasswordHash');
    var passwordSaltIndex = indexOf(headers, 'PasswordSalt');

    if (passwordHashIndex === -1) {
      sheet.insertColumnAfter(passwordIndex + 1);
      sheet.getRange(1, passwordIndex + 2).setValue('PasswordHash');
      passwordHashIndex = passwordIndex + 1;
    }

    if (passwordSaltIndex === -1) {
      sheet.insertColumnAfter(passwordHashIndex + 1);
      sheet.getRange(1, passwordHashIndex + 2).setValue('PasswordSalt');
      passwordSaltIndex = passwordHashIndex + 1;
    }

    // 4. Hash passwords and update rows
    for (var i = 1; i < data.length; i++) {
      var password = data[i][passwordIndex];
      if (password) {
        var salt = generateSalt();
        var hash = hashPassword(password, salt);
        sheet.getRange(i + 1, passwordHashIndex + 1).setValue(hash);
        sheet.getRange(i + 1, passwordSaltIndex + 1).setValue(salt);
      }
    }

    // 5. Remove the old password column
    sheet.deleteColumn(passwordIndex + 1);

    return { success: true, message: 'Migration complete. Please verify the UserConfig sheet.' };
  } catch (error) {
    logError('Failed to migrate users', error);
    return { success: false, error: 'Failed to migrate users: ' + error.toString() };
  }
}
