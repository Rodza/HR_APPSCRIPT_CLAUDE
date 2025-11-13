/**
 * UTILS.GS - Utility Functions
 * HR Payroll System for SA Grinding Wheels & Scorpio Abrasives
 *
 * This file contains shared utility functions used throughout the system:
 * - Sheet access and locator functions
 * - Date formatting
 * - Currency formatting
 * - UUID generation
 * - Validation helpers
 * - Logging helpers
 * - Audit trail management
 */

/**
 * Get all sheets in the spreadsheet with flexible name matching
 * This function provides a centralized way to access sheets regardless of their exact names
 *
 * @returns {Object} Object containing sheet references
 * @property {Sheet} salary - MASTERSALARY sheet
 * @property {Sheet} loans - EmployeeLoans sheet
 * @property {Sheet} empdetails - EMPLOYEE DETAILS sheet
 * @property {Sheet} leave - LEAVE sheet
 * @property {Sheet} pending - PendingTimesheets sheet
 */
function getSheets() {
  Logger.log('📋 Scanning spreadsheet for sheets...');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  const sheetMap = {};

  allSheets.forEach(sheet => {
    const originalName = sheet.getName();
    const name = originalName.toLowerCase().replace(/\s+/g, '').replace(/-/g, '');

    // Match MASTERSALARY or salary records
    if (name === 'mastersalary' || name === 'salaryrecords' || name === 'payroll') {
      sheetMap.salary = sheet;
      Logger.log('  📊 Found salary sheet: ' + originalName);
    }
    // Match EmployeeLoans
    else if (name === 'employeeloans' || name === 'loantransactions' || name === 'loans') {
      sheetMap.loans = sheet;
      Logger.log('  📊 Found loans sheet: ' + originalName);
    }
    // Match EMPLOYEE DETAILS
    else if (name === 'empdetails' || name === 'employeedetails' || name === 'employees') {
      sheetMap.empdetails = sheet;
      Logger.log('  📊 Found employee details sheet: ' + originalName);
    }
    // Match LEAVE
    else if (name === 'leave' || name === 'leaverecords') {
      sheetMap.leave = sheet;
      Logger.log('  📊 Found leave sheet: ' + originalName);
    }
    // Match PendingTimesheets
    else if (name === 'pendingtimesheets' || name === 'timeapproval' || name === 'pending') {
      sheetMap.pending = sheet;
      Logger.log('  📊 Found pending timesheets sheet: ' + originalName);
    }
  });

  Logger.log('✅ Sheet mapping complete. Found: ' + Object.keys(sheetMap).join(', '));

  return sheetMap;
}

/**
 * Get the currently logged-in user's email
 *
 * @returns {string} User email address
 */
function getCurrentUser() {
  return Session.getActiveUser().getEmail();
}

/**
 * Format date for display
 * Example: "17 October 2025"
 *
 * @param {Date|string|number} dateValue - Date to format
 * @returns {string} Formatted date string
 */
function formatDate(dateValue) {
  try {
    const date = parseDate(dateValue);
    const day = date.getDate();
    const month = date.toLocaleString('en-US', { month: 'long' });
    const year = date.getFullYear();
    return `${day} ${month} ${year}`;
  } catch (error) {
    Logger.log('⚠️ Error formatting date: ' + error.message);
    return String(dateValue);
  }
}

/**
 * Format date for short display (tables, lists)
 * Example: "2025-10-17"
 *
 * @param {Date|string|number} dateValue - Date to format
 * @returns {string} Formatted date string
 */
function formatDateShort(dateValue) {
  try {
    const date = parseDate(dateValue);
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  } catch (error) {
    Logger.log('⚠️ Error formatting date: ' + error.message);
    return String(dateValue);
  }
}

/**
 * Format currency with South African Rand symbol and comma separators
 * Example: "R1,341.42"
 *
 * @param {number} amount - Amount to format
 * @returns {string} Formatted currency string
 */
function formatCurrency(amount) {
  if (amount === null || amount === undefined || amount === '') {
    return 'R0.00';
  }

  const num = parseFloat(amount);
  if (isNaN(num)) {
    return 'R0.00';
  }

  // Round to 2 decimal places
  const rounded = Math.round(num * 100) / 100;

  // Format with comma separators
  const formatted = rounded.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');

  return 'R' + formatted;
}

/**
 * Generate a unique UUID
 * Used for creating unique IDs for records
 *
 * @returns {string} UUID string (8 characters)
 */
function generateUUID() {
  return Utilities.getUuid().substring(0, 8);
}

/**
 * Generate a full UUID (36 characters)
 * Used for loan IDs and other records requiring longer unique identifiers
 *
 * @returns {string} Full UUID string
 */
function generateFullUUID() {
  return Utilities.getUuid();
}

/**
 * Validate South African ID Number
 * Must be 13 digits
 *
 * @param {string} idNumber - ID number to validate
 * @returns {boolean} True if valid
 */
function validateSAIdNumber(idNumber) {
  if (!idNumber) return false;

  const cleaned = String(idNumber).replace(/\s/g, '');
  return /^\d{13}$/.test(cleaned);
}

/**
 * Validate South African phone number
 * Accepts formats: 0821234567, +27821234567
 *
 * @param {string} phoneNumber - Phone number to validate
 * @returns {boolean} True if valid
 */
function validatePhoneNumber(phoneNumber) {
  if (!phoneNumber) return false;

  const cleaned = String(phoneNumber).replace(/[\s\-\(\)]/g, '');
  return /^(\+27|0)[0-9]{9}$/.test(cleaned);
}

/**
 * Validate email address
 *
 * @param {string} email - Email to validate
 * @returns {boolean} True if valid
 */
function validateEmail(email) {
  if (!email) return false;

  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

/**
 * Parse date from various formats
 *
 * @param {Date|string|number} dateValue - Date value to parse
 * @returns {Date} Parsed date object
 * @throws {Error} If date cannot be parsed
 */
function parseDate(dateValue) {
  if (dateValue instanceof Date) {
    return dateValue;
  }

  if (typeof dateValue === 'string') {
    const parsed = new Date(dateValue);
    if (!isNaN(parsed.getTime())) {
      return parsed;
    }
  }

  if (typeof dateValue === 'number') {
    const parsed = new Date(dateValue);
    if (!isNaN(parsed.getTime())) {
      return parsed;
    }
  }

  throw new Error('Invalid date format: ' + dateValue);
}

/**
 * Add audit fields to data object
 * Adds USER, TIMESTAMP on create
 * Adds MODIFIED_BY, LAST_MODIFIED on update
 *
 * @param {Object} data - Data object to enrich
 * @param {boolean} isCreate - True if creating new record, false if updating
 * @returns {Object} Enriched data object
 */
function addAuditFields(data, isCreate = true) {
  const user = getCurrentUser();
  const now = new Date();

  if (isCreate) {
    data.USER = user;
    data.TIMESTAMP = now;
  }

  data.MODIFIED_BY = user;
  data.LAST_MODIFIED = now;

  return data;
}

/**
 * Sanitize user input
 * Removes HTML tags and trims whitespace
 *
 * @param {string} input - Input string to sanitize
 * @returns {string} Sanitized string
 */
function sanitizeInput(input) {
  if (typeof input !== 'string') {
    return input;
  }

  // Remove HTML tags
  let sanitized = input.replace(/<[^>]*>/g, '');

  // Trim whitespace
  sanitized = sanitized.trim();

  return sanitized;
}

/**
 * Calculate number of days between two dates
 *
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @returns {number} Number of days (inclusive)
 */
function calculateDaysBetween(startDate, endDate) {
  const start = parseDate(startDate);
  const end = parseDate(endDate);

  const diffMs = end.getTime() - start.getTime();
  const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));

  return diffDays + 1; // Inclusive of both days
}

/**
 * Find column index by header name
 *
 * @param {Array} headers - Array of header strings
 * @param {string} columnName - Name of column to find
 * @returns {number} Column index (0-based) or -1 if not found
 */
function findColumnIndex(headers, columnName) {
  const normalizedName = columnName.toLowerCase().trim();

  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i]).toLowerCase().trim();
    if (header === normalizedName) {
      return i;
    }
  }

  return -1;
}

/**
 * Get column value from row by column name
 *
 * @param {Array} row - Data row array
 * @param {Array} headers - Header row array
 * @param {string} columnName - Column name to get
 * @returns {*} Column value or null if not found
 */
function getColumnValue(row, headers, columnName) {
  const index = findColumnIndex(headers, columnName);
  return index >= 0 ? row[index] : null;
}

/**
 * Convert row array to object using headers
 *
 * @param {Array} row - Data row array
 * @param {Array} headers - Header row array
 * @returns {Object} Object with headers as keys
 */
function rowToObject(row, headers) {
  const obj = {};

  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    obj[header] = row[i];
  }

  return obj;
}

/**
 * Convert object to row array using headers
 *
 * @param {Object} obj - Object with data
 * @param {Array} headers - Header row array
 * @returns {Array} Row array matching header order
 */
function objectToRow(obj, headers) {
  const row = [];

  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    row.push(obj[header] !== undefined ? obj[header] : '');
  }

  return row;
}

/**
 * Round number to specified decimal places
 *
 * @param {number} value - Value to round
 * @param {number} decimals - Number of decimal places (default: 2)
 * @returns {number} Rounded value
 */
function roundTo(value, decimals = 2) {
  const multiplier = Math.pow(10, decimals);
  return Math.round(value * multiplier) / multiplier;
}

/**
 * Check if value is empty (null, undefined, empty string)
 *
 * @param {*} value - Value to check
 * @returns {boolean} True if empty
 */
function isEmpty(value) {
  return value === null || value === undefined || value === '';
}

/**
 * Get week ending date (Friday) from any date in the week
 *
 * @param {Date|string} dateValue - Date in the week
 * @returns {Date} Friday of that week
 */
function getWeekEnding(dateValue) {
  const date = parseDate(dateValue);
  const dayOfWeek = date.getDay(); // 0 = Sunday, 5 = Friday

  // Calculate days until Friday
  let daysUntilFriday;
  if (dayOfWeek <= 5) {
    daysUntilFriday = 5 - dayOfWeek;
  } else {
    daysUntilFriday = 7 - dayOfWeek + 5;
  }

  const friday = new Date(date);
  friday.setDate(date.getDate() + daysUntilFriday);

  return friday;
}

/**
 * Lock script execution to prevent concurrent modifications
 *
 * @param {Function} callback - Function to execute with lock
 * @param {number} timeout - Timeout in milliseconds (default: 30000)
 * @returns {*} Result of callback function
 */
function withLock(callback, timeout = 30000) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(timeout);
    const result = callback();
    return result;
  } catch (error) {
    Logger.log('❌ Lock timeout or error: ' + error.message);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Test all utility functions
 * Call this to verify utilities are working
 */
function testUtilities() {
  Logger.log('\n========== TESTING UTILITIES ==========');

  try {
    // Test UUID generation
    const uuid = generateUUID();
    Logger.log('✅ UUID generated: ' + uuid);

    // Test date formatting
    const today = new Date();
    const formatted = formatDate(today);
    const shortFormatted = formatDateShort(today);
    Logger.log('✅ Date formatted: ' + formatted);
    Logger.log('✅ Date short: ' + shortFormatted);

    // Test currency formatting
    const currency = formatCurrency(1234.567);
    Logger.log('✅ Currency formatted: ' + currency);

    // Test ID validation
    const validId = validateSAIdNumber('9401015800081');
    const invalidId = validateSAIdNumber('123');
    Logger.log('✅ ID validation: ' + validId + ' (should be true)');
    Logger.log('✅ ID validation: ' + invalidId + ' (should be false)');

    // Test phone validation
    const validPhone = validatePhoneNumber('0821234567');
    const invalidPhone = validatePhoneNumber('123');
    Logger.log('✅ Phone validation: ' + validPhone + ' (should be true)');
    Logger.log('✅ Phone validation: ' + invalidPhone + ' (should be false)');

    // Test date calculation
    const days = calculateDaysBetween(new Date('2025-10-01'), new Date('2025-10-05'));
    Logger.log('✅ Days between: ' + days + ' (should be 5)');

    // Test rounding
    const rounded = roundTo(1.2345, 2);
    Logger.log('✅ Rounded: ' + rounded + ' (should be 1.23)');

    Logger.log('✅ ALL UTILITY TESTS PASSED');

  } catch (error) {
    Logger.log('❌ TEST FAILED: ' + error.message);
  }

  Logger.log('========== UTILITIES TEST COMPLETE ==========\n');
}
