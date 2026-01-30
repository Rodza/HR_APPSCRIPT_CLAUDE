/**
 * Utils.gs - Utility Functions for HR System
 * 
 * Contains helper functions for validation, formatting, logging, etc.
 * Phase 1 focus: Core utilities for employee management
 */

// ============================================================================
// LOGGING UTILITIES
// ============================================================================

/**
 * Log function start
 * @param {string} functionName - Name of the function
 * @param {Object} params - Parameters passed to function
 */
function logFunctionStart(functionName, params) {
  console.log('→ ' + functionName + ' started', params || {});
}

/**
 * Log function end
 * @param {string} functionName - Name of the function  
 * @param {Object} result - Result data
 */
function logFunctionEnd(functionName, result) {
  console.log('← ' + functionName + ' completed', result || {});
}

/**
 * Log success message
 * @param {string} message - Success message
 * @param {Object} data - Additional data
 */
function logSuccess(message, data) {
  console.log('✓ ' + message, data || {});
}

/**
 * Log warning message
 * @param {string} message - Warning message
 * @param {Object} data - Additional data
 */
function logWarning(message, data) {
  console.warn('⚠ ' + message, data || {});
}

/**
 * Log error message
 * @param {string} message - Error message
 * @param {Error} error - Error object
 */
function logError(message, error) {
  console.error('✗ ' + message, error ? error.toString() : '');
  
  // Also log to Logger for Apps Script logging
  if (typeof Logger !== 'undefined') {
    Logger.log('ERROR: ' + message + ' - ' + (error ? error.toString() : ''));
  }
}

/**
 * Log info message
 * @param {string} message - Info message
 * @param {Object} data - Additional data
 */
function logInfo(message, data) {
  console.info('ℹ ' + message, data || {});
}

// ============================================================================
// VALIDATION UTILITIES
// ============================================================================

/**
 * Check if value is empty
 * Enhanced to handle all data types safely
 * 
 * @param {any} value - Value to check
 * @returns {boolean} True if empty
 */
function isEmpty(value) {
  // Handle null and undefined
  if (value === null || value === undefined) {
    return true;
  }
  
  // Handle strings
  if (typeof value === 'string') {
    return value.trim() === '';
  }
  
  // Handle numbers (0 is not empty)
  if (typeof value === 'number') {
    return false;
  }
  
  // Handle booleans
  if (typeof value === 'boolean') {
    return false;
  }
  
  // Handle dates
  if (value instanceof Date) {
    return false;
  }
  
  // Handle arrays
  if (Array.isArray(value)) {
    return value.length === 0;
  }
  
  // Handle objects
  if (typeof value === 'object') {
    return Object.keys(value).length === 0;
  }
  
  // Default to not empty for other types
  return false;
}

/**
 * Validate South African ID Number
 * 
 * @param {string} idNumber - 13-digit SA ID number
 * @returns {boolean} True if valid
 */
function isValidSAIdNumber(idNumber) {
  // Convert to string and remove spaces
  idNumber = String(idNumber || '').replace(/\s/g, '');
  
  // Must be 13 digits
  if (!/^\d{13}$/.test(idNumber)) {
    return false;
  }
  
  // Extract date components
  var year = parseInt(idNumber.substring(0, 2));
  var month = parseInt(idNumber.substring(2, 4));
  var day = parseInt(idNumber.substring(4, 6));
  
  // Determine century (00-30 = 2000s, 31-99 = 1900s)
  var fullYear = year <= 30 ? 2000 + year : 1900 + year;
  
  // Validate date
  if (month < 1 || month > 12 || day < 1 || day > 31) {
    return false;
  }
  
  // Check if date is valid
  var testDate = new Date(fullYear, month - 1, day);
  if (testDate.getFullYear() !== fullYear || 
      testDate.getMonth() !== month - 1 || 
      testDate.getDate() !== day) {
    return false;
  }
  
  // Validate gender digits (4 digits, 0000-4999 = female, 5000-9999 = male)
  var genderDigits = parseInt(idNumber.substring(6, 10));
  if (genderDigits < 0 || genderDigits > 9999) {
    return false;
  }
  
  // Validate citizenship (0 = SA citizen, 1 = permanent resident)
  var citizenship = parseInt(idNumber.charAt(10));
  if (citizenship !== 0 && citizenship !== 1) {
    return false;
  }
  
  // Digit 11 is usually 8 or 9 (obsolete)
  
  // Validate checksum (Luhn algorithm)
  var sum = 0;
  var alternate = false;
  
  for (var i = idNumber.length - 1; i >= 0; i--) {
    var digit = parseInt(idNumber.charAt(i));
    
    if (alternate) {
      digit *= 2;
      if (digit > 9) {
        digit = (digit % 10) + 1;
      }
    }
    
    sum += digit;
    alternate = !alternate;
  }
  
  return (sum % 10) === 0;
}

/**
 * Validate phone number (South African format)
 * Accepts: 0123456789, 012 345 6789, +27123456789, etc.
 * 
 * @param {string} phone - Phone number to validate
 * @returns {boolean} True if valid
 */
function isValidPhoneNumber(phone) {
  // Convert to string and remove spaces, dashes, parentheses
  phone = String(phone || '').replace(/[\s\-\(\)]/g, '');
  
  // Check for South African formats
  // Local: 0XXXXXXXXX (10 digits starting with 0)
  // International: +27XXXXXXXXX or 0027XXXXXXXXX
  
  if (phone.startsWith('+27')) {
    phone = '0' + phone.substring(3);
  } else if (phone.startsWith('0027')) {
    phone = '0' + phone.substring(4);
  } else if (phone.startsWith('27')) {
    phone = '0' + phone.substring(2);
  }
  
  // Should now be 10 digits starting with 0
  return /^0\d{9}$/.test(phone);
}

/**
 * Format phone number to standard format
 * 
 * @param {string} phone - Phone number to format
 * @returns {string} Formatted phone number
 */
function formatPhoneNumber(phone) {
  // Clean and validate
  phone = String(phone || '').replace(/[\s\-\(\)]/g, '');
  
  if (!isValidPhoneNumber(phone)) {
    return phone; // Return as-is if invalid
  }
  
  // Convert to local format
  if (phone.startsWith('+27')) {
    phone = '0' + phone.substring(3);
  } else if (phone.startsWith('0027')) {
    phone = '0' + phone.substring(4);
  } else if (phone.startsWith('27')) {
    phone = '0' + phone.substring(2);
  }
  
  // Format as 0XX XXX XXXX
  return phone.substring(0, 3) + ' ' + 
         phone.substring(3, 6) + ' ' + 
         phone.substring(6);
}

// ============================================================================
// SHEET UTILITIES
// ============================================================================

/**
 * Cache for getSheets() to avoid repeated expensive spreadsheet API calls
 * This cache persists for the duration of a single script execution
 */
var _SHEETS_CACHE = null;

/**
 * Get all sheets with proper error handling - VERSION 2 (Cache-busting rename)
 *
 * This is the actual implementation. The original getSheets() delegates to this.
 *
 * @returns {Object} Object with sheet references
 */
function getSheets_v2() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = {};
    var allSheets = ss.getSheets();

    // Map sheets by name (case-insensitive)
    for (var i = 0; i < allSheets.length; i++) {
      var sheet = allSheets[i];
      var sheetName = sheet.getName().toLowerCase().replace(/\s+/g, ''); // Remove spaces

      // Map to expected keys - ORDER MATTERS! Most specific first
      if (sheetName.indexOf('clockinimports') >= 0 || sheetName.indexOf('clock_in_imports') >= 0) {
        sheets.clockImports = sheet;
      }
      else if (sheetName.indexOf('rawclockdata') >= 0 || sheetName.indexOf('raw_clock_data') >= 0 || sheetName.indexOf('clockdata') >= 0) {
        sheets.rawClockData = sheet;
      }
      else if (sheetName.indexOf('loan') >= 0) {
        sheets.loans = sheet;
      }
      else if (sheetName.indexOf('employeedetails') >= 0 || sheetName === 'empdetails') {
        sheets.empdetails = sheet;
      }
      else if (sheetName.indexOf('salary') >= 0 || sheetName.indexOf('payroll') >= 0) {
        sheets.salary = sheet;
      }
      else if (sheetName.indexOf('leave') >= 0) {
        sheets.leave = sheet;
      }
      else if ((sheetName.indexOf('pending') >= 0 && sheetName.indexOf('timesheet') >= 0) &&
               sheetName.indexOf('recon') < 0 && sheetName.indexOf('detail') < 0) {
        if (!sheets.pending || sheetName === 'pending_timesheets') {
          sheets.pending = sheet;
        }
      }
      else if (sheetName.indexOf('userconfig') >= 0 || sheetName.indexOf('authorizedusers') >= 0) {
        sheets.userConfig = sheet;
      }
    }

    return sheets;

  } catch (error) {
    logError('Failed to get sheets', error);
    throw new Error('Failed to access spreadsheet: ' + error.toString());
  }
}

/**
 * Get all sheets with proper error handling (WRAPPER - calls _v2 with caching)
 *
 * This wrapper caches the result to avoid repeated expensive spreadsheet API calls.
 * The cache persists for the duration of a single script execution.
 *
 * @returns {Object} Object with sheet references
 */
function getSheets() {
  if (_SHEETS_CACHE === null) {
    _SHEETS_CACHE = getSheets_v2();
  }
  return _SHEETS_CACHE;
}

/**
 * Find column index by header name (case-insensitive)
 * 
 * @param {Array} headers - Array of header values
 * @param {string} columnName - Column name to find
 * @returns {number} Column index or -1 if not found
 */
function indexOf(headers, columnName) {
  if (!headers || !columnName) return -1;
  
  columnName = columnName.toLowerCase();
  
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).toLowerCase() === columnName) {
      return i;
    }
  }
  
  return -1;
}

/**
 * Build object from row data and headers
 * 
 * @param {Array} row - Row data
 * @param {Array} headers - Header row
 * @returns {Object} Object with header keys
 */
function buildObjectFromRow(row, headers) {
  var obj = {};

  for (var i = 0; i < headers.length; i++) {
    var header = String(headers[i]);
    var value = row[i];

    // Handle dates - validate that Date objects are valid
    if (value instanceof Date) {
      // Check if the Date is valid
      if (isNaN(value.getTime())) {
        // Invalid date - set to null
        obj[header] = null;
      } else {
        obj[header] = value;
      }
    }
    // Handle numbers
    else if (typeof value === 'number') {
      obj[header] = value;
    }
    // Handle strings
    else if (typeof value === 'string') {
      obj[header] = value;
    }
    // Handle booleans
    else if (typeof value === 'boolean') {
      obj[header] = value;
    }
    // Default to empty string for null/undefined
    else {
      obj[header] = '';
    }
  }

  return obj;
}

/**
 * Convert object to array row based on headers
 * Reverse operation of buildObjectFromRow
 *
 * @param {Object} obj - Object to convert to row
 * @param {Array} headers - Array of header names (column order)
 * @returns {Array} Row array in header order
 */
function objectToRow(obj, headers) {
  var row = [];

  for (var i = 0; i < headers.length; i++) {
    var header = String(headers[i]);
    var value = obj[header];

    // Use the value from the object if it exists, otherwise empty string
    if (value !== undefined && value !== null) {
      row.push(value);
    } else {
      row.push('');
    }
  }

  return row;
}

/**
 * Filter array by field value
 * 
 * @param {Array} array - Array to filter
 * @param {string} field - Field name
 * @param {string} value - Value to match
 * @returns {Array} Filtered array
 */
function filterByField(array, field, value) {
  return array.filter(function(item) {
    return item[field] === value;
  });
}

/**
 * Filter array by search term in field
 * 
 * @param {Array} array - Array to filter
 * @param {string} field - Field name
 * @param {string} search - Search term
 * @returns {Array} Filtered array
 */
function filterBySearch(array, field, search) {
  var searchLower = search.toLowerCase();
  
  return array.filter(function(item) {
    var fieldValue = String(item[field] || '').toLowerCase();
    return fieldValue.indexOf(searchLower) >= 0;
  });
}

// ============================================================================
// DATE UTILITIES
// ============================================================================

/**
 * Normalize a date to midnight (00:00:00) to strip time components
 * This ensures consistent date-only comparisons and storage
 *
 * @param {Date} date - Date object to normalize
 * @returns {Date} Date object set to midnight
 */
function normalizeToDateOnly(date) {
  if (!date || !(date instanceof Date)) {
    return date;
  }

  var normalized = new Date(date);
  normalized.setHours(0, 0, 0, 0);
  return normalized;
}

/**
 * Parse date from various input formats
 * Returns date normalized to midnight (00:00:00) for consistent date-only handling
 *
 * @param {Date|string|number} dateValue - Date in various formats
 * @returns {Date} Parsed Date object normalized to midnight
 * @throws {Error} If date format is invalid
 */
function parseDate(dateValue) {
  if (!dateValue) {
    throw new Error('Date value is required');
  }

  var parsedDate;

  // Already a Date object
  if (dateValue instanceof Date) {
    parsedDate = dateValue;
  }
  // String or number - convert to Date
  else if (typeof dateValue === 'string' || typeof dateValue === 'number') {
    parsedDate = new Date(dateValue);

    // Check if valid date
    if (isNaN(parsedDate.getTime())) {
      throw new Error('Invalid date format: ' + dateValue);
    }
  }
  else {
    throw new Error('Invalid date format: ' + typeof dateValue);
  }

  // Normalize to midnight to strip time components
  return normalizeToDateOnly(parsedDate);
}

/**
 * Calculate days between two dates
 *
 * @param {Date|string} startDate - Start date
 * @param {Date|string} endDate - End date
 * @returns {number} Number of days between dates (not inclusive)
 */
function calculateDaysBetween(startDate, endDate) {
  var start = parseDate(startDate);
  var end = parseDate(endDate);

  // Calculate difference in milliseconds
  var diffMs = end - start;

  // Convert to days
  var diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));

  return diffDays;
}

/**
 * Calculate working days between two dates (Monday to Friday only)
 * Excludes weekends (Saturday and Sunday)
 *
 * @param {Date|string} startDate - Start date (inclusive)
 * @param {Date|string} endDate - End date (inclusive)
 * @returns {number} Number of working days between dates (inclusive)
 */
function calculateWorkingDaysBetween(startDate, endDate) {
  var start = parseDate(startDate);
  var end = parseDate(endDate);

  // Ensure start is before end
  if (start > end) {
    return 0;
  }

  var workingDays = 0;
  var currentDate = new Date(start);

  // Loop through each day from start to end (inclusive)
  while (currentDate <= end) {
    var dayOfWeek = currentDate.getDay(); // 0 = Sunday, 6 = Saturday

    // Count only Monday (1) to Friday (5)
    if (dayOfWeek >= 1 && dayOfWeek <= 5) {
      workingDays++;
    }

    // Move to next day
    currentDate.setDate(currentDate.getDate() + 1);
  }

  return workingDays;
}

/**
 * Round number to specified decimal places
 *
 * @param {number} value - Number to round
 * @param {number} decimals - Number of decimal places (default: 2)
 * @returns {number} Rounded number
 */
function roundTo(value, decimals) {
  if (decimals === undefined) decimals = 2;

  var number = parseFloat(value);
  if (isNaN(number)) {
    return 0;
  }

  var multiplier = Math.pow(10, decimals);
  return Math.round(number * multiplier) / multiplier;
}

/**
 * Add audit fields to a data object
 *
 * @param {Object} data - The data object to add audit fields to
 * @param {boolean} isNew - True for new records (adds TIMESTAMP), false for updates (adds MODIFIED_BY, LAST_MODIFIED)
 */
function addAuditFields(data, isNew) {
  var currentUser = getCurrentUser();
  var currentTime = getCurrentTimestamp();

  if (isNew) {
    // For new records
    data.TIMESTAMP = currentTime;
    data.USER = currentUser;
    data.MODIFIED_BY = currentUser;
    data.LAST_MODIFIED = currentTime;
  } else {
    // For updates - preserve original TIMESTAMP and USER
    data.MODIFIED_BY = currentUser;
    data.LAST_MODIFIED = currentTime;
  }
}

/**
 * Format number as currency (South African Rand)
 *
 * @param {number} amount - Amount to format
 * @returns {string} Formatted currency string (e.g., "R1,234.56")
 */
function formatCurrency(amount) {
  if (amount === null || amount === undefined || amount === '') {
    return 'R0.00';
  }

  var number = parseFloat(amount);
  if (isNaN(number)) {
    return 'R0.00';
  }

  // Format with 2 decimal places and thousands separator
  var formatted = number.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return 'R' + formatted;
}

/**
 * Format date to YYYY-MM-DD
 *
 * @param {Date} date - Date to format
 * @returns {string} Formatted date string
 */
function formatDate(date) {
  if (!date) return '';

  if (!(date instanceof Date)) {
    date = new Date(date);
  }

  // Check if date is valid
  if (isNaN(date.getTime())) {
    return '';
  }

  var year = date.getFullYear();
  var month = ('0' + (date.getMonth() + 1)).slice(-2);
  var day = ('0' + date.getDate()).slice(-2);

  return year + '-' + month + '-' + day;
}

/**
 * Format date to DD-MM-YYYY (European/South African format)
 *
 * @param {Date} date - Date to format
 * @returns {string} Formatted date string
 */
function formatDateDDMMYYYY(date) {
  if (!date) return '';

  if (!(date instanceof Date)) {
    date = new Date(date);
  }

  // Check if date is valid
  if (isNaN(date.getTime())) {
    return '';
  }

  var year = date.getFullYear();
  var month = ('0' + (date.getMonth() + 1)).slice(-2);
  var day = ('0' + date.getDate()).slice(-2);

  return day + '-' + month + '-' + year;
}

/**
 * Get current timestamp
 *
 * @returns {string} Current timestamp in readable format
 */
function getCurrentTimestamp() {
  return new Date().toLocaleString('en-ZA', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit'
  });
}

/**
 * Find column index by header name (case-insensitive)
 * Alias for indexOf() for consistency with other modules
 *
 * @param {Array} headers - Array of header values
 * @param {string} columnName - Column name to find
 * @returns {number} Column index or -1 if not found
 */
function findColumnIndex(headers, columnName) {
  return indexOf(headers, columnName);
}

/**
 * Ensures sheet headers match the expected column definitions.
 * If headers are missing or misaligned, updates them to match.
 * This prevents data/header misalignment when column definitions are updated.
 *
 * @param {Sheet} sheet - The Google Sheet to check
 * @param {Array<string>} expectedColumns - Array of expected column header names
 */
function ensureSheetHeaders(sheet, expectedColumns) {
  if (!sheet || !expectedColumns || expectedColumns.length === 0) return;

  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    // Empty sheet - write headers
    var headerRange = sheet.getRange(1, 1, 1, expectedColumns.length);
    headerRange.setValues([expectedColumns]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f3f3f3');
    Logger.log('✅ Headers written to empty sheet "' + sheet.getName() + '" (' + expectedColumns.length + ' columns)');
    return;
  }

  var checkCols = Math.max(lastCol, expectedColumns.length);
  var currentHeaders = sheet.getRange(1, 1, 1, checkCols).getValues()[0];

  // Check if headers match expected columns
  var needsUpdate = false;
  for (var i = 0; i < expectedColumns.length; i++) {
    if (String(currentHeaders[i] || '').trim() !== expectedColumns[i]) {
      needsUpdate = true;
      break;
    }
  }

  if (needsUpdate) {
    Logger.log('⚠️ Sheet "' + sheet.getName() + '" headers do not match expected columns, updating...');
    Logger.log('   Expected ' + expectedColumns.length + ' columns: ' + expectedColumns.join(', '));
    Logger.log('   Found ' + lastCol + ' columns: ' + currentHeaders.slice(0, lastCol).join(', '));

    var newHeaderRange = sheet.getRange(1, 1, 1, expectedColumns.length);
    newHeaderRange.setValues([expectedColumns]);
    newHeaderRange.setFontWeight('bold');
    newHeaderRange.setBackground('#f3f3f3');

    Logger.log('✅ Headers updated to ' + expectedColumns.length + ' columns');
  }
}

// ============================================================================
// ID GENERATION
// ============================================================================

/**
 * Generate unique ID
 *
 * @returns {string} 8-character unique ID
 */
function generateId() {
  return Utilities.getUuid().substring(0, 8);
}

/**
 * Generate full UUID
 *
 * @returns {string} Full UUID string
 */
function generateFullUUID() {
  return Utilities.getUuid();
}

// ============================================================================
// DATA SANITIZATION FOR WEB SERIALIZATION
// ============================================================================

/**
 * Sanitize object for serialization to web app
 * Converts dates, handles null values, ensures all data is serializable
 * This is a general-purpose version that works for any object type
 *
 * @param {Object} obj - Object to sanitize
 * @returns {Object} Sanitized object
 */
function sanitizeForWeb(obj) {
  var sanitized = {};

  for (var key in obj) {
    if (obj.hasOwnProperty(key)) {
      var value = obj[key];

      // Convert Date objects to DD-MM-YYYY format (date only, no time)
      if (value instanceof Date) {
        sanitized[key] = formatDateDDMMYYYY(value);
      }
      // Convert null/undefined to empty string
      else if (value === null || value === undefined) {
        sanitized[key] = '';
      }
      // Keep numbers, strings, booleans as-is
      else if (typeof value === 'number' || typeof value === 'string' || typeof value === 'boolean') {
        sanitized[key] = value;
      }
      // Recursively sanitize nested objects
      else if (typeof value === 'object' && !Array.isArray(value)) {
        sanitized[key] = sanitizeForWeb(value);
      }
      // Handle arrays
      else if (Array.isArray(value)) {
        sanitized[key] = value.map(function(item) {
          if (typeof item === 'object') {
            return sanitizeForWeb(item);
          }
          return item;
        });
      }
      // Convert everything else to string
      else {
        sanitized[key] = String(value);
      }
    }
  }

  return sanitized;
}

// ============================================================================
// TIMESHEET UTILITIES
// ============================================================================

/**
 * Generate hash of file content for duplicate detection
 *
 * @param {string} content - File content or data to hash
 * @returns {string} MD5 hash of content
 */
function generateImportHash(content) {
  try {
    // Convert to string if not already
    var dataString = typeof content === 'string' ? content : JSON.stringify(content);

    // Generate MD5 hash
    var hash = Utilities.computeDigest(
      Utilities.DigestAlgorithm.MD5,
      dataString,
      Utilities.Charset.UTF_8
    );

    // Convert to hex string
    var hashString = hash.map(function(byte) {
      var v = (byte < 0) ? 256 + byte : byte;
      return ('0' + v.toString(16)).slice(-2);
    }).join('');

    return hashString;

  } catch (error) {
    logError('Failed to generate import hash', error);
    return '';
  }
}

/**
 * Format duration in minutes to readable string
 *
 * @param {number} minutes - Duration in minutes
 * @param {boolean} [verbose=false] - Use verbose format (e.g., "2 hours 30 minutes")
 * @returns {string} Formatted duration
 */
function formatDuration(minutes, verbose) {
  if (minutes === null || minutes === undefined || minutes === '') {
    return '0h 0m';
  }

  var totalMinutes = Math.abs(parseFloat(minutes));
  if (isNaN(totalMinutes)) {
    return '0h 0m';
  }

  var hours = Math.floor(totalMinutes / 60);
  var mins = Math.round(totalMinutes % 60);

  if (verbose) {
    var parts = [];
    if (hours > 0) {
      parts.push(hours + (hours === 1 ? ' hour' : ' hours'));
    }
    if (mins > 0 || hours === 0) {
      parts.push(mins + (mins === 1 ? ' minute' : ' minutes'));
    }
    return parts.join(' ');
  }

  return hours + 'h ' + mins + 'm';
}

/**
 * Parse Excel date serial number to JavaScript Date
 *
 * @param {number|Date|string} value - Excel serial number, Date object, or date string
 * @returns {Date} JavaScript Date object
 */
function parseExcelDate(value) {
  // Already a Date
  if (value instanceof Date) {
    return value;
  }

  // Excel serial number (days since 1900-01-01)
  if (typeof value === 'number') {
    // Excel incorrectly treats 1900 as a leap year
    var excelEpoch = new Date(1899, 11, 30); // December 30, 1899
    var milliseconds = value * 24 * 60 * 60 * 1000;
    return new Date(excelEpoch.getTime() + milliseconds);
  }

  // String - try to parse
  if (typeof value === 'string') {
    var parsed = new Date(value);
    if (!isNaN(parsed.getTime())) {
      return parsed;
    }
  }

  throw new Error('Cannot parse date value: ' + value);
}

/**
 * Get week ending date (Saturday) for a given date
 *
 * @param {Date|string} date - Date to get week ending for
 * @returns {Date} Saturday of that week
 */
function getWeekEnding(date) {
  var d = parseDate(date);

  // Get day of week (0 = Sunday, 6 = Saturday)
  var dayOfWeek = d.getDay();

  // Calculate days until Saturday
  var daysUntilSaturday = (6 - dayOfWeek + 7) % 7;

  // If already Saturday, use current date
  if (daysUntilSaturday === 0 && dayOfWeek === 6) {
    return d;
  }

  // Add days to get to Saturday
  var saturday = new Date(d);
  saturday.setDate(d.getDate() + daysUntilSaturday);

  return saturday;
}

/**
 * Get week starting date (Sunday) for a given date
 *
 * @param {Date|string} date - Date to get week starting for
 * @returns {Date} Sunday of that week
 */
function getWeekStarting(date) {
  var d = parseDate(date);

  // Get day of week (0 = Sunday, 6 = Saturday)
  var dayOfWeek = d.getDay();

  // Calculate days back to Sunday
  var daysBackToSunday = dayOfWeek;

  // If already Sunday, use current date
  if (daysBackToSunday === 0) {
    return d;
  }

  // Subtract days to get to Sunday
  var sunday = new Date(d);
  sunday.setDate(d.getDate() - daysBackToSunday);

  return sunday;
}

/**
 * Check if a date is a Friday
 *
 * @param {Date|string} date - Date to check
 * @returns {boolean} True if Friday
 */
function isFriday(date) {
  var d = parseDate(date);
  return d.getDay() === 5;
}

/**
 * Get first Friday of a month
 *
 * @param {Date|string} date - Any date in the month
 * @returns {Date} First Friday of that month
 */
function getFirstFridayOfMonth(date) {
  var d = parseDate(date);

  // Set to first day of month
  var firstDay = new Date(d.getFullYear(), d.getMonth(), 1);

  // Get day of week for first day (0 = Sunday, 5 = Friday)
  var dayOfWeek = firstDay.getDay();

  // Calculate days until first Friday
  var daysUntilFriday;
  if (dayOfWeek <= 5) {
    // Friday is in the same week
    daysUntilFriday = 5 - dayOfWeek;
  } else {
    // Friday is in the next week (Saturday = 6, need to go to next Friday)
    daysUntilFriday = 6; // From Saturday to Friday
  }

  // Add days to get to first Friday
  var firstFriday = new Date(firstDay);
  firstFriday.setDate(firstDay.getDate() + daysUntilFriday);

  return firstFriday;
}

/**
 * Get last Friday of a month
 *
 * @param {Date|string} date - Any date in the month
 * @returns {Date} Last Friday of that month
 */
function getLastFridayOfMonth(date) {
  var d = parseDate(date);

  // Set to last day of month
  var lastDay = new Date(d.getFullYear(), d.getMonth() + 1, 0);

  // Get day of week for last day (0 = Sunday, 5 = Friday)
  var dayOfWeek = lastDay.getDay();

  // Calculate days back to last Friday
  var daysBackToFriday;
  if (dayOfWeek >= 5) {
    // Friday is in the same week
    daysBackToFriday = dayOfWeek - 5;
  } else {
    // Friday is in the previous week
    daysBackToFriday = dayOfWeek + 2; // e.g., if Thursday (4), go back 6 days
  }

  // Subtract days to get to last Friday
  var lastFriday = new Date(lastDay);
  lastFriday.setDate(lastDay.getDate() - daysBackToFriday);

  return lastFriday;
}

/**
 * Get unique values from array
 *
 * @param {Array} array - Array to get unique values from
 * @returns {Array} Array with unique values only
 */
function getUniqueValues(array) {
  var seen = {};
  var result = [];

  for (var i = 0; i < array.length; i++) {
    var value = array[i];
    var key = JSON.stringify(value);

    if (!seen[key]) {
      seen[key] = true;
      result.push(value);
    }
  }

  return result;
}

/**
 * Parse CSV string to array of objects
 *
 * @param {string} csvString - CSV data as string
 * @param {boolean} [hasHeaders=true] - First row contains headers
 * @returns {Array} Array of objects
 */
function parseCSV(csvString, hasHeaders) {
  if (hasHeaders === undefined) hasHeaders = true;

  var lines = csvString.split('\n');
  var headers = [];
  var result = [];

  if (hasHeaders && lines.length > 0) {
    headers = lines[0].split(',').map(function(h) {
      return h.trim();
    });
  }

  var startIndex = hasHeaders ? 1 : 0;

  for (var i = startIndex; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;

    var values = line.split(',').map(function(v) {
      return v.trim();
    });

    if (hasHeaders) {
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        obj[headers[j]] = values[j] || '';
      }
      result.push(obj);
    } else {
      result.push(values);
    }
  }

  return result;
}

// ============================================================================
// USER AUTHORIZATION UTILITIES
// ============================================================================

/**
 * Generates a random salt for password hashing.
 * @returns {string} A unique salt.
 */
function generateSalt() {
  return Utilities.getUuid();
}

/**
 * Hashes a password with a given salt.
 * @param {string} password The plain text password.
 * @param {string} salt The salt to use for hashing.
 * @returns {string} The hashed password.
 */
function hashPassword(password, salt) {
  var saltedPassword = password + salt;
  var signature = Utilities.computeHmacSha256Signature(saltedPassword, salt);
  return Utilities.base64Encode(signature);
}

/**
 * Verifies a password against a salt and hash.
 * @param {string} password The plain text password to verify.
 * @param {string} salt The salt used for the original hash.
 * @param {string} hash The stored password hash.
 * @returns {boolean} True if the password is correct, false otherwise.
 */
function verifyPassword(password, salt, hash) {
  var newHash = hashPassword(password, salt);
  return newHash === hash;
}

/**
 * Get list of authorized user emails from UserConfig sheet
 * First tries to load from Script Properties (doesn't require spreadsheet access)
 * Falls back to reading from UserConfig sheet if properties not set
 *
 * @returns {Array<string>} Array of authorized email addresses
 */
function getAuthorizedUsers() {
  try {
    var scriptProperties = PropertiesService.getScriptProperties();
    var userConfigJson = scriptProperties.getProperty('USER_CONFIG_DATA');

    if (userConfigJson) {
      var users = JSON.parse(userConfigJson);
      return users.map(function(u) { return u.Email.toLowerCase(); });
    }
    return [];
  } catch (error) {
    logError('Failed to get authorized users from Script Properties', error);
    return [];
  }
}

/**
 * Check if a user email is in the authorized users list
 *
 * @param {string} userEmail - Email address to check
 * @returns {boolean} True if user is authorized
 */
function isUserAuthorized(userEmail) {
  if (!userEmail || typeof userEmail !== 'string') {
    return false;
  }
  var authorizedUsers = getAuthorizedUsers();
  return authorizedUsers.indexOf(userEmail.trim().toLowerCase()) !== -1;
}

/**
 * Create UserConfig sheet with headers if it doesn't exist
 * This is a helper function for initial setup
 *
 * @returns {Object} Result object with success status
 */
function setupUserConfigSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = getSheets();

    // Check if sheet already exists
    if (sheets.userConfig) {
      return {
        success: true,
        message: 'UserConfig sheet already exists',
        sheetName: sheets.userConfig.getName()
      };
    }

    // Create new sheet
    var sheet = ss.insertSheet('UserConfig');

    // Add headers
    var headers = USER_CONFIG_COLUMNS;
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Format header row
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');

    // Set column widths
    sheet.setColumnWidth(1, 200); // Name
    sheet.setColumnWidth(2, 300); // Email
    sheet.setColumnWidth(3, 150); // Password

    // Freeze header row
    sheet.setFrozenRows(1);

    logSuccess('UserConfig sheet created successfully');

    return {
      success: true,
      message: 'UserConfig sheet created successfully',
      sheetName: 'UserConfig'
    };

  } catch (error) {
    logError('Failed to create UserConfig sheet', error);
    return {
      success: false,
      error: 'Failed to create UserConfig sheet: ' + error.toString()
    };
  }
}

/**
 * Sync UserConfig sheet to Script Properties
 * This allows the whitelist to be accessed without requiring spreadsheet access
 * Call this function after updating the UserConfig sheet
 *
 * @returns {Object} Result object with success status
 */
function syncUserConfigToProperties() {
  try {
    var sheets = getSheets();

    if (!sheets.userConfig) {
      return {
        success: false,
        error: 'UserConfig sheet not found. Create it first using "Setup UserConfig Sheet".'
      };
    }

    var sheet = sheets.userConfig;
    var data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return {
        success: false,
        error: 'UserConfig sheet is empty. Add authorized users first.'
      };
    }

    var headers = data[0];
    var users = [];
    for (var i = 1; i < data.length; i++) {
      var user = buildObjectFromRow(data[i], headers);
      if (user.Email) { // only add users with email
        users.push(user);
      }
    }

    if (users.length === 0) {
      return {
        success: false,
        error: 'No valid users found in UserConfig sheet'
      };
    }

    // Save to Script Properties
    var scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('USER_CONFIG_DATA', JSON.stringify(users));
    scriptProperties.setProperty('USER_CONFIG_LAST_SYNC', new Date().toISOString());

    // Clean up old property
    scriptProperties.deleteProperty('USER_WHITELIST');
    scriptProperties.deleteProperty('USER_WHITELIST_LAST_SYNC');

    logSuccess('Synced ' + users.length + ' users to Script Properties');

    return {
      success: true,
      message: 'Successfully synced ' + users.length + ' authorized users',
      count: users.length,
      users: users.map(function(u) { return u.Email; })
    };

  } catch (error) {
    logError('Failed to sync UserConfig to properties', error);
    return {
      success: false,
      error: 'Failed to sync: ' + error.toString()
    };
  }
}

/**
 * Clear user whitelist from Script Properties
 * Use this to reset the whitelist
 *
 * @returns {Object} Result object with success status
 */
function clearUserWhitelist() {
  try {
    var scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.deleteProperty('USER_WHITELIST');
    scriptProperties.deleteProperty('USER_WHITELIST_LAST_SYNC');

    logSuccess('User whitelist cleared from Script Properties');

    return {
      success: true,
      message: 'User whitelist cleared successfully'
    };

  } catch (error) {
    logError('Failed to clear user whitelist', error);
    return {
      success: false,
      error: 'Failed to clear whitelist: ' + error.toString()
    };
  }
}

/**
 * Authenticate user with email and password
 * Checks credentials against UserConfig sheet
 *
 * @param {string} email - User's email address
 * @param {string} password - User's password
 * @returns {Object} Result object with success status and user info
 */
function authenticateUser(email, password) {
  try {
    if (!email || !password) {
      return { success: false, error: 'Email and password are required' };
    }

    var scriptProperties = PropertiesService.getScriptProperties();
    var userConfigJson = scriptProperties.getProperty('USER_CONFIG_DATA');

    if (!userConfigJson) {
      return { success: false, error: 'User configuration not found. Please sync users.' };
    }

    var users = JSON.parse(userConfigJson);
    var normalizedEmail = email.trim().toLowerCase();
    var user = users.find(function(u) {
      return u.Email.toLowerCase() === normalizedEmail;
    });

    if (user) {
      if (verifyPassword(password, user.PasswordSalt, user.PasswordHash)) {
        logSuccess('User authenticated: ' + email);
        return {
          success: true,
          user: {
            email: user.Email,
            name: user.Name
          }
        };
      } else {
        logWarning('Failed login attempt for: ' + email);
        return { success: false, error: 'Invalid password' };
      }
    } else {
      logWarning('Login attempt for unknown user: ' + email);
      return { success: false, error: 'User not found' };
    }
  } catch (error) {
    logError('Authentication error', error);
    return { success: false, error: 'Authentication failed: ' + error.toString() };
  }
}

/**
 * Create session for authenticated user
 * Stores user email in CacheService for 6 hours
 *
 * @param {string} email - User's email address
 * @returns {string} Session token
 */
function createUserSession(email) {
  var sessionToken = Utilities.getUuid();
  var cache = CacheService.getScriptCache();

  // Store for 6 hours (21600 seconds)
  cache.put('session_' + sessionToken, email, 21600);

  logInfo('Session created for: ' + email);
  return sessionToken;
}

/**
 * Get user email from session token
 *
 * @param {string} sessionToken - Session token
 * @returns {string|null} User email or null if session invalid
 */
function getUserFromSession(sessionToken) {
  if (!sessionToken) {
    return null;
  }

  var cache = CacheService.getScriptCache();
  var email = cache.get('session_' + sessionToken);

  return email;
}

/**
 * Destroy user session
 *
 * @param {string} sessionToken - Session token to destroy
 */
function destroySession(sessionToken) {
  if (!sessionToken) {
    return;
  }

  var cache = CacheService.getScriptCache();
  cache.remove('session_' + sessionToken);

  logInfo('Session destroyed');
}

/**
 * Creates a new user with a salted and hashed password.
 * @param {string} name The user's full name.
 * @param {string} email The user's email address.
 * @param {string} password The user's plain text password.
 * @returns {Object} A result object indicating success or failure.
 */
function createNewUser(name, email, password) {
  try {
    var salt = generateSalt();
    var hash = hashPassword(password, salt);

    var sheets = getSheets();
    if (!sheets.userConfig) {
      setupUserConfigSheet();
      sheets = getSheets();
    }
    var sheet = sheets.userConfig;
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var emailIndex = indexOf(headers, 'Email');

    // Check if user already exists
    for (var i = 1; i < data.length; i++) {
      if (data[i][emailIndex] === email) {
        return { success: false, error: 'User already exists' };
      }
    }

    var newRow = objectToRow({
      'Name': name,
      'Email': email,
      'PasswordHash': hash,
      'PasswordSalt': salt
    }, headers);

    sheet.appendRow(newRow);

    // Sync properties after adding the new user
    syncUserConfigToProperties();

    return { success: true, user: { name: name, email: email } };
  } catch (error) {
    logError('Failed to create new user', error);
    return { success: false, error: 'Failed to create new user: ' + error.toString() };
  }
}

// ============================================================================
// PASSWORD RESET UTILITIES
// ============================================================================

/**
 * Generate a password reset token
 * Stores token with email and expiry in Script Properties
 *
 * @param {string} email - User's email address
 * @returns {Object} Result with token if successful
 */
function generatePasswordResetToken(email) {
  try {
    if (!email) {
      return { success: false, error: 'Email is required' };
    }

    // Check if user exists
    var scriptProperties = PropertiesService.getScriptProperties();
    var userConfigJson = scriptProperties.getProperty('USER_CONFIG_DATA');

    if (!userConfigJson) {
      return { success: false, error: 'User configuration not found' };
    }

    var users = JSON.parse(userConfigJson);
    var normalizedEmail = email.trim().toLowerCase();
    var user = users.find(function(u) {
      return u.Email.toLowerCase() === normalizedEmail;
    });

    if (!user) {
      // Don't reveal that user doesn't exist (security best practice)
      return { success: true, message: 'If this email exists, a reset link has been sent' };
    }

    // Check rate limiting
    var rateLimitKey = 'reset_rate_' + normalizedEmail;
    var rateLimitData = scriptProperties.getProperty(rateLimitKey);

    if (rateLimitData) {
      var rateLimit = JSON.parse(rateLimitData);
      var now = new Date().getTime();

      // Clean up old requests (older than 1 hour)
      rateLimit.requests = rateLimit.requests.filter(function(timestamp) {
        return now - timestamp < 3600000; // 1 hour
      });

      // Check if exceeded max requests
      if (rateLimit.requests.length >= PASSWORD_RESET_MAX_REQUESTS_PER_HOUR) {
        return {
          success: false,
          error: 'Too many reset requests. Please try again later.'
        };
      }

      rateLimit.requests.push(now);
      scriptProperties.setProperty(rateLimitKey, JSON.stringify(rateLimit));
    } else {
      scriptProperties.setProperty(rateLimitKey, JSON.stringify({
        requests: [new Date().getTime()]
      }));
    }

    // Generate unique token
    var token = Utilities.getUuid();
    var expiry = new Date().getTime() + PASSWORD_RESET_TOKEN_EXPIRY;

    // Store token with email and expiry
    var tokenKey = 'reset_token_' + token;
    var tokenData = {
      email: user.Email,
      expiry: expiry,
      used: false
    };

    scriptProperties.setProperty(tokenKey, JSON.stringify(tokenData));

    logInfo('Password reset token generated for: ' + email);

    return {
      success: true,
      token: token,
      userName: user.Name,
      userEmail: user.Email
    };

  } catch (error) {
    logError('Failed to generate password reset token', error);
    return { success: false, error: 'Failed to generate reset token: ' + error.toString() };
  }
}

/**
 * Validate a password reset token
 * Checks if token exists, is not expired, and has not been used
 *
 * @param {string} token - Reset token to validate
 * @returns {Object} Result with email if valid
 */
function validatePasswordResetToken(token) {
  try {
    if (!token) {
      return { success: false, error: 'Token is required' };
    }

    var scriptProperties = PropertiesService.getScriptProperties();
    var tokenKey = 'reset_token_' + token;
    var tokenDataJson = scriptProperties.getProperty(tokenKey);

    if (!tokenDataJson) {
      return { success: false, error: 'Invalid or expired reset token' };
    }

    var tokenData = JSON.parse(tokenDataJson);
    var now = new Date().getTime();

    // Check if token has expired
    if (now > tokenData.expiry) {
      scriptProperties.deleteProperty(tokenKey);
      return { success: false, error: 'Reset token has expired' };
    }

    // Check if token has been used
    if (tokenData.used) {
      return { success: false, error: 'Reset token has already been used' };
    }

    return {
      success: true,
      email: tokenData.email
    };

  } catch (error) {
    logError('Failed to validate password reset token', error);
    return { success: false, error: 'Failed to validate token: ' + error.toString() };
  }
}

/**
 * Reset user password using a valid token
 * Marks token as used after successful reset
 *
 * @param {string} token - Reset token
 * @param {string} newPassword - New password
 * @returns {Object} Result indicating success or failure
 */
function resetPasswordWithToken(token, newPassword) {
  try {
    logInfo('=== resetPasswordWithToken START ===');
    logInfo('Token: ' + token);
    logInfo('New password length: ' + (newPassword ? newPassword.length : 0));

    if (!token || !newPassword) {
      logError('Missing token or password');
      return { success: false, error: 'Token and new password are required' };
    }

    // Validate token first
    logInfo('Validating token...');
    var validation = validatePasswordResetToken(token);
    logInfo('Token validation result: ' + JSON.stringify(validation));

    if (!validation.success) {
      logError('Token validation failed');
      return validation;
    }

    var email = validation.email;
    logInfo('Email from token: ' + email);

    // Update password in UserConfig sheet
    logInfo('Getting sheets...');
    var sheets = getSheets();
    if (!sheets.userConfig) {
      logError('UserConfig sheet not found');
      return { success: false, error: 'UserConfig sheet not found' };
    }
    logInfo('UserConfig sheet found');

    var sheet = sheets.userConfig;
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    logInfo('Headers: ' + JSON.stringify(headers));

    var emailIndex = indexOf(headers, 'Email');
    var hashIndex = indexOf(headers, 'PasswordHash');
    var saltIndex = indexOf(headers, 'PasswordSalt');
    logInfo('Column indices - Email: ' + emailIndex + ', Hash: ' + hashIndex + ', Salt: ' + saltIndex);

    // Find user row
    var userRowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][emailIndex].toLowerCase() === email.toLowerCase()) {
        userRowIndex = i;
        logInfo('User found at row: ' + userRowIndex);
        break;
      }
    }

    if (userRowIndex === -1) {
      logError('User not found in UserConfig: ' + email);
      return { success: false, error: 'User not found' };
    }

    // Generate new salt and hash
    logInfo('Generating new salt and hash...');
    var newSalt = generateSalt();
    var newHash = hashPassword(newPassword, newSalt);
    logInfo('New salt length: ' + (newSalt ? newSalt.length : 0));
    logInfo('New hash length: ' + (newHash ? newHash.length : 0));

    // Update the sheet
    logInfo('Updating sheet at row ' + (userRowIndex + 1) + ', hash col ' + (hashIndex + 1) + ', salt col ' + (saltIndex + 1));
    sheet.getRange(userRowIndex + 1, hashIndex + 1).setValue(newHash);
    sheet.getRange(userRowIndex + 1, saltIndex + 1).setValue(newSalt);
    logInfo('Sheet updated successfully');

    // Mark token as used
    logInfo('Marking token as used...');
    var scriptProperties = PropertiesService.getScriptProperties();
    var tokenKey = 'reset_token_' + token;
    var tokenDataJson = scriptProperties.getProperty(tokenKey);
    var tokenData = JSON.parse(tokenDataJson);
    tokenData.used = true;
    scriptProperties.setProperty(tokenKey, JSON.stringify(tokenData));
    logInfo('Token marked as used');

    // Sync to Script Properties
    logInfo('Syncing to properties...');
    syncUserConfigToProperties();
    logInfo('Sync complete');

    logSuccess('Password reset successful for: ' + email);
    logInfo('=== resetPasswordWithToken END ===');

    return {
      success: true,
      message: 'Password reset successfully'
    };

  } catch (error) {
    logError('Failed to reset password', error);
    logError('Error stack: ' + error.stack);
    return { success: false, error: 'Failed to reset password: ' + error.toString() };
  }
}

/**
 * Send password reset email
 * Generates token and sends email with reset link
 *
 * @param {string} email - User's email address
 * @param {string} resetUrl - Base URL for password reset page
 * @returns {Object} Result indicating success or failure
 */
function sendPasswordResetEmail(email, resetUrl) {
  try {
    logInfo('=== sendPasswordResetEmail START ===');
    logInfo('Email: ' + email);
    logInfo('Reset URL: ' + resetUrl);

    // Generate token
    logInfo('Generating token...');
    var tokenResult = generatePasswordResetToken(email);
    logInfo('Token result: ' + JSON.stringify(tokenResult));

    if (!tokenResult.success) {
      logWarning('Token generation failed or user not found');
      // If user doesn't exist, still return success to not reveal user existence
      if (!tokenResult.token) {
        return { success: true, message: 'If this email exists, a reset link has been sent' };
      }
      return tokenResult;
    }

    // Build reset link
    logInfo('Building reset link...');
    var separator = resetUrl.indexOf('?') > -1 ? '&' : '?';
    var resetLink = resetUrl + separator + 'page=reset&token=' + tokenResult.token;

    logInfo('Reset URL base: ' + resetUrl);
    logInfo('Reset link generated: ' + resetLink);

    // Get email template
    logInfo('Getting email template...');
    var template = EMAIL_TEMPLATES.passwordReset;
    var subject = template.subject;
    var body = template.getBody(resetLink, tokenResult.userName);

    logInfo('Email subject: ' + subject);
    logInfo('Email to: ' + tokenResult.userEmail);

    // Send email
    logInfo('Sending email via MailApp...');
    MailApp.sendEmail({
      to: tokenResult.userEmail,
      subject: subject,
      body: body
    });

    logSuccess('Password reset email sent to: ' + email);
    logInfo('=== sendPasswordResetEmail END ===');

    return {
      success: true,
      message: 'Password reset email sent successfully'
    };

  } catch (error) {
    logError('Failed to send password reset email', error);
    logError('Error details: ' + error.toString());
    logError('Error stack: ' + error.stack);
    return { success: false, error: 'Failed to send reset email: ' + error.toString() };
  }
}

// ============================================================================
// EXPORT UTILITIES FOR TESTING
// ============================================================================

// Make functions available globally for testing
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    isEmpty: isEmpty,
    isValidSAIdNumber: isValidSAIdNumber,
    isValidPhoneNumber: isValidPhoneNumber,
    formatPhoneNumber: formatPhoneNumber,
    formatDate: formatDate,
    formatDuration: formatDuration,
    generateId: generateId,
    generateImportHash: generateImportHash,
    parseExcelDate: parseExcelDate,
    getWeekEnding: getWeekEnding,
    getWeekStarting: getWeekStarting,
    logFunctionStart: logFunctionStart,
    logFunctionEnd: logFunctionEnd,
    logSuccess: logSuccess,
    logWarning: logWarning,
    logError: logError,
    logInfo: logInfo
  };
}
