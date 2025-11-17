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
 * Get all sheets with proper error handling
 *
 * @returns {Object} Object with sheet references
 */
function getSheets() {
  try {
    console.log('=== getSheets() START ===');

    // Step 1: Get active spreadsheet
    console.log('Step 1: Getting active spreadsheet...');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    console.log('✓ Spreadsheet retrieved:', ss.getName());
    console.log('  Spreadsheet ID:', ss.getId());

    var sheets = {};

    // Step 2: Get all sheets
    console.log('Step 2: Getting all sheets...');
    var allSheets = ss.getSheets();
    console.log('✓ Found', allSheets.length, 'sheets total');

    // Step 3: List all sheet names
    console.log('Step 3: Listing all sheet names:');
    for (var i = 0; i < allSheets.length; i++) {
      console.log('  - Sheet ' + (i + 1) + ':', allSheets[i].getName());
    }

    // Step 4: Map sheets by name (case-insensitive)
    console.log('Step 4: Mapping sheets to expected keys...');
    for (var i = 0; i < allSheets.length; i++) {
      var sheet = allSheets[i];
      var sheetName = sheet.getName().toLowerCase().replace(/\s+/g, ''); // Remove spaces
      console.log('  Processing sheet:', sheet.getName(), '(normalized:', sheetName + ')');

      // Map to expected keys - ORDER MATTERS! Most specific first
      // Check for clock-in imports (most specific first)
      if (sheetName.indexOf('clockinimports') >= 0 || sheetName.indexOf('clock_in_imports') >= 0) {
        sheets.clockImports = sheet;
        console.log('    ✓ Mapped to: clockImports');
      }
      // Check for raw clock data
      else if (sheetName.indexOf('rawclockdata') >= 0 || sheetName.indexOf('raw_clock_data') >= 0 || sheetName.indexOf('clockdata') >= 0) {
        sheets.rawClockData = sheet;
        console.log('    ✓ Mapped to: rawClockData');
      }
      // Check loans BEFORE employee to avoid conflict
      else if (sheetName.indexOf('loan') >= 0) {
        sheets.loans = sheet;
        console.log('    ✓ Mapped to: loans');
      }
      // Check for employee details/empdetails (but not loans)
      else if (sheetName.indexOf('employeedetails') >= 0 || sheetName === 'empdetails') {
        sheets.empdetails = sheet;
        console.log('    ✓ Mapped to: empdetails');
      }
      // Check salary/payroll
      else if (sheetName.indexOf('salary') >= 0 || sheetName.indexOf('payroll') >= 0) {
        sheets.salary = sheet;
        console.log('    ✓ Mapped to: salary');
      }
      // Check leave
      else if (sheetName.indexOf('leave') >= 0) {
        sheets.leave = sheet;
        console.log('    ✓ Mapped to: leave');
      }
      // Check pending/timesheets
      else if (sheetName.indexOf('pending') >= 0 || sheetName.indexOf('timesheet') >= 0) {
        sheets.pending = sheet;
        console.log('    ✓ Mapped to: pending');
      } else {
        console.log('    - Not mapped (no matching key)');
      }
    }

    // Step 5: Summary
    console.log('Step 5: Summary of mapped sheets:');
    var mappedKeys = Object.keys(sheets);
    console.log('  Mapped', mappedKeys.length, 'sheets:', mappedKeys.join(', '));

    if (sheets.empdetails) {
      console.log('  ✓ empdetails sheet found:', sheets.empdetails.getName());
    } else {
      console.warn('  ✗ empdetails sheet NOT found!');
    }

    console.log('=== getSheets() END ===');
    return sheets;

  } catch (error) {
    console.error('=== getSheets() ERROR ===');
    console.error('Error type:', error.name);
    console.error('Error message:', error.message);
    console.error('Full error:', error.toString());
    console.error('Stack trace:', error.stack);
    logError('Failed to get sheets', error);
    throw new Error('Failed to access spreadsheet: ' + error.toString());
  }
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

    // Handle dates
    if (value instanceof Date) {
      obj[header] = value;
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
 * Parse date from various input formats
 *
 * @param {Date|string|number} dateValue - Date in various formats
 * @returns {Date} Parsed Date object
 * @throws {Error} If date format is invalid
 */
function parseDate(dateValue) {
  if (!dateValue) {
    throw new Error('Date value is required');
  }

  // Already a Date object
  if (dateValue instanceof Date) {
    return dateValue;
  }

  // String or number - convert to Date
  if (typeof dateValue === 'string' || typeof dateValue === 'number') {
    var parsed = new Date(dateValue);

    // Check if valid date
    if (isNaN(parsed.getTime())) {
      throw new Error('Invalid date format: ' + dateValue);
    }

    return parsed;
  }

  throw new Error('Invalid date format: ' + typeof dateValue);
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

  var year = date.getFullYear();
  var month = ('0' + (date.getMonth() + 1)).slice(-2);
  var day = ('0' + date.getDate()).slice(-2);

  return year + '-' + month + '-' + day;
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

      // Convert Date objects to ISO strings
      if (value instanceof Date) {
        sanitized[key] = value.toISOString();
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
