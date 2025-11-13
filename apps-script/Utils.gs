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
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheets = {};
    
    // Get all sheets
    var allSheets = ss.getSheets();
    
    // Map sheets by name (case-insensitive)
    for (var i = 0; i < allSheets.length; i++) {
      var sheet = allSheets[i];
      var sheetName = sheet.getName().toLowerCase();
      
      // Map to expected keys
      if (sheetName.indexOf('employee') >= 0 || sheetName === 'empdetails') {
        sheets.empdetails = sheet;
      } else if (sheetName.indexOf('salary') >= 0) {
        sheets.salary = sheet;
      } else if (sheetName.indexOf('loan') >= 0) {
        sheets.loans = sheet;
      } else if (sheetName.indexOf('leave') >= 0) {
        sheets.leave = sheet;
      }
    }
    
    return sheets;
    
  } catch (error) {
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
    generateId: generateId,
    logFunctionStart: logFunctionStart,
    logFunctionEnd: logFunctionEnd,
    logSuccess: logSuccess,
    logWarning: logWarning,
    logError: logError,
    logInfo: logInfo
  };
}
