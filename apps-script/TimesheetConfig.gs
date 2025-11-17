/**
 * TIMESHEETCONFIG.GS - Configurable Time Processing Rules
 * HR Payroll System for SA Grinding Wheels & Scorpio Abrasives
 *
 * This module manages configurable time processing rules for the clock-in system.
 * Rules are stored in PropertiesService for persistence across sessions.
 */

// ==================== DEFAULT TIME CONFIGURATION ====================

/**
 * Default time processing configuration
 * These values are used when no custom configuration exists
 */
var DEFAULT_TIME_CONFIG = {
  // Standard times
  standardStartTime: '07:30',
  standardEndTime: '16:30',
  fridayEndTime: '13:00',

  // Buffers & Grace periods
  graceMinutes: 5,
  endBufferMinutes: 5,
  lunchBufferMinutes: 5,

  // Lunch rules
  minLunchMinutes: 20,
  maxLunchMinutes: 35,
  standardLunchMinutes: 30,
  applyLunchOnFriday: false,
  lateMorningThreshold: '11:00',

  // Other rules
  dailyBathroomThreshold: 30,
  earlyChangeThreshold: 10,
  longBathroomThreshold: 15,
  projectIncompleteDays: true
};

// ==================== GET CONFIGURATION ====================

/**
 * Get current time processing configuration
 * Loads from PropertiesService or returns defaults if not set
 *
 * @returns {Object} Current time configuration
 */
function getTimeConfig() {
  try {
    var props = PropertiesService.getScriptProperties();
    var configString = props.getProperty('TIME_CONFIG');

    if (configString) {
      // Parse stored configuration
      var config = JSON.parse(configString);

      // Merge with defaults to ensure all properties exist
      return mergeWithDefaults(config);
    }

    // Return defaults if no config stored
    return JSON.parse(JSON.stringify(DEFAULT_TIME_CONFIG));

  } catch (error) {
    Logger.log('⚠️ Error loading time config: ' + error.message);
    Logger.log('Returning default configuration');
    return JSON.parse(JSON.stringify(DEFAULT_TIME_CONFIG));
  }
}

/**
 * Merge custom config with defaults to ensure all properties exist
 *
 * @param {Object} customConfig - Custom configuration
 * @returns {Object} Merged configuration
 */
function mergeWithDefaults(customConfig) {
  var merged = JSON.parse(JSON.stringify(DEFAULT_TIME_CONFIG));

  for (var key in customConfig) {
    if (customConfig.hasOwnProperty(key)) {
      merged[key] = customConfig[key];
    }
  }

  return merged;
}

// ==================== UPDATE CONFIGURATION ====================

/**
 * Update time processing configuration
 * Validates and saves configuration to PropertiesService
 *
 * @param {Object} newConfig - New configuration values
 * @returns {Object} Result with success flag
 */
function updateTimeConfig(newConfig) {
  try {
    Logger.log('\n========== UPDATE TIME CONFIG ==========');
    Logger.log('New config: ' + JSON.stringify(newConfig));

    // Validate configuration
    var validation = validateTimeConfig(newConfig);
    if (!validation.valid) {
      throw new Error('Invalid configuration: ' + validation.errors.join(', '));
    }

    // Merge with defaults to ensure all properties exist
    var fullConfig = mergeWithDefaults(newConfig);

    // Save to PropertiesService
    var props = PropertiesService.getScriptProperties();
    props.setProperty('TIME_CONFIG', JSON.stringify(fullConfig));

    Logger.log('✅ Configuration saved successfully');
    Logger.log('========== UPDATE TIME CONFIG COMPLETE ==========\n');

    return {
      success: true,
      data: fullConfig
    };

  } catch (error) {
    Logger.log('❌ ERROR in updateTimeConfig: ' + error.message);
    Logger.log('========== UPDATE TIME CONFIG FAILED ==========\n');
    return {
      success: false,
      error: error.message
    };
  }
}

// ==================== RESET CONFIGURATION ====================

/**
 * Reset time processing configuration to defaults
 *
 * @returns {Object} Result with success flag
 */
function resetTimeConfig() {
  try {
    Logger.log('\n========== RESET TIME CONFIG ==========');

    var props = PropertiesService.getScriptProperties();
    props.deleteProperty('TIME_CONFIG');

    Logger.log('✅ Configuration reset to defaults');
    Logger.log('========== RESET TIME CONFIG COMPLETE ==========\n');

    return {
      success: true,
      data: JSON.parse(JSON.stringify(DEFAULT_TIME_CONFIG))
    };

  } catch (error) {
    Logger.log('❌ ERROR in resetTimeConfig: ' + error.message);
    Logger.log('========== RESET TIME CONFIG FAILED ==========\n');
    return {
      success: false,
      error: error.message
    };
  }
}

// ==================== VALIDATION ====================

/**
 * Validate time configuration
 *
 * @param {Object} config - Configuration to validate
 * @returns {Object} Validation result with valid flag and errors array
 */
function validateTimeConfig(config) {
  var errors = [];

  // Validate time formats (HH:MM)
  var timePattern = /^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$/;

  if (config.standardStartTime && !timePattern.test(config.standardStartTime)) {
    errors.push('Invalid standardStartTime format (use HH:MM)');
  }

  if (config.standardEndTime && !timePattern.test(config.standardEndTime)) {
    errors.push('Invalid standardEndTime format (use HH:MM)');
  }

  if (config.fridayEndTime && !timePattern.test(config.fridayEndTime)) {
    errors.push('Invalid fridayEndTime format (use HH:MM)');
  }

  if (config.lateMorningThreshold && !timePattern.test(config.lateMorningThreshold)) {
    errors.push('Invalid lateMorningThreshold format (use HH:MM)');
  }

  // Validate numeric values
  if (config.graceMinutes !== undefined) {
    var graceMin = parseFloat(config.graceMinutes);
    if (isNaN(graceMin) || graceMin < 0 || graceMin > 60) {
      errors.push('graceMinutes must be between 0 and 60');
    }
  }

  if (config.endBufferMinutes !== undefined) {
    var endBuffer = parseFloat(config.endBufferMinutes);
    if (isNaN(endBuffer) || endBuffer < 0 || endBuffer > 60) {
      errors.push('endBufferMinutes must be between 0 and 60');
    }
  }

  if (config.lunchBufferMinutes !== undefined) {
    var lunchBuffer = parseFloat(config.lunchBufferMinutes);
    if (isNaN(lunchBuffer) || lunchBuffer < 0 || lunchBuffer > 60) {
      errors.push('lunchBufferMinutes must be between 0 and 60');
    }
  }

  if (config.minLunchMinutes !== undefined) {
    var minLunch = parseFloat(config.minLunchMinutes);
    if (isNaN(minLunch) || minLunch < 0 || minLunch > 120) {
      errors.push('minLunchMinutes must be between 0 and 120');
    }
  }

  if (config.maxLunchMinutes !== undefined) {
    var maxLunch = parseFloat(config.maxLunchMinutes);
    if (isNaN(maxLunch) || maxLunch < 0 || maxLunch > 120) {
      errors.push('maxLunchMinutes must be between 0 and 120');
    }
  }

  if (config.standardLunchMinutes !== undefined) {
    var stdLunch = parseFloat(config.standardLunchMinutes);
    if (isNaN(stdLunch) || stdLunch < 0 || stdLunch > 120) {
      errors.push('standardLunchMinutes must be between 0 and 120');
    }
  }

  if (config.dailyBathroomThreshold !== undefined) {
    var bathThreshold = parseFloat(config.dailyBathroomThreshold);
    if (isNaN(bathThreshold) || bathThreshold < 0 || bathThreshold > 240) {
      errors.push('dailyBathroomThreshold must be between 0 and 240');
    }
  }

  if (config.earlyChangeThreshold !== undefined) {
    var earlyThreshold = parseFloat(config.earlyChangeThreshold);
    if (isNaN(earlyThreshold) || earlyThreshold < 0 || earlyThreshold > 60) {
      errors.push('earlyChangeThreshold must be between 0 and 60');
    }
  }

  if (config.longBathroomThreshold !== undefined) {
    var longBathroom = parseFloat(config.longBathroomThreshold);
    if (isNaN(longBathroom) || longBathroom < 0 || longBathroom > 120) {
      errors.push('longBathroomThreshold must be between 0 and 120');
    }
  }

  // Validate boolean values
  if (config.applyLunchOnFriday !== undefined && typeof config.applyLunchOnFriday !== 'boolean') {
    errors.push('applyLunchOnFriday must be true or false');
  }

  if (config.projectIncompleteDays !== undefined && typeof config.projectIncompleteDays !== 'boolean') {
    errors.push('projectIncompleteDays must be true or false');
  }

  // Validate logical relationships
  if (config.minLunchMinutes && config.maxLunchMinutes) {
    if (parseFloat(config.minLunchMinutes) > parseFloat(config.maxLunchMinutes)) {
      errors.push('minLunchMinutes cannot be greater than maxLunchMinutes');
    }
  }

  return {
    valid: errors.length === 0,
    errors: errors
  };
}

// ==================== EXPORT/IMPORT CONFIGURATION ====================

/**
 * Export current configuration as JSON string
 *
 * @returns {Object} Result with configuration JSON
 */
function exportTimeConfig() {
  try {
    var config = getTimeConfig();

    return {
      success: true,
      data: JSON.stringify(config, null, 2)
    };

  } catch (error) {
    Logger.log('❌ ERROR in exportTimeConfig: ' + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Import configuration from JSON string
 *
 * @param {string} jsonString - JSON configuration string
 * @returns {Object} Result with success flag
 */
function importTimeConfig(jsonString) {
  try {
    Logger.log('\n========== IMPORT TIME CONFIG ==========');

    // Parse JSON
    var config = JSON.parse(jsonString);

    // Update configuration (includes validation)
    var result = updateTimeConfig(config);

    Logger.log('========== IMPORT TIME CONFIG COMPLETE ==========\n');
    return result;

  } catch (error) {
    Logger.log('❌ ERROR in importTimeConfig: ' + error.message);
    Logger.log('========== IMPORT TIME CONFIG FAILED ==========\n');
    return {
      success: false,
      error: error.message
    };
  }
}

// ==================== HELPER FUNCTIONS ====================

/**
 * Parse time string to minutes since midnight
 *
 * @param {string} timeString - Time in HH:MM format
 * @returns {number} Minutes since midnight
 */
function timeToMinutes(timeString) {
  var parts = timeString.split(':');
  var hours = parseInt(parts[0], 10);
  var minutes = parseInt(parts[1], 10);
  return hours * 60 + minutes;
}

/**
 * Convert minutes since midnight to time string
 *
 * @param {number} minutes - Minutes since midnight
 * @returns {string} Time in HH:MM format
 */
function minutesToTime(minutes) {
  var hours = Math.floor(minutes / 60);
  var mins = minutes % 60;
  return pad(hours) + ':' + pad(mins);
}

/**
 * Pad single digit with leading zero
 *
 * @param {number} num - Number to pad
 * @returns {string} Padded string
 */
function pad(num) {
  return num < 10 ? '0' + num : '' + num;
}

/**
 * Get current user email
 *
 * @returns {string} User email
 */
function getCurrentUser() {
  try {
    return Session.getActiveUser().getEmail();
  } catch (error) {
    return 'system';
  }
}

// ==================== TEST FUNCTIONS ====================

/**
 * Test function for time configuration
 */
function test_timeConfig() {
  Logger.log('\n========== TEST: TIME CONFIG ==========');

  // Test 1: Get default config
  var config = getTimeConfig();
  Logger.log('Current config: ' + JSON.stringify(config));

  // Test 2: Update config
  var newConfig = {
    standardStartTime: '08:00',
    graceMinutes: 10
  };
  var updateResult = updateTimeConfig(newConfig);
  Logger.log('Update result: ' + (updateResult.success ? 'SUCCESS' : 'FAILED - ' + updateResult.error));

  // Test 3: Validate invalid config
  var invalidConfig = {
    standardStartTime: '25:00',  // Invalid time
    graceMinutes: -5             // Invalid value
  };
  var validation = validateTimeConfig(invalidConfig);
  Logger.log('Validation result: ' + (validation.valid ? 'VALID' : 'INVALID'));
  Logger.log('Errors: ' + validation.errors.join(', '));

  // Test 4: Reset config
  var resetResult = resetTimeConfig();
  Logger.log('Reset result: ' + (resetResult.success ? 'SUCCESS' : 'FAILED'));

  Logger.log('========== TEST COMPLETE ==========\n');
}
