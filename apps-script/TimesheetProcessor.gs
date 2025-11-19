/**
 * TIMESHEETPROCESSOR.GS - Time Processing Logic
 * HR Payroll System for SA Grinding Wheels & Scorpio Abrasives
 *
 * This module processes raw clock-in data and calculates worked time
 * with configurable rules for grace periods, lunch deductions, etc.
 *
 * Ported from HTML Time Sheet Analyzer
 */

// ==================== PARSE TIME ====================

/**
 * Parse Excel serial date/time to JavaScript Date
 *
 * @param {number|Date|string} value - Excel serial number or date
 * @returns {Date} JavaScript Date object
 */
function parseTime(value) {
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

  throw new Error('Cannot parse time value: ' + value);
}

/**
 * Format time from Date object to HH:MM string
 *
 * @param {Date} date - Date object
 * @returns {string} Time in HH:MM format
 */
function formatTime(date) {
  var hours = date.getHours();
  var minutes = date.getMinutes();
  return pad(hours) + ':' + pad(minutes);
}

/**
 * Pad number with leading zero
 *
 * @param {number} num - Number to pad
 * @returns {string} Padded string
 */
function pad(num) {
  return num < 10 ? '0' + num : '' + num;
}

// ==================== BUFFER APPLICATIONS ====================

/**
 * Apply start buffer (grace period) to clock-in time
 * If employee clocks in early or within grace period, adjust to standard start
 *
 * @param {Date} clockIn - Clock-in time
 * @param {Date} standardStart - Standard start time
 * @param {number} graceMinutes - Grace period in minutes
 * @returns {Object} Adjusted time and applied flag
 */
function applyStartBuffer(clockIn, standardStart, graceMinutes) {
  var clockInMs = clockIn.getTime();
  var standardMs = standardStart.getTime();
  var graceMs = graceMinutes * 60 * 1000;

  // Employee clocked in early → use standard start
  if (clockInMs < standardMs) {
    return {
      adjustedTime: new Date(standardMs),
      bufferApplied: true,
      reason: 'Early clock-in adjusted to standard time'
    };
  }

  // Employee clocked in within grace period → use standard start
  if (clockInMs <= standardMs + graceMs) {
    return {
      adjustedTime: new Date(standardMs),
      bufferApplied: true,
      reason: 'Within grace period, adjusted to standard time'
    };
  }

  // Late clock-in → use actual time
  return {
    adjustedTime: clockIn,
    bufferApplied: false,
    reason: null
  };
}

/**
 * Apply end buffer to clock-out time
 * Allow small buffer for clocking out early
 *
 * @param {Date} clockOut - Clock-out time
 * @param {Date} standardEnd - Standard end time
 * @param {number} bufferMinutes - Buffer in minutes
 * @returns {Object} Adjusted time and applied flag
 */
function applyEndBuffer(clockOut, standardEnd, bufferMinutes) {
  var clockOutMs = clockOut.getTime();
  var standardMs = standardEnd.getTime();
  var bufferMs = bufferMinutes * 60 * 1000;

  // Clocked out after standard → use actual time
  if (clockOutMs > standardMs) {
    return {
      adjustedTime: clockOut,
      bufferApplied: false,
      reason: null
    };
  }

  // Clocked out within buffer → use standard end
  if (clockOutMs >= standardMs - bufferMs) {
    return {
      adjustedTime: new Date(standardMs),
      bufferApplied: true,
      reason: 'Within end buffer, adjusted to standard time'
    };
  }

  // Clocked out too early → use actual time (will show as short day)
  return {
    adjustedTime: clockOut,
    bufferApplied: false,
    reason: 'Early clock-out, no buffer applied'
  };
}

// ==================== LUNCH CALCULATION ====================

/**
 * Calculate lunch break from clock punches
 * Detects gaps between clock-out and clock-in that indicate lunch
 *
 * @param {Array} punches - Array of punch objects {time: Date, type: 'in'|'out'}
 * @param {Object} config - Time configuration
 * @param {boolean} isFriday - Is this a Friday
 * @returns {Object} Lunch details
 */
function calculateLunchBreak(punches, config, isFriday) {
  Logger.log('\n  🍽️  DIAGNOSTIC: Calculating lunch break...');
  Logger.log('    isFriday: ' + isFriday + ', applyLunchOnFriday: ' + config.applyLunchOnFriday);

  // Don't apply lunch on Friday if configured
  if (isFriday && !config.applyLunchOnFriday) {
    Logger.log('    ❌ No lunch on Friday (config)');
    return {
      lunchTaken: false,
      lunchMinutes: 0,
      lunchStart: null,
      lunchEnd: null,
      reason: 'No lunch on Friday (config)'
    };
  }

  // Need at least 2 clock-ins and 2 clock-outs to detect lunch
  var clockIns = punches.filter(function(p) { return p.type === 'in'; });
  var clockOuts = punches.filter(function(p) { return p.type === 'out'; });

  Logger.log('    Total punches: ' + punches.length);
  Logger.log('    Clock-ins: ' + clockIns.length + ', Clock-outs: ' + clockOuts.length);
  Logger.log('    Lunch detection requires: 2+ INs AND 2+ OUTs');

  // Try to detect lunch from multiple punches
  if (clockIns.length >= 2 && clockOuts.length >= 2) {
    Logger.log('    ✓ Sufficient punches for lunch detection');
    Logger.log('    Analyzing gaps between OUT/IN pairs...');

    // Find the largest gap between consecutive out/in pairs
    var largestGap = 0;
    var gapStart = null;
    var gapEnd = null;

    for (var i = 0; i < clockOuts.length; i++) {
      var out = clockOuts[i].time;

      // Find next clock-in after this clock-out
      for (var j = 0; j < clockIns.length; j++) {
        var in_time = clockIns[j].time;

        if (in_time.getTime() > out.getTime()) {
          var gapMinutes = (in_time.getTime() - out.getTime()) / (60 * 1000);

          Logger.log('      Gap #' + (i + 1) + ': ' + formatTime(out) + ' → ' + formatTime(in_time) +
                     ' = ' + Math.round(gapMinutes) + ' minutes');

          if (gapMinutes > largestGap) {
            largestGap = gapMinutes;
            gapStart = out;
            gapEnd = in_time;
          }
          break;
        }
      }
    }

    Logger.log('    Largest gap: ' + Math.round(largestGap) + ' minutes');
    Logger.log('    Valid lunch range: ' + config.minLunchMinutes + '-' + config.maxLunchMinutes + ' minutes');

    // Check if gap is within lunch range
    if (largestGap >= config.minLunchMinutes && largestGap <= config.maxLunchMinutes) {
      Logger.log('    ✅ LUNCH DETECTED: ' + Math.round(largestGap) + ' min gap → deducting ' + config.standardLunchMinutes + ' min');
      return {
        lunchTaken: true,
        lunchMinutes: config.standardLunchMinutes,
        lunchStart: gapStart,
        lunchEnd: gapEnd,
        actualGapMinutes: Math.round(largestGap),
        reason: 'Lunch detected (' + Math.round(largestGap) + ' min gap)'
      };
    } else {
      Logger.log('    ❌ Gap outside valid range - no lunch deducted');
    }
  } else {
    Logger.log('    ❌ INSUFFICIENT PUNCHES: Need 2+ INs AND 2+ OUTs, but got ' +
               clockIns.length + ' INs and ' + clockOuts.length + ' OUTs');
  }

  // Check if employee clocked in late (after late morning threshold)
  if (clockIns.length > 0) {
    var firstClockIn = clockIns[0].time;
    var lateThreshold = parseTimeString(config.lateMorningThreshold);
    lateThreshold.setFullYear(firstClockIn.getFullYear());
    lateThreshold.setMonth(firstClockIn.getMonth());
    lateThreshold.setDate(firstClockIn.getDate());

    Logger.log('    Checking late arrival rule...');
    Logger.log('      First clock-in: ' + formatTime(firstClockIn));
    Logger.log('      Late threshold: ' + config.lateMorningThreshold);

    if (firstClockIn.getTime() >= lateThreshold.getTime()) {
      Logger.log('    ✅ LATE ARRIVAL: Lunch assumed taken, deducting ' + config.standardLunchMinutes + ' min');
      return {
        lunchTaken: true,
        lunchMinutes: config.standardLunchMinutes,
        lunchStart: null,
        lunchEnd: null,
        reason: 'Late arrival - lunch assumed taken'
      };
    }
  }

  // No lunch detected
  Logger.log('    ❌ NO LUNCH DETECTED - No deduction applied');
  return {
    lunchTaken: false,
    lunchMinutes: 0,
    lunchStart: null,
    lunchEnd: null,
    reason: 'No lunch detected'
  };
}

/**
 * Parse time string (HH:MM) to Date object (today)
 *
 * @param {string} timeString - Time in HH:MM format
 * @returns {Date} Date object with today's date and specified time
 */
function parseTimeString(timeString) {
  var parts = timeString.split(':');
  var hours = parseInt(parts[0], 10);
  var minutes = parseInt(parts[1], 10);

  var date = new Date();
  date.setHours(hours);
  date.setMinutes(minutes);
  date.setSeconds(0);
  date.setMilliseconds(0);

  return date;
}

// ==================== BATHROOM TIME CALCULATION ====================

/**
 * Calculate bathroom time from bathroom entry/exit punches
 *
 * @param {Array} bathroomPunches - Array of bathroom punch objects
 * @param {Object} config - Time configuration
 * @returns {Object} Bathroom time details
 */
function calculateBathroomTime(bathroomPunches, config) {
  if (!bathroomPunches || bathroomPunches.length === 0) {
    return {
      totalMinutes: 0,
      breaks: [],
      warnings: []
    };
  }

  var breaks = [];
  var totalMinutes = 0;
  var warnings = [];

  // Sort punches by time
  bathroomPunches.sort(function(a, b) {
    return a.time.getTime() - b.time.getTime();
  });

  // Pair entry/exit punches
  for (var i = 0; i < bathroomPunches.length; i += 2) {
    if (i + 1 < bathroomPunches.length) {
      var entry = bathroomPunches[i];
      var exit = bathroomPunches[i + 1];

      var duration = (exit.time.getTime() - entry.time.getTime()) / (60 * 1000);
      totalMinutes += duration;

      var breakInfo = {
        entry: entry.time,
        exit: exit.time,
        minutes: Math.round(duration)
      };

      // Check for long bathroom break
      if (duration > config.longBathroomThreshold) {
        breakInfo.warning = 'Long bathroom break (' + Math.round(duration) + ' min)';
        warnings.push(breakInfo.warning);
      }

      breaks.push(breakInfo);
    }
  }

  // Check daily threshold
  if (totalMinutes > config.dailyBathroomThreshold) {
    warnings.push('Daily bathroom time exceeds threshold (' + Math.round(totalMinutes) + ' min)');
  }

  return {
    totalMinutes: Math.round(totalMinutes),
    breaks: breaks,
    warnings: warnings
  };
}

// ==================== MAIN PROCESSING FUNCTION ====================

/**
 * Process clock data for an employee for one week
 *
 * @param {Array} clockData - Array of clock punch records
 * @param {Object} config - Time configuration
 * @returns {Object} Processed timesheet data
 */
function processClockData(clockData, config) {
  try {
    Logger.log('\n========== PROCESS CLOCK DATA ==========');
    Logger.log('Processing ' + clockData.length + ' clock records');

    // Group punches by date
    var punchesByDate = {};

    // CRITICAL: Sort clock data by date and time FIRST before assigning in/out types
    // The alternating pattern only works if punches are processed in chronological order
    clockData.sort(function(a, b) {
      var timeA = parseTime(a.PUNCH_TIME);
      var timeB = parseTime(b.PUNCH_TIME);
      return timeA.getTime() - timeB.getTime();
    });

    Logger.log('\n🔍 DIAGNOSTIC: Grouping and classifying punches (sorted by time)...');
    for (var i = 0; i < clockData.length; i++) {
      var record = clockData[i];
      var punchTime = parseTime(record.PUNCH_TIME);
      var dateKey = formatDate(punchTime);

      if (!punchesByDate[dateKey]) {
        punchesByDate[dateKey] = {
          date: dateKey,
          clockInPunches: [],
          bathroomPunches: []
        };
      }

      var deviceName = record.DEVICE_NAME || '';

      if (deviceName.toLowerCase().indexOf('bathroom') >= 0) {
        punchesByDate[dateKey].bathroomPunches.push({
          time: punchTime,
          device: deviceName
        });
      } else {
        // Determine if clock-in or clock-out
        var type = 'unknown';
        var currentCount = punchesByDate[dateKey].clockInPunches.length;
        var deviceLower = deviceName.toLowerCase();

        // Check for SPECIFIC device names that indicate in/out
        // Must be exact matches like "Clock In" or "Clock Out", not substrings
        if (deviceLower === 'clock in' || deviceLower === 'clock-in' || deviceLower === 'clockin') {
          type = 'in';
        } else if (deviceLower === 'clock out' || deviceLower === 'clock-out' || deviceLower === 'clockout') {
          type = 'out';
        } else {
          // For any other device (like "Main Unit"), use alternating pattern
          // 1st scan = in, 2nd = out, 3rd = in, 4th = out
          type = currentCount % 2 === 0 ? 'in' : 'out';
        }

        Logger.log('  📌 ' + dateKey + ' @ ' + formatTime(punchTime) +
                   ' | Device: "' + deviceName + '" | Count: ' + currentCount +
                   ' | Type: ' + type.toUpperCase());

        punchesByDate[dateKey].clockInPunches.push({
          time: punchTime,
          type: type,
          device: deviceName
        });
      }
    }

    // Log summary for each day
    Logger.log('\n📊 DIAGNOSTIC: Daily punch summary:');
    for (var dateKey in punchesByDate) {
      var dayData = punchesByDate[dateKey];
      var inCount = dayData.clockInPunches.filter(function(p) { return p.type === 'in'; }).length;
      var outCount = dayData.clockInPunches.filter(function(p) { return p.type === 'out'; }).length;
      Logger.log('  ' + dateKey + ': ' + inCount + ' INs, ' + outCount + ' OUTs, ' +
                 dayData.bathroomPunches.length + ' bathroom punches');
    }

    // Process each day
    var dailyBreakdown = [];
    var totalMinutes = 0;
    var totalLunchDeduction = 0;
    var totalBathroomTime = 0;
    var appliedRules = [];
    var allWarnings = [];

    for (var dateKey in punchesByDate) {
      var dayData = punchesByDate[dateKey];
      var dayResult = processDayData(dayData, config);

      dailyBreakdown.push(dayResult);
      totalMinutes += dayResult.totalMinutes;
      totalLunchDeduction += dayResult.lunchMinutes;
      totalBathroomTime += dayResult.bathroomMinutes;

      if (dayResult.appliedRules) {
        dayResult.appliedRules.forEach(function(rule) {
          if (appliedRules.indexOf(rule) === -1) {
            appliedRules.push(rule);
          }
        });
      }

      if (dayResult.warnings) {
        allWarnings = allWarnings.concat(dayResult.warnings);
      }
    }

    // Calculate hours and minutes
    var totalHours = Math.floor(totalMinutes / 60);
    var remainingMinutes = totalMinutes % 60;

    // Log weekly summary
    Logger.log('\n📊 DIAGNOSTIC: WEEKLY SUMMARY');
    Logger.log('============================================');
    dailyBreakdown.forEach(function(day) {
      var dayHours = Math.floor(day.totalMinutes / 60);
      var dayMins = day.totalMinutes % 60;
      Logger.log('  ' + day.date + ': ' + dayHours + 'h ' + dayMins + 'm' +
                 (day.lunchMinutes > 0 ? ' (lunch: -' + day.lunchMinutes + 'm)' : ' (NO LUNCH)'));
    });
    Logger.log('============================================');
    Logger.log('  TOTAL PAID TIME: ' + totalHours + 'h ' + remainingMinutes + 'm');
    Logger.log('  Total lunch deducted: ' + totalLunchDeduction + ' minutes');
    Logger.log('============================================');

    var result = {
      success: true,
      data: {
        calculatedTotalHours: totalHours,
        calculatedTotalMinutes: remainingMinutes,
        rawTotalMinutes: totalMinutes,
        lunchDeductionMinutes: totalLunchDeduction,
        bathroomTimeMinutes: totalBathroomTime,
        dailyBreakdown: dailyBreakdown,
        appliedRules: appliedRules,
        warnings: allWarnings
      }
    };

    Logger.log('✅ Processing complete: ' + totalHours + 'h ' + remainingMinutes + 'm');
    Logger.log('========== PROCESS CLOCK DATA COMPLETE ==========\n');

    return result;

  } catch (error) {
    Logger.log('❌ ERROR in processClockData: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== PROCESS CLOCK DATA FAILED ==========\n');
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Process clock data for a single day
 *
 * @param {Object} dayData - Day's punch data
 * @param {Object} config - Time configuration
 * @returns {Object} Day's processed data
 */
function processDayData(dayData, config) {
  var date = new Date(dayData.date);
  var isFriday = date.getDay() === 5;

  Logger.log('\n📅 DIAGNOSTIC: Processing day: ' + dayData.date + (isFriday ? ' (FRIDAY)' : ''));

  var warnings = [];
  var appliedRules = [];

  // Sort punches by time
  dayData.clockInPunches.sort(function(a, b) {
    return a.time.getTime() - b.time.getTime();
  });

  Logger.log('  Sorted punches (' + dayData.clockInPunches.length + ' total):');
  for (var i = 0; i < dayData.clockInPunches.length; i++) {
    Logger.log('    ' + (i + 1) + '. ' + formatTime(dayData.clockInPunches[i].time) +
               ' (' + dayData.clockInPunches[i].type.toUpperCase() + ')');
  }

  // Get first clock-in and last clock-out
  var firstIn = null;
  var lastOut = null;

  for (var i = 0; i < dayData.clockInPunches.length; i++) {
    if (dayData.clockInPunches[i].type === 'in' && !firstIn) {
      firstIn = dayData.clockInPunches[i].time;
    }
    if (dayData.clockInPunches[i].type === 'out') {
      lastOut = dayData.clockInPunches[i].time;
    }
  }

  Logger.log('  First IN: ' + (firstIn ? formatTime(firstIn) : 'NONE'));
  Logger.log('  Last OUT: ' + (lastOut ? formatTime(lastOut) : 'NONE'));

  // If incomplete day and projectIncompleteDays is enabled
  if (config.projectIncompleteDays && (!firstIn || !lastOut)) {
    warnings.push('Incomplete day - missing clock-in or clock-out');

    // Use defaults
    if (!firstIn) {
      firstIn = parseTimeString(config.standardStartTime);
      firstIn.setFullYear(date.getFullYear());
      firstIn.setMonth(date.getMonth());
      firstIn.setDate(date.getDate());
      warnings.push('No clock-in found, using standard start time');
    }

    if (!lastOut) {
      var endTime = isFriday ? config.fridayEndTime : config.standardEndTime;
      lastOut = parseTimeString(endTime);
      lastOut.setFullYear(date.getFullYear());
      lastOut.setMonth(date.getMonth());
      lastOut.setDate(date.getDate());
      warnings.push('No clock-out found, using standard end time');
    }

    appliedRules.push('projectIncompleteDays');
  }

  if (!firstIn || !lastOut) {
    return {
      date: dayData.date,
      clockIn: null,
      clockOut: null,
      adjustedIn: null,
      adjustedOut: null,
      lunchTaken: false,
      lunchMinutes: 0,
      bathroomMinutes: 0,
      totalMinutes: 0,
      warnings: warnings,
      appliedRules: appliedRules
    };
  }

  // Apply start buffer
  var standardStart = parseTimeString(config.standardStartTime);
  standardStart.setFullYear(date.getFullYear());
  standardStart.setMonth(date.getMonth());
  standardStart.setDate(date.getDate());

  var startResult = applyStartBuffer(firstIn, standardStart, config.graceMinutes);
  var adjustedIn = startResult.adjustedTime;
  if (startResult.bufferApplied) {
    appliedRules.push('graceBufferApplied');
  }

  // Apply end buffer
  var endTime = isFriday ? config.fridayEndTime : config.standardEndTime;
  var standardEnd = parseTimeString(endTime);
  standardEnd.setFullYear(date.getFullYear());
  standardEnd.setMonth(date.getMonth());
  standardEnd.setDate(date.getDate());

  var endResult = applyEndBuffer(lastOut, standardEnd, config.endBufferMinutes);
  var adjustedOut = endResult.adjustedTime;
  if (endResult.bufferApplied) {
    appliedRules.push('endBufferApplied');
  }

  Logger.log('\n  ⏰ DIAGNOSTIC: Time adjustments...');
  Logger.log('    Adjusted IN: ' + formatTime(adjustedIn));
  Logger.log('    Adjusted OUT: ' + formatTime(adjustedOut));

  // Calculate lunch
  var lunchResult = calculateLunchBreak(dayData.clockInPunches, config, isFriday);
  if (lunchResult.lunchTaken) {
    appliedRules.push('lunchDeducted');
  }

  // Calculate bathroom time
  var bathroomResult = calculateBathroomTime(dayData.bathroomPunches, config);
  if (bathroomResult.warnings.length > 0) {
    warnings = warnings.concat(bathroomResult.warnings);
  }

  // Calculate total minutes
  var workedMs = adjustedOut.getTime() - adjustedIn.getTime();
  var workedMinutes = workedMs / (60 * 1000);

  Logger.log('\n  💰 DIAGNOSTIC: Final calculation...');
  Logger.log('    Adjusted OUT - Adjusted IN = Worked minutes');
  Logger.log('    ' + formatTime(adjustedOut) + ' - ' + formatTime(adjustedIn) + ' = ' +
             Math.round(workedMinutes) + ' minutes (' + (Math.round(workedMinutes / 60 * 10) / 10) + ' hours)');

  // Handle invalid data: clock-out before clock-in (results in negative time)
  if (workedMinutes < 0) {
    warnings.push('Invalid data: Clock-out (' + formatTime(adjustedOut) + ') before clock-in (' + formatTime(adjustedIn) + ')');
    workedMinutes = 0;
  }

  var totalMinutes = workedMinutes - lunchResult.lunchMinutes;

  Logger.log('    Lunch deduction: -' + lunchResult.lunchMinutes + ' minutes');
  Logger.log('    PAID TIME: ' + Math.round(totalMinutes) + ' minutes (' +
             Math.floor(totalMinutes / 60) + 'h ' + Math.round(totalMinutes % 60) + 'm)');

  // Ensure total doesn't go negative after lunch deduction
  if (totalMinutes < 0) {
    totalMinutes = 0;
  }

  return {
    date: dayData.date,
    clockIn: formatTime(firstIn),
    clockOut: formatTime(lastOut),
    adjustedIn: formatTime(adjustedIn),
    adjustedOut: formatTime(adjustedOut),
    lunchTaken: lunchResult.lunchTaken,
    lunchMinutes: lunchResult.lunchMinutes,
    bathroomMinutes: bathroomResult.totalMinutes,
    bathroomBreaks: bathroomResult.breaks,
    totalMinutes: Math.round(totalMinutes),
    warnings: warnings,
    appliedRules: appliedRules
  };
}

/**
 * Format date to YYYY-MM-DD
 *
 * @param {Date} date - Date to format
 * @returns {string} Formatted date string
 */
function formatDate(date) {
  var year = date.getFullYear();
  var month = pad(date.getMonth() + 1);
  var day = pad(date.getDate());
  return year + '-' + month + '-' + day;
}

// ==================== VALIDATION ====================

/**
 * Validate employee clock references against employee table
 *
 * @param {Array} clockRefs - Array of clock reference numbers from import
 * @returns {Object} Validation result with matched and unmatched refs
 */
function validateClockRefs(clockRefs) {
  try {
    Logger.log('\n========== VALIDATE CLOCK REFS ==========');
    Logger.log('Validating ' + clockRefs.length + ' clock references');

    var sheets = getSheets();
    var empSheet = sheets.empdetails;

    if (!empSheet) {
      throw new Error('Employee details sheet not found');
    }

    var empData = empSheet.getDataRange().getValues();
    var headers = empData[0];
    var clockRefCol = findColumnIndex(headers, 'ClockInRef');

    if (clockRefCol === -1) {
      throw new Error('ClockInRef column not found in employee table');
    }

    // Build set of valid clock references
    var validRefs = {};
    for (var i = 1; i < empData.length; i++) {
      var clockRef = String(empData[i][clockRefCol]).trim();
      if (clockRef) {
        validRefs[clockRef] = true;
      }
    }

    // Check each imported clock ref
    var matched = [];
    var unmatched = [];

    clockRefs.forEach(function(ref) {
      var refStr = String(ref).trim();
      if (validRefs[refStr]) {
        matched.push(refStr);
      } else {
        unmatched.push(refStr);
      }
    });

    Logger.log('✅ Validation complete: ' + matched.length + ' matched, ' + unmatched.length + ' unmatched');
    Logger.log('========== VALIDATE CLOCK REFS COMPLETE ==========\n');

    return {
      success: true,
      data: {
        matched: matched,
        unmatched: unmatched,
        totalValidated: clockRefs.length
      }
    };

  } catch (error) {
    Logger.log('❌ ERROR in validateClockRefs: ' + error.message);
    Logger.log('========== VALIDATE CLOCK REFS FAILED ==========\n');
    return {
      success: false,
      error: error.message
    };
  }
}

// ==================== TEST FUNCTIONS ====================

/**
 * Test function for time processing
 */
function test_timesheetProcessor() {
  Logger.log('\n========== TEST: TIMESHEET PROCESSOR ==========');

  // Test data
  var testClockData = [
    {
      PUNCH_TIME: new Date(2025, 10, 17, 7, 25, 0),
      DEVICE_NAME: 'Clock In',
      EMPLOYEE_CLOCK_REF: '1234'
    },
    {
      PUNCH_TIME: new Date(2025, 10, 17, 16, 35, 0),
      DEVICE_NAME: 'Clock Out',
      EMPLOYEE_CLOCK_REF: '1234'
    }
  ];

  var configResult = getTimeConfig();
  if (!configResult.success) {
    Logger.log('❌ Failed to get config');
    return;
  }
  var config = configResult.data;

  // Test processing
  var result = processClockData(testClockData, config);
  Logger.log('Process result: ' + (result.success ? 'SUCCESS' : 'FAILED - ' + result.error));
  if (result.success) {
    Logger.log('Total hours: ' + result.data.calculatedTotalHours);
    Logger.log('Total minutes: ' + result.data.calculatedTotalMinutes);
  }

  Logger.log('========== TEST COMPLETE ==========\n');
}
