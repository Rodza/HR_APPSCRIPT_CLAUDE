/**
 * TIMESHEETPROCESSOR.GS - Time Processing Logic (Revised Rules)
 * HR Payroll System for SA Grinding Wheels & Scorpio Abrasives
 *
 * This module processes raw clock-in data according to the revised clock rules specification.
 *
 * CLOCK IN/OUT RULES:
 * - Monday-Thursday: 4 clocks (Morning IN, Lunch OUT, Lunch RETURN, Afternoon OUT)
 * - Friday: 2 clocks (Morning IN, Afternoon OUT)
 * - Grace periods apply to Clock 1 (5 min) and Clock 2 (2 min)
 * - Lunch deducted ONLY when Clock 3 is present
 * - Bathroom time counted only during work periods
 * - Duplicate detection: 2 min for main clocks, 60 sec for bathroom
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

/**
 * Parse time string (HH:MM) to Date object (today)
 *
 * @param {string} timeString - Time in HH:MM format
 * @param {Date} referenceDate - Reference date to use
 * @returns {Date} Date object with reference date and specified time
 */
function parseTimeString(timeString, referenceDate) {
  var parts = timeString.split(':');
  var hours = parseInt(parts[0], 10);
  var minutes = parseInt(parts[1], 10);

  var date = new Date(referenceDate || new Date());
  date.setHours(hours);
  date.setMinutes(minutes);
  date.setSeconds(0);
  date.setMilliseconds(0);

  return date;
}

/**
 * Get time in minutes since midnight
 *
 * @param {Date} date - Date object
 * @returns {number} Minutes since midnight
 */
function getMinutesSinceMidnight(date) {
  return date.getHours() * 60 + date.getMinutes();
}

// ==================== CLOCK CLASSIFICATION ====================

/**
 * Classify clock punches into Clock 1, 2, 3, 4 based on time windows
 *
 * @param {Array} punches - Array of punch objects {time: Date, type: 'in'|'out', device: string}
 * @param {Object} config - Time configuration
 * @param {boolean} isFriday - Is this a Friday
 * @returns {Object} Classified clocks {clock1, clock2, clock3, clock4}
 */
function classifyClockPunches(punches, config, isFriday) {
  Logger.log('\n  🔍 CLASSIFYING CLOCK PUNCHES...');

  var result = {
    clock1: null,  // Morning IN
    clock2: null,  // Lunch OUT
    clock3: null,  // Lunch RETURN
    clock4: null   // Afternoon OUT
  };

  if (punches.length === 0) {
    return result;
  }

  // Parse time windows
  var clock1MaxMin = timeToMinutes(config.clock1MaxTime);
  var clock2StartMin = timeToMinutes(config.clock2WindowStart);
  var clock2EndMin = timeToMinutes(config.clock2WindowEnd);
  var clock3StartMin = timeToMinutes(config.clock3WindowStart);
  var clock3EndMin = timeToMinutes(config.clock3WindowEnd);
  var clock4MinMin = timeToMinutes(config.clock4MinTime);

  // Separate IN and OUT punches
  var inPunches = punches.filter(function(p) { return p.type === 'in'; });
  var outPunches = punches.filter(function(p) { return p.type === 'out'; });

  Logger.log('    IN punches: ' + inPunches.length + ', OUT punches: ' + outPunches.length);

  if (isFriday) {
    // Friday: Clock 1 = first IN, Clock 2 = last OUT
    if (inPunches.length > 0) {
      result.clock1 = inPunches[0];
      Logger.log('    Clock 1 (Friday IN): ' + formatTime(result.clock1.time));
    }
    if (outPunches.length > 0) {
      result.clock2 = outPunches[outPunches.length - 1];
      Logger.log('    Clock 2 (Friday OUT): ' + formatTime(result.clock2.time));
    }
  } else {
    // Monday-Thursday: Classify by time windows

    // Clock 1: First IN before clock1MaxTime
    for (var i = 0; i < inPunches.length; i++) {
      var minutes = getMinutesSinceMidnight(inPunches[i].time);
      if (minutes < clock1MaxMin) {
        result.clock1 = inPunches[i];
        Logger.log('    Clock 1 (Morning IN): ' + formatTime(result.clock1.time));
        break;
      }
    }

    // Clock 2: First OUT in window 12:00-12:10
    for (var i = 0; i < outPunches.length; i++) {
      var minutes = getMinutesSinceMidnight(outPunches[i].time);
      if (minutes >= clock2StartMin && minutes <= clock2EndMin) {
        result.clock2 = outPunches[i];
        Logger.log('    Clock 2 (Lunch OUT): ' + formatTime(result.clock2.time));
        break;
      }
    }

    // Clock 3: First IN in window 12:10-13:00
    for (var i = 0; i < inPunches.length; i++) {
      var minutes = getMinutesSinceMidnight(inPunches[i].time);
      if (minutes >= clock3StartMin && minutes <= clock3EndMin) {
        result.clock3 = inPunches[i];
        Logger.log('    Clock 3 (Lunch RETURN): ' + formatTime(result.clock3.time));
        break;
      }
    }

    // Clock 4: Last OUT after clock4MinTime
    for (var i = outPunches.length - 1; i >= 0; i--) {
      var minutes = getMinutesSinceMidnight(outPunches[i].time);
      if (minutes >= clock4MinMin) {
        result.clock4 = outPunches[i];
        Logger.log('    Clock 4 (Afternoon OUT): ' + formatTime(result.clock4.time));
        break;
      }
    }
  }

  return result;
}

/**
 * Convert time string to minutes since midnight
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

// ==================== GRACE PERIOD ADJUSTMENTS ====================

/**
 * Apply grace period adjustments according to revised rules
 *
 * @param {Object} clocks - Classified clock punches
 * @param {Object} config - Time configuration
 * @param {Date} referenceDate - Reference date
 * @param {boolean} isFriday - Is this a Friday
 * @returns {Object} Adjusted clocks and flags
 */
function applyGracePeriods(clocks, config, referenceDate, isFriday) {
  Logger.log('\n  ⏰ APPLYING GRACE PERIODS...');

  var adjusted = {
    clock1: clocks.clock1 ? clocks.clock1.time : null,
    clock2: clocks.clock2 ? clocks.clock2.time : null,
    clock3: clocks.clock3 ? clocks.clock3.time : null,
    clock4: clocks.clock4 ? clocks.clock4.time : null
  };

  var flags = [];

  // Clock 1: Grace period 07:30-07:35
  if (clocks.clock1) {
    var standardStart = parseTimeString(config.standardStartTime, referenceDate);
    var lateAfter = parseTimeString(config.flagLateAfter, referenceDate);
    var clock1Time = clocks.clock1.time;

    if (clock1Time.getTime() < standardStart.getTime()) {
      // Before 07:30: Cap at 07:30
      adjusted.clock1 = standardStart;
      Logger.log('    Clock 1: ' + formatTime(clock1Time) + ' → ' + formatTime(adjusted.clock1) + ' (capped at standard time)');
    } else if (clock1Time.getTime() <= lateAfter.getTime()) {
      // Between 07:30 and 07:35: Adjust to 07:30
      adjusted.clock1 = standardStart;
      Logger.log('    Clock 1: ' + formatTime(clock1Time) + ' → ' + formatTime(adjusted.clock1) + ' (grace period applied)');
    } else {
      // After 07:35: Use actual time and flag as late
      adjusted.clock1 = clock1Time;
      flags.push('Late Clock In (' + formatTime(clock1Time) + ')');
      Logger.log('    Clock 1: ' + formatTime(clock1Time) + ' (LATE - no adjustment)');
    }
  }

  // Clock 2: Grace period 12:00-12:02
  if (clocks.clock2) {
    var lunchStart = parseTimeString(config.lunchStart, referenceDate);
    var lunchGraceEnd = new Date(lunchStart.getTime() + config.lunchOutGraceMinutes * 60 * 1000);
    var clock2Time = clocks.clock2.time;

    if (clock2Time.getTime() >= lunchStart.getTime() && clock2Time.getTime() <= lunchGraceEnd.getTime()) {
      // Between 12:00 and 12:02: Adjust to 12:00
      adjusted.clock2 = lunchStart;
      Logger.log('    Clock 2: ' + formatTime(clock2Time) + ' → ' + formatTime(adjusted.clock2) + ' (grace period applied)');
    } else {
      adjusted.clock2 = clock2Time;
      Logger.log('    Clock 2: ' + formatTime(clock2Time) + ' (no adjustment)');
    }
  }

  // Clock 3: After 12:30 flag as late
  if (clocks.clock3) {
    var lunchEnd = parseTimeString(config.lunchEnd, referenceDate);
    var clock3Time = clocks.clock3.time;

    adjusted.clock3 = clock3Time;

    if (clock3Time.getTime() > lunchEnd.getTime()) {
      flags.push('Late lunch return - review required (' + formatTime(clock3Time) + ')');
      Logger.log('    Clock 3: ' + formatTime(clock3Time) + ' (LATE RETURN)');
    } else {
      Logger.log('    Clock 3: ' + formatTime(clock3Time) + ' (no adjustment)');
    }
  }

  // Clock 4: After 16:30 flag as overtime
  if (clocks.clock4 && !isFriday) {
    var standardEnd = parseTimeString(config.standardEndTime, referenceDate);
    var clock4Time = clocks.clock4.time;

    adjusted.clock4 = clock4Time;

    if (clock4Time.getTime() > standardEnd.getTime()) {
      flags.push('Overtime - manual review (' + formatTime(clock4Time) + ')');
      Logger.log('    Clock 4: ' + formatTime(clock4Time) + ' (OVERTIME)');
    } else {
      Logger.log('    Clock 4: ' + formatTime(clock4Time) + ' (no adjustment)');
    }
  }

  // Friday Clock 2: Before 13:00 use actual time
  if (isFriday && clocks.clock2) {
    var fridayEnd = parseTimeString(config.fridayEndTime, referenceDate);
    adjusted.clock2 = clocks.clock2.time;
    Logger.log('    Clock 2 (Friday): ' + formatTime(adjusted.clock2) + ' (actual time)');
  }

  return {
    adjusted: adjusted,
    flags: flags
  };
}

// ==================== PAID TIME CALCULATION ====================

/**
 * Calculate paid time based on punch scenario (revised rules)
 *
 * @param {Object} clocks - Classified clock punches
 * @param {Object} adjusted - Adjusted clock times
 * @param {Object} config - Time configuration
 * @param {boolean} isFriday - Is this a Friday
 * @returns {Object} Paid time calculation results
 */
function calculatePaidTime(clocks, adjusted, config, isFriday) {
  Logger.log('\n  💰 CALCULATING PAID TIME...');

  var paidMinutes = 0;
  var lunchMinutes = 0;
  var scenario = '';
  var flags = [];

  // Determine which clocks are present
  var hasC1 = clocks.clock1 !== null;
  var hasC2 = clocks.clock2 !== null;
  var hasC3 = clocks.clock3 !== null;
  var hasC4 = clocks.clock4 !== null;

  if (isFriday) {
    // ========== FRIDAY SCENARIOS ==========
    if (hasC1 && hasC2) {
      // Scenario: 1, 2 → Clock 2 - Clock 1
      paidMinutes = (adjusted.clock2.getTime() - adjusted.clock1.getTime()) / (60 * 1000);
      scenario = 'Friday: Complete day (Clock 2 - Clock 1)';
      Logger.log('    Scenario: Friday complete (1,2)');
      Logger.log('    ' + formatTime(adjusted.clock2) + ' - ' + formatTime(adjusted.clock1) + ' = ' + Math.round(paidMinutes) + ' min');
    } else if (hasC1 && !hasC2) {
      // Scenario: 1 only → 0 paid time
      paidMinutes = 0;
      scenario = 'Friday: Only morning scan';
      flags.push('Missing Friday out - manual adjustment required');
      Logger.log('    Scenario: Missing Friday out (1 only)');
    } else if (!hasC1 && hasC2) {
      // Scenario: 2 only → 0 paid time
      paidMinutes = 0;
      scenario = 'Friday: Only afternoon scan';
      flags.push('Missing Friday in - manual adjustment required');
      Logger.log('    Scenario: Missing Friday in (2 only)');
    } else {
      // No punches
      paidMinutes = 0;
      scenario = 'Friday: No punches';
      flags.push('No Friday punches recorded');
      Logger.log('    Scenario: No Friday punches');
    }
  } else {
    // ========== MONDAY-THURSDAY SCENARIOS ==========

    if (hasC1 && hasC2 && hasC3 && hasC4) {
      // Scenario: 1, 2, 3, 4 → (Clock 4 - Clock 1) - 30 min
      var totalMinutes = (adjusted.clock4.getTime() - adjusted.clock1.getTime()) / (60 * 1000);
      lunchMinutes = config.standardLunchMinutes;
      paidMinutes = totalMinutes - lunchMinutes;
      scenario = 'Complete day (1,2,3,4)';
      Logger.log('    Scenario: Complete day (1,2,3,4)');
      Logger.log('    (' + formatTime(adjusted.clock4) + ' - ' + formatTime(adjusted.clock1) + ') - ' + lunchMinutes + ' min = ' + Math.round(paidMinutes) + ' min');

    } else if (hasC1 && hasC3 && hasC4 && !hasC2) {
      // Scenario: 1, 3, 4 → (Clock 4 - Clock 1) - 30 min
      var totalMinutes = (adjusted.clock4.getTime() - adjusted.clock1.getTime()) / (60 * 1000);
      lunchMinutes = config.standardLunchMinutes;
      paidMinutes = totalMinutes - lunchMinutes;
      scenario = 'Missing lunch out (1,3,4)';
      flags.push('Missing lunch out scan - lunch deducted');
      Logger.log('    Scenario: Missing lunch out (1,3,4)');
      Logger.log('    (' + formatTime(adjusted.clock4) + ' - ' + formatTime(adjusted.clock1) + ') - ' + lunchMinutes + ' min = ' + Math.round(paidMinutes) + ' min');

    } else if (hasC1 && hasC4 && !hasC2 && !hasC3) {
      // Scenario: 1, 4 only → 0 (no lunch recorded)
      paidMinutes = 0;
      lunchMinutes = 0;
      scenario = 'No lunch recorded (1,4)';
      flags.push('No lunch recorded - investigation required');
      Logger.log('    Scenario: No lunch recorded (1,4 only) → 0 paid time');

    } else if (hasC1 && hasC2 && hasC4 && !hasC3) {
      // Scenario: 1, 2, 4 → 0 (missing lunch return)
      paidMinutes = 0;
      lunchMinutes = 0;
      scenario = 'Missing lunch return (1,2,4)';
      flags.push('Missing lunch return - manual adjustment required');
      Logger.log('    Scenario: Missing lunch return (1,2,4) → 0 paid time');

    } else if (hasC2 && hasC3 && hasC4 && !hasC1) {
      // Scenario: 2, 3, 4 → 0 (no morning scan)
      paidMinutes = 0;
      lunchMinutes = 0;
      scenario = 'No morning scan (2,3,4)';
      flags.push('No morning scan - manual adjustment required');
      Logger.log('    Scenario: No morning scan (2,3,4) → 0 paid time');

    } else if (hasC1 && hasC2 && hasC3 && !hasC4) {
      // Scenario: 1, 2, 3 → 0 (no afternoon out)
      paidMinutes = 0;
      lunchMinutes = 0;
      scenario = 'No afternoon out (1,2,3)';
      flags.push('No afternoon out - manual adjustment required');
      Logger.log('    Scenario: No afternoon out (1,2,3) → 0 paid time');

    } else if (hasC1 && !hasC2 && !hasC3 && !hasC4) {
      // Scenario: 1 only → 0
      paidMinutes = 0;
      lunchMinutes = 0;
      scenario = 'Only morning scan (1)';
      flags.push('Only morning scan - manual adjustment required');
      Logger.log('    Scenario: Only morning scan (1) → 0 paid time');

    } else if (!hasC1 && !hasC2 && !hasC3 && hasC4) {
      // Scenario: 4 only → 0
      paidMinutes = 0;
      lunchMinutes = 0;
      scenario = 'Only afternoon scan (4)';
      flags.push('Only afternoon scan - manual adjustment required');
      Logger.log('    Scenario: Only afternoon scan (4) → 0 paid time');

    } else if (!hasC1 && hasC2 && hasC3 && !hasC4) {
      // Scenario: 2, 3 only → 0
      paidMinutes = 0;
      lunchMinutes = 0;
      scenario = 'Only lunch scans (2,3)';
      flags.push('Only lunch scans - manual adjustment required');
      Logger.log('    Scenario: Only lunch scans (2,3) → 0 paid time');

    } else if (!hasC1 && hasC2 && !hasC3 && !hasC4) {
      // Scenario: 2 only → 0
      paidMinutes = 0;
      lunchMinutes = 0;
      scenario = 'Only lunch out scan (2)';
      flags.push('Only lunch out scan - manual adjustment required');
      Logger.log('    Scenario: Only lunch out (2) → 0 paid time');

    } else if (!hasC1 && !hasC2 && hasC3 && !hasC4) {
      // Scenario: 3 only → 0
      paidMinutes = 0;
      lunchMinutes = 0;
      scenario = 'Only lunch return scan (3)';
      flags.push('Only lunch return scan - manual adjustment required');
      Logger.log('    Scenario: Only lunch return (3) → 0 paid time');

    } else {
      // Any other combination → 0
      paidMinutes = 0;
      lunchMinutes = 0;
      scenario = 'Irregular punch pattern';
      flags.push('Irregular punch pattern - manual adjustment required');
      Logger.log('    Scenario: Irregular pattern → 0 paid time');
    }
  }

  // Ensure paid time doesn't go negative
  if (paidMinutes < 0) {
    paidMinutes = 0;
  }

  return {
    paidMinutes: Math.round(paidMinutes),
    lunchMinutes: lunchMinutes,
    scenario: scenario,
    flags: flags
  };
}

// ==================== BATHROOM TIME CALCULATION ====================

/**
 * Calculate bathroom time from bathroom entry/exit punches (revised rules)
 *
 * @param {Array} bathroomPunches - Array of bathroom punch objects
 * @param {Object} config - Time configuration
 * @param {Object} clocks - Classified clock punches
 * @param {Object} adjusted - Adjusted clock times
 * @param {boolean} isFriday - Whether this is a Friday
 * @returns {Object} Bathroom time details
 */
function calculateBathroomTime(bathroomPunches, config, clocks, adjusted, isFriday) {
  if (!bathroomPunches || bathroomPunches.length === 0) {
    return {
      totalMinutes: 0,
      breaks: [],
      warnings: [],
      unpairedEntries: [],
      unpairedExits: []
    };
  }

  Logger.log('\n  🚻 CALCULATING BATHROOM TIME...');

  var breaks = [];
  var totalMinutes = 0;
  var warnings = [];
  var unpairedEntries = [];
  var unpairedExits = [];

  // Determine work periods
  var workPeriods = [];

  if (isFriday) {
    // Friday: Clock 1 to Clock 2
    if (adjusted.clock1 && adjusted.clock2) {
      workPeriods.push({
        start: adjusted.clock1,
        end: adjusted.clock2,
        name: 'Friday full day'
      });
    }
  } else {
    // Mon-Thu: Morning (Clock 1 to Clock 2) and Afternoon (Clock 3 to Clock 4)
    if (adjusted.clock1 && adjusted.clock2) {
      workPeriods.push({
        start: adjusted.clock1,
        end: adjusted.clock2,
        name: 'Morning'
      });
    }
    if (adjusted.clock3 && adjusted.clock4) {
      workPeriods.push({
        start: adjusted.clock3,
        end: adjusted.clock4,
        name: 'Afternoon'
      });
    }
  }

  Logger.log('    Work periods: ' + workPeriods.length);
  workPeriods.forEach(function(p, i) {
    Logger.log('      ' + (i + 1) + '. ' + p.name + ': ' + formatTime(p.start) + ' - ' + formatTime(p.end));
  });

  // Helper function to check if time is within work periods
  function isWithinWorkPeriods(time) {
    for (var i = 0; i < workPeriods.length; i++) {
      var period = workPeriods[i];
      if (time.getTime() >= period.start.getTime() && time.getTime() <= period.end.getTime()) {
        return true;
      }
    }
    return false;
  }

  // Sort bathroom punches by time
  bathroomPunches.sort(function(a, b) {
    return a.time.getTime() - b.time.getTime();
  });

  // Separate entries and exits
  var entries = [];
  var exits = [];

  for (var i = 0; i < bathroomPunches.length; i++) {
    var punch = bathroomPunches[i];
    var deviceLower = (punch.device || '').toLowerCase();

    if (deviceLower.indexOf('entry') >= 0) {
      entries.push(punch);
    } else if (deviceLower.indexOf('exit') >= 0) {
      exits.push(punch);
    }
  }

  Logger.log('    Raw bathroom punches - Entries: ' + entries.length + ', Exits: ' + exits.length);

  // Filter duplicates (60 seconds threshold for bathroom)
  function filterBathroomDuplicates(punches) {
    if (punches.length === 0) return [];

    var filtered = [punches[0]];
    var thresholdMs = config.bathroomDuplicateSeconds * 1000;

    for (var i = 1; i < punches.length; i++) {
      var lastTime = filtered[filtered.length - 1].time.getTime();
      var currentTime = punches[i].time.getTime();

      if (currentTime - lastTime >= thresholdMs) {
        filtered.push(punches[i]);
      } else {
        Logger.log('    🚻 Duplicate bathroom scan at ' + formatTime(punches[i].time) + ' - skipped');
      }
    }

    return filtered;
  }

  entries = filterBathroomDuplicates(entries);
  exits = filterBathroomDuplicates(exits);

  Logger.log('    After duplicate filtering - Entries: ' + entries.length + ', Exits: ' + exits.length);

  // Check for early bathroom (within 10 min of Clock 1 or Clock 3)
  if (entries.length > 0) {
    var earlyThresholdMs = config.earlyBathroomThreshold * 60 * 1000;

    if (adjusted.clock1) {
      var firstEntryAfterClock1 = null;
      for (var i = 0; i < entries.length; i++) {
        if (entries[i].time.getTime() > adjusted.clock1.getTime()) {
          firstEntryAfterClock1 = entries[i];
          break;
        }
      }
      if (firstEntryAfterClock1 && (firstEntryAfterClock1.time.getTime() - adjusted.clock1.getTime()) <= earlyThresholdMs) {
        warnings.push('Early bathroom after morning clock in (' + formatTime(firstEntryAfterClock1.time) + ')');
      }
    }

    if (adjusted.clock3) {
      var firstEntryAfterClock3 = null;
      for (var i = 0; i < entries.length; i++) {
        if (entries[i].time.getTime() > adjusted.clock3.getTime()) {
          firstEntryAfterClock3 = entries[i];
          break;
        }
      }
      if (firstEntryAfterClock3 && (firstEntryAfterClock3.time.getTime() - adjusted.clock3.getTime()) <= earlyThresholdMs) {
        warnings.push('Early bathroom after lunch return (' + formatTime(firstEntryAfterClock3.time) + ')');
      }
    }
  }

  // Match entries with exits
  var usedExits = [];

  for (var i = 0; i < entries.length; i++) {
    var entry = entries[i];
    var matched = false;

    // Find the first exit that comes after this entry
    for (var j = 0; j < exits.length; j++) {
      if (usedExits.indexOf(j) >= 0) continue;

      var exit = exits[j];
      if (exit.time.getTime() > entry.time.getTime()) {
        var duration = (exit.time.getTime() - entry.time.getTime()) / (60 * 1000);

        var breakInfo = {
          entry: formatTime(entry.time),
          exit: formatTime(exit.time),
          minutes: Math.round(duration)
        };

        // Check if both entry and exit are within work periods
        var entryInWork = isWithinWorkPeriods(entry.time);
        var exitInWork = isWithinWorkPeriods(exit.time);

        if (entryInWork && exitInWork) {
          // Count this break
          totalMinutes += duration;

          // Check for long bathroom break
          if (duration > config.longBathroomThreshold) {
            breakInfo.warning = 'Long bathroom break (' + Math.round(duration) + ' min)';
            warnings.push(breakInfo.warning);
          }

          breaks.push(breakInfo);
          Logger.log('    ✓ Break: ' + breakInfo.entry + '-' + breakInfo.exit + ' (' + breakInfo.minutes + 'm)');
        } else {
          Logger.log('    ✗ Break outside work: ' + breakInfo.entry + '-' + breakInfo.exit + ' (not counted)');
        }

        usedExits.push(j);
        matched = true;
        break;
      }
    }

    if (!matched && isWithinWorkPeriods(entry.time)) {
      unpairedEntries.push(formatTime(entry.time));
      warnings.push('Bathroom entry without matching exit (' + formatTime(entry.time) + ')');
    }
  }

  // Find exits without entries
  for (var j = 0; j < exits.length; j++) {
    if (usedExits.indexOf(j) === -1 && isWithinWorkPeriods(exits[j].time)) {
      unpairedExits.push(formatTime(exits[j].time));
      warnings.push('Bathroom exit without matching entry (' + formatTime(exits[j].time) + ')');
    }
  }

  // Check daily threshold
  if (totalMinutes > config.dailyBathroomThreshold) {
    warnings.push('Daily bathroom threshold exceeded: ' + Math.round(totalMinutes) + ' minutes');
  }

  Logger.log('    Total bathroom time: ' + Math.round(totalMinutes) + ' min');

  return {
    totalMinutes: Math.round(totalMinutes),
    breaks: breaks,
    warnings: warnings,
    unpairedEntries: unpairedEntries,
    unpairedExits: unpairedExits
  };
}

// ==================== DUPLICATE DETECTION ====================

/**
 * Filter duplicate main clock punches (2 minute threshold)
 *
 * @param {Array} punches - Array of punch objects
 * @param {Object} config - Time configuration
 * @returns {Array} Filtered punches
 */
function filterDuplicateMainPunches(punches, config) {
  if (punches.length === 0) return [];

  var filtered = [punches[0]];
  var thresholdMs = config.mainClockDuplicateMinutes * 60 * 1000;
  var duplicatesRemoved = 0;

  for (var i = 1; i < punches.length; i++) {
    var lastTime = filtered[filtered.length - 1].time.getTime();
    var currentTime = punches[i].time.getTime();

    if (currentTime - lastTime >= thresholdMs) {
      filtered.push(punches[i]);
    } else {
      duplicatesRemoved++;
      Logger.log('  ⚠️ DUPLICATE MAIN SCAN: ' + formatTime(punches[i].time) +
                 ' (only ' + Math.round((currentTime - lastTime) / 1000) + 's after previous) - SKIPPED');
    }
  }

  if (duplicatesRemoved > 0) {
    Logger.log('  📋 Filtered ' + duplicatesRemoved + ' duplicate main clock scans');
  }

  return filtered;
}

// ==================== DEVICE RECOGNITION ====================

/**
 * Determine punch type (in/out) from device name
 *
 * @param {string} deviceName - Device name
 * @param {number} punchIndex - Index in sequence (for alternating pattern)
 * @returns {string} 'in' or 'out'
 */
function determinePunchType(deviceName, punchIndex) {
  var deviceLower = (deviceName || '').toLowerCase().replace(/\s+/g, '');

  // Explicit IN devices
  if (deviceLower === 'clockin' || deviceLower === 'clock-in' || deviceLower === 'clock in') {
    return 'in';
  }

  // Explicit OUT devices
  if (deviceLower === 'clockout' || deviceLower === 'clock-out' || deviceLower === 'clock out') {
    return 'out';
  }

  // Unknown device: Use alternating pattern (even index = in, odd = out)
  return punchIndex % 2 === 0 ? 'in' : 'out';
}

// ==================== MAIN PROCESSING FUNCTION ====================

/**
 * Process clock data for an employee for one week (revised rules)
 *
 * @param {Array} clockData - Array of clock punch records
 * @param {Object} config - Time configuration
 * @returns {Object} Processed timesheet data
 */
function processClockData(clockData, config) {
  try {
    Logger.log('\n========== PROCESS CLOCK DATA (REVISED RULES) ==========');
    Logger.log('Processing ' + clockData.length + ' clock records');

    // Filter out records with empty or invalid PUNCH_TIME
    var validClockData = [];
    for (var i = 0; i < clockData.length; i++) {
      var record = clockData[i];
      if (record.PUNCH_TIME && record.PUNCH_TIME.toString().trim() !== '') {
        validClockData.push(record);
      } else {
        Logger.log('  ⚠️ Skipping record with empty PUNCH_TIME');
      }
    }

    if (validClockData.length !== clockData.length) {
      Logger.log('  📋 Filtered out ' + (clockData.length - validClockData.length) + ' records with empty PUNCH_TIME');
    }

    clockData = validClockData;

    // Sort by time
    clockData.sort(function(a, b) {
      try {
        var timeA = parseTime(a.PUNCH_TIME);
        var timeB = parseTime(b.PUNCH_TIME);
        return timeA.getTime() - timeB.getTime();
      } catch (e) {
        Logger.log('  ❌ ERROR sorting: ' + e.message);
        return 0;
      }
    });

    // Separate main and bathroom punches
    var mainPunches = [];
    var bathroomPunches = [];

    for (var i = 0; i < clockData.length; i++) {
      var record = clockData[i];
      var deviceName = record.DEVICE_NAME || '';

      if (deviceName.toLowerCase().indexOf('bathroom') >= 0) {
        bathroomPunches.push(record);
      } else {
        mainPunches.push(record);
      }
    }

    Logger.log('\n📋 Separated punches - Main: ' + mainPunches.length + ', Bathroom: ' + bathroomPunches.length);

    // Group by date
    var punchesByDate = {};

    for (var i = 0; i < mainPunches.length; i++) {
      var record = mainPunches[i];
      var punchTime = parseTime(record.PUNCH_TIME);
      var dateKey = formatDate(punchTime);

      if (!punchesByDate[dateKey]) {
        punchesByDate[dateKey] = {
          date: dateKey,
          mainPunches: [],
          bathroomPunches: []
        };
      }

      // Determine type using device recognition
      var type = determinePunchType(record.DEVICE_NAME, punchesByDate[dateKey].mainPunches.length);

      punchesByDate[dateKey].mainPunches.push({
        time: punchTime,
        type: type,
        device: record.DEVICE_NAME
      });
    }

    // Add bathroom punches to dates
    for (var i = 0; i < bathroomPunches.length; i++) {
      var record = bathroomPunches[i];
      var punchTime = parseTime(record.PUNCH_TIME);
      var dateKey = formatDate(punchTime);

      if (!punchesByDate[dateKey]) {
        punchesByDate[dateKey] = {
          date: dateKey,
          mainPunches: [],
          bathroomPunches: []
        };
      }

      punchesByDate[dateKey].bathroomPunches.push({
        time: punchTime,
        device: record.DEVICE_NAME
      });
    }

    // Process each day
    var dailyBreakdown = [];
    var totalMinutes = 0;
    var totalLunchDeduction = 0;
    var totalBathroomTime = 0;
    var allWarnings = [];

    for (var dateKey in punchesByDate) {
      var dayData = punchesByDate[dateKey];

      // Filter duplicates from main punches
      dayData.mainPunches = filterDuplicateMainPunches(dayData.mainPunches, config);

      var dayResult = processDayData(dayData, config);

      dailyBreakdown.push(dayResult);
      totalMinutes += dayResult.totalMinutes;
      totalLunchDeduction += dayResult.lunchMinutes;
      totalBathroomTime += dayResult.bathroomMinutes;

      if (dayResult.warnings) {
        allWarnings = allWarnings.concat(dayResult.warnings);
      }
    }

    // Calculate hours and minutes
    var totalHours = Math.floor(totalMinutes / 60);
    var remainingMinutes = totalMinutes % 60;

    // Log summary
    Logger.log('\n📊 WEEKLY SUMMARY');
    Logger.log('============================================');
    dailyBreakdown.forEach(function(day) {
      var dayHours = Math.floor(day.totalMinutes / 60);
      var dayMins = day.totalMinutes % 60;
      Logger.log('  ' + day.date + ': ' + dayHours + 'h ' + dayMins + 'm' +
                 (day.lunchMinutes > 0 ? ' (lunch: -' + day.lunchMinutes + 'm)' : ''));
    });
    Logger.log('============================================');
    Logger.log('  TOTAL PAID TIME: ' + totalHours + 'h ' + remainingMinutes + 'm');
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
 * Process clock data for a single day (revised rules)
 *
 * @param {Object} dayData - Day's punch data
 * @param {Object} config - Time configuration
 * @returns {Object} Day's processed data
 */
function processDayData(dayData, config) {
  var date = new Date(dayData.date);
  var isFriday = date.getDay() === 5;

  Logger.log('\n📅 Processing: ' + dayData.date + (isFriday ? ' (FRIDAY)' : ''));
  Logger.log('  Main punches: ' + dayData.mainPunches.length + ', Bathroom: ' + dayData.bathroomPunches.length);

  var warnings = [];

  // Classify clocks into Clock 1, 2, 3, 4
  var clocks = classifyClockPunches(dayData.mainPunches, config, isFriday);

  // Apply grace periods
  var graceResult = applyGracePeriods(clocks, config, date, isFriday);
  var adjusted = graceResult.adjusted;
  warnings = warnings.concat(graceResult.flags);

  // Calculate paid time
  var paidResult = calculatePaidTime(clocks, adjusted, config, isFriday);
  warnings = warnings.concat(paidResult.flags);

  // Calculate bathroom time
  var bathroomResult = calculateBathroomTime(dayData.bathroomPunches, config, clocks, adjusted, isFriday);
  warnings = warnings.concat(bathroomResult.warnings);

  // Collect punch times for display
  var punchTimes = [];
  if (clocks.clock1) punchTimes.push(formatTime(clocks.clock1.time));
  if (clocks.clock2) punchTimes.push(formatTime(clocks.clock2.time));
  if (clocks.clock3) punchTimes.push(formatTime(clocks.clock3.time));
  if (clocks.clock4) punchTimes.push(formatTime(clocks.clock4.time));

  return {
    date: dayData.date,
    scenario: paidResult.scenario,
    clockIn: clocks.clock1 ? formatTime(clocks.clock1.time) : null,
    clockOut: (isFriday ? (clocks.clock2 ? formatTime(clocks.clock2.time) : null) : (clocks.clock4 ? formatTime(clocks.clock4.time) : null)),
    adjustedIn: adjusted.clock1 ? formatTime(adjusted.clock1) : null,
    adjustedOut: (isFriday ? (adjusted.clock2 ? formatTime(adjusted.clock2) : null) : (adjusted.clock4 ? formatTime(adjusted.clock4) : null)),
    punchTimes: punchTimes,
    lunchTaken: paidResult.lunchMinutes > 0,
    lunchMinutes: paidResult.lunchMinutes,
    bathroomMinutes: bathroomResult.totalMinutes,
    bathroomBreaks: bathroomResult.breaks,
    bathroomUnpairedEntries: bathroomResult.unpairedEntries,
    bathroomUnpairedExits: bathroomResult.unpairedExits,
    totalMinutes: paidResult.paidMinutes,
    warnings: warnings
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
