/**
 * EMPLOYEES.GS - Employee Management
 * HR Payroll System for SA Grinding Wheels & Scorpio Abrasives
 *
 * This file handles all employee CRUD operations:
 * - Add new employee
 * - Update existing employee
 * - Get employee details
 * - List/search employees
 * - Terminate employee
 * - Validation
 */

/**
 * Add a new employee to the system
 *
 * @param {Object} data - Employee data
 * @returns {Object} Result object with success flag and data/error
 */
function addEmployee(data) {
  try {
    Logger.log('\n========== ADD EMPLOYEE ==========');
    Logger.log('ℹ️ Input: ' + JSON.stringify(data));

    // Sanitize inputs
    data.employeeName = sanitizeInput(data.employeeName);
    data.surname = sanitizeInput(data.surname);

    // Validate employee data
    validateEmployee(data);

    // Get sheet
    const sheets = getSheets();
    const empSheet = sheets.empdetails;
    if (!empSheet) {
      throw new Error('Employee sheet not found');
    }

    // Generate ID and REFNAME
    data.id = generateUUID();
    data.REFNAME = generateRefName(data.employeeName, data.surname);

    // Add audit fields
    addAuditFields(data, true);

    // Get headers
    const headers = empSheet.getRange(1, 1, 1, empSheet.getLastColumn()).getValues()[0];

    // Convert object to row
    const row = objectToRow(data, headers);

    // Append row
    empSheet.appendRow(row);
    SpreadsheetApp.flush();

    Logger.log('✅ Employee added successfully: ' + data.id);
    Logger.log('========== ADD EMPLOYEE COMPLETE ==========\n');

    return { success: true, data: { id: data.id, refName: data.REFNAME } };

  } catch (error) {
    Logger.log('❌ ERROR in addEmployee: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== ADD EMPLOYEE FAILED ==========\n');

    return { success: false, error: error.message };
  }
}

/**
 * Update an existing employee
 *
 * @param {string} id - Employee ID
 * @param {Object} data - Updated employee data
 * @returns {Object} Result object
 */
function updateEmployee(id, data) {
  try {
    Logger.log('\n========== UPDATE EMPLOYEE ==========');
    Logger.log('ℹ️ Employee ID: ' + id);
    Logger.log('ℹ️ Update data: ' + JSON.stringify(data));

    // Sanitize inputs
    if (data.employeeName) data.employeeName = sanitizeInput(data.employeeName);
    if (data.surname) data.surname = sanitizeInput(data.surname);

    // Get sheet
    const sheets = getSheets();
    const empSheet = sheets.empdetails;
    if (!empSheet) {
      throw new Error('Employee sheet not found');
    }

    // Find employee
    const allData = empSheet.getDataRange().getValues();
    const headers = allData[0];
    const rows = allData.slice(1);

    const idCol = findColumnIndex(headers, 'id');
    const rowIndex = rows.findIndex(row => row[idCol] === id);

    if (rowIndex === -1) {
      throw new Error('Employee not found: ' + id);
    }

    // Get current employee data
    const currentEmployee = rowToObject(rows[rowIndex], headers);

    // Merge with updates
    const updatedEmployee = Object.assign({}, currentEmployee, data);

    // Update REFNAME if name changed
    if (data.employeeName || data.surname) {
      updatedEmployee.REFNAME = generateRefName(
        updatedEmployee['EMPLOYEE NAME'],
        updatedEmployee.SURNAME
      );
    }

    // Validate updated data
    validateEmployee(updatedEmployee, id);

    // Add audit fields (update)
    addAuditFields(updatedEmployee, false);

    // Convert to row
    const updatedRow = objectToRow(updatedEmployee, headers);

    // Update sheet (row index + 2 because: +1 for header, +1 for 1-based indexing)
    const sheetRowIndex = rowIndex + 2;
    empSheet.getRange(sheetRowIndex, 1, 1, headers.length).setValues([updatedRow]);
    SpreadsheetApp.flush();

    Logger.log('✅ Employee updated successfully: ' + id);
    Logger.log('========== UPDATE EMPLOYEE COMPLETE ==========\n');

    return { success: true, data: updatedEmployee };

  } catch (error) {
    Logger.log('❌ ERROR in updateEmployee: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== UPDATE EMPLOYEE FAILED ==========\n');

    return { success: false, error: error.message };
  }
}

/**
 * Get employee by ID
 *
 * @param {string} id - Employee ID
 * @returns {Object} Result object
 */
function getEmployeeById(id) {
  try {
    Logger.log('🔍 Getting employee by ID: ' + id);

    const sheets = getSheets();
    const empSheet = sheets.empdetails;
    if (!empSheet) {
      throw new Error('Employee sheet not found');
    }

    const allData = empSheet.getDataRange().getValues();
    const headers = allData[0];
    const rows = allData.slice(1);

    const idCol = findColumnIndex(headers, 'id');
    const row = rows.find(r => r[idCol] === id);

    if (!row) {
      throw new Error('Employee not found: ' + id);
    }

    const employee = rowToObject(row, headers);

    Logger.log('✅ Employee found: ' + employee.REFNAME);

    return { success: true, data: employee };

  } catch (error) {
    Logger.log('❌ ERROR in getEmployeeById: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Get employee by REFNAME
 *
 * @param {string} refName - Employee REFNAME (e.g., "John Smith")
 * @returns {Object} Result object
 */
function getEmployeeByName(refName) {
  try {
    Logger.log('🔍 Getting employee by name: ' + refName);

    const sheets = getSheets();
    const empSheet = sheets.empdetails;
    if (!empSheet) {
      throw new Error('Employee sheet not found');
    }

    const allData = empSheet.getDataRange().getValues();
    const headers = allData[0];
    const rows = allData.slice(1);

    const refNameCol = findColumnIndex(headers, 'REFNAME');
    const row = rows.find(r => r[refNameCol] === refName);

    if (!row) {
      throw new Error('Employee not found: ' + refName);
    }

    const employee = rowToObject(row, headers);

    Logger.log('✅ Employee found: ' + employee.id);

    return { success: true, data: employee };

  } catch (error) {
    Logger.log('❌ ERROR in getEmployeeByName: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * List all employees with optional filtering
 *
 * @param {Object} filters - Filter options
 * @param {string} filters.employer - Filter by employer
 * @param {string} filters.status - Filter by employment status
 * @param {string} filters.search - Search by name
 * @param {boolean} filters.activeOnly - Show only active (no termination date)
 * @returns {Object} Result object
 */
function listEmployees(filters = {}) {
  try {
    Logger.log('\n========== LIST EMPLOYEES ==========');
    Logger.log('ℹ️ Filters: ' + JSON.stringify(filters));

    const sheets = getSheets();
    const empSheet = sheets.empdetails;
    if (!empSheet) {
      throw new Error('Employee sheet not found');
    }

    const allData = empSheet.getDataRange().getValues();
    const headers = allData[0];
    let rows = allData.slice(1);

    // Convert rows to objects
    let employees = rows.map(row => rowToObject(row, headers));

    // Apply filters
    if (filters.employer) {
      employees = employees.filter(emp => emp.EMPLOYER === filters.employer);
    }

    if (filters.status) {
      employees = employees.filter(emp => emp['EMPLOYMENT STATUS'] === filters.status);
    }

    if (filters.activeOnly) {
      employees = employees.filter(emp => !emp['TERMINATION DATE']);
    }

    if (filters.search) {
      const searchLower = filters.search.toLowerCase();
      employees = employees.filter(emp => {
        const refName = (emp.REFNAME || '').toLowerCase();
        return refName.includes(searchLower);
      });
    }

    Logger.log('✅ Found ' + employees.length + ' employees');
    Logger.log('========== LIST EMPLOYEES COMPLETE ==========\n');

    return { success: true, data: employees };

  } catch (error) {
    Logger.log('❌ ERROR in listEmployees: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    Logger.log('========== LIST EMPLOYEES FAILED ==========\n');

    return { success: false, error: error.message };
  }
}

/**
 * Terminate an employee (set termination date)
 *
 * @param {string} id - Employee ID
 * @param {Date|string} terminationDate - Termination date
 * @returns {Object} Result object
 */
function terminateEmployee(id, terminationDate) {
  try {
    Logger.log('\n========== TERMINATE EMPLOYEE ==========');
    Logger.log('ℹ️ Employee ID: ' + id);
    Logger.log('ℹ️ Termination Date: ' + terminationDate);

    const date = parseDate(terminationDate);

    const result = updateEmployee(id, {
      'TERMINATION DATE': date
    });

    if (result.success) {
      Logger.log('✅ Employee terminated successfully');
    }

    Logger.log('========== TERMINATE EMPLOYEE COMPLETE ==========\n');

    return result;

  } catch (error) {
    Logger.log('❌ ERROR in terminateEmployee: ' + error.message);
    Logger.log('========== TERMINATE EMPLOYEE FAILED ==========\n');

    return { success: false, error: error.message };
  }
}

/**
 * Validate employee data
 *
 * @param {Object} data - Employee data to validate
 * @param {string} excludeId - Employee ID to exclude from uniqueness checks
 * @throws {Error} If validation fails
 */
function validateEmployee(data, excludeId = null) {
  const errors = [];

  // Check required fields
  if (!data.employeeName && !data['EMPLOYEE NAME']) {
    errors.push('Employee Name is required');
  }
  if (!data.surname && !data.SURNAME) {
    errors.push('Surname is required');
  }
  if (!data.employer && !data.EMPLOYER) {
    errors.push('Employer is required');
  }
  if (!data.hourlyRate && !data['HOURLY RATE']) {
    errors.push('Hourly Rate is required');
  }
  if (!data.idNumber && !data['ID NUMBER']) {
    errors.push('ID Number is required');
  }
  if (!data.contactNumber && !data['CONTACT NUMBER']) {
    errors.push('Contact Number is required');
  }
  if (!data.address && !data.ADDRESS) {
    errors.push('Address is required');
  }
  if (!data.employmentStatus && !data['EMPLOYMENT STATUS']) {
    errors.push('Employment Status is required');
  }
  if (!data.clockInRef && !data.ClockInRef) {
    errors.push('Clock Number is required');
  }

  // Get values (handle both camelCase and UPPERCASE keys)
  const hourlyRate = data.hourlyRate || data['HOURLY RATE'];
  const employer = data.employer || data.EMPLOYER;
  const employmentStatus = data.employmentStatus || data['EMPLOYMENT STATUS'];
  const idNumber = data.idNumber || data['ID NUMBER'];
  const clockInRef = data.clockInRef || data.ClockInRef;

  // Validate hourly rate
  if (hourlyRate && hourlyRate <= 0) {
    errors.push('Hourly Rate must be greater than 0');
  }

  // Validate employer
  if (employer && !isValidEnum(employer, EMPLOYER_LIST)) {
    errors.push('Invalid employer. Must be one of: ' + EMPLOYER_LIST.join(', '));
  }

  // Validate employment status
  if (employmentStatus && !isValidEnum(employmentStatus, EMPLOYMENT_STATUS_LIST)) {
    errors.push('Invalid employment status. Must be one of: ' + EMPLOYMENT_STATUS_LIST.join(', '));
  }

  // Validate ID number format
  if (idNumber && !validateSAIdNumber(idNumber)) {
    errors.push('Invalid ID Number format. Must be 13 digits');
  }

  // Check unique constraints
  if (idNumber && isIdNumberUsed(idNumber, excludeId)) {
    errors.push('ID Number already exists: ' + idNumber);
  }

  if (clockInRef && isClockInRefUsed(clockInRef, excludeId)) {
    errors.push('Clock Number already exists: ' + clockInRef);
  }

  if (errors.length > 0) {
    throw new Error(errors.join('; '));
  }
}

/**
 * Check if ID Number is already used
 *
 * @param {string} idNumber - ID number to check
 * @param {string} excludeId - Employee ID to exclude from check
 * @returns {boolean} True if ID is already used
 */
function isIdNumberUsed(idNumber, excludeId = null) {
  try {
    const sheets = getSheets();
    const empSheet = sheets.empdetails;
    if (!empSheet) return false;

    const allData = empSheet.getDataRange().getValues();
    const headers = allData[0];
    const rows = allData.slice(1);

    const idCol = findColumnIndex(headers, 'id');
    const idNumberCol = findColumnIndex(headers, 'ID NUMBER');

    for (let row of rows) {
      if (row[idNumberCol] === idNumber && row[idCol] !== excludeId) {
        return true;
      }
    }

    return false;

  } catch (error) {
    Logger.log('⚠️ Error checking ID number: ' + error.message);
    return false;
  }
}

/**
 * Check if Clock Number is already used
 *
 * @param {string} clockInRef - Clock number to check
 * @param {string} excludeId - Employee ID to exclude from check
 * @returns {boolean} True if clock number is already used
 */
function isClockInRefUsed(clockInRef, excludeId = null) {
  try {
    const sheets = getSheets();
    const empSheet = sheets.empdetails;
    if (!empSheet) return false;

    const allData = empSheet.getDataRange().getValues();
    const headers = allData[0];
    const rows = allData.slice(1);

    const idCol = findColumnIndex(headers, 'id');
    const clockCol = findColumnIndex(headers, 'ClockInRef');

    for (let row of rows) {
      if (row[clockCol] === clockInRef && row[idCol] !== excludeId) {
        return true;
      }
    }

    return false;

  } catch (error) {
    Logger.log('⚠️ Error checking clock number: ' + error.message);
    return false;
  }
}

/**
 * Generate REFNAME from first and last name
 *
 * @param {string} firstName - First name
 * @param {string} surname - Surname
 * @returns {string} REFNAME (e.g., "John Smith")
 */
function generateRefName(firstName, surname) {
  return firstName + ' ' + surname;
}

// ========== TEST FUNCTIONS ==========

/**
 * Test adding an employee
 */
function test_addEmployee() {
  Logger.log('\n========== TEST: ADD EMPLOYEE ==========');

  const testData = {
    employeeName: 'Test',
    surname: 'Employee',
    employer: 'SA Grinding Wheels',
    hourlyRate: 35.00,
    idNumber: '9501015800081',
    contactNumber: '0821234567',
    address: '123 Test Street, Johannesburg',
    employmentStatus: 'Permanent',
    clockInRef: '999'
  };

  const result = addEmployee(testData);

  if (result.success) {
    Logger.log('✅ TEST PASSED: Employee added with ID ' + result.data.id);
  } else {
    Logger.log('❌ TEST FAILED: ' + result.error);
  }

  Logger.log('========== TEST COMPLETE ==========\n');
}

/**
 * Test listing employees
 */
function test_listEmployees() {
  Logger.log('\n========== TEST: LIST EMPLOYEES ==========');

  const result = listEmployees({ activeOnly: true });

  if (result.success) {
    Logger.log('✅ TEST PASSED: Found ' + result.data.length + ' employees');
    result.data.forEach(emp => {
      Logger.log('  - ' + emp.REFNAME + ' (' + emp.EMPLOYER + ')');
    });
  } else {
    Logger.log('❌ TEST FAILED: ' + result.error);
  }

  Logger.log('========== TEST COMPLETE ==========\n');
}
