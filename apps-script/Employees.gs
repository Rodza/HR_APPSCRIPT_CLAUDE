/**
 * Employees.gs - Employee Management Functions
 *
 * Contains all server-side functions for employee CRUD operations
 * with comprehensive logging for debugging
 */

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Ensure consistent response format
 * @param {boolean} success - Whether the operation was successful
 * @param {any} data - The data to return
 * @param {string} error - Error message if any
 * @returns {Object} Properly formatted response
 */
function formatResponse(success, data, error) {
  return {
    success: success,
    data: data || null,
    error: error || null
  };
}

// ============================================================================
// TEST FUNCTION - Simple response test
// ============================================================================

/**
 * Test function to verify response format works from web app
 * @returns {Object} Simple test response
 */
function testSimpleResponse() {
  console.log('testSimpleResponse called');
  var result = {
    success: true,
    data: [{name: 'Test Employee', id: '123'}],
    error: null
  };
  console.log('Returning:', JSON.stringify(result));
  return result;
}

/**
 * Minimal test of listEmployees - just the basics
 * @returns {Object} Response
 */
function testListEmployeesMinimal() {
  console.log('=== testListEmployeesMinimal START ===');

  try {
    console.log('Step 1: Calling formatResponse directly...');
    var testResponse = formatResponse(true, [], null);
    console.log('formatResponse result:', JSON.stringify(testResponse));

    console.log('Step 2: Trying to get sheets...');
    var sheets = getSheets();
    console.log('Got sheets:', Object.keys(sheets).join(', '));

    console.log('Step 3: Checking empdetails...');
    if (!sheets.empdetails) {
      console.error('No empdetails sheet!');
      return formatResponse(false, [], 'No empdetails sheet');
    }
    console.log('empdetails exists:', sheets.empdetails.getName());

    console.log('Step 4: Getting data...');
    var data = sheets.empdetails.getDataRange().getValues();
    console.log('Got', data.length, 'rows');

    console.log('Step 5: Creating response...');
    var response = formatResponse(true, [{test: 'data'}], null);
    console.log('Response created:', JSON.stringify(response));

    console.log('=== testListEmployeesMinimal END ===');
    return response;

  } catch (error) {
    console.error('=== testListEmployeesMinimal ERROR ===');
    console.error('Error:', error.toString());
    console.error('Stack:', error.stack);
    return formatResponse(false, [], error.toString());
  }
}

// ============================================================================
// LIST EMPLOYEES
// ============================================================================

/**
 * Sanitize employee object for serialization to web app
 * Converts dates, handles null values, ensures all data is serializable
 * @param {Object} emp - Employee object from sheet
 * @returns {Object} Sanitized employee object
 */
function sanitizeEmployeeForWeb(emp) {
  var sanitized = {};

  for (var key in emp) {
    if (emp.hasOwnProperty(key)) {
      var value = emp[key];

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
      // Convert everything else to string
      else {
        sanitized[key] = String(value);
      }
    }
  }

  return sanitized;
}

/**
 * List all employees with optional filters
 * @param {Object} filters - Optional filters {employer, status, search, activeOnly}
 * @returns {Object} Response with employee list
 */
function listEmployees(filters) {
  try {
    console.log('=== listEmployees() START (from web app) ===');
    console.log('Filters received:', filters);

    // Step 1: Get sheets
    console.log('Step 1: Getting sheets...');
    var sheets = getSheets();
    console.log('✓ Sheets retrieved');

    // Step 2: Get employee sheet
    console.log('Step 2: Getting employee sheet...');
    var empSheet = sheets.empdetails;

    if (!empSheet) {
      console.error('✗ Employee Details sheet not found!');
      console.error('Available sheets:', Object.keys(sheets));
      throw new Error('Employee Details sheet not found');
    }
    console.log('✓ Employee sheet found:', empSheet.getName());

    // Step 3: Get data range
    console.log('Step 3: Getting data range...');
    var data = empSheet.getDataRange().getValues();
    console.log('✓ Data retrieved:', data.length, 'rows total');

    if (data.length === 0) {
      console.warn('Sheet is empty!');
      return formatResponse(true, [], null);
    }

    // Step 4: Extract headers
    console.log('Step 4: Extracting headers...');
    var headers = data[0];
    console.log('✓ Headers:', headers.join(', '));

    // Step 5: Build employee objects
    console.log('Step 5: Building employee objects...');
    var employees = [];
    var skippedRows = 0;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) { // Skip empty rows
        try {
          var employee = buildObjectFromRow(data[i], headers);
          employees.push(employee);
        } catch (rowError) {
          console.error('Error building row ' + (i + 1) + ':', rowError);
          skippedRows++;
        }
      } else {
        skippedRows++;
      }
    }

    console.log('✓ Built', employees.length, 'employee objects');
    if (skippedRows > 0) {
      console.log('  Skipped', skippedRows, 'empty/invalid rows');
    }

    // Step 6: Apply filters
    var originalCount = employees.length;

    if (filters) {
      console.log('Step 6: Applying filters...');

      // Filter by employer
      if (filters.employer && filters.employer !== '') {
        console.log('  Filtering by employer:', filters.employer);
        employees = filterByField(employees, 'EMPLOYER', filters.employer);
        console.log('  After employer filter:', employees.length, 'employees');
      }

      // Filter by search (name)
      if (filters.search && filters.search !== '') {
        console.log('  Filtering by search term:', filters.search);
        employees = filterBySearch(employees, 'REFNAME', filters.search);
        console.log('  After search filter:', employees.length, 'employees');
      }

      // Filter by status
      if (filters.status && filters.status !== '') {
        console.log('  Filtering by status:', filters.status);
        employees = filterByField(employees, 'EMPLOYMENT STATUS', filters.status);
        console.log('  After status filter:', employees.length, 'employees');
      }

      // Filter active only
      if (filters.activeOnly) {
        console.log('  Filtering active only');
        employees = employees.filter(function(emp) {
          return emp['EMPLOYMENT STATUS'] !== 'Terminated' &&
                 emp['EMPLOYMENT STATUS'] !== 'Resigned';
        });
        console.log('  After active filter:', employees.length, 'employees');
      }
    } else {
      console.log('Step 6: No filters provided, returning all employees');
    }

    console.log('✓ Final count:', employees.length, '/', originalCount, 'employees');

    // CRITICAL FIX: Sanitize employee data for web serialization
    console.log('Step 7: Sanitizing employee data for web...');
    var sanitizedEmployees = [];
    for (var i = 0; i < employees.length; i++) {
      try {
        var sanitized = sanitizeEmployeeForWeb(employees[i]);
        sanitizedEmployees.push(sanitized);
      } catch (sanitizeError) {
        console.error('Error sanitizing employee ' + (i + 1) + ':', sanitizeError);
      }
    }
    console.log('✓ Sanitized', sanitizedEmployees.length, 'employees');

    console.log('=== listEmployees() END - SUCCESS ===');
    console.log('Preparing response...');

    var response = formatResponse(true, sanitizedEmployees, null);
    console.log('Response created:', response ? 'OK' : 'NULL');
    console.log('Response.success:', response.success);
    console.log('Response.data length:', response.data ? response.data.length : 'NULL');

    return response;

  } catch (error) {
    console.error('=== listEmployees() ERROR ===');
    console.error('Error type:', error.name);
    console.error('Error message:', error.message);
    console.error('Full error:', error.toString());
    console.error('Stack trace:', error.stack);

    var errorResponse = formatResponse(false, [], error.toString());
    console.log('Error response created:', errorResponse ? 'OK' : 'NULL');
    return errorResponse;
  }
}

/**
 * Full implementation - temporarily disabled for testing
 */
function listEmployeesFull(filters) {
  try {
    console.log('=== listEmployees() START (from web app) ===');
    console.log('Filters received:', filters);

    // Step 1: Get sheets
    console.log('Step 1: Getting sheets...');
    var sheets = getSheets();
    console.log('✓ Sheets retrieved');

    // Step 2: Get employee sheet
    console.log('Step 2: Getting employee sheet...');
    var empSheet = sheets.empdetails;

    if (!empSheet) {
      console.error('✗ Employee Details sheet not found!');
      console.error('Available sheets:', Object.keys(sheets));
      throw new Error('Employee Details sheet not found');
    }
    console.log('✓ Employee sheet found:', empSheet.getName());

    // Step 3: Get data range
    console.log('Step 3: Getting data range...');
    var data = empSheet.getDataRange().getValues();
    console.log('✓ Data retrieved:', data.length, 'rows total');

    if (data.length === 0) {
      console.warn('Sheet is empty!');
      return formatResponse(true, [], null);
    }

    // Step 4: Extract headers
    console.log('Step 4: Extracting headers...');
    var headers = data[0];
    console.log('✓ Headers:', headers.join(', '));

    // Step 5: Build employee objects
    console.log('Step 5: Building employee objects...');
    var employees = [];
    var skippedRows = 0;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) { // Skip empty rows
        try {
          var employee = buildObjectFromRow(data[i], headers);
          employees.push(employee);
        } catch (rowError) {
          console.error('Error building row ' + (i + 1) + ':', rowError);
          skippedRows++;
        }
      } else {
        skippedRows++;
      }
    }

    console.log('✓ Built', employees.length, 'employee objects');
    if (skippedRows > 0) {
      console.log('  Skipped', skippedRows, 'empty/invalid rows');
    }

    // Step 6: Apply filters
    var originalCount = employees.length;

    if (filters) {
      console.log('Step 6: Applying filters...');

      // Filter by employer
      if (filters.employer && filters.employer !== '') {
        console.log('  Filtering by employer:', filters.employer);
        employees = filterByField(employees, 'EMPLOYER', filters.employer);
        console.log('  After employer filter:', employees.length, 'employees');
      }

      // Filter by search (name)
      if (filters.search && filters.search !== '') {
        console.log('  Filtering by search term:', filters.search);
        employees = filterBySearch(employees, 'REFNAME', filters.search);
        console.log('  After search filter:', employees.length, 'employees');
      }

      // Filter by status
      if (filters.status && filters.status !== '') {
        console.log('  Filtering by status:', filters.status);
        employees = filterByField(employees, 'EMPLOYMENT STATUS', filters.status);
        console.log('  After status filter:', employees.length, 'employees');
      }

      // Filter active only
      if (filters.activeOnly) {
        console.log('  Filtering active only');
        employees = employees.filter(function(emp) {
          return emp['EMPLOYMENT STATUS'] !== 'Terminated' &&
                 emp['EMPLOYMENT STATUS'] !== 'Resigned';
        });
        console.log('  After active filter:', employees.length, 'employees');
      }
    } else {
      console.log('Step 6: No filters provided, returning all employees');
    }

    console.log('✓ Final count:', employees.length, '/', originalCount, 'employees');

    // Don't use logging functions before return - they might interfere
    // logSuccess('Found ' + employees.length + ' employees');
    // logFunctionEnd('listEmployees', {count: employees.length});

    console.log('=== listEmployees() END - SUCCESS ===');
    console.log('Preparing response...');

    var response = formatResponse(true, employees, null);
    console.log('Response created:', response ? 'OK' : 'NULL');
    console.log('Response.success:', response.success);
    console.log('Response.data length:', response.data ? response.data.length : 'NULL');

    return response;

  } catch (error) {
    console.error('=== listEmployees() ERROR ===');
    console.error('Error type:', error.name);
    console.error('Error message:', error.message);
    console.error('Full error:', error.toString());
    console.error('Stack trace:', error.stack);

    // Don't use logError before return
    // logError('Failed to list employees', error);

    var errorResponse = formatResponse(false, [], error.toString());
    console.log('Error response created:', errorResponse ? 'OK' : 'NULL');
    return errorResponse;
  }
}

// ============================================================================
// GET EMPLOYEE BY ID
// ============================================================================

/**
 * Get a single employee by ID
 * @param {string} id - Employee ID
 * @returns {Object} Response with employee data
 */
function getEmployeeById(id) {
  try {
    console.log('=== getEmployeeById() START ===');
    console.log('Looking for employee ID:', id);
    logFunctionStart('getEmployeeById', {id: id});

    if (!id) {
      throw new Error('Employee ID is required');
    }

    // Step 1: Get sheets
    console.log('Step 1: Getting sheets...');
    var sheets = getSheets();
    var empSheet = sheets.empdetails;

    if (!empSheet) {
      throw new Error('Employee Details sheet not found');
    }
    console.log('✓ Employee sheet found:', empSheet.getName());

    // Step 2: Get data
    console.log('Step 2: Getting data...');
    var data = empSheet.getDataRange().getValues();
    var headers = data[0];
    console.log('✓ Data retrieved:', data.length, 'rows');

    // Step 3: Find ID column
    console.log('Step 3: Finding ID column...');
    var idCol = indexOf(headers, 'id');

    if (idCol === -1) {
      throw new Error('ID column not found in sheet');
    }
    console.log('✓ ID column index:', idCol);

    // Step 4: Search for employee
    console.log('Step 4: Searching for employee...');
    for (var i = 1; i < data.length; i++) {
      if (data[i][idCol] === id) {
        var employee = buildObjectFromRow(data[i], headers);
        console.log('✓ Employee found at row', (i + 1) + ':', employee.REFNAME);

        // Step 5: Sanitize employee data for web serialization
        console.log('Step 5: Sanitizing employee data for web...');
        var sanitizedEmployee = sanitizeEmployeeForWeb(employee);
        console.log('✓ Employee data sanitized');

        logSuccess('Found employee: ' + employee.REFNAME);
        logFunctionEnd('getEmployeeById', {found: true});

        console.log('=== getEmployeeById() END - SUCCESS ===');
        return formatResponse(true, sanitizedEmployee, null);
      }
    }

    console.log('✗ Employee not found with ID:', id);
    logWarning('Employee not found: ' + id);
    logFunctionEnd('getEmployeeById', {found: false});

    console.log('=== getEmployeeById() END - NOT FOUND ===');
    return formatResponse(false, null, 'Employee not found');

  } catch (error) {
    console.error('=== getEmployeeById() ERROR ===');
    console.error('Full error:', error.toString());
    logError('Failed to get employee', error);
    return formatResponse(false, null, error.toString());
  }
}

// ============================================================================
// FIELD NAME TRANSFORMATION
// ============================================================================

/**
 * Transform camelCase field names from frontend to uppercase sheet column names
 * @param {Object} data - Employee data with camelCase field names
 * @returns {Object} Employee data with uppercase field names
 */
function transformEmployeeFieldNames(data) {
  var transformed = {};

  // Define field mapping from camelCase (frontend) to UPPERCASE (backend/sheet)
  var fieldMapping = {
    'employeeName': 'EMPLOYEE NAME',
    'surname': 'SURNAME',
    'idNumber': 'ID NUMBER',
    'dateOfBirth': 'DATE OF BIRTH',
    'contactNumber': 'CONTACT NUMBER',
    'alternativeContact': 'ALTERNATIVE CONTACT',
    'altContactName': 'ALT CONTACT NAME',
    'address': 'ADDRESS',
    'employer': 'EMPLOYER',
    'employmentStatus': 'EMPLOYMENT STATUS',
    'hourlyRate': 'HOURLY RATE',
    'clockInRef': 'ClockInRef',
    'employmentDate': 'EMPLOYMENT DATE',
    'incomeTaxNumber': 'INCOME TAX NUMBER',
    'terminationDate': 'TERMINATION DATE',
    'overalSize': 'OveralSize',
    'shoeSize': 'ShoeSize',
    'notes': 'NOTES'
  };

  // Transform camelCase fields to uppercase
  for (var key in data) {
    if (data.hasOwnProperty(key)) {
      var mappedKey = fieldMapping[key] || key;
      transformed[mappedKey] = data[key];
    }
  }

  return transformed;
}

// ============================================================================
// CREATE EMPLOYEE
// ============================================================================

/**
 * Create a new employee
 * @param {Object} employeeData - Employee data object
 * @returns {Object} Response with created employee
 */
function createEmployee(employeeData) {
  try {
    console.log('=== createEmployee() START ===');
    console.log('Employee data received:', JSON.stringify(employeeData));
    logFunctionStart('createEmployee', employeeData);

    // Step 0: Transform camelCase fields to uppercase sheet column names
    console.log('Step 0: Transforming field names...');
    employeeData = transformEmployeeFieldNames(employeeData);
    console.log('Transformed data:', JSON.stringify(employeeData));

    // Step 1: Validate required fields
    console.log('Step 1: Validating required fields...');
    var requiredFields = getConfig('EMPLOYEE_REQUIRED_FIELDS');
    var missingFields = [];

    for (var i = 0; i < requiredFields.length; i++) {
      var field = requiredFields[i];
      if (isEmpty(employeeData[field])) {
        missingFields.push(field);
      }
    }

    if (missingFields.length > 0) {
      throw new Error('Missing required fields: ' + missingFields.join(', '));
    }
    console.log('✓ All required fields present');

    // Step 2: Get sheets
    console.log('Step 2: Getting sheets...');
    var sheets = getSheets();
    var empSheet = sheets.empdetails;

    if (!empSheet) {
      throw new Error('Employee Details sheet not found');
    }
    console.log('✓ Employee sheet found');

    // Step 3: Check for duplicates
    console.log('Step 3: Checking for duplicates...');
    var data = empSheet.getDataRange().getValues();
    var headers = data[0];
    var idNumCol = indexOf(headers, 'ID NUMBER');
    var clockCol = indexOf(headers, 'ClockInRef');

    for (var i = 1; i < data.length; i++) {
      if (data[i][idNumCol] === employeeData['ID NUMBER']) {
        throw new Error('ID Number already exists: ' + employeeData['ID NUMBER']);
      }
      if (data[i][clockCol] === employeeData['ClockInRef']) {
        throw new Error('Clock Number already exists: ' + employeeData['ClockInRef']);
      }
    }
    console.log('✓ No duplicates found');

    // Step 4: Generate ID and REFNAME
    console.log('Step 4: Generating ID and REFNAME...');
    employeeData.id = generateId();
    employeeData.REFNAME = employeeData['EMPLOYEE NAME'] + ' ' + employeeData['SURNAME'];
    employeeData.USER = getCurrentUser();
    employeeData.TIMESTAMP = getCurrentTimestamp();
    console.log('✓ Generated ID:', employeeData.id);
    console.log('✓ Generated REFNAME:', employeeData.REFNAME);

    // Step 5: Build row data
    console.log('Step 5: Building row data...');
    var newRow = [];
    for (var i = 0; i < headers.length; i++) {
      var header = headers[i];
      newRow.push(employeeData[header] || '');
    }

    // Step 6: Append to sheet
    console.log('Step 6: Appending to sheet...');
    empSheet.appendRow(newRow);
    console.log('✓ Employee added to sheet');

    logSuccess('Employee created: ' + employeeData.REFNAME);
    logFunctionEnd('createEmployee', {id: employeeData.id});

    console.log('=== createEmployee() END - SUCCESS ===');
    return formatResponse(true, employeeData, null);

  } catch (error) {
    console.error('=== createEmployee() ERROR ===');
    console.error('Full error:', error.toString());
    logError('Failed to create employee', error);
    return formatResponse(false, null, error.toString());
  }
}

// ============================================================================
// UPDATE EMPLOYEE
// ============================================================================

/**
 * Update an existing employee
 * @param {string} id - Employee ID
 * @param {Object} employeeData - Updated employee data
 * @returns {Object} Response with updated employee
 */
function updateEmployee(id, employeeData) {
  try {
    console.log('=== updateEmployee() START ===');
    console.log('Updating employee ID:', id);
    console.log('Update data:', JSON.stringify(employeeData));
    logFunctionStart('updateEmployee', {id: id, data: employeeData});

    if (!id) {
      throw new Error('Employee ID is required');
    }

    // Transform camelCase fields to uppercase sheet column names
    console.log('Step 0: Transforming field names...');
    employeeData = transformEmployeeFieldNames(employeeData);
    console.log('Transformed data:', JSON.stringify(employeeData));

    // Get sheets
    var sheets = getSheets();
    var empSheet = sheets.empdetails;

    if (!empSheet) {
      throw new Error('Employee Details sheet not found');
    }

    // Find employee row
    var data = empSheet.getDataRange().getValues();
    var headers = data[0];
    var idCol = indexOf(headers, 'id');
    var rowIndex = -1;

    for (var i = 1; i < data.length; i++) {
      if (data[i][idCol] === id) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error('Employee not found: ' + id);
    }

    console.log('✓ Found employee at row', (rowIndex + 1));

    // Update audit fields
    employeeData.MODIFIED_BY = getCurrentUser();
    employeeData.LAST_MODIFIED = getCurrentTimestamp();

    // Update REFNAME if name changed
    if (employeeData['EMPLOYEE NAME'] || employeeData['SURNAME']) {
      var firstName = employeeData['EMPLOYEE NAME'] || data[rowIndex][indexOf(headers, 'EMPLOYEE NAME')];
      var surname = employeeData['SURNAME'] || data[rowIndex][indexOf(headers, 'SURNAME')];
      employeeData.REFNAME = firstName + ' ' + surname;
    }

    // Build updated row
    var updatedRow = data[rowIndex];
    for (var i = 0; i < headers.length; i++) {
      var header = headers[i];
      if (employeeData.hasOwnProperty(header)) {
        updatedRow[i] = employeeData[header];
      }
    }

    // Write back to sheet
    var range = empSheet.getRange(rowIndex + 1, 1, 1, headers.length);
    range.setValues([updatedRow]);

    console.log('✓ Employee updated');
    logSuccess('Employee updated: ' + employeeData.REFNAME);
    logFunctionEnd('updateEmployee', {id: id});

    console.log('=== updateEmployee() END - SUCCESS ===');
    return formatResponse(true, buildObjectFromRow(updatedRow, headers), null);

  } catch (error) {
    console.error('=== updateEmployee() ERROR ===');
    console.error('Full error:', error.toString());
    logError('Failed to update employee', error);
    return formatResponse(false, null, error.toString());
  }
}

// ============================================================================
// DELETE/TERMINATE EMPLOYEE
// ============================================================================

/**
 * Terminate an employee (soft delete)
 * @param {string} id - Employee ID
 * @param {string} terminationDate - Termination date
 * @returns {Object} Response
 */
function terminateEmployee(id, terminationDate) {
  try {
    console.log('=== terminateEmployee() START ===');
    console.log('Terminating employee ID:', id);
    logFunctionStart('terminateEmployee', {id: id, date: terminationDate});

    var updateData = {
      'EMPLOYMENT STATUS': 'Terminated',
      'TERMINATION DATE': terminationDate || new Date()
    };

    var result = updateEmployee(id, updateData);

    if (result.success) {
      logSuccess('Employee terminated: ' + id);
      console.log('=== terminateEmployee() END - SUCCESS ===');
    }

    return result;

  } catch (error) {
    console.error('=== terminateEmployee() ERROR ===');
    console.error('Full error:', error.toString());
    logError('Failed to terminate employee', error);
    return formatResponse(false, null, error.toString());
  }
}

// ============================================================================
// TEST FUNCTION
// ============================================================================

/**
 * Test employee functions with comprehensive logging
 */
function testEmployeeFunctions() {
  console.log('====================================');
  console.log('TESTING EMPLOYEE FUNCTIONS');
  console.log('====================================');

  try {
    // Test 1: List employees
    console.log('\n--- Test 1: List Employees ---');
    var listResult = listEmployees(null);
    console.log('Result:', JSON.stringify(listResult, null, 2));

    // Test 2: List with filters
    console.log('\n--- Test 2: List with Filters ---');
    var filterResult = listEmployees({employer: 'SA Grinding Wheels'});
    console.log('Result:', JSON.stringify(filterResult, null, 2));

    console.log('\n====================================');
    console.log('ALL TESTS COMPLETED');
    console.log('====================================');

    return true;

  } catch (error) {
    console.error('TEST FAILED:', error);
    return false;
  }
}
