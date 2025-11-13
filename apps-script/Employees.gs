// CRITICAL FIX FOR Employees.gs
// Add this wrapper function at the top of your Employees.gs file:

/**
 * Ensure consistent response format
 * @param {any} data - The data to wrap
 * @returns {Object} Properly formatted response
 */
function formatResponse(success, data, error) {
  return {
    success: success,
    data: data || null,
    error: error || null
  };
}

// UPDATED listEmployees function - Replace your existing function with this:
function listEmployees(filters) {
  try {
    logFunctionStart('listEmployees', filters);
    
    var sheets = getSheets();
    var empSheet = sheets.empdetails;
    
    if (!empSheet) {
      throw new Error('Employee Details sheet not found');
    }
    
    var data = empSheet.getDataRange().getValues();
    var headers = data[0];
    var employees = [];
    
    // Convert all rows to objects
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) { // Skip empty rows
        var employee = buildObjectFromRow(data[i], headers);
        employees.push(employee);
      }
    }
    
    // Apply filters if provided
    if (filters) {
      // Filter by employer
      if (filters.employer && filters.employer !== '') {
        employees = filterByField(employees, 'EMPLOYER', filters.employer);
      }
      
      // Filter by search (name)
      if (filters.search && filters.search !== '') {
        employees = filterBySearch(employees, 'REFNAME', filters.search);
      }
      
      // Filter by status
      if (filters.status && filters.status !== '') {
        employees = filterByField(employees, 'EMPLOYMENT STATUS', filters.status);
      }
      
      // Filter active only
      if (filters.activeOnly) {
        employees = employees.filter(function(emp) {
          return emp['EMPLOYMENT STATUS'] !== 'Terminated' && 
                 emp['EMPLOYMENT STATUS'] !== 'Resigned';
        });
      }
    }
    
    logSuccess('Found ' + employees.length + ' employees');
    logFunctionEnd('listEmployees', {count: employees.length});
    
    // FIX: Always return consistent format
    return formatResponse(true, employees, null);
    
  } catch (error) {
    logError('Failed to list employees', error);
    // FIX: Return proper error response
    return formatResponse(false, [], error.toString());
  }
}

// ADDITIONAL FIX for getEmployeeById:
function getEmployeeById(id) {
  try {
    logFunctionStart('getEmployeeById', {id: id});
    
    if (!id) {
      throw new Error('Employee ID is required');
    }
    
    var sheets = getSheets();
    var empSheet = sheets.empdetails;
    
    if (!empSheet) {
      throw new Error('Employee Details sheet not found');
    }
    
    var data = empSheet.getDataRange().getValues();
    var headers = data[0];
    var idCol = indexOf(headers, 'id');
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][idCol] === id) {
        var employee = buildObjectFromRow(data[i], headers);
        logSuccess('Found employee: ' + employee.REFNAME);
        logFunctionEnd('getEmployeeById', {found: true});
        
        // FIX: Return consistent format
        return formatResponse(true, employee, null);
      }
    }
    
    logWarning('Employee not found: ' + id);
    logFunctionEnd('getEmployeeById', {found: false});
    
    // FIX: Return proper not found response
    return formatResponse(false, null, 'Employee not found');
    
  } catch (error) {
    logError('Failed to get employee', error);
    return formatResponse(false, null, error.toString());
  }
}
