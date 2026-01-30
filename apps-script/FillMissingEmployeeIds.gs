/**
 * Fill missing employee IDs
 * This script checks for employees without IDs and generates them
 */

function fillMissingEmployeeIds() {
  console.log('=== FILLING MISSING EMPLOYEE IDs ===\n');

  try {
    var sheets = getSheets();
    var empSheet = sheets.empdetails;

    if (!empSheet) {
      console.error('Employee Details sheet not found!');
      return {success: false, error: 'Sheet not found'};
    }

    var data = empSheet.getDataRange().getValues();
    var headers = data[0];

    // Find column indices
    var idColIndex = -1;
    var refNameColIndex = -1;

    for (var i = 0; i < headers.length; i++) {
      if (headers[i] === 'id') idColIndex = i;
      if (headers[i] === 'REFNAME') refNameColIndex = i;
    }

    if (idColIndex === -1) {
      console.error('id column not found!');
      return {success: false, error: 'id column not found'};
    }

    if (refNameColIndex === -1) {
      console.error('REFNAME column not found!');
      return {success: false, error: 'REFNAME column not found'};
    }

    console.log('id column index:', idColIndex);
    console.log('REFNAME column index:', refNameColIndex);
    console.log('\n');

    var updatedCount = 0;
    var rowsToUpdate = [];

    // Check each employee row
    for (var i = 1; i < data.length; i++) {
      var currentId = data[i][idColIndex];
      var refName = data[i][refNameColIndex];

      // If employee has a REFNAME but no ID, generate one
      if (refName && refName.toString().trim() !== '' && (!currentId || currentId.toString().trim() === '')) {
        var newId = generateId();
        console.log('Row', (i + 1), ':', refName, '- Generating ID:', newId);

        rowsToUpdate.push({
          row: i + 1,
          refName: refName,
          newId: newId
        });

        // Update the data array
        data[i][idColIndex] = newId;
        updatedCount++;
      }
    }

    console.log('\nTotal employees needing IDs:', updatedCount);

    if (updatedCount > 0) {
      console.log('\nUpdating sheet...');

      // Write all data back to sheet
      var range = empSheet.getRange(1, 1, data.length, data[0].length);
      range.setValues(data);

      console.log('✓ Sheet updated successfully!');
      console.log('\nUpdated employees:');
      for (var i = 0; i < rowsToUpdate.length; i++) {
        console.log('  -', rowsToUpdate[i].refName, '(Row', rowsToUpdate[i].row + ')');
      }

      return {
        success: true,
        updatedCount: updatedCount,
        updatedEmployees: rowsToUpdate
      };
    } else {
      console.log('\n✓ All employees already have IDs!');
      return {
        success: true,
        updatedCount: 0,
        message: 'All employees already have IDs'
      };
    }

  } catch (error) {
    console.error('Error:', error);
    console.error('Stack:', error.stack);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Check which employees are missing IDs without updating
 */
function checkMissingEmployeeIds() {
  console.log('=== CHECKING FOR MISSING EMPLOYEE IDs ===\n');

  try {
    var sheets = getSheets();
    var empSheet = sheets.empdetails;

    if (!empSheet) {
      console.error('Employee Details sheet not found!');
      return;
    }

    var data = empSheet.getDataRange().getValues();
    var headers = data[0];

    // Find column indices
    var idColIndex = -1;
    var refNameColIndex = -1;

    for (var i = 0; i < headers.length; i++) {
      if (headers[i] === 'id') idColIndex = i;
      if (headers[i] === 'REFNAME') refNameColIndex = i;
    }

    if (idColIndex === -1 || refNameColIndex === -1) {
      console.error('Required columns not found!');
      return;
    }

    var missingIds = [];

    // Check each employee row
    for (var i = 1; i < data.length; i++) {
      var currentId = data[i][idColIndex];
      var refName = data[i][refNameColIndex];

      // If employee has a REFNAME but no ID
      if (refName && refName.toString().trim() !== '' && (!currentId || currentId.toString().trim() === '')) {
        missingIds.push({
          row: i + 1,
          refName: refName
        });
      }
    }

    console.log('Total rows in sheet:', data.length - 1);
    console.log('Employees missing IDs:', missingIds.length);
    console.log('\n');

    if (missingIds.length > 0) {
      console.log('Employees without IDs:');
      for (var i = 0; i < missingIds.length; i++) {
        console.log('  Row', missingIds[i].row + ':', missingIds[i].refName);
      }
    } else {
      console.log('✓ All employees have IDs!');
    }

  } catch (error) {
    console.error('Error:', error);
    console.error('Stack:', error.stack);
  }
}
