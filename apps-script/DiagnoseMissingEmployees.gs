/**
 * Diagnostic script to find missing employees
 * This script checks which employees exist in the sheet but might not be showing in the app
 */

function diagnoseMissingEmployees() {
  console.log('=== DIAGNOSING MISSING EMPLOYEES ===\n');

  // List of employees that should be in the sheet
  var expectedEmployees = [
    'Aaron Mohlolo Lesolle',
    'Amos Dames-Fabriek',
    'Amos Dames-Tuin',
    'Andrew Molokola',
    'April Itumeleng Kemorwe',
    'April Scorpio',
    'Archie Patrick',
    'Boitumelo Mosimane Senokoane',
    'Casparus Johan Muller',
    'Daniel K Mokwena',
    'Darrion Keenan Frazenburg',
    'Delpert Morongoa Lehong',
    'Douglas Boytjie Afrika',
    'Eric Moses Sithole',
    'Evan Joe Nelson',
    'Gerald Maasdorp',
    'Goodwill Maetla',
    'Granno Gushwen Williams',
    'John Vaughan Adams',
    'John Vuisile Kgol',
    'Joshua Sephuma',
    'Judith Elizabeth Phefo',
    'Keagan Tremaine Anderson',
    'Kennedy Gaolekwe Mogole',
    'Leonard Masiea',
    'Leroy Segeel',
    'Lesego Phechudi',
    'Lesley Mogomotsiemang Malgase',
    'Leston Arthur Hartney',
    'Lukas Lesenyeho',
    'Lydia Mantutu Lesenyeho',
    'Manfid Bowker',
    'Manfid Bowker (cancel)',
    'Marcia Nonkululeko Biyela',
    'Maria Nomadlozi Mahlangu',
    'Mc Donald Mackenzie',
    'Melvyn Anderson',
    'Mzamo Ngqina',
    'Nicholas Molale',
    'Patric Manthi',
    'Peter Andre Graham Patrick',
    'Richard Tshkiso Mokgosi',
    'Rodney Segeel',
    'Ronaldo Patrick',
    'Ronnie Bronkhorst',
    'Roslyn Thebejane',
    'SAG Yvonne K',
    'Samuel (fabriek) Nhlapo',
    'Samuel Nhlapo Scorpio',
    'Samuel Ramathiba Mokaila',
    'Samuel Tuin',
    'Stephen Chimane Plaatjie',
    'Tapiwa Malvin Zulu',
    'Ticky William Lekoko',
    'Walter Hebert Coglin',
    'Walter Raphiri',
    'Yvonne Kedikgotse'
  ];

  // Employees that user says are missing from app
  var missingFromApp = [
    'Samuel Tuin',
    'Stephen Chimane Plaatjie',
    'Tapiwa Malvin Zulu',
    'Ticky William Lekoko',
    'Walter Hebert Coglin',
    'Walter Raphiri',
    'Yvonne Kedikgotse'
  ];

  try {
    var sheets = getSheets();
    var empSheet = sheets.empdetails;

    if (!empSheet) {
      console.error('Employee Details sheet not found!');
      return;
    }

    var data = empSheet.getDataRange().getValues();
    var headers = data[0];

    console.log('Total rows in sheet (including header):', data.length);
    console.log('Headers:', headers.join(', '));
    console.log('\n');

    // Find the REFNAME column index
    var refNameCol = -1;
    var statusCol = -1;
    var firstCol = 0;

    for (var i = 0; i < headers.length; i++) {
      if (headers[i] === 'REFNAME') refNameCol = i;
      if (headers[i] === 'EMPLOYMENT STATUS') statusCol = i;
    }

    console.log('REFNAME column index:', refNameCol);
    console.log('EMPLOYMENT STATUS column index:', statusCol);
    console.log('First column name:', headers[0]);
    console.log('\n');

    // Check each employee row
    console.log('=== CHECKING ALL ROWS IN SHEET ===\n');
    var foundInSheet = [];
    var rowDetails = [];

    for (var i = 1; i < data.length; i++) {
      var refname = refNameCol >= 0 ? data[i][refNameCol] : '';
      var status = statusCol >= 0 ? data[i][statusCol] : '';
      var firstColValue = data[i][firstCol];

      if (refname) {
        foundInSheet.push(refname);
        rowDetails.push({
          row: i + 1,
          refname: refname,
          status: status,
          firstCol: firstColValue,
          firstColEmpty: !firstColValue
        });
      }
    }

    console.log('Total employees found in sheet with REFNAME:', foundInSheet.length);
    console.log('\n');

    // Check specifically for missing employees
    console.log('=== CHECKING MISSING EMPLOYEES ===\n');

    for (var i = 0; i < missingFromApp.length; i++) {
      var empName = missingFromApp[i];
      var found = false;

      for (var j = 0; j < rowDetails.length; j++) {
        if (rowDetails[j].refname === empName) {
          found = true;
          console.log('FOUND: ' + empName);
          console.log('  Row:', rowDetails[j].row);
          console.log('  Status:', rowDetails[j].status);
          console.log('  First column value:', rowDetails[j].firstCol);
          console.log('  First column empty?', rowDetails[j].firstColEmpty);
          console.log('');
          break;
        }
      }

      if (!found) {
        console.log('NOT FOUND IN SHEET: ' + empName);
        console.log('');
      }
    }

    // Check for empty first column
    console.log('=== CHECKING FOR EMPTY FIRST COLUMN ===\n');
    var emptyFirstCol = [];

    for (var i = 0; i < rowDetails.length; i++) {
      if (rowDetails[i].firstColEmpty) {
        emptyFirstCol.push(rowDetails[i]);
      }
    }

    if (emptyFirstCol.length > 0) {
      console.log('Found', emptyFirstCol.length, 'employees with empty first column:');
      for (var i = 0; i < emptyFirstCol.length; i++) {
        console.log('  Row ' + emptyFirstCol[i].row + ':', emptyFirstCol[i].refname);
      }
    } else {
      console.log('No employees with empty first column');
    }
    console.log('\n');

    // Check employment status distribution
    console.log('=== EMPLOYMENT STATUS DISTRIBUTION ===\n');
    var statusCount = {};

    for (var i = 0; i < rowDetails.length; i++) {
      var status = rowDetails[i].status || '(empty)';
      if (!statusCount[status]) {
        statusCount[status] = 0;
      }
      statusCount[status]++;
    }

    for (var status in statusCount) {
      console.log(status + ':', statusCount[status]);
    }
    console.log('\n');

    // Now simulate the listEmployees function
    console.log('=== SIMULATING listEmployees() ===\n');

    var employees = [];
    var skippedRows = [];

    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) { // Check first column
        try {
          var employee = buildObjectFromRow(data[i], headers);
          employees.push(employee);
        } catch (rowError) {
          skippedRows.push({row: i + 1, error: rowError.toString()});
        }
      } else {
        skippedRows.push({row: i + 1, error: 'Empty first column'});
      }
    }

    console.log('Employees successfully built:', employees.length);
    console.log('Rows skipped:', skippedRows.length);

    if (skippedRows.length > 0) {
      console.log('\nSkipped rows details:');
      for (var i = 0; i < skippedRows.length; i++) {
        console.log('  Row ' + skippedRows[i].row + ':', skippedRows[i].error);
      }
    }
    console.log('\n');

    // Check which expected employees are in the built list
    console.log('=== CHECKING WHICH EMPLOYEES MADE IT TO THE LIST ===\n');

    for (var i = 0; i < missingFromApp.length; i++) {
      var empName = missingFromApp[i];
      var foundInList = false;

      for (var j = 0; j < employees.length; j++) {
        if (employees[j].REFNAME === empName) {
          foundInList = true;
          break;
        }
      }

      if (foundInList) {
        console.log('✓ ' + empName + ' - IN LIST');
      } else {
        console.log('✗ ' + empName + ' - NOT IN LIST');
      }
    }

    console.log('\n=== DIAGNOSIS COMPLETE ===');

  } catch (error) {
    console.error('Error during diagnosis:', error);
    console.error('Stack:', error.stack);
  }
}
