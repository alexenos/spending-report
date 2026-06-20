// Dax Garner
// 10-3-2012
// SpendingReport Version 4: Multiple Accounts

/***************************************************************************************************/
/*                                                                              Global Definitions */
/***************************************************************************************************/

/***************************************************************************************************/
/*                                                                                 Month Functions */
/***************************************************************************************************/

function eoySpendAsmt() { //    EOY Spending Assessment: Execute operations to assess EOY spending //
  // Select Active Spreadsheet
  var eoy = SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EOY"));
  // EOY Check
  if(validEOYCheck(eoy)) {               
    // Clear EOY Data
    clearEOYData();
    // Get Monthly Totals
    getMonthlyTotals();
    // Update Categories
    updateEOYCategories();
    // Get Category Totals
    getEOYCategoryTotals();
    // Calculate EOY Totals/Average/Percentages
    eoyTotals();
    // Format
    eoyFormat();
  } else { // Error
    eoyErrorMsg();
  } // end if
}

function clearEOYData() { //                               Clear Data: Clears all EOY computations //
  // Select Active Spreadsheet
  var eoy = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // EOY Check
  if(validEOYCheck(eoy)) {
    // Clear all computation data
    var numRows = getNumCat(def) + 3;
    // Clear data
    eoy.getRange(2, 2, numRows, 15).clear();
  } else { // Error
    eoyErrorMsg();
  } // end if
}    

function getMonthlyTotals() { //                                                 Get Montly Totals //
  // Select Active Spreadsheet
  var eoy = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Get All Sheets
  var allMonths = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  // EOY Check
  if(validEOYCheck(eoy)) {
    // Instantiate
    var totals;
    // Get Jan Index
    var index = getJanIndex();
    // Initialize while
    var begBal = allMonths[index].getRange(3, 8).getValue();
    var endBal = allMonths[index].getRange(4, 8).getValue();
    var col = 2;
    // While loop
    while(validMonthCheck(allMonths[index]) && (begBal != "" || endBal != "")) {
      // Get Monthly Totals
      totals = allMonths[index].getRange(5, 8, 3); // Totals
      // Set Monthly Totals in EOY sheet
      totals.copyTo(eoy.getRange(2, col));
      // Increment
      index++;
      col++;
      // Get new values
      begBal = allMonths[index].getRange(3, 8).getValue();
      endBal = allMonths[index].getRange(4, 8).getValue();
    }
  } else { // Error
    eoyErrorMsg();
  } // end if
}

function updateEOYCategories() { //                                 Update Categories on EOY Sheet //
  // Select Active Spreadsheet
  var eoy = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Instantiate
  var name;
  var abb;
  // EOY Check
  if(validEOYCheck(eoy)) {
    // Clear Category Data
    clearEOYCatData(eoy, def);
    // Initialize For
    var row = 5; var col = 1;
    // Update Entries
    for(var i = 0; i < 2; i++) { // For each category type
      for(var j = 0; j < def.cat[i].length; j++) { // For each category
        name = def.name[i][j];
        abb = def.cat[i][j];
        eoy.getRange(row, col).setValue(name + " (" + abb + ")");
        // Increment
        row++;
      } // end for
    } // end for
  } else { // Error
    eoyErrorMsg();
  } // end if
}

function getEOYCategoryTotals() { //                Update Categories Spending Totals on EOY Sheet //
  // Select Active Spreadsheet
  var eoy = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Get All Sheets
  var allMonths = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  // EOY Check
  if(validEOYCheck(eoy)) {
    // Instantiate
    var totals;
    // Get Jan Index
    var index = getJanIndex();
    // Get Spreadsheet
    var spendReport = SpreadsheetApp.getActiveSpreadsheet();
    // Initialize while
    var begBal = allMonths[index].getRange(3, 8).getValue();
    var endBal = allMonths[index].getRange(4, 8).getValue();
    var col = 2;
    // While loop
    while(validMonthCheck(allMonths[index]) && (begBal != "" || endBal != "")) {
      // --- Set Active Sheet to month of interest --- //
      spendReport.setActiveSheet(allMonths[index]);
      // Update Category and Category Spending for the Month of interest
      // updateCategories(); // too many calculations
      // monthlyCatSpending(); // too many calculations
      setMonthCatBorders(allMonths[index], def);
      allMonths[index].getRange(10, 7, getNumCat(def)).setHorizontalAlignment("left"); // Categories
      allMonths[index].getRange(10, 8, getNumCat(def)).setHorizontalAlignment("right"); // Totals
      allMonths[index].getRange(10, 9, getNumCat(def)).setHorizontalAlignment("right"); // Percentage
      // --- Return Active Sheet to EOY --- //
      spendReport.setActiveSheet(spendReport.getSheetByName("EOY"));
      // Get Monthly Totals
      totals = allMonths[index].getRange(10, 8, getNumCat(def)); // Totals
      // Set Monthly Totals in EOY sheet
      totals.copyTo(eoy.getRange(5, col));
      // Increment
      index++;
      col++;
      // Get new values
      begBal = allMonths[index].getRange(3, 8).getValue();
      endBal = allMonths[index].getRange(4, 8).getValue();
    }
  } else { // Error
    eoyErrorMsg();
  } // end if
}

function eoyTotals() { //                            Compute EOY Averages, Totals, and Percentages //
  // Select Active Spreadsheet
  var eoy = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // EOY Check
  if(validEOYCheck(eoy)) {
    // Instantiate
    var col;
    var sum;
    var avg;
    var pct;
    var cnt;
    // Initialize for loop
    var row = 2;
    var numOfRow = 3 + getNumCat(def);
    // For loop
    for(var i = 0; i < numOfRow; i++) { // For each row
      // Initialize while
      col = 2;
      sum = 0;
      cnt = 0;
      // While
      var val = eoy.getRange(2, col).getValue();
      while(val !== "" && cnt < 12) { // While the month has been processed.
        // Aggregate
        sum = sum + eoy.getRange(row, col).getValue();
        // Increment
        col++;
        cnt++; 
        val = eoy.getRange(2, col).getValue();
      } // end while
      // Cnt Check
      if(cnt <= 0) { cnt = 1; }
      // Compute Average
      avg = sum/cnt;
      //Compute Percentage
      pct = computeEOYPercent(eoy, row, sum, def);
      // Place values in sheet
      eoy.getRange(row, 14).setValue(sum) ;    // Sum
      eoy.getRange(row, 14 + 1).setValue(avg); // Average
      eoy.getRange(row, 14 + 2).setValue(pct); // Percentage
      // Increment
      row++;
    } // end for
  } else { // Error
    eoyErrorMsg();
  } // end if
}
      
function eoyFormat() { //                                                         Format EOY Sheet //
  // Select Active Spreadsheet
  var eoy = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // EOY Check
  if(validEOYCheck(eoy)) {
    // Alignment
    eoy.getRange(2, 16, 3 + getNumCat(def)).setHorizontalAlignment("right"); // Percentages
    // Clear Borders
    eoy.getRange(1, 1, 100, 16).setBorder(true, false, false, false, false, false); // Clear Borders
    // Initial Borders
    eoy.getRange(1, 1, 1, 16).setBorder(true, true, true, true, false, false); // Headings
    eoy.getRange(2, 1, 3 + getNumCat(def), 16).setBorder(true, true, true, true, false, false); // Outside Box
    // Borders
    eoyBordersByCol(eoy, 1, 1, def); // Category Headings
    eoyBordersByCol(eoy, 2, 12, def); // Months
    eoyBordersByCol(eoy, 14, 3, def); // EOY Totals
  } else { // Error
    eoyErrorMsg();
  } // end if
}

/***************************************************************************************************/
/*                                                                                 Month Utilities */
/***************************************************************************************************/

function validEOYCheck(eoy) { //            EOY Check: Confirms the active sheet is the EOY sheet. //
  // Initialize
  var name = eoy.getName();
  // Comparisions
  var isEOY = name == "EOY";
  var isNotEOY = (name == "Cat" || name == "Accts" || name == "Wealth" || name == "Jan" || name == "Feb" || name == "Mar" || name == "Apr" || name == "May" || name == "Jun" || name == "Jul" || name == "Aug" || name == "Sep" || name == "Oct" || name == "Nov" || name == "Dec" );
  // Return
  return (isEOY && !isNotEOY);
}
    
function getJanIndex() { //                                                           Find January //
  // Get All Sheets
  var allMonths = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  // Find January
  for(var i = 0; i < allMonths.length; i++) { // For each sheet
    if(allMonths[i].getName() == "Jan") { // If the sheet is titled "Jan"
      return i; // Return the index
    } // end if
  }
  // If it gets here there is an error.
  Browser.msgBox("There is no month of January! You need to start over.");
}
    
function clearEOYCatData(eoy, def) { //                          Clear Category Totals/Percentages //
  var numRows = getNumCat(def);
  // Clear data
  eoy.getRange(5, 1, numRows, 13).clear();
}

function computeEOYPercent(eoy, row, total, def) { //    Compute EOY Percentage for each category. //
  // Instantiate
  var pct;
  // Compute Percentage
  if(row > 4 && row < (5 + getNumCat(def))) { // Excludes the Earned/Spent/Saved and Utility Cateogories
    if(row <= 4 + def.cat[0].length) { // Income Categories
      pct = total/eoy.getRange(2, 14).getValue()*100;
    } else { // Spending Category
      pct = total/eoy.getRange(3, 14).getValue()*100;
    } // end if
  } else { // For Earned/Spent/Saved and Utility Cateogories
    pct = "--";
  } //end if
  // Return
  return pct;
}
  
function eoyBordersByCol(eoy, col, numCol, def) { //               Format Border for EOY by Column //
  eoy.getRange(2, col, 3, numCol).setBorder(true, true, true, true, false, false); // Earned/Spent/Saved
  eoy.getRange(5, col, def.cat[0].length, numCol).setBorder(true, true, true, true, false, false); // Income
  eoy.getRange(5 + def.cat[0].length, col, def.cat[1].length, numCol).setBorder(true, true, true, true, false, false); // Spending
}
    
function eoyErrorMsg() { //                                                         Not EOY sheet! //
  Browser.msgBox("Please select/create the EOY sheet in the lower left-hand corner before running this function.");
}