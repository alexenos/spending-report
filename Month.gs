// Dax Garner
// 10-3-2012
// SpendingReport Version 4: Multiple Accounts

/***************************************************************************************************/
/*                                                                              Global Definitions */
/***************************************************************************************************/


/***************************************************************************************************/
/*                                                                                 Month Functions */
/***************************************************************************************************/
function monthSpendAsmt() { //                                         Monthly Spending Assessment //
  // Select Active Spreadsheet
  var month = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Month Check
  if(validMonthCheck(month)) {  
    // Clear Data
    clearMonthData();    
    // Compute Transaction Balances
    transacBalances_M();    
    // Compute Monthly Totals
    monthlyTotals();    
    // Update Categories
    updateCategories_M();    
    // Category Spending
    monthlyCatSpending();    
    // Format
    monthlyFormat();    
  } else { // Error
    monthErrorMsg();
  } // end if  
}

function clearMonthData() { //                         Clear Data: Clears all monthly computations //
  // Select Active Spreadsheet
  var month = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Month Check
  if(validMonthCheck(month)) {
    // Clear all computation data
    clearBalanceData(month);
    clearTotalsData(month);
    clearCatData(month, def);
  } else { // Error
    monthErrorMsg();
  } // end if
}

function transacBalances_M() { //                                       Compute Transaction Balances //
  // Select Active Spreadsheet
  var month = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Instantiate
  var trn;  var utl;
  // Month Check
  if(validMonthCheck(month)) {
    // Get Beginning Balance
    var bal = getBegBal(month); 
    // Set Begininng Balances
    for(var i = 0; i < def.cat[3].length + 1; i++) {
      month.getRange(3, 8 + i).setValue(bal[i]);    
    } // end for
    // Get Transactions
    trn = getTransac(month);    
    // Intialize For loop
    var row = 2;    
    // Compute Balance for each transaction
    for(var i = 0; i < trn.amt.length; i++) { // For each transaction
      // Compute Balance
      utl = addOrSubtract(trn.cat[i], def); // Determine whether to add or subtract transaction and from which balance.
      bal[0] = bal[0] + utl.fct*(trn.amt[i]); // Always add/subtract transaction from checking
      if(utl.acct != 0) { // If account is not checking also add/subtract from external account balance, but opposite.
        bal[utl.acct] = bal[utl.acct] - utl.fct*(trn.amt[i]);    
      } // end if
      // Set Balance
      month.getRange(row, 5).setValue(bal[0]);    
      // Increment
      row++;
    } // end for    
    // Set Ending Balance
    for(var i = 0; i < def.cat[3].length + 1; i++) {
      month.getRange(4, 8 + i).setValue(bal[i]);    
    } // end for
  } else { // Error
    monthErrorMsg();
  } // end if  
}

function monthlyTotals() { //  Monthly Totals: Computes total spent, total earned and total saved. //
  // Select Active Spreadsheet
  var month = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Month Check
  if(validMonthCheck(month)) {
    // Get Transactions
    var trn = getTransac(month); 
    var numAccts = def.cat[3].length + 1;
    // Initialize for loop
    var earned = new Array();
    var spent = new Array();
    // For loop
    for(var i = 0; i < numAccts; i++) { // For each account
      earned[i] = 0; // zero out beginning amount
      spent[i] = 0;
    } // end for
    // For loop
    for(var i = 0; i < trn.amt.length; i++) { // For each transaction
      // Account Info on transaction
      var utl = AccountInfo(trn.cat[i], def)
      // If income
      if(isIncome(trn.cat[i], def) || (utl.isAcct && utl.type == "TI")) { // If the transaction is income or external account transfer in
        earned[utl.idx] = earned[utl.idx] + trn.amt[i]; // Remember utl.idx defautls to checking even if isAcct = false.
      } // end if
      // If Expense
      if(isExpense(trn.cat[i], def) || (utl.isAcct && utl.type == "TO")) { // If the transaction is an expense or external account transfer out
        spent[utl.idx] = spent[utl.idx] + trn.amt[i];
      } // end if
    } // end for    
    // Populate totals
    for(var i  = 0; i < numAccts; i++) {
      month.getRange(5, 8 + i).setValue(earned[i]);            // Total Earnings
      month.getRange(6, 8 + i).setValue(spent[i]);             // Total Spending
      month.getRange(7, 8 + i).setValue(earned[i] - spent[i]); // Total Savings    
    } // end for
  } else { // Error
    monthErrorMsg();
  } // end if
}

function updateCategories_M() { //                          Update Categories: Updates category list //
  // Select Active Spreadsheet
  var month = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); 
  // Instantiate local variables
  var name; var abb;  
  // Month Check
  if(validMonthCheck(month)) {    
    // Clear Category Data
    clearCatData(month, def);   
    // Update Category Entries
    var row = 10; var col = 7; 
    for(var i = 0; i < 2; i++) { // For each category type
      for(var j = 0; j < def.cat[i].length; j++) { // For each category
        name = def.name[i][j];
        abb = def.cat[i][j];
        month.getRange(row, col).setValue(name + " (" + abb + ")");        
        // Increment
        row++;
      } // end for
    } // end for
    // Update Account Entries
    var i = 3;
    row = 2; col = 9;
    month.getRange(row, 8).setValue("Checking");
    for(var j = 0; j < def.cat[i].length; j++) { // For each account
      name = def.name[i][j];
      abb = def.cat[i][j];
      month.getRange(row, col).setValue(name + " (" + abb + ")");        
      // Increment
      col++;
    } // end for
    // Update Investment Entries
    i = 2;
    row = 9; col = 12;    
    for(var j = 0; j < def.cat[i].length; j++) { // For each investment
      name = def.name[i][j];
      abb = def.cat[i][j];
      month.getRange(row, col).setValue(name + " (" + abb + ")");        
      // Increment
      col++;
    } // end for
  } else { // Error
    monthErrorMsg();
  } // end if
}

function monthlyCatSpending() { // Monthly Category Spending: Assess spending within each category //
  // Select Active Spreadsheet
  var month = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Month Check
  if(validMonthCheck(month)) {
    // Instantiate
    var total;
    var sum;
    var pct;
    // Get Transactions
    var trn = getTransac(month);
    // Initialize for loop 1
    var row = 10;
    // For loop 1
    for(var i = 0; i < 2; i++) { // For each category type
      // For loop 2
      for(var j = 0; j < def.cat[i].length; j++) { // For each category
        // Initialize Category Sum
        sum = 0;
        // For loop 3
        for(var k = 0; k < trn.amt.length; k++) { // For each transaction
          if(trn.cat[k] == def.cat[i][j]) { // If the transaction category matches the category of interest.
            sum = sum + trn.amt[k];
          } // end if
        } // end for 3
        // Get Total
        total = getCatTypeTotal(month, def, def.cat[i][j]);
        // Compute Percentage
        pct = computePct_M(sum, total);
        // Set Spreadsheet values
        month.getRange(row, 8).setValue(sum);
        month.getRange(row, 9).setValue(pct);
        // Increment
        row++;
      }
    }     
  } else { // Error
    monthErrorMsg();
  } // end if
}
      
function monthlyFormat() { //                                                  Format Monthly Data //
  // Select Active Spreadsheet
  var month = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Month Check
  if(validMonthCheck(month)) {
    // Get Transactions
    var trn = getTransac(month);
    if(trn.cat.length > 0) { // If there are any transactions
      // Bold
      month.getRange(1, 1, 1, 5).setFontWeight("bold"); // Transaction Headings
      month.getRange(3, 7, 5, 1).setFontWeight("bold"); // Balances and Totals
      month.getRange(2, 8, 1, def.cat[3].length + 1).setFontWeight("bold"); // Acount Headings
      month.getRange(9, 12, 1, def.cat[2].length).setFontWeight("bold"); // Investment Headings
      month.getRange(10, 11, 3, 1).setFontWeight("bold"); // Investment Headings
      month.getRange(9, 7, 1, 3).setFontWeight("bold"); // Category Headings
      // Alignment
      month.getRange(2, 1, trn.cat.length).setHorizontalAlignment("right"); // Dates
      month.getRange(2, 2, trn.cat.length).setHorizontalAlignment("left"); // Description
      month.getRange(2, 3, trn.cat.length).setHorizontalAlignment("center"); // Cat
      month.getRange(2, 4, trn.cat.length).setHorizontalAlignment("right"); // Amount
      month.getRange(2, 5, trn.cat.length).setHorizontalAlignment("right"); // Balance
      month.getRange(3, 7, 5).setHorizontalAlignment("left"); // Balances and Totals
      month.getRange(3, 8, 5).setHorizontalAlignment("right"); // Balances and Totals
      month.getRange(10, 7, getNumCat(def)).setHorizontalAlignment("left"); // Categories
      month.getRange(10, 8, getNumCat(def)).setHorizontalAlignment("right"); // Totals
      month.getRange(10, 9, getNumCat(def)).setHorizontalAlignment("right"); // Percentage
      // Clear Borders
      month.getRange(2, 1, 100, 5).setBorder(true, false, false, false, false, false); // Clear Transaction Borders
      month.getRange(10, 7, 100, 3).setBorder(true, false, false, false, false, false); // Clear Category Borders
      // Borders
      month.getRange(1, 1, 1, 5).setBorder(true, true, true, true, true, false); // Trans Headings
      month.getRange(2, 1, trn.cat.length, 5).setBorder(true, true, true, true, true, false); // Trans
      month.getRange(2, 8, 6, def.cat[3].length + 1).setBorder(false, true, true, true, false, false); // Accounts
      month.getRange(2, 8, 3, def.cat[3].length + 1).setBorder(false, true, true, true, false, false); // Accounts
      month.getRange(2, 9, 1, def.cat[3].length).setBorder(false, true, true, true, false, false); // Accounts
      month.getRange(2, 8, 6, 1).setBorder(false, true, true, true, false, false); // Accounts
      month.getRange(2, 8, 3, 1).setBorder(false, true, true, true, false, false); // Accounts
      month.getRange(2, 8, 1, 1).setBorder(false, true, true, true, false, false); // Accounts
      month.getRange(10, 11, 3, 1).setBorder(true, true, true, true, false, false); // Investments  
      month.getRange(10, 11, 2, 1).setBorder(true, true, true, true, false, false); // Investments 
      month.getRange(9, 12, 1, def.cat[2].length).setBorder(true, true, true, true, false, false); // Investments 
      month.getRange(10, 12, 2, def.cat[2].length).setBorder(true, true, true, true, false, false); // Investments 
      month.getRange(12, 12, 1, def.cat[2].length).setBorder(true, true, true, true, false, false); // Investments 
      setMonthCatBorders(month, def);
    } else { // Just clear the borders
      // Clear Borders
      month.getRange(2, 1, 100, 5).setBorder(true, false, false, false, false, false); // Clear Transaction Borders
      month.getRange(10, 7, 100, 3).setBorder(true, false, false, false, false, false); // Clear Category Borders
    }
  } else { // Error
    monthErrorMsg();
  } // end if
}


/***************************************************************************************************/
/*                                                                                 Month Utilities */
/***************************************************************************************************/

function validMonthCheck(month) { //      Month Check: Confirms the active sheet is a month sheet. //
  // Initialize
  var name = month.getName();  
  // Comparisions
  var isMonth = (name == "Jan" || name == "Feb" || name == "Mar" || name == "Apr" || name == "May" || name == "Jun" || name == "Jul" || name == "Aug" || name == "Sep" || name == "Oct" || name == "Nov" || name == "Dec" );
  var isNotMonth = (name == "Def" || name == "EOY" || name == "Utl");  
  // Return
  return (isMonth && !isNotMonth);
}

function clearBalanceData(month) { //               Clear Balance: Clears all balance computations //
  var numRows = getTransac(month).cat.length;
  // Zero conditional
  if(numRows == 0) {
    return;
  }  
  // Clear data
  month.getRange(2, 5, numRows, 1).clear();
}

function clearTotalsData(month) { //                   Clear Totals: Clears all total computations //
  var numCols = def.cat[3].length + 1;
  // Clear data
  month.getRange(2, 8, 1, numCols).clear(); // Headings
  month.getRange(4, 8, 4, numCols).clear(); // Computed values
  if(month.getName() != "Jan") {
    month.getRange(3, 8, 1, numCols).clear(); // Beg balance
  }
}

function clearCatData(month, def) { //                           Clear Category Totals/Percentages //
  var numRows = getNumCat(def); 
  // Clear data
  month.getRange(10, 7, numRows, 3).clear(); // Categories
  month.getRange(2, 9, 1, def.cat[3].length).clear(); // Accounts
  month.getRange(9, 12, 1, def.cat[2].length).clear(); // Investments, just names no numbers!
}

function getTransac(month) { //                                                   Get Transactions //
  // Instantiate
  var cat = new Array();
  var amt = new Array();
  // Initialize while
  var row = 2;
  var cnt = 0;
  // While there is a description or a category value, there is a transaction.
  while(month.getRange(row, 2).getValue() != "" || month.getRange(row, 3).getValue() != "") {
    cat[cnt] = month.getRange(row, 3).getValue();
    amt[cnt] = month.getRange(row, 4).getValue();
    // Increment
    row++; cnt++;
  }
  // Return
  return {cat : cat, amt : amt};
}

function getBegBal(month) { //   Get Beginning Balance: Gets beginning balance depending on month. //
  // Instantiate
  var begBal = new Array(); 
  var allMonths;
  var numAccts = def.cat[3].length + 1;
  if(month.getName() == "Jan") { // If the active month is January.
    if(month.getRange(3, 8).getValue() != "") { // There is an initial balance!
      for(var i = 0; i < numAccts; i++) {
        begBal[i] = month.getRange(3, 8 + i).getValue();
      }
    } else { // Error: Must have an initial balance! 
      Browser.msgBox("Please enter a beginning account balance in cell H3 for the month of January.");
    } // end if
  } else { // Otherwise grab the previous months ending balance.
    allMonths = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for(var i = 0; i < numAccts; i++) {
      begBal[i] = allMonths[month.getIndex() - 2].getRange(4, 8 + i).getValue();
    }
  } // end if
  // Return
  return begBal;  
}

function addOrSubtract(trn, def) { //       Add or Subtract: Determines whether transaction is +/- //
  // Initialize
  var fct = -1; // Default to subtract: expect more expenditures than income transactions.
  var acct = 0; // Default to checking
  // Determine Category Type
  if(isExpense(trn, def)) { // Is expense
    fct = -1; // Subtract
    acct = 0; // Checking
  } else {// Not sure yet
    if(isIncome(trn, def)) { // Is income
      fct = 1; // Add
      acct = 0; // Checking
    } else { // Not sure yet
      var utl = AccountInfo(trn, def);
      if(utl.isAcct) { // Is account
        acct = utl.idx
        if(utl.type == "TI") { // Transfer In
          fct = -1; // Transfer Out of Checking and into External Account.
        } else { // Double check
          if(utl.type == "TO") { // Transfer Out
            fct = 1; // Transfer out of external account into checking
          } else { // Error: Accounts must end in TO or TI
            catErrorMsg(trn);
          } // end if
        } // end if
      } else { // Not sure yet
        trn_sub = trn.substring(0, trn.length - 2);
        if(isInvest(trn_sub, def)) {
          fct = -1;
        } else { // Not sure yet
          if(trn == "") { // Blank.
            fct = 0;
          } else { // Error: No Category recognized
            catErrorMsg(trn);
          } // end if
        } // end if
      } // end if
    } // end if
  } // end if  
  // Return
  return {fct : fct, acct : acct};
}
                      
function getCatTypeTotal(month, def, cat) { //     Gets the Earned or Saved Totals for percentages //
  var total;
  if(isExpense(cat, def)) { // If the category is an expense
    total = month.getRange(6, 8).getValue();
  } else { // Not sure yet
    if(isIncome(cat, def)) { // If the category is an income category
      total = month.getRange(5, 8).getValue();
    } else { // Neither
      total = "--";
    } // end if
  } // end if               
  // Return
  return total;     
}
                       
function computePct_M(sum, total) { //                             Compute percentage based on total //
  // Instantiate
  var pct;                     
  if(total == "--") { // If there is no total
    pct = "--";
  } else { 
    if(total == 0) { // If there is no total
      pct = "--";
    } else {// Compute percentage
      pct = sum/total*100;
    } // end if
  } // end if                       
  // Return
  return pct;
}

function setMonthCatBorders(month, def) { //                 Category Boarders in the Month Sheets //
  month.getRange(9, 7, 1, 3).setBorder(true, true, true, true, true, false); // Cat Headings
  month.getRange(10, 7, getNumCat(def), 3).setBorder(true, true, true, true, true, false); // Cat
  month.getRange(10, 7, def.cat[0].length, 3).setBorder(true, true, true, true, true, false); // Income Spending Division
  //month.getRange(10 + getNumCat(def) - def.cat[2].length, 7, def.cat[2].length, 3).setBorder(true, true, true, true, true, false); // Spending Utility Division - Obsolete
}

function monthErrorMsg() { //                                                         Not a month! //
  Browser.msgBox("Please select one of the month tabs (Jan - Dec) in the lower left-hand corner before running this function.");
}

function catErrorMsg(trn) { //                                                     Not a category! //
  Browser.msgBox(trn + " is not a defined category! Please see the definitions sheet (Def) for valid spending categories.");
}

