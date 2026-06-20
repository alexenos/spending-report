// Dax Garner
// 10-3-2012
// SpendingReport Version 4: Multiple Accounts

/***************************************************************************************************/
/*                                                                              Global Definitions */
/***************************************************************************************************/

/***************************************************************************************************/
/*                                                                                Wealth Functions */
/***************************************************************************************************/

function wealthAsmt() {
  // Select Active Spreadsheet
  var wealth = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Wealth Check
  if(validWealthCheck(wealth)) {  
    // Clear Data
    clearWealthData();
    // Add Accounts
    updateAccounts();
    // Get Current Balances
    getBalances();
    // Compute Totals
    computeTotals();
    // Compute Percentages
    computePct_W();
    // Format
    wealthFormat();
  }
}

function clearWealthData() {
  // Select Active Spreadsheet
  var wealth = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Wealth Check
  if(validWealthCheck(wealth)) {
    wealth.getRange(3, 1, 100, 4).clear();
    wealth.getRange(3, 6, 1, 2).clear();
  }
}

function updateAccounts() {
  // Select Active Spreadsheet
  var wealth = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Wealth Check
  if(validWealthCheck(wealth)) {
    // Clear Data
    clearWealthData();
    // Update Accounts
    wealth.getRange(3, 1).setValue("Checking"); // Checking is special
    var name; var abb;
    var row = 4; var col = 1;
    for(var i = 3; i > 1; i--) {
      for(var j = 0; j < def.cat[i].length; j++) { // For each category/account
        name = def.name[i][j];
        abb = def.cat[i][j];
        if(i == 3) { // External Accounts
          wealth.getRange(row, col).setValue(name + " (" + abb + ")");
        }
        if(i == 2) { // Investments
          wealth.getRange(row, col).setValue(name + " (" + abb + "): Balance");
          row++;
          wealth.getRange(row, col).setValue(name + " (" + abb + "): Mkt Value");
        }
        // Increment
        row++;
      } // end for
    }
  }
}

function getBalances() {
  // Select Active Spreadsheet
  var wealth = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Get All Sheets
  var allMonths = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  // Wealth Check
  if(validWealthCheck(wealth)) {
    // Get Jan Index
    var index = getJanIndex();
    // Initialize Balances
    var begBal = allMonths[index].getRange(3, 8).getValue();
    var endBal = allMonths[index].getRange(4, 8).getValue();
    // While loop
    while(validMonthCheck(allMonths[index]) && (begBal != "" || endBal != "")) {
      begBal = allMonths[index].getRange(3, 8).getValue();
      endBal = allMonths[index].getRange(4, 8).getValue();
      index++;
    }
    // Initialize for loops
    var lastIndex = index - 2;
    var bal = new Array();
    var k = 0;
    // Get Balances
    for(var i = 0; i < def.cat[3].length + 1; i++) { // For external accounts 
      var name = allMonths[lastIndex].getName();
      bal[k] = allMonths[lastIndex].getRange(3, 8 + i).getValue();
      k++;
    }
    for(var i = 0; i < def.cat[2].length; i++) { // For investments 
      for(var j = 0; j < 2; j++) { // balance and mkt value
        bal[k] = allMonths[lastIndex].getRange(10 + j, 12 + i).getValue();
        k++;
      }
    }
    // Populate Balances
    for(var i = 0; i < (def.cat[3].length + 1 + 2*def.cat[2].length); i++) { // For each balance
      wealth.getRange(3 + i, 2).setValue(bal[i]);
    }
  }
}

function computeTotals() {
  // Select Active Spreadsheet
  var wealth = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Wealth Check
  if(validWealthCheck(wealth)) {
    // Start Sum
    var k = 0;
    var nom_sum = 0;
    for(var i = 0; i < def.cat[3].length + 1; i++) { // External Accounts
      nom_sum = nom_sum + wealth.getRange(3 + k, 2).getValue();
      k++;
    }
    // Nominal Sum = Market Sum when only adding external accounts
    var mkt_sum = nom_sum;
    for(var i = 0; i < def.cat[2].length; i++) {
      nom_sum = nom_sum + wealth.getRange(3 + k, 2).getValue();
      mkt_sum = mkt_sum + wealth.getRange(3 + (k + 1), 2).getValue();
      k = k + 2;
    }
    // Set Totals
    wealth.getRange(3, 6).setValue(nom_sum);
    wealth.getRange(3, 7).setValue(mkt_sum);
  }
}

function computePct_W() {
  // Select Active Spreadsheet
  var wealth = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Wealth Check
  if(validWealthCheck(wealth)) {
    // Get Totals
    var nom_sum = wealth.getRange(3, 6).getValue();
    var mkt_sum = wealth.getRange(3, 7).getValue();
    var k = 0;
    for(var i = 0; i < def.cat[3].length + 1; i++) { // External Accounts
      // Account Balance
      var bal = wealth.getRange(3 + k, 2).getValue();
      // Nominal Pct
      var nom_pct = bal/nom_sum*100;
      wealth.getRange(3 + k, 3).setValue(nom_pct);
      // Market Pct
      var mkt_pct = bal/mkt_sum*100;
      wealth.getRange(3 + k, 4).setValue(mkt_pct);
      k++;
    }
    for(var i = 0; i < def.cat[2].length; i++) { // Investments
      // Account Balance
      var bal_nom = wealth.getRange(3 + k, 2).getValue();
      var bal_mkt = wealth.getRange(3 + (k + 1), 2).getValue();
      // Nominal Pct
      var nom_pct = bal_nom/nom_sum*100;
      wealth.getRange(3 + k, 3).setValue(nom_pct);
      wealth.getRange(3 + k, 4).setValue("--");
      // Market Pct
      var mkt_pct = bal_mkt/mkt_sum*100;
      wealth.getRange(3 + (k + 1), 3).setValue("--");
      wealth.getRange(3 + (k + 1), 4).setValue(mkt_pct);
      k = k + 2;
    }
  }
}

function wealthFormat() {
  // Select Active Spreadsheet
  var wealth = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Wealth Check
  if(validWealthCheck(wealth)) {
    var num_rows = def.cat[3].length + 1 + 2*def.cat[2].length;
    // Alignment
    var k = def.cat[3].length + 1;
    for(var i = 0; i < def.cat[2].length; i++) { // Investments
      // Alignment
      wealth.getRange(3 + (k + 1), 3).setHorizontalAlignment("right"); // Percentages
      wealth.getRange(3 + k, 4).setHorizontalAlignment("right"); // Percentages
      k = k + 2;
    }
    // Borders
    wealth.getRange(3, 1, 100, 4).setBorder(true, false, false, false, false, false); // Clear Borders
    wealth.getRange(3, 1, num_rows, 4).setBorder(true, true, true, true, false, false);
    wealth.getRange(3, 2, num_rows, 1).setBorder(true, true, true, true, false, false);
    wealth.getRange(3 + (def.cat[3].length + 1), 1, 2*def.cat[2].length, 4).setBorder(true, true, true, true, false, false);
    wealth.getRange(3 + (def.cat[3].length + 1), 2, 2*def.cat[2].length, 1).setBorder(true, true, true, true, false, false);
  }
}

/***************************************************************************************************/
/*                                                                                Wealth Utilities */
/***************************************************************************************************/
function validWealthCheck(sht) {
  // Initialize
  var name = sht.getName();  
  // Comparisions
  var isWealth = (name == "Wealth");
  var isNotWealth = (name == "Cat" || name == "Accts" || name == "EOY" || name == "Jan" || name == "Feb" || name == "Mar" || name == "Apr" || name == "May" || name == "Jun" || name == "Jul" || name == "Aug" || name == "Sep" || name == "Oct" || name == "Nov" || name == "Dec" );  
  // Return
  return (isWealth && !isNotWealth);
}