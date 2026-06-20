// Dax Garner
// 10-3-2012
// SpendingReport Version 4: Multiple Accounts

/***************************************************************************************************/
/*                                                                              Global Definitions */
/***************************************************************************************************/

/***************************************************************************************************/
/*                                                                            Definition Functions */
/***************************************************************************************************/

function getDef() { //     Get Definitions: Gets the category names and abbrv. from the Def sheet. //
  // Def Sheets
  var catSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cat");
  var acctSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Accts");
  // Instantiate
  var cat = new Array();
  var nam = new Array();  
  // Define Categories
  cat[0] = defCycle(catSheet,  5); // Income Categories            (Column E)
  cat[1] = defCycle(catSheet,  2); // Spending Categories          (Column B)
  cat[2] = defCycle(acctSheet, 5); // Investment Categories        (Column E) 
  cat[3] = defCycle(acctSheet, 2); // External Account Categories  (Column B)
  // Define Category Names
  nam[0] = defCycle(catSheet,  4); // Income Categories            (Column D)
  nam[1] = defCycle(catSheet,  1); // Spending Categories          (Column A)
  nam[2] = defCycle(acctSheet, 4); // Investment Categories        (Column D)
  nam[3] = defCycle(acctSheet, 1); // External Account Categories  (Column A)
  return {cat : cat, name : nam}; 
}

function defCycle(defSheet, col) { //   Definition Sheet Cycle: Returns all values within a specified column //
  // Instantiate
  var defVal = new Array();
  // Initialize While
  var row = 2; var cnt = 0;
  // While Loop
  while(defSheet.getRange(row, col).getValue() != "") { // While there is a value
    defVal[cnt] = defSheet.getRange(row, col).getValue();
    row++; cnt++;
  }
  // Return
  return defVal;
}

function getNumCat(def) { //         Get Total Number of Categories: Returns number of categories. //
  // Initialize
  var num = 0;
  // Count categories
  for(var i = 0; i < 2; i++) { // For each category type
    num = num + def.cat[i].length;
  }
  // Return
  return num;
}

function defCompare(trn, cat) { //     Compares the category in to the categories within the type. //
  // Initialize
  var isCatType = false; // Default to false
  for(var i = 0; i < cat.length; i++) { // For each category defined in the category type
      if(trn == cat[i]) { // If the transaction category matches one of the categories
        isCatType = true;
      } // end if
  } // end for
  // Return
  return isCatType;
}

function returnAcctIdx(trn, cat) { //     Compares the category in to the accounts. //
  // Initialize
  var AcctIdx = 0; // Default to Checking
  for(var i = 0; i < cat.length; i++) { // For each category defined in the category type
      if(trn == cat[i]) { // If the transaction category matches one of the categories
        AcctIdx = i + 1; // Return index
      } // end if
  } // end for
  // Return
  return AcctIdx;
}
    
function isIncome(trn, def) { //     Is Income? : Determines if the category is an income category. //
  var isInc = defCompare(trn, def.cat[0]);
  return isInc;
}

function isExpense(trn, def) { //  Is Expense? : Determines if the category is an expense category. //
  var isExp = defCompare(trn, def.cat[1]);
  return isExp;
}
  
function isAcct(trn, def) { // Is Account? : Determines if the category is an account utility. //
  var isAcct = defCompare(trn, def.cat[3]);
  return isAcct;
}

function isInvest(trn, def) { // Is Investment? : Determines if the category is an investment . //
  var isInv = defCompare(trn, def.cat[2]);
  return isInv;
}

function whichAccount(trn, def) { // Which Account? : Returns account index. //
  var AcctIdx = returnAcctIdx(trn, def.cat[3]);
  return AcctIdx;
}

function AccountInfo(trn, def) { // Is, Which Account and what transaction
  var cat = new String();
  var type = new String();
  cat = trn.substring(0, trn.length - 2);
  var isAcct = defCompare(cat, def.cat[3]);
  var AcctIdx = returnAcctIdx(cat, def.cat[3]);
  type = trn.substring(trn.length - 2);
  return {isAcct : isAcct, idx : AcctIdx, type : type}
}

function isCat(cat, def) { // Is a valid category
  return (isIncome(cat, def) || isExpense(cat, def) || isAcct(cat, def) || isInvest(cat, def))
}