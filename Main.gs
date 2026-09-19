// Dax Garner
// 10-3-2012
// SpendingReport Version 4: Multiple Accounts

/***************************************************************************************************/
/*                                                                              Global Definitions */
/***************************************************************************************************/

// Category Definitions
var def = getDef();

/***************************************************************************************************/
/*                                                                                  Main Functions */
/***************************************************************************************************/
function onOpen() { //                                                             Menu Defintions //
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Month")
    .addItem("Monthly Spending",   "monthSpendAsmt")
    .addItem("Clear Data",         "clearMonthData")
    .addItem("Calculate Balances", "transacBalances_M")
    .addItem("Calculate Totals",   "monthlyTotals")
    .addItem("Update Categories",  "updateCategories_M")
    .addItem("Category Spending",  "monthlyCatSpending")
    .addItem("Format",             "monthlyFormat")
    .addToUi();
  ui.createMenu("Eoy")
    .addItem("EOY Spending",       "eoySpendAsmt")
    .addItem("Clear Data",         "clearEOYData")
    .addItem("Monthly Totals",     "getMonthlyTotals")
    .addItem("Update Categories",  "updateEOYCategories")
    .addItem("Category Spending",  "getEOYCategoryTotals")
    .addItem("EOY Totals",         "eoyTotals")
    .addItem("Format",             "eoyFormat")
    .addToUi();
  ui.createMenu("Utl")
    .addItem("Wealth Assessment",  "wealthAsmt")
    .addItem("Itemize Category",   "itemizeCat")
    .addToUi();
}