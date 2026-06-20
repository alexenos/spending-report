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
  var spendReport = SpreadsheetApp.getActiveSpreadsheet();
  var monthMenu = [ {name: "Monthly Spending",   functionName: "monthSpendAsmt"}, 
                    {name: "Clear Data",         functionName: "clearMonthData"},
                    {name: "Calculate Balances", functionName: "transacBalances"},
                    {name: "Calculate Totals",   functionName: "monthlyTotals"},
                    {name: "Update Categories",  functionName: "updateCategories"},
                    {name: "Category Spending",  functionName: "monthlyCatSpending"},
                    {name: "Format",             functionName: "monthlyFormat"}
                  ];                    
  var eoyMenu   = [ {name: "EOY Spending",       functionName: "eoySpendAsmt"},
                    {name: "Clear Data",         functionName: "clearEOYData"},
                    {name: "Monthly Totals",     functionName: "getMonthlyTotals"},
                    {name: "Update Categories",  functionName: "updateEOYCategories"},
                    {name: "Category Spending",  functionName: "getEOYCategoryTotals"},
                    {name: "EOY Totals",         functionName: "eoyTotals"},
                    {name: "Format",             functionName: "eoyFormat"}
                  ];
  var utlMenu   = [ {name: "Wealth Assessment",  functionName: "wealthAsmt"},
                    {name: "Itemize Category",   functionName: "itemizeCat"}
                  ];
  
  spendReport.addMenu("Month", monthMenu);
  spendReport.addMenu("Eoy",   eoyMenu);
  spendReport.addMenu("Utl",   utlMenu);  
}