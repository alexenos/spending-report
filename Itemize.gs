// Dax Garner
// 10-3-2012
// SpendingReport Version 4: Multiple Accounts

/***************************************************************************************************/
/*                                                                              Global Definitions */
/***************************************************************************************************/

/***************************************************************************************************/
/*                                                                           Itemization Functions */
/***************************************************************************************************/

function itemizeCat() {
  // Select Active Spreadsheet
  var itmSht = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Itemize Sheet Check
  if(validItemSheetCheck(itmSht)) {
    // Clear current contents
    itmSht.getRange(2, 1, 100, 3).clear();
    // Get Category of interest
    var cat = itmSht.getRange(2, 5).getValue();
    if(isCat(cat, def)) {
      var allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
      var idx = getJanIndex();
      var rowI = 2;
      var sum = 0;
      for(var i = 0; i < 12; i++) { // For each month
        // Get Sheet
        var month = allSheets[idx];
        // While there is an expenditure
        rowM = 2;
        while(month.getRange(rowM, 2).getValue() != "" || month.getRange(rowM, 3).getValue() != "") {
          if(month.getRange(rowM, 3).getValue() == cat) { // transaction is of category of interest
            month.getRange(rowM, 1, 1, 2).copyTo(itmSht.getRange(rowI, 1, 1, 2));
            month.getRange(rowM, 4, 1, 1).copyTo(itmSht.getRange(rowI, 3, 1, 1));
            sum = sum + itmSht.getRange(rowI, 3, 1, 1).getValue();
            rowI++;
          }
          rowM++;
        }
        idx++;
      }
      // Print Sum
      itmSht.getRange(2, 6).setValue(sum);
    } else {
      catErrorMsg(cat);
    }
  }
}

/***************************************************************************************************/
/*                                                                           Itemization Utilities */
/***************************************************************************************************/
function validItemSheetCheck(sht) {
  // Initialize
  var name = sht.getName();  
  // Comparisions
  var isItem = (name == "Itemize");
  var isNotItem = (name == "Wealth" || name == "Cat" || name == "Accts" || name == "EOY" || name == "Jan" || name == "Feb" || name == "Mar" || name == "Apr" || name == "May" || name == "Jun" || name == "Jul" || name == "Aug" || name == "Sep" || name == "Oct" || name == "Nov" || name == "Dec" );  
  // Return
  return (isItem && !isNotItem);
}