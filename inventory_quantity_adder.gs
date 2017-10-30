/*
    Inventory Quantity Adder
    
    Add the total number of items of each Product ID in "Inventory List" and output the
    quantity in "Inventory Count"
*/

var InventoryList = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inventory List");
var InventoryCount = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inventory Count");

function countItems() {

    InventoryList.getRange(4, 3, InventoryList.getLastRow()-3, 8).sort([{
        column: 3,
        ascending: true
    }]);
  
    var masterList = InventoryList.getRange(4, 3, InventoryList.getLastRow()-3, 3).getValues();
    var countList = InventoryCount.getRange(2, 1, InventoryCount.getLastRow()-1, 2).getValues();
    var lenCountList = countList.length;

    countList.sort(function(a, b) {
        return a - b
    });
    
    for (var i = 0; i < lenCountList; i++) {
        var productID = countList[i][0];
        var quantity = countList[i][1];
      
        quantity = 0;
      
        for (var j = 0; j < masterList.length; j++) {
            if (masterList[j][0] == productID) {
                quantity += Number(parseInt((masterList[j][2] / 1), 10));
            }
        }
      
        countList[i][1] = quantity;
    }
  
    Logger.log(countList);
  
    InventoryCount.getRange(2, 1, lenCountList, 2).setValues(countList);
    
    InventoryList.getRange(4, 3, InventoryList.getLastRow()-3, 8).sort([{
        column: 6,
        ascending: true
    }]);

}
