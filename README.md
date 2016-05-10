// These codes are for inventory managment between a Master sheet and slave sheet.
// It allows for statuses of inventory items to be changed, which color codes the master and populates the slave
// Written mostly by Taylor Adams with moral/logical support and vision casting by Luke Manning

function sortfunction() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetS = ss.getSheets()[1];
  var lastRowS = sheetS.getLastRow();
  var lastColumnS = sheetS.getLastColumn();
  var UPCS = sheetS.getRange(2, 1, lastRowS, lastColumnS);
  
  UPCS.sort(1);
}  

function cellModifier() {
  
  //NOTES
  // The denoted last column variable must be the column -1 that contains the variable. eg "Sold", "Pending",etc
  // The colors can be changed to suit the users needs and more variables can be added with little effort
  // Anything about the cell and text within the cell can be modified
  
  //Defines which sheet to run program on. Must be Sheet #1
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  
  var lastRow = sheet.getLastRow(); //Defines active range variables 
  var lastColumn = 14; //Column that the program checks for the 'Sold' string. A=1; B=2, etc.
  
  var active_row = 1; //Dont change this, it makes it start on row #1
  
  while (active_row <= lastRow) {
    var cell2test = sheet.getRange(active_row, lastColumn); //
    
    
    //Settings for "Sold" items
    if (cell2test.getValue() == "Sold") {
     var change_range = sheet.getRange(active_row, 1, 1, lastColumn);
      change_range.setBackground("White"); //function will accept color names or their hex equivilent
      change_range.setFontStyle("italic"); //normal, bold, italic
      change_range.setFontColor("#d3d3d3");} //names and hexidecimal
      continue;
    //change_range.setFontLine("line-through")}
    
    //Settings for "In Progress" items
    if (cell2test.getValue() == "In Progress") {
     var change_range = sheet.getRange(active_row, 1, 1, lastColumn);
      change_range.setBackground("#f6f0ba");
      change_range.setFontStyle("normal");
      change_range.setFontColor("Black");}
      continue;
    
    //Settings for "For Sale" items
    if (cell2test.getValue() == "For Sale") {
     var change_range = sheet.getRange(active_row, 1, 1, lastColumn);
      change_range.setBackground("#97e197");
      change_range.setFontStyle("normal");
      change_range.setFontColor("Black");}
      continue;
    
    //Settings for "Rented" items
    if (cell2test.getValue() == "Rented") {
     var change_range = sheet.getRange(active_row, 1, 1, lastColumn);
      change_range.setBackground("#add8e6");
      change_range.setFontStyle("normal");
      change_range.setFontColor("Black");}
      continue;
   
    //Settings for "Liquidated" items
    if (cell2test.getValue() == "Liquidated") {
     var change_range = sheet.getRange(active_row, 1, 1, lastColumn);
      change_range.setBackground("White");
      change_range.setFontStyle("italic"); 
      change_range.setFontColor("#d3d3d3");} 
      //change_range.setFontLine("line-through")}
      continue;
    
    //Settings for "Unlocking" items
    if (cell2test.getValue() == "Unlocking") {
     var change_range = sheet.getRange(active_row, 1, 1, lastColumn);
      change_range.setBackground("#ff7f7f");
      change_range.setFontStyle("normal");
      change_range.setFontColor("Black");}
      continue;
    
    
    active_row++ ;} //close while loop - iterates the 'active_row' variable
}

// This function will write and erase data from first sheet to the second sheet based on statuses from first sheet.
// The goal here was to make a "For Sale" sheet which could be published and embedded
// It currently Works very well with one client. 

function copypasta() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheets()[0];
  var lastRow1 = sheet1.getLastRow(); //Defines active range variables 
  var lastColumn1 = 14; //Column that the program checks for the 'Sold' string. A=0; B=1, etc.
  var active_row1 = 1; //Dont change this, it makes it start on row #1
  var active_row2 = 1;
  
  var sheet2 = ss.getSheets()[1];
  
  while (active_row1 <= lastRow1) {
    var cell2test1 = sheet1.getRange(active_row1, lastColumn1); 
    var lastRow2 = sheet2.getLastRow();
    var FirstEmptyRow = lastRow2+1;
    var endCopyColumn = 9; 
    var InSlaveColumn = 18;
    var InSlaveValue = sheet1.getRange(active_row1, InSlaveColumn).getValue();
    var InSlaveCell = sheet1.getRange(active_row1, InSlaveColumn);
    
    if (cell2test1.getValue() == "For Sale") {
      
      if (InSlaveValue == 0) {     
        var storage = sheet1.getRange(active_row1, 1, 1, endCopyColumn).getValues();
      
        var range2paste = sheet2.getRange(FirstEmptyRow, 1, 1, endCopyColumn)
        range2paste.setValues(storage);
        sheet1.getRange(active_row1, InSlaveColumn).setValue(1);
        sortfunction();
      }
    }
    
    if (cell2test1.getValue() == "Sold" && InSlaveValue == 1) {
      var UPC1 = sheet1.getRange(active_row1, 1).getValue();
      
      while (active_row2 <= lastRow2) {
        var UPC2 = sheet2.getRange(active_row2, 1).getValue();
        if (UPC2 == UPC1) {
          sheet2.deleteRow(active_row2);
          InSlaveCell.setValue(0);
          sortfunction();
          break;
        }
        
      active_row2++; }
    }
    
  active_row1++; }
}
    

  
    
    
    
    
    
    
    
    
    
