

// function to create and add Menu in spreadsheet
function onOpenSpreadsheet() {
  // get the active worksheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // generate sub menu and add it to menu
  var menuEntries = [{name: "Add Employee one by one", functionName: "addEmpSheet"},
    {name: "Add Employee's at once (FIFO)", functionName: "addAllEmpSheets"},
    {name: "Add Employee's at once (LIFO)", functionName: "addAllEmpSheetsLIFO"},
    {name: "Delete single Sheets", functionName: "deleteEmpSheet"},
    {name: "Delete Sheets", functionName: "delAllEmpSheets"}
  ];
  ss.addMenu("Maintenance", menuEntries);
}


// function to add individual employee sheet
// Place curson on employee row and select 'Add Employee one by one' from 'Maintenance'
function addEmpSheet() {
  //get active worksheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //get active sheet
  var sh = ss.getActiveSheet();
  // get the row number of active cell
  var row = sh.getActiveRange().getRowIndex();
  // get the data of active row
  var rData = sh.getRange(row, 1, 1, 3).getValues();
 
  // check the active row should not be header row
  if (row == 1) {
    ss.toast("This is the header");    
    return
  }

  // check for data in active row, if there is no data in first 3 cell end with error msg
  if(rData[0][0] != null || rData[0][1] != null || rData[0][2] != null) {
    try {      
      // get template sheet (sheet at index 1)
      var sheet = ss.getSheets()[1];
      // copy template sheet to active workbook
      var sh2 =sheet.copyTo(ss);
      // set new sheet name to first cell in active row i.e, Name of employee
      sh2.setName(rData[0][0]);
      //Browser.msgBox(rData[0][1]);
      
      // logic to add protection in sheet, 
      // other employee cannot edit that sheet
      var protection = sh2.protect().setDescription('Sample protected sheet');
      // add current employee mail id as editor of sheet
      protection.addEditor(rData[0][1]);
      
      //ss.insertSheet(rData[i][2]);
      
      // set first sheet as active  and update current date in the D column of selected employee
      ss.setActiveSheet(ss.getSheets()[0]);
      sh.getRange("D"+(row)).setValue(new Date());
    } catch(e) {
      throw 'This employee allready has a sheet. Try another sheet name.'+ e;      
    }
  }
}


// Similarly add Sheets for all employee
function addAllEmpSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getActiveSheet();
  // read available data in sheet and loop to add sheet for all
  var rData = sh.getDataRange().getValues();

  var message = [];    
  for(var i=1, len=rData.length; i<len; i++) {
    if(rData[i][3] == null || rData[i][3] == "") {   
      if(rData[i][0] != null || rData[i][1] != null || rData[i][2] != null) {
        try {    
          var sheet = ss.getSheets()[1];
          var sh2 =sheet.copyTo(ss);
          sh2.setName(rData[i][0]);
          //Browser.msgBox(rData[i][1]);
          var protection = sh2.protect().setDescription('Sample protected sheet');
          protection.addEditor(rData[i][1]);
          
          //ss.insertSheet(rData[i][2]);
          
          
          ss.setActiveSheet(ss.getSheets()[0]);
          sh.getRange("D"+(i+1)).setValue(new Date());
        } catch(e) {
          message.push("row " + (i+1));
        }
      }    
    }
  }
  ss.toast("These sheets allready exist: " + message);
  ss.setActiveSheet(ss.getSheets()[0]);
}



// Add Sheets for all employee in Reverse order
function addAllEmpSheetsLIFO() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getActiveSheet();
  var rData = sh.getDataRange().getValues();

  var message = [], i=rData.length; 
  while(i--) {
    if(rData[i][3] == null || rData[i][3] == "") {   
      if(rData[i][0] != null || rData[i][1] != null || rData[i][2] != null) {
        try {      
          var sheet = ss.getSheets()[1];
          var sh2 =sheet.copyTo(ss);
          sh2.setName(rData[i][0]);
          //ss.insertSheet(rData[i][2]);
          ss.setActiveSheet(ss.getSheets()[0]);
          sh.getRange("D"+(i+1)).setValue(new Date());
        } catch(e) {
          message.push("row " + (i+1));
        }
      }    
    }
  }
  ss.toast("These sheets allready exist: " + message);
  ss.setActiveSheet(ss.getSheets()[0]);
}

// Delete Active row employee sheet
function deleteEmpSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // get active sheet
  var sh = ss.getActiveSheet();
  // get active row number
  var row = sh.getActiveRange().getRowIndex();
  var rData = sh.getRange(row, 1, 1, 3).getValues();
    if (row == 1) {
    ss.toast("This is the header");    
    return
  }
  // read employee name to delete that sheet
  var shName = rData[0][0];
  // search sheet by name
  ss.setActiveSheet(ss.getSheetByName(shName));
  // delete sheet
  ss.deleteActiveSheet();
  
  // delete date from that row
  ss.setActiveSheet(ss.getSheets()[0]);
  ss.getRange("D"+(row)).clear(); 
  
}


// delete all employee sheet
function delAllEmpSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shs = ss.getNumSheets();

  for(var i=shs-1;i>1;i--){
    ss.setActiveSheet(ss.getSheets()[i]);
    ss.deleteActiveSheet();
  }
  ss.setActiveSheet(ss.getSheets()[0]);
  ss.getRange("D2:D").clear();
}
