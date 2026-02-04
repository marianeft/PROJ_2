// ATR GSHEET TOOL
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Training Tools")
    .addItem("Generate Report by Control Number", "generateReportByControlNumber")
    .addToUi();
}

//Prompt user for control number and generate report
function generateReportByControlNumber() {
  try {
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt("Enter Training Control Number (e.g., 2025DL_Ilo_01):");
    var controlNumber = response.getResponseText().trim();

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sample"); 
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

    for (var i = 0; i < data.length; i++) {
      if (data[i][5] == controlNumber) { // Column F = Training Code
        
        //Create or reuse folder
        var folders = DriveApp.getFoldersByName(controlNumber);
        var folder;
        if (folders.hasNext()) {
          folder = folders.next();
        } else {
          folder = DriveApp.createFolder(controlNumber);
        }

        //Make copy of the tmeplate Document
        var templateId = "1qQuIoGSClhl27qoA-Tn3wUhzUCnC4Lc-IwL3-pxCets";
        var copy = DriveApp.getFileById(templateId).makeCopy(controlNumber + "_Report", folder);

        // Open Google Doc template and replace placholders
        var doc = DocumentApp.openById(copy.getId());
        var body = doc.getBody();


        //Fill placehodler with values
        trainingDB(body, data[i]);
        trainingDBGender(body, controlNumber);
        trainingDBSector(body, controlNumber);

        doc.saveAndClose();        

        ui.alert("Report generated for " + controlNumber);
        Logger.log("Report generated for " + controlNumber);
        return;
      }
    }

    ui.alert("Training with Control Number " + controlNumber + " not found.");
  } catch (err) {
    Logger.log("Error: " + err.message);
  }
}


// Sample 1 (Training Info)

function trainingDB(body, values) {

  // Replace placeholders with values
  body.replaceText("{{Training Title}}", values[10]); // Column K
  body.replaceText("{{Course Code}}", values[5]); // Column F
  body.replaceText("{{Date}}", values[6]); // Column G
  // body.replaceText("{{Time}}", values[]); // Column 
  // body.replaceText("{{Duration}}", values[]); // Column 
  // body.replaceText("{{Venue}}", values[]); // Column 
  // body.replaceText("{{Train Specialist}}", values[]); // Column 
  // body.replaceText("{{Train Support}}", values[]); // Column 
  body.replaceText("{{Resource Person}}", values[13]); // Column N
  //body.replaceText("{{Status}}", values[8]);    // Column I
  //body.replaceText("{{Platform}}", values[]); // Column 
  body.replaceText("{{Mode of Conduct}}", values[14]);    // Column I
  //body.replaceText("{{No. of Training Hours}}", values[9]);    // Column J
  //body.replaceText("{{Total}}", values[16]);    // Column Q
  //body.replaceText("{{Sector}}", values[15]);    // Column P    

  Logger.log("Training info updated for " + values[5]);
}

// Sample 2 (Gender Section)

function trainingDBGender(body) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sample 2");
  if (!sheet) throw new Error("Sheet not found.");

  var data = sheet.getDataRange().getValues();
  var filtered = data.filter((row, i) => i > 0 && row[1]);

  var maleCount = 0;
  var femaleCount = 0;

  // Gender Filter
  filtered.forEach(row => {
    var gender = row[2];
    if (typeof gender === "string") {
    gender = gender.trim().toLowerCase();
    if (gender === "male") {
      maleCount++;
    } else if (gender === "female") {
      femaleCount++;
    }
    }
  });
  
  // Replace placeholder values
  body.replaceText("{{Total Attendants Male}}", String(maleCount));
  body.replaceText("{{Total Attendants Female}}", String(femaleCount));
  body.replaceText("{{Total Attendants}}", String(maleCount + femaleCount));

  Logger.log("Document Gender Section Updated" );
}

// Sample 2 (Sector Section )

function trainingDBSector(body, controlNumber) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sample 2");
  if (!sheet) throw new Error("Sheet not found.");

  var data = sheet.getDataRange().getValues();
  var filtered = data.filter((row, i) => i > 0 && row[1]);

  // Initialize sectors counters
  var sectorTotals = {
    LGU: 0,
    NGA: 0,
    PWD: 0,
    Others: 0
  };
  
  // Sector Filter
  filtered.forEach(row => {
    var sectorRow = row[3];
    if (typeof sectorRow === "string") {
      var sectors = sectorRow.split(",").map(s => s.trim());
      sectors.forEach(sec => {
        if (sectorTotals.hasOwnProperty(sec)) {
          sectorTotals[sec]++;
          } else {
            sectorTotals.Others++; // catch-all bucket
          }
        });
      }
  });

  // Replace placeholder values
  body.replaceText("{{Total LGU}}", sectorTotals.LGU);
  body.replaceText("{{Total NGA}}", sectorTotals.NGA);
  body.replaceText("{{Total Others}}", sectorTotals.Others);
  body.replaceText("{{Total PWD}}", sectorTotals.PWD);

  Logger.log("Document Sector Updated" );
}