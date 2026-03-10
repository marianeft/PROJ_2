// MAIN
function generateReportByCN(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const formSheet = ss.getSheetByName("ATR Request Log");
    const detailsSheet = ss.getSheetByName("Sample v2");
    const participantSheet = ss.getSheetByName("Sample2 v2");

    if (!formSheet) throw new Error("ATR Request Log sheet not found.");
    if (!detailsSheet) throw new Error("Sample v2 sheet not found.");
    if (!participantSheet) throw new Error("Sample2 v2 sheet not found.");

    const controlNumber = extractControlNumber(e, formSheet);
    if (!controlNumber) throw new Error("Control Number is empty.");

    const trainingData = getTrainingRow(detailsSheet, controlNumber);
    if (!trainingData) {
      throw new Error("No training found for Control Number: " + controlNumber);
    }

    const participantData = participantSheet.getDataRange().getValues();

    validateTrainingRow(trainingData);

    const folder = getOrCreateFolder(controlNumber);

    const templateId = "1qQuIoGSClhl27qoA-Tn3wUhzUCnC4Lc-IwL3-pxCets";
    const copy = DriveApp.getFileById(templateId)
      .makeCopy(controlNumber + "_Report", folder);

    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();

    fillTrainingDetails(body, trainingData);
    fillGenderSection(body, controlNumber, participantData);
    fillSectorSection(body, controlNumber, participantData);
    fillReportDetails(body, trainingData);

    replaceSafe(body, "{{Rationale}}",       "");
    replaceSafe(body, "{{Issues Concerns}}", "");
    replaceSafe(body, "{{Recommendation}}",  "");
    //aiFillSections(body, trainingData);
    //getReportDetails(trainingTitle);

    logMissingPlaceholders(body);

    doc.saveAndClose();

    Logger.log("Report successfully generated for " + controlNumber);

  } catch (err) {
    Logger.log("ERROR: " + err.message);
    throw err;
  }
}

// CONTROL NUMBER EXTRACTION
function extractControlNumber(e, formSheet) {
  if (e && e.values && e.values.length > 2) {
    return String(e.values[2]).trim();
  }
  return String(formSheet.getRange(formSheet.getLastRow(), 3).getValue()).trim();
}

// TRAINING DATA FETCH
function getTrainingRow(sheet, controlNumber) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const data = sheet
    .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .getValues();

  for (let i = 0; i < data.length; i++) {
    const code = String(data[i][0] || "").trim();
    if (code === String(controlNumber).trim()) {
      return data[i];
    }
  }
  return null;
}

// VALIDATION LAYER
function validateTrainingRow(row) {
  const requiredFields = {
    "Course Code": row[0],
    "Start Date": row[6],
    "Training Title": row[10],
    "Resource Person": row[13]
  };

  const missing = [];

  Object.keys(requiredFields).forEach(key => {
    if (!requiredFields[key]) {
      missing.push(key);
    }
  });

  if (missing.length > 0) {
    throw new Error("Missing required training fields: " + missing.join(", "));
  }
}