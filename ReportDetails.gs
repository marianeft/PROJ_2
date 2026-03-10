function getReportDetails(trainingTitle) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Sample LIST OF POIs");
  if (!sheet) throw new Error("Sample LIST OF POIs sheet not found.");

  const data = sheet.getDataRange().getValues();
  const normalizedTitle = trainingTitle.toString().trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    const rowTitle = String(data[i][2] || "").trim().toLowerCase(); // Column C
    if (rowTitle === normalizedTitle) {
      Logger.log("POI Match found — Row " + (i + 1) + ": [" + data[i][2] + "]");
      return {
        objectives:    String(data[i][4] || "").trim(), // Column E
        topicsCovered: String(data[i][5] || "").trim()  // Column F
      };
    }
  }

  // No match — log all available titles to help diagnose
  Logger.log("No POI match for: [" + trainingTitle + "]");
  Logger.log("Available titles in POI sheet:");
  for (let i = 1; i < data.length; i++) {
    if (data[i][2]) Logger.log("  Row " + (i + 1) + ": [" + data[i][2] + "]");
  }

  return null;
}

function fillReportDetails(body, trainingRow) {
  const title = safe(trainingRow[10]);

  const poiData = getReportDetails(title);

  Logger.log("=== REPORT DETAILS ===");
  Logger.log("Training Title: [" + title + "]");
  Logger.log("POI Match     : " + (poiData ? "YES" : "NO"));
  Logger.log("======================");

  if (!poiData) {
    throw new Error(
      "No POI match found for: \"" + title + "\". " +
      "Add it to 'Sample LIST OF POIs' before generating."
    );
  }

  replaceSafe(body, "{{Objectives}}",     poiData.objectives);
  replaceSafe(body, "{{Topics Covered}}", poiData.topicsCovered);

  Logger.log("Objectives and Topics filled from DATABASE");
}