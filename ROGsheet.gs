function updateROGSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rogSheet        = ss.getSheetByName("Sample ROG");
  const trainingSheet   = ss.getSheetByName("Sample v2");
  const attendanceSheet = ss.getSheetByName("Sample2 v2");

  if (!rogSheet)        throw new Error("Sample ROG sheet not found.");
  if (!trainingSheet)   throw new Error("Sample v2 sheet not found.");
  if (!attendanceSheet) throw new Error("Sample2 v2 sheet not found.");

  // ── TRAINING MAP: courseCode → training hours ─────────────────
  // Sample v2: A[0]=Course Code, J[9]=No. of Training Hours
  const trainingData = trainingSheet
    .getRange(2, 1, trainingSheet.getLastRow() - 1, trainingSheet.getLastColumn())
    .getValues();

  const trainingMap = {};

  trainingData.forEach(row => {
    const courseCode = String(row[0] || "").trim();            // Column A
    const rawHours   = String(row[9] || "").replace(/[^\d.]/g, ""); // Column J
    const hours      = parseFloat(rawHours);

    if (courseCode && !isNaN(hours)) {
      trainingMap[courseCode] = hours;
    }
  });

  Logger.log("Training map: " + Object.keys(trainingMap).length + " entries");

  // ── ATTENDANCE MAP: traineeId_trainingCode → hours attended ───
  // Sample2 v2: B[1]=Trainee ID, N[13]=Training Code, P[15]=Training Hours
  const attendanceData = attendanceSheet
    .getRange(2, 1, attendanceSheet.getLastRow() - 1, attendanceSheet.getLastColumn())
    .getValues();

  const attendanceMap = {};

  attendanceData.forEach(row => {
    const traineeId    = String(row[1]  || "").trim();  // Column B — Trainee ID
    const trainingCode = String(row[13] || "").trim();  // Column N — Training Code
    const rawHours     = String(row[15] || "").replace(/[^\d.]/g, ""); // Column P
    const hours        = parseFloat(rawHours);

    if (!traineeId || !trainingCode || isNaN(hours)) return;

    const key = traineeId + "_" + trainingCode;
    attendanceMap[key] = (attendanceMap[key] || 0) + hours;
  });

  Logger.log("Attendance map: " + Object.keys(attendanceMap).length + " entries");

  // ── UPDATE ROG SHEET ──────────────────────────────────────────
  // Sample ROG: A[0]=Training Code, B[1]=Trainee ID,
  //             C[2]=Total Hours Attended, D[3]=Training Hours,
  //             E[4]=Attendance(%), F[5]=Status
  const lastRow  = rogSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("⚠ ROG sheet has no data rows.");
    return;
  }

  const rogRange = rogSheet.getRange(2, 1, lastRow - 1, 6);
  const rogData  = rogRange.getValues();

  rogData.forEach((row, k) => {
    const courseCode = String(row[0] || "").trim();  // Column A
    const traineeId  = String(row[1] || "").trim();  // Column B

    if (!courseCode || !traineeId) return;

    const key          = traineeId + "_" + courseCode;
    const attended     = attendanceMap[key]    || 0;
    const totalHours   = trainingMap[courseCode] || 0;

    rogData[k][2] = attended;   // Column C — Total Hours Attended
    rogData[k][3] = totalHours; // Column D — Training Hours

    if (totalHours > 0) {
      const pct     = (attended / totalHours) * 100;
      rogData[k][4] = Math.round(pct * 100) / 100;      // Column E — Attendance (%)
      rogData[k][5] = pct >= 80 ? "Pass" : "Fail";      // Column F — Status
    } else {
      rogData[k][4] = 0;
      rogData[k][5] = "Fail";
    }
  });

  rogSheet.getRange(2, 1, rogData.length, 6).setValues(rogData);
  Logger.log("ROG updated — " + rogData.length + " rows processed.");
}