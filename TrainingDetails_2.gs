// TRAINING DETAILS SECTION
function fillTrainingDetails(body, values) {
  const tz = Session.getScriptTimeZone();

  replaceSafe(body, "{{Training Title}}", values[10]); // Column K
  replaceSafe(body, "{{Course Code}}", values[0]); // Column A
  replaceSafe(body, "{{Date}}", formatDateRange(values[6], values[7], tz)); // Column G - H
  replaceSafe(body, "{{Resource Person}}", values[13]); // Column N
  replaceSafe(body, "{{Mode of Conduct}}", values[14]); // Column O
  replaceSafe(body, "{{No. of Training Hours}}", values[9]); // Column J
  replaceSafe(body, "{{Sector}}", values[15]); // Column P
  //replaceSafe(body, "{{ICT Track}}", values[12]); // Column M
  replaceSafe(body, "{{Sector}}", values[15]); // Column P
  // replaceSafe(body, "{{Training ID}}", values[]); // Column 
  // replaceSafe(body, "{{Training Owner}}", values[]); // Column 
}

// GENDER SECTION
function fillGenderSection(body, controlNumber, data) {
  const seen = new Set();
  let male = 0;
  let female = 0;

  data.forEach((row, i) => {
    if (i === 0) return;

    const trainee = (row[2] || "").toString().trim();
    const code    = String(row[13] || "").trim();

    if (!trainee || code !== controlNumber) return;

    const key = trainee.toLowerCase() + "_" + code;
    if (seen.has(key)) return;
    seen.add(key);

    if (typeof row[6] === "string") { // Column G Gender
      const gender = row[6].trim().toLowerCase();
      if (gender === "male")        male++;
      else if (gender === "female") female++;
    }
  });

  replaceSafe(body, "{{Total Attendants Male}}",   male);
  replaceSafe(body, "{{Total Attendants Female}}", female);
  replaceSafe(body, "{{Total Attendants}}",        male + female);
}

//SECTOR SECTION
function fillSectorSection(body, controlNumber, data) {
  const seen = new Set();

  const totals = {
    LGU: 0,
    NGA: 0,
    PWD: 0,
    Others: 0
  };

  data.forEach((row, i) => {
    if (i === 0) return;

    const trainee = (row[1] || "").toString().trim();
    const code = String(row[13] || "").trim(); 
    const sectorCell = (row[8] || "").toString(); // Column I Sector

    if (!trainee || code !== controlNumber) return;

    const key = trainee.toLowerCase() + "_" + code;
    if (seen.has(key)) return;
    seen.add(key);

    const uniqueSectors = [...new Set(
      sectorCell.split(",").map(s => s.trim())
    )];

    uniqueSectors.forEach(sec => {
      if (totals.hasOwnProperty(sec)) totals[sec]++;
      else if (sec) totals.Others++;
    });
  });

  replaceSafe(body, "{{Total LGU}}", totals.LGU);
  replaceSafe(body, "{{Total NGA}}", totals.NGA);
  replaceSafe(body, "{{Total PWD}}", totals.PWD);
  replaceSafe(body, "{{Total Others}}", totals.Others);
}

// PLACEHOLDER ERROR LOGGING
function logMissingPlaceholders(body) {
  const text = body.getText();
  const regex = /\{\{.*?\}\}/g;
  const matches = text.match(regex);

  if (matches && matches.length > 0) {
    Logger.log("Unreplaced placeholders detected:");
    matches.forEach(m => Logger.log(" - " + m));
  }
}   

// DATE FORMATTER
function formatDateRange(startDate, endDate, tz) {
  if (!startDate) return "";

  const start = new Date(startDate);
  const end = endDate ? new Date(endDate) : null;

  const startFormatted = Utilities.formatDate(start, tz, "MMMM d, yyyy");

  if (!end || start.getTime() === end.getTime()) {
    return startFormatted;
  }

  const endFormatted = Utilities.formatDate(end, tz, "MMMM d, yyyy");
  return startFormatted + " to " + endFormatted;
}

// FOLDER CREATION
function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}   