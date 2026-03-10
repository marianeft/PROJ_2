function aiFillSections(body, trainingRow) {
  if (!trainingRow || trainingRow.length < 10) {
    throw new Error("Invalid training data passed to AI function.");
  }

  const tz = Session.getScriptTimeZone();

  const data = {
    courseCode: safe(trainingRow[0]),
    province:   safe(trainingRow[4]),
    owner:      safe(trainingRow[5]),
    startDate:  formatDate(trainingRow[6], tz),
    endDate:    formatDate(trainingRow[7], tz),
    status:     safe(trainingRow[8]),
    hours:      safe(trainingRow[9]),
    title:      safe(trainingRow[10])
  };

  validateAIInputs(data);

  // ── Pull objectives & topics strictly from database ───────────────
  const poiData = getReportDetails(data.title);

  // Logger.log("=== TITLE LOOKUP DEBUG ===");
  // Logger.log("Course Code   : [" + data.courseCode + "]");
  // Logger.log("Training Title: [" + data.title + "]");
  // Logger.log("POI Match     : " + (poiData ? "YES " : "NO "));
  // Logger.log("==========================");

  // Hard stop — do not proceed without DB data
  if (!poiData) {
    throw new Error(
      "No POI match found for training: \"" + data.title + "\". " +
      "Please add this training to the 'Sample LIST OF POIs' sheet before generating a report."
    );
  }
  // ─────────────────────────────────────────────────────────────────

  // ── Ask Gemini only for rationale, issues, recommendations ───────
  const prompt = buildPrompt(data);

  const apiKey = PropertiesService.getScriptProperties().getProperty("API");
  if (!apiKey) throw new Error("Gemini API key not found in Script Properties.");

  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=" + apiKey;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.3,
      maxOutputTokens: 2000,
      responseMimeType: "application/json"
    }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  let aiJson;

  for (let attempt = 1; attempt <= 2; attempt++) {
    const response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() !== 200) {
      throw new Error("Gemini API Error: " + response.getContentText());
    }

    const raw = JSON.parse(response.getContentText());
    const aiTextRaw = raw?.candidates?.[0]?.content?.parts?.[0]?.text;

    Logger.log("=== GEMINI RAW (Attempt " + attempt + ") ===");
    Logger.log(aiTextRaw ? aiTextRaw.substring(0, 300) : "EMPTY");
    Logger.log("========================================");

    if (!aiTextRaw) {
      if (attempt === 2) throw new Error("AI returned empty response after retry.");
      continue;
    }

    const aiTextClean = sanitizeAIResponse(aiTextRaw);

    try {
      aiJson = JSON.parse(aiTextClean);
      validateAIOutput(aiJson);
      break;
    } catch (err) {
      Logger.log("AI JSON Parse Failed (Attempt " + attempt + "): " + err.message);
      if (attempt === 2) throw new Error("Invalid JSON returned by Gemini after retry.");
    }
  }

  // ── Replace placeholders ─────────────────────────────────────────
  replaceSafe(body, "{{Rationale}}",       aiJson.rationale);
  replaceSafe(body, "{{Objectives}}",      poiData.objectives);     // always from DB
  replaceSafe(body, "{{Topics Covered}}",  poiData.topicsCovered);  // always from DB
  replaceSafe(body, "{{Issues Concerns}}", aiJson.issuesConcerns);
  replaceSafe(body, "{{Recommendation}}",  aiJson.recommendations);

  Logger.log("Report sections filled for " + data.courseCode);
  Logger.log("Objectives source : DATABASE ");
  Logger.log("Topics source     : DATABASE ");
  Logger.log("Rationale source  : AI");
}