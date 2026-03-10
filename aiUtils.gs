function safe(value) {
  return value ? String(value).trim() : "";
}

function replaceSafe(body, placeholder, value) {
  body.replaceText(placeholder, value ? String(value) : "");
}

function formatDate(dateValue, tz) {
  if (!dateValue) return "";
  return Utilities.formatDate(new Date(dateValue), tz, "MMMM d, yyyy");
}

function sanitizeAIResponse(text) {
  return text
    .replace(/```json/gi, "")
    .replace(/```/g, "")
    .trim();
}

function validateAIInputs(data) {
  const required = ["courseCode", "province", "startDate", "hours"];
  const missing = required.filter(key => !data[key]);
  if (missing.length > 0) {
    throw new Error("Missing required AI fields: " + missing.join(", "));
  }
}

function validateAIOutput(json) {
  const requiredFields = [
    "rationale",
    //"objectives",
    //"topicsCovered",
    "issuesConcerns",
    "recommendations",
    "plansActionItems"
  ];
  if (includeAllFields) {
    requiredFields.push(
    "objectives", 
    "topicsCovered"
    );
  }
  const missing = requiredFields.filter(f => !json[f]);
  if (missing.length > 0) {
    throw new Error("AI output missing fields: " + missing.join(", "));
  }
}

function buildPrompt(data) {
  return `You are generating content for an AFTER TRAINING REPORT (ATR).

CRITICAL INSTRUCTIONS:
- Your ENTIRE response must be a single valid JSON object.
- Start your response with { and end with }
- Do NOT include any text before or after the JSON.
- Do NOT use markdown, backticks, or code fences.
- Do NOT include newlines inside string values — use spaces instead.

Training Details:
- Course Code: ${data.courseCode}
- Province: ${data.province}
- Training Owner: ${data.owner}
- Dates: ${data.startDate} to ${data.endDate}
- Duration: ${data.hours} hours
- Status: ${data.status}

JSON format:
{
  "rationale": "...",
  "objectives": "...",
  "topicsCovered": "...",
  "issuesConcerns": "...",
  "recommendations": "..."
  "plansActionItems": "..."
}

Writing rules:
- Formal government report tone
- Third person
- 1–2 short paragraphs per section
- Assume a successful regional training unless status indicates otherwise
`;
}