const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");

const DOCX_MIME =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const DEFAULT_VALUE = "待教师补充";

const FIELD_NAMES = [
  "lesson_no",
  "hours",
  "teaching_content",
  "key_points",
  "difficult_points",
  "time_1_content",
  "time_1_minutes",
  "time_2_content",
  "time_2_minutes",
  "time_3_content",
  "time_3_minutes",
  "time_4_content",
  "time_4_minutes",
  "time_5_content",
  "time_5_minutes",
  "org_course_intro",
  "org_lead_in",
  "org_content_1",
  "org_content_2",
  "org_practice_interaction",
  "org_summary",
  "homework_report",
  "teacher_signature_date"
];

function getTemplatePath() {
  return path.join(process.cwd(), "templates", "lesson_plan_template.docx");
}

function parsePayload(body) {
  if (!body) return {};
  if (typeof body === "object" && !Buffer.isBuffer(body)) return body;
  if (typeof body === "string") {
    const trimmed = body.trim();
    return trimmed ? parseJsonLenient(trimmed) : {};
  }
  return parseJsonLenient(body.toString("utf8"));
}

function parseJsonLenient(text) {
  const candidates = [];
  const trimmed = text.trim();
  candidates.push(trimmed);

  const withoutFences = trimmed
    .replace(/```(?:json|JSON)?/g, "")
    .replace(/```/g, "")
    .trim();
  if (withoutFences !== trimmed) candidates.push(withoutFences);

  const extracted = extractFirstJsonObject(withoutFences);
  if (extracted && !candidates.includes(extracted)) candidates.push(extracted);

  for (const candidate of candidates) {
    try {
      return JSON.parse(candidate);
    } catch (_) {
      // Try the next repair candidate.
    }
  }

  const preview = trimmed.slice(0, 300).replace(/\s+/g, " ");
  throw new Error(`Invalid JSON. Body preview: ${preview}`);
}

function extractFirstJsonObject(text) {
  const start = text.indexOf("{");
  if (start < 0) return "";

  let depth = 0;
  let inString = false;
  let escaped = false;

  for (let index = start; index < text.length; index += 1) {
    const char = text[index];

    if (inString) {
      if (escaped) {
        escaped = false;
      } else if (char === "\\") {
        escaped = true;
      } else if (char === '"') {
        inString = false;
      }
      continue;
    }

    if (char === '"') {
      inString = true;
    } else if (char === "{") {
      depth += 1;
    } else if (char === "}") {
      depth -= 1;
      if (depth === 0) return text.slice(start, index + 1);
    }
  }

  return "";
}

function collectData(payload) {
  const source =
    payload && typeof payload.data === "object" && payload.data !== null
      ? payload.data
      : payload && typeof payload.fields === "object" && payload.fields !== null
        ? payload.fields
        : payload || {};

  const data = {};
  for (const name of FIELD_NAMES) {
    const value = source[name];
    data[name] = normalizeValue(value);
  }
  return data;
}

function normalizeValue(value) {
  if (value === undefined || value === null) return DEFAULT_VALUE;
  const text = String(value).replace(/\r\n/g, "\n").replace(/\r/g, "\n").trim();
  return text || DEFAULT_VALUE;
}

function escapeXml(value) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function toWordTextXml(value) {
  return value
    .split("\n")
    .map((line) => escapeXml(line))
    .join("</w:t><w:br/><w:t>");
}

function safeFileName(filename) {
  const requested = String(filename || "教案.docx").trim() || "教案.docx";
  const cleaned = requested.replace(/[<>:"/\\|?*\u0000-\u001F]+/g, "_");
  const limited = cleaned.slice(0, 120);
  return limited.toLowerCase().endsWith(".docx") ? limited : `${limited}.docx`;
}

function contentDisposition(filename) {
  const fallback = filename
    .replace(/[^\x20-\x7E]+/g, "_")
    .replace(/["\\;]/g, "_");
  return `attachment; filename="${fallback}"; filename*=UTF-8''${encodeURIComponent(
    filename
  )}`;
}

async function fillLessonPlanDocx(payload) {
  const templatePath = getTemplatePath();
  if (!fs.existsSync(templatePath)) {
    throw new Error(`Template not found: ${templatePath}`);
  }

  const template = fs.readFileSync(templatePath);
  const zip = await JSZip.loadAsync(template);
  const documentFile = zip.file("word/document.xml");
  if (!documentFile) {
    throw new Error("Invalid DOCX template: word/document.xml not found");
  }

  const data = collectData(payload);
  let xml = await documentFile.async("string");

  for (const [key, value] of Object.entries(data)) {
    xml = xml.split(`{{${key}}}`).join(toWordTextXml(value));
  }

  const remaining = xml.match(/\{\{[a-zA-Z0-9_]+\}\}/g);
  if (remaining && remaining.length) {
    throw new Error(`Unfilled template placeholders: ${remaining.join(", ")}`);
  }

  zip.file("word/document.xml", xml);
  const buffer = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE"
  });

  return {
    buffer,
    filename: safeFileName(payload.filename || payload.output_filename),
    mimeType: DOCX_MIME
  };
}

module.exports = {
  DOCX_MIME,
  FIELD_NAMES,
  collectData,
  contentDisposition,
  fillLessonPlanDocx,
  parsePayload,
  safeFileName
};
