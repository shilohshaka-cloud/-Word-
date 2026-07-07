const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");
const zlib = require("zlib");

const DOCX_MIME =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const DEFAULT_VALUE = "";

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
  if (!body) throw new Error("Empty request body. Check the Dify HTTP node Body variable.");
  if (typeof body === "object" && !Buffer.isBuffer(body)) return body;
  if (typeof body === "string") {
    const trimmed = body.trim();
    if (!trimmed) {
      throw new Error("Empty request body. Check the Dify HTTP node Body variable.");
    }
    return parseJsonLenient(trimmed);
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
  const source = resolveDataSource(payload);

  const data = {};
  for (const name of FIELD_NAMES) {
    const value = source[name];
    data[name] = normalizeFieldValue(name, value);
  }
  return data;
}

function resolveDataSource(payload) {
  if (!payload || typeof payload !== "object") return {};

  if (payload.data && typeof payload.data === "object") return payload.data;
  if (payload.fields && typeof payload.fields === "object") return payload.fields;

  if (typeof payload.data === "string") {
    try {
      const parsed = parseJsonLenient(payload.data);
      if (parsed && typeof parsed.data === "object") return parsed.data;
      if (parsed && typeof parsed === "object") return parsed;
    } catch (_) {
      return payload;
    }
  }

  if (typeof payload.fields === "string") {
    try {
      const parsed = parseJsonLenient(payload.fields);
      if (parsed && typeof parsed.data === "object") return parsed.data;
      if (parsed && typeof parsed === "object") return parsed;
    } catch (_) {
      return payload;
    }
  }

  return payload;
}

function normalizeValue(value) {
  if (value === undefined || value === null) return DEFAULT_VALUE;
  const text = String(value).replace(/\r\n/g, "\n").replace(/\r/g, "\n").trim();
  return text || DEFAULT_VALUE;
}

function normalizeFieldValue(name, value) {
  const text = normalizeValue(value);
  if (!text) return text;

  if (name === "lesson_no") return cleanLessonNo(text);
  if (name === "hours") return cleanHours(text);
  if (/^time_\d+_minutes$/.test(name)) return cleanMinutes(text);

  return normalizeListBreaks(text);
}

function cleanLessonNo(value) {
  return value
    .replace(/\s+/g, "")
    .replace(/第/g, "")
    .replace(/次课|课次|次|课/g, "")
    .replace(/[：:]/g, "")
    .trim();
}

function cleanHours(value) {
  const compact = value.replace(/\s+/g, "");
  const hourMatch = compact.match(/(\d+(?:\.\d+)?)(?=学时|课时|小时|h\b)/i);
  if (hourMatch) return formatNumber(hourMatch[1]);

  const minuteMatch = compact.match(/(\d+(?:\.\d+)?)(?=分钟|分|min\b|mins\b|minute\b)/i);
  if (minuteMatch) return formatNumber(Number(minuteMatch[1]) / 45);

  const numberMatch = compact.match(/\d+(?:\.\d+)?/);
  if (numberMatch) return formatNumber(numberMatch[0]);

  return compact.replace(/学时|课时|小时/g, "");
}

function cleanMinutes(value) {
  const compact = value.replace(/\s+/g, "");
  const numberMatch = compact.match(/\d+(?:\.\d+)?/);
  if (numberMatch) return formatNumber(numberMatch[0]);
  return compact.replace(/分钟|分|min|mins|minute/gi, "");
}

function formatNumber(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) return String(value);
  return Number.isInteger(number) ? String(number) : String(Number(number.toFixed(1)));
}

function normalizeListBreaks(value) {
  return value
    .replace(/([。；;])\s*((?:\d+|[一二三四五六七八九十]+)[、.．])/g, "$1\n$2")
    .replace(/\s+((?:[2-9]|1[0-9])[.．、])/g, "\n$1");
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

function encodeDownloadToken(payload) {
  const normalized = {
    filename: safeFileName(payload.filename || payload.output_filename),
    data: collectData(payload)
  };
  const json = JSON.stringify(normalized);
  return zlib.deflateRawSync(Buffer.from(json, "utf8")).toString("base64url");
}

function decodeDownloadToken(token) {
  const json = zlib
    .inflateRawSync(Buffer.from(String(token || ""), "base64url"))
    .toString("utf8");
  return JSON.parse(json);
}

function getPublicBaseUrl(req) {
  const forwardedProto = req.headers["x-forwarded-proto"];
  const proto = Array.isArray(forwardedProto)
    ? forwardedProto[0]
    : forwardedProto || "https";
  const host = req.headers["x-forwarded-host"] || req.headers.host;
  return `${proto}://${host}`;
}

module.exports = {
  DOCX_MIME,
  FIELD_NAMES,
  collectData,
  contentDisposition,
  decodeDownloadToken,
  encodeDownloadToken,
  fillLessonPlanDocx,
  getPublicBaseUrl,
  parsePayload,
  safeFileName
};
