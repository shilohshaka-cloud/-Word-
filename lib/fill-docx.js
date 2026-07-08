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
  "teaching_resources",
  "homework_report",
  "teacher_signature_date"
];

const TIME_CONTENT_LABELS = {
  time_1_content: ["课程介绍"],
  time_2_content: ["导入新课"],
  time_3_content: ["教学内容1", "教学内容一"],
  time_4_content: ["教学内容2", "教学内容二"],
  time_5_content: ["小结"]
};

const ORGANIZATION_LABELS = {
  org_course_intro: ["一、课程介绍", "课程介绍"],
  org_lead_in: ["二、导入新课", "导入新课"],
  org_content_1: ["三、教学内容1", "教学内容1", "教学内容一"],
  org_content_2: ["四、教学内容2", "教学内容2", "教学内容二"],
  org_practice_interaction: ["五、课堂练习/互动", "课堂练习/互动", "课堂练习", "互动"],
  org_summary: ["六、小结", "小结"]
};

const ORGANIZATION_FIELDS = new Set(Object.keys(ORGANIZATION_LABELS));
const NO_BOLD_BODY_FIELDS = new Set([...ORGANIZATION_FIELDS, "homework_report"]);

const CIRCLED_NUMBERS = [
  "①",
  "②",
  "③",
  "④",
  "⑤",
  "⑥",
  "⑦",
  "⑧",
  "⑨",
  "⑩"
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

  const preview = describeTextForError(trimmed);
  throw new Error(`Invalid JSON. Body preview: ${preview}`);
}

function describeTextForError(text) {
  const visible = text
    .slice(0, 300)
    .replace(/\s+/g, " ")
    .replace(/[\u200B-\u200D\uFEFF]/g, "");
  const codepoints = Array.from(text.slice(0, 40))
    .map((char) => `U+${char.codePointAt(0).toString(16).toUpperCase().padStart(4, "0")}`)
    .join(" ");
  return `${visible || "<no visible characters>"}; length=${text.length}; codepoints=${codepoints || "none"}`;
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
  if (name === "teacher_signature_date") return normalizeSignatureDate(text);
  if (!text) return text;

  if (name === "lesson_no") return cleanLessonNo(text);
  if (name === "hours") return cleanHours(text);
  if (/^time_\d+_minutes$/.test(name)) return cleanMinutes(text);
  if (TIME_CONTENT_LABELS[name]) {
    return normalizeListBreaks(stripLeadingLabels(text, TIME_CONTENT_LABELS[name]));
  }
  if (ORGANIZATION_FIELDS.has(name)) {
    return normalizeOrganizationText(stripLeadingLabels(text, ORGANIZATION_LABELS[name]));
  }
  if (name === "teaching_resources") return normalizeStructuredBodyText(stripTeachingResourcesLabel(text));
  if (name === "homework_report") return normalizeStructuredBodyText(stripHomeworkLabel(text));

  return normalizeBodyText(text);
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

function stripLeadingLabels(value, labels) {
  let text = value.trim();
  let changed = true;

  while (changed) {
    changed = false;
    for (const label of labels) {
      const escaped = escapeRegExp(label).replace(/\\\//g, "\\s*[/／]\\s*");
      const pattern = new RegExp(
        `^\\s*(?:[一二三四五六七八九十]+[、.．]\\s*)?${escaped}\\s*[：:：、.．\\-—]*\\s*`
      );
      const next = text.replace(pattern, "");
      if (next !== text) {
        text = next.trimStart();
        changed = true;
      }
    }
  }

  return text;
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function formatNumber(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) return String(value);
  return Number.isInteger(number) ? String(number) : String(Number(number.toFixed(1)));
}

function normalizeListBreaks(value) {
  return normalizeBulletMarkers(value)
    .replace(/([。；;])\s*((?:\d+|[一二三四五六七八九十]+)[、.．)])/g, "$1\n$2")
    .replace(/([。；;])\s*([（(](?:\d+|[一二三四五六七八九十]+)[)）])/g, "$1\n$2")
    .replace(/([。；;])\s*([②③④⑤⑥⑦⑧⑨⑩])/g, "$1\n$2")
    .replace(/([。；;])\s*([•●·]\s*)/g, "$1\n$2")
    .replace(/\s+((?:[2-9]|1[0-9])[.．、)](?=\s|[\u4e00-\u9fff]))/g, "\n$1")
    .replace(/\s+([（(](?:[2-9]|1[0-9]|[二三四五六七八九十]+)[)）])/g, "\n$1")
    .replace(/\s+([②③④⑤⑥⑦⑧⑨⑩])/g, "\n$1");
}

function normalizeBulletMarkers(value) {
  return value.replace(/(^|\n)\s*(?:[-*•●·])\s+/g, "$1• ");
}

function normalizeOrganizationText(value) {
  return normalizeStructuredBodyText(value);
}

function normalizeBodyText(value) {
  return preferPrimaryArabicMarkers(normalizeListBreaks(value));
}

function normalizeStructuredBodyText(value) {
  const withLineBreaks = normalizeListBreaks(value);
  const withPlainItemsNumbered = enumeratePlainDelimitedItems(withLineBreaks);
  return preferPrimaryArabicMarkers(withPlainItemsNumbered);
}

function enumeratePlainDelimitedItems(value) {
  if (hasExplicitListMarkers(value)) return value;

  const parts = value
    .split(/[；;]/)
    .map((part) => part.trim())
    .filter(Boolean);
  if (parts.length < 2) return value;

  return parts.map((part, index) => `${index + 1}、${part}`).join("\n");
}

function hasExplicitListMarkers(value) {
  return /(^|\n)\s*(?:\d{1,2}|[一二三四五六七八九十]{1,3})[、.．)]\s*/.test(value) ||
    /(^|\n)\s*[（(](?:\d{1,2}|[一二三四五六七八九十]{1,3})[)）]\s*/.test(value) ||
    /(^|\n)\s*[①②③④⑤⑥⑦⑧⑨⑩]\s*/.test(value);
}

function preferPrimaryArabicMarkers(value) {
  const hasPrimaryArabic = /(^|\n)\s*\d{1,2}[、.．)]\s*/.test(value);

  return value
    .replace(/(^|\n)(\s*)[（(](\d{1,2}|[一二三四五六七八九十]{1,3})[)）]\s*/g, (match, prefix, indent, numberText) => {
      const number = /^\d+$/.test(numberText) ? formatNumber(numberText) : chineseOrdinalToNumber(numberText);
      return number ? `${prefix}${indent}${number}、` : match;
    })
    .replace(/(^|\n)(\s*)([一二三四五六七八九十]{1,3})[、.．]\s*/g, (match, prefix, indent, numberText) => {
      const number = chineseOrdinalToNumber(numberText);
      return number ? `${prefix}${indent}${number}、` : match;
    })
    .replace(/(^|\n)(\s*)(\d{1,2})[.．、)]\s*/g, (match, prefix, indent, numberText) => {
      return `${prefix}${indent}${formatNumber(numberText)}、`;
    })
    .replace(/(^|\n)(\s*)([①②③④⑤⑥⑦⑧⑨⑩])\s*/g, (match, prefix, indent, marker) => {
      if (hasPrimaryArabic || indent) return `${prefix}${indent}${marker}`;
      const index = CIRCLED_NUMBERS.indexOf(marker);
      return `${prefix}${index + 1}、`;
    });
}

function chineseOrdinalToNumber(value) {
  const digits = {
    一: 1,
    二: 2,
    三: 3,
    四: 4,
    五: 5,
    六: 6,
    七: 7,
    八: 8,
    九: 9
  };

  if (value === "十") return 10;
  if (value.startsWith("十")) return 10 + (digits[value[1]] || 0);
  if (value.endsWith("十")) return (digits[value[0]] || 0) * 10;
  if (value.includes("十")) {
    const [tens, ones] = value.split("十");
    return (digits[tens] || 0) * 10 + (digits[ones] || 0);
  }
  return digits[value] || 0;
}

function stripHomeworkLabel(value) {
  return stripLeadingLabels(value, ["思考题及实验实训报告", "思考题", "实验实训报告"]);
}

function stripTeachingResourcesLabel(value) {
  return stripLeadingLabels(value, ["教学资源", "教学材料", "教学工具", "资源与材料", "学习资源"]);
}

function normalizeSignatureDate(value) {
  const compact = value.replace(/\s+/g, "");
  if (
    !compact ||
    /待.*补充|年月日|年\s*月\s*日|教师签名|签名/.test(value)
  ) {
    return getTodayChinaDateText();
  }

  const dateMatch = compact.match(/(\d{4})[年\-/.](\d{1,2})[月\-/.](\d{1,2})日?/);
  if (dateMatch) {
    return `${dateMatch[1]}年${Number(dateMatch[2])}月${Number(dateMatch[3])}日`;
  }

  return value;
}

function getTodayChinaDateText() {
  const parts = new Intl.DateTimeFormat("zh-CN", {
    timeZone: "Asia/Shanghai",
    year: "numeric",
    month: "numeric",
    day: "numeric"
  }).formatToParts(new Date());
  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${values.year}年${Number(values.month)}月${Number(values.day)}日`;
}

function runPropertiesXml({ bold = false } = {}) {
  const boldXml = bold ? "<w:b/><w:bCs/>" : '<w:b w:val="0"/><w:bCs w:val="0"/>';
  return `<w:rPr><w:rFonts w:ascii="宋体" w:eastAsia="宋体" w:hAnsi="宋体" w:cs="宋体" w:hint="eastAsia"/>${boldXml}<w:sz w:val="24"/></w:rPr>`;
}

function textOrPreservedTextXml(value, preserveSpace = false) {
  const spaceAttribute = preserveSpace ? ' xml:space="preserve"' : "";
  return `<w:t${spaceAttribute}>${escapeXml(value)}</w:t>`;
}

function replacementRunXml(value, options = {}) {
  const lines = value.split("\n");
  const textXml = lines
    .map((line, index) => textOrPreservedTextXml(line, options.preserveSpace && index === 0))
    .join("<w:br/>");
  const leadingBreak = options.leadingBreak ? "<w:br/>" : "";

  return `</w:t></w:r><w:r>${runPropertiesXml(options)}${leadingBreak}${textXml}</w:r><w:r><w:t>`;
}

function paragraphXml(text, options = {}) {
  return `<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="7080"/></w:tabs><w:spacing w:before="120" w:line="360" w:lineRule="exact"/></w:pPr><w:r>${runPropertiesXml(options)}${text
    .split("\n")
    .map((line) => textOrPreservedTextXml(line))
    .join("<w:br/>")}</w:r></w:p>`;
}

function teachingResourcesSectionXml(value) {
  if (!value) return "";
  return `${paragraphXml("教学资源", { bold: true })}${paragraphXml(value, { bold: false })}`;
}

function insertTeachingResourcesSection(xml, value) {
  const sectionXml = teachingResourcesSectionXml(value);
  if (!sectionXml) return xml;

  const markerIndex = xml.indexOf("思考题及实验实训报告：");
  if (markerIndex < 0) return xml;

  const paragraphStart = xml.lastIndexOf("<w:p", markerIndex);
  if (paragraphStart < 0) return xml;

  return `${xml.slice(0, paragraphStart)}${sectionXml}${xml.slice(paragraphStart)}`;
}

function toTeacherSignatureDateXml(value) {
  return `</w:t></w:r><w:r>${runPropertiesXml({
    bold: true
  })}<w:tab/>${textOrPreservedTextXml(value)}</w:r><w:r><w:t>`;
}

function toNoBoldWordTextXml(value) {
  return replacementRunXml(value, { bold: false });
}

function toHomeworkReportXml(value) {
  return replacementRunXml(value, { bold: false, leadingBreak: true });
}

function toWordTextXmlForField(name, value) {
  if (name === "teacher_signature_date") return toTeacherSignatureDateXml(value);
  if (name === "homework_report") return toHomeworkReportXml(value);
  if (NO_BOLD_BODY_FIELDS.has(name)) return toNoBoldWordTextXml(value);
  return toWordTextXml(value);
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
    xml = xml.split(`{{${key}}}`).join(toWordTextXmlForField(key, value));
  }

  xml = insertTeachingResourcesSection(xml, data.teaching_resources);

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
