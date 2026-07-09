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

const LOOSE_PAYLOAD_FIELD_NAMES = [
  ...FIELD_NAMES.slice(0, FIELD_NAMES.indexOf("homework_report")),
  "teaching_resources",
  ...FIELD_NAMES.slice(FIELD_NAMES.indexOf("homework_report"))
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

const ORGANIZATION_ROLE_GROUPS = [
  {
    label: "教师活动",
    aliases: ["教师活动", "教师动作", "教师具体动作"]
  },
  {
    label: "学生活动",
    aliases: ["学生活动", "学生任务", "学生具体任务"]
  },
  {
    label: "使用材料",
    aliases: ["使用材料", "使用材料/工具", "资源工具", "教学资源", "材料工具", "使用资源"]
  },
  {
    label: "关键提问",
    aliases: ["关键提问", "关键问题", "关键问题/课堂产出", "关键提问或产出", "问题/产出"]
  },
  {
    label: "即时评价",
    aliases: ["即时评价", "反馈评价", "即时反馈", "评价反馈", "反馈方式", "即时评价或反馈方式"]
  }
];

const ORGANIZATION_ROLE_ALIAS_TO_LABEL = new Map(
  ORGANIZATION_ROLE_GROUPS.flatMap((group) =>
    group.aliases.map((alias) => [alias, group.label])
  )
);

const ORGANIZATION_ROLE_PATTERN = ORGANIZATION_ROLE_GROUPS
  .flatMap((group) => group.aliases)
  .sort((a, b) => b.length - a.length)
  .map(escapeRegExp)
  .join("|");

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

  for (const candidate of [...candidates]) {
    const withoutTrailingCommas = removeTrailingJsonCommas(candidate);
    if (withoutTrailingCommas !== candidate && !candidates.includes(withoutTrailingCommas)) {
      candidates.push(withoutTrailingCommas);
    }

    const escapedControls = escapeRawControlCharactersInJsonStrings(withoutTrailingCommas);
    if (escapedControls !== withoutTrailingCommas && !candidates.includes(escapedControls)) {
      candidates.push(escapedControls);
    }
  }

  for (const candidate of candidates) {
    try {
      return JSON.parse(candidate);
    } catch (_) {
      // Try the next repair candidate.
    }
  }

  const loosePayload = parseExpectedPayloadLoose(extracted || withoutFences || trimmed);
  if (loosePayload) return loosePayload;

  const markdownPayload = parseMarkdownLessonPlanPayload(withoutFences || trimmed);
  if (markdownPayload) return markdownPayload;

  const preview = describeTextForError(trimmed);
  throw new Error(`Invalid JSON. Body preview: ${preview}`);
}

function removeTrailingJsonCommas(text) {
  return text.replace(/,\s*([}\]])/g, "$1");
}

function escapeRawControlCharactersInJsonStrings(text) {
  let result = "";
  let inString = false;
  let escaped = false;

  for (const char of text) {
    if (!inString) {
      if (char === "\"") inString = true;
      result += char;
      continue;
    }

    if (escaped) {
      result += char;
      escaped = false;
      continue;
    }

    if (char === "\\") {
      result += char;
      escaped = true;
      continue;
    }

    if (char === "\"") {
      inString = false;
      result += char;
      continue;
    }

    if (char === "\n") {
      result += "\\n";
      continue;
    }

    if (char === "\r") {
      result += "\\r";
      continue;
    }

    if (char === "\t") {
      result += "\\t";
      continue;
    }

    const code = char.charCodeAt(0);
    if (code < 0x20) {
      result += `\\u${code.toString(16).toUpperCase().padStart(4, "0")}`;
      continue;
    }

    result += char;
  }

  return result;
}

function parseExpectedPayloadLoose(text) {
  const source = text.trim();
  if (!source.includes("\"data\"")) return null;

  const filename = extractLooseStringBeforeKey(source, "filename", "data");
  const data = {};
  let foundField = false;

  for (const field of LOOSE_PAYLOAD_FIELD_NAMES) {
    const value = extractLooseStringField(source, field);
    if (value === null) continue;
    data[field] = value;
    if (FIELD_NAMES.includes(field)) foundField = true;
  }

  if (!foundField) return null;

  return {
    filename: filename || "高校课程教案.docx",
    data
  };
}

function extractLooseStringBeforeKey(source, field, nextField) {
  const startRegex = new RegExp(`"${escapeRegExp(field)}"\\s*:\\s*"`, "m");
  const startMatch = startRegex.exec(source);
  if (!startMatch) return "";

  const start = startMatch.index + startMatch[0].length;
  const endRegex = new RegExp(`"\\s*,\\s*"${escapeRegExp(nextField)}"\\s*:`, "gm");
  endRegex.lastIndex = start;
  const endMatch = endRegex.exec(source);
  if (!endMatch) return "";

  return decodeLooseJsonString(source.slice(start, endMatch.index));
}

function extractLooseStringField(source, field) {
  const startRegex = new RegExp(`"${escapeRegExp(field)}"\\s*:\\s*"`, "m");
  const startMatch = startRegex.exec(source);
  if (!startMatch) return null;

  const start = startMatch.index + startMatch[0].length;
  let endIndex = -1;

  for (const nextField of LOOSE_PAYLOAD_FIELD_NAMES) {
    if (nextField === field) continue;
    const endRegex = new RegExp(`"\\s*,\\s*"${escapeRegExp(nextField)}"\\s*:`, "gm");
    endRegex.lastIndex = start;
    const endMatch = endRegex.exec(source);
    if (endMatch && (endIndex < 0 || endMatch.index < endIndex)) {
      endIndex = endMatch.index;
    }
  }

  if (endIndex < 0) {
    const finalFieldRegex = /"\s*}\s*}?\s*$/gm;
    finalFieldRegex.lastIndex = start;
    const finalMatch = finalFieldRegex.exec(source);
    if (finalMatch) endIndex = finalMatch.index;
  }

  if (endIndex < 0) return null;

  return decodeLooseJsonString(source.slice(start, endIndex));
}

function decodeLooseJsonString(raw) {
  return raw
    .replace(/\\u([0-9a-fA-F]{4})/g, (_, hex) => String.fromCharCode(parseInt(hex, 16)))
    .replace(/\\r\\n/g, "\n")
    .replace(/\\n/g, "\n")
    .replace(/\\r/g, "\n")
    .replace(/\\t/g, "\t")
    .replace(/\\"/g, "\"")
    .replace(/\\\\/g, "\\")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .trim();
}

function parseMarkdownLessonPlanPayload(text) {
  const markdown = stripThinkBlocks(String(text || "")).trim();
  if (!looksLikeLessonPlanMarkdown(markdown)) return null;

  const courseName = firstMatch(markdown, /#\s*《([^》]+)》课程教案/) || "高校课程";
  const lessonNo = firstMatch(markdown, /第\s*([一二三四五六七八九十\d]+)\s*次课/);
  const hours = firstMatch(markdown, /第\s*[一二三四五六七八九十\d]+\s*次课\s*([0-9.一二三四五六七八九十]+)\s*学时/);
  const teachingContent = extractMarkdownTableValue(markdown, "授课内容") || extractHeadingTopic(markdown) || "课程教案";

  const data = {
    lesson_no: lessonNo || "",
    hours: hours || "",
    teaching_content: teachingContent,
    key_points: extractMarkdownTableValue(markdown, "教学重点"),
    difficult_points: extractMarkdownTableValue(markdown, "教学难点"),
    ...extractMarkdownTimeRows(markdown),
    org_course_intro: extractMarkdownOrganizationSection(markdown, "一、课程介绍"),
    org_lead_in: extractMarkdownOrganizationSection(markdown, "二、导入新课"),
    org_content_1: extractMarkdownOrganizationSection(markdown, "三、教学内容1") ||
      extractMarkdownOrganizationSection(markdown, "三、教学内容一"),
    org_content_2: extractMarkdownOrganizationSection(markdown, "四、教学内容2") ||
      extractMarkdownOrganizationSection(markdown, "四、教学内容二"),
    org_practice_interaction: "",
    org_summary: extractMarkdownOrganizationSection(markdown, "五、小结") ||
      extractMarkdownOrganizationSection(markdown, "六、小结"),
    homework_report: extractMarkdownHomeworkSection(markdown),
    teacher_signature_date: ""
  };

  return {
    filename: safeFileName(`${courseName}-${teachingContent}-教案.docx`),
    data
  };
}

function looksLikeLessonPlanMarkdown(text) {
  return /课程教案/.test(text) && (
    /###\s*[一二三四五六七八九十]+[、.．]\s*教学组织/.test(text) ||
    /####\s*[一二三四五六七八九十]+[、.．]\s*课程介绍/.test(text) ||
    /\|\s*项目\s*\|\s*内容\s*\|/.test(text)
  );
}

function stripThinkBlocks(text) {
  let cleaned = text.replace(/<think>[\s\S]*?<\/think>/gi, "");
  if (/^\s*<think>/i.test(cleaned)) {
    const firstHeading = cleaned.search(/\n\s*#\s+/);
    cleaned = firstHeading >= 0 ? cleaned.slice(firstHeading + 1) : cleaned.replace(/^\s*<think>/i, "");
  }
  return cleaned.trim();
}

function firstMatch(text, pattern) {
  const match = text.match(pattern);
  return match ? normalizeMarkdownCell(match[1]) : "";
}

function extractHeadingTopic(markdown) {
  const match = markdown.match(/^#+\s*[^#\n]*?[:：]\s*(.+)$/m);
  return match ? normalizeMarkdownCell(match[1]) : "";
}

function extractMarkdownTableValue(markdown, label) {
  const escaped = escapeRegExp(label);
  const pattern = new RegExp(`^\\|\\s*${escaped}\\s*\\|([\\s\\S]*?)\\|\\s*$`, "m");
  const match = markdown.match(pattern);
  return match ? normalizeMarkdownCell(match[1]) : "";
}

function extractMarkdownTimeRows(markdown) {
  const section = extractMarkdownSection(markdown, "课堂教学时间分配");
  const data = {
    time_1_content: "",
    time_1_minutes: "",
    time_2_content: "",
    time_2_minutes: "",
    time_3_content: "",
    time_3_minutes: "",
    time_4_content: "",
    time_4_minutes: "",
    time_5_content: "",
    time_5_minutes: ""
  };

  const rows = section
    .split(/\n/)
    .map((line) => line.trim())
    .filter((line) => line.startsWith("|") && !/^\|\s*-+/.test(line) && !/\|\s*序号\s*\|/.test(line));

  let index = 1;
  for (const row of rows) {
    const cells = splitMarkdownRow(row);
    if (cells.length < 3) continue;
    const order = normalizeMarkdownCell(cells[0]);
    const content = normalizeMarkdownCell(cells[1]);
    const minutes = normalizeMarkdownCell(cells[2]);
    if (!/^\d+$/.test(order) || !content || !minutes || index > 5) continue;
    data[`time_${index}_content`] = content;
    data[`time_${index}_minutes`] = minutes;
    index += 1;
  }

  return data;
}

function extractMarkdownOrganizationSection(markdown, title) {
  const escaped = escapeRegExp(title).replace(/\\\d/g, "\\s*\\d");
  const pattern = new RegExp(`^####\\s*${escaped}(?:[^\\n]*)\\n([\\s\\S]*?)(?=^####\\s+|^###\\s+|\\s*$)`, "m");
  const match = markdown.match(pattern);
  return match ? normalizeMarkdownBlock(match[1]) : "";
}

function extractMarkdownHomeworkSection(markdown) {
  const section = extractMarkdownSection(markdown, "思考题");
  return normalizeMarkdownBlock(section);
}

function extractMarkdownSection(markdown, titleKeyword) {
  const escaped = escapeRegExp(titleKeyword);
  const pattern = new RegExp(`^###\\s*[^\\n]*${escaped}[^\\n]*\\n([\\s\\S]*?)(?=^###\\s+|\\s*$)`, "m");
  const match = markdown.match(pattern);
  return match ? match[1] : "";
}

function splitMarkdownRow(row) {
  return row
    .replace(/^\|/, "")
    .replace(/\|$/, "")
    .split("|")
    .map((cell) => cell.trim());
}

function normalizeMarkdownCell(value) {
  return String(value || "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/\*\*/g, "")
    .replace(/&nbsp;/gi, " ")
    .replace(/\s+\n/g, "\n")
    .replace(/\n\s+/g, "\n")
    .replace(/[ \t]+/g, " ")
    .trim();
}

function normalizeMarkdownBlock(value) {
  return normalizeMarkdownCell(value)
    .replace(/^#{1,6}\s*/gm, "")
    .replace(/^\s*[-*]\s+/gm, "• ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
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
  return value
    .replace(/([。；;])\s*((?:\d+|[一二三四五六七八九十]+)[、.．)])/g, "$1\n$2")
    .replace(/([。；;])\s*([（(](?:\d+|[一二三四五六七八九十]+)[)）])/g, "$1\n$2")
    .replace(/([。；;])\s*([②③④⑤⑥⑦⑧⑨⑩])/g, "$1\n$2")
    .replace(/([^\n])([•●·]\s*)/g, "$1\n$2")
    .replace(/([^\n])([①②③④⑤⑥⑦⑧⑨⑩]\s*)/g, "$1\n$2")
    .replace(/([^\n])((?:[2-9]|[1-9]\d)[、.．)]\s*)/g, "$1\n$2")
    .replace(/\s+((?:[2-9]|1[0-9])[.．、)])/g, "\n$1")
    .replace(/\s+([（(](?:[2-9]|1[0-9]|[二三四五六七八九十]+)[)）])/g, "\n$1")
    .replace(/\s+([②③④⑤⑥⑦⑧⑨⑩])/g, "\n$1");
}

function normalizeOrganizationText(value) {
  return normalizeTeachingOrganizationText(value);
}

function normalizeBodyText(value) {
  return preferPrimaryArabicMarkers(normalizeListBreaks(value));
}

function normalizeStructuredBodyText(value) {
  const withLineBreaks = normalizeListBreaks(value);
  const withPlainItemsNumbered = enumeratePlainDelimitedItems(withLineBreaks);
  return preferPrimaryArabicMarkers(withPlainItemsNumbered);
}

function normalizeTeachingOrganizationText(value) {
  const withLineBreaks = normalizeListBreaks(value);
  const withRoleBreaks = insertOrganizationRoleBreaks(withLineBreaks);
  const roleBlocks = collectOrganizationRoleBlocks(withRoleBreaks);

  if (roleBlocks.length < 2) {
    return normalizeStructuredBodyText(value);
  }

  return roleBlocks
    .map((block, index) => {
      const content = normalizeOrganizationRoleContent(block.content);
      return content
        ? `${index + 1}、${block.label}：\n${content}`
        : `${index + 1}、${block.label}：`;
    })
    .join("\n");
}

function insertOrganizationRoleBreaks(value) {
  if (!ORGANIZATION_ROLE_PATTERN) return value;

  const roleRegex = new RegExp(
    `(?:^|\\s*)(?:[•●·]\\s*)?(?:\\d{1,2}[、.．)]\\s*)?(${ORGANIZATION_ROLE_PATTERN})[：:]`,
    "g"
  );

  return value
    .replace(roleRegex, (match, roleLabel, offset) => {
      const prefix = offset === 0 ? "" : "\n";
      const canonical = ORGANIZATION_ROLE_ALIAS_TO_LABEL.get(roleLabel) || roleLabel;
      return `${prefix}${canonical}：`;
    })
    .trim();
}

function collectOrganizationRoleBlocks(value) {
  if (!ORGANIZATION_ROLE_PATTERN) return [];

  const roleStartRegex = new RegExp(`^(${ORGANIZATION_ROLE_PATTERN})[：:]\\s*(.*)$`);
  const blocks = [];
  let current = null;

  for (const rawLine of value.split("\n")) {
    const line = rawLine.trim();
    if (!line) continue;

    const match = line.match(roleStartRegex);
    if (match) {
      if (current) blocks.push(current);
      current = {
        label: ORGANIZATION_ROLE_ALIAS_TO_LABEL.get(match[1]) || match[1],
        content: match[2].trim()
      };
      continue;
    }

    if (current) {
      current.content = current.content ? `${current.content}\n${line}` : line;
    }
  }

  if (current) blocks.push(current);

  const canonicalOrder = ORGANIZATION_ROLE_GROUPS.map((group) => group.label);
  return blocks.sort((a, b) => {
    const orderA = canonicalOrder.indexOf(a.label);
    const orderB = canonicalOrder.indexOf(b.label);
    return (orderA < 0 ? 99 : orderA) - (orderB < 0 ? 99 : orderB);
  });
}

function normalizeOrganizationRoleContent(value) {
  const lines = normalizeListBreaks(value)
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  if (!lines.length) return "";

  return lines
    .map((line) => normalizeOrganizationRoleLine(line))
    .join("\n");
}

function normalizeOrganizationRoleLine(line) {
  const withoutBullet = line.replace(/^[•●·]\s*/, "").trim();

  const circledMatch = withoutBullet.match(/^([①②③④⑤⑥⑦⑧⑨⑩])\s*(.*)$/);
  if (circledMatch) return `${circledMatch[1]}${circledMatch[2].trim()}`;

  const parenNumberMatch = withoutBullet.match(/^[（(](\d{1,2}|[一二三四五六七八九十]{1,3})[)）]\s*(.*)$/);
  if (parenNumberMatch) {
    return `• ${parenNumberMatch[2].trim()}`;
  }

  const arabicMatch = withoutBullet.match(/^(\d{1,2})[、.．)]\s*(.*)$/);
  if (arabicMatch) {
    return `• ${arabicMatch[2].trim()}`;
  }

  const chineseMatch = withoutBullet.match(/^([一二三四五六七八九十]{1,3})[、.．)]\s*(.*)$/);
  if (chineseMatch) {
    return `• ${chineseMatch[2].trim()}`;
  }

  return `• ${withoutBullet}`;
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
      return `${prefix}${indent}${marker}`;
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

function normalizeSignatureDate(value) {
  const compact = value.replace(/\s+/g, "");
  const dateMatch = compact.match(/(\d{4})[年\-/.](\d{1,2})[月\-/.](\d{1,2})日?/);
  if (dateMatch) {
    return signatureLineWithDate(
      `${dateMatch[1]}年${Number(dateMatch[2])}月${Number(dateMatch[3])}日`
    );
  }

  if (
    !compact ||
    /待.*补充|年月日|年\s*月\s*日|教师签名|签名|_+/.test(value)
  ) {
    return signatureLineWithDate(getTodayChinaDateText());
  }

  return value;
}

function signatureLineWithDate(dateText) {
  return `  ${dateText}`;
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

function toTeacherSignatureDateXml(value) {
  return `</w:t></w:r><w:r>${runPropertiesXml({
    bold: true
  })}${textOrPreservedTextXml(value, true)}</w:r><w:r><w:t>`;
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
