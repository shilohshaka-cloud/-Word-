const {
  encodeDownloadToken,
  fillLessonPlanDocx,
  getPublicBaseUrl,
  parsePayload
} = require("../lib/fill-docx");

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Accept");

  function sendText(statusCode, text) {
    res.statusCode = statusCode;
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.end(text);
  }

  if (req.method === "OPTIONS") {
    res.statusCode = 204;
    res.end();
    return;
  }

  if (req.method !== "POST") {
    sendText(405, "Method Not Allowed");
    return;
  }

  try {
    const payload = parsePayload(req.body);
    const { filename } = await fillLessonPlanDocx(payload);
    const token = encodeDownloadToken({ ...payload, filename });
    const downloadUrl = `${getPublicBaseUrl(
      req
    )}/api/download-lesson-plan-docx?t=${encodeURIComponent(token)}`;

    sendText(200, `[下载 ${filename}](${downloadUrl})`);
  } catch (error) {
    sendText(400, `DOCX_LINK_FAILED: ${error.message}`);
  }
};
