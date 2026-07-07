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

  if (req.method === "OPTIONS") {
    res.status(204).end();
    return;
  }

  if (req.method !== "POST") {
    res.status(405).type("text/plain").send("Method Not Allowed");
    return;
  }

  try {
    const payload = parsePayload(req.body);
    const { filename } = await fillLessonPlanDocx(payload);
    const token = encodeDownloadToken({ ...payload, filename });
    const downloadUrl = `${getPublicBaseUrl(
      req
    )}/api/download-lesson-plan-docx?t=${encodeURIComponent(token)}`;

    res
      .status(200)
      .type("text/plain")
      .send(`[下载 ${filename}](${downloadUrl})`);
  } catch (error) {
    res.status(400).type("text/plain").send(`DOCX_LINK_FAILED: ${error.message}`);
  }
};
