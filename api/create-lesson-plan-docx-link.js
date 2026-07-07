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
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  try {
    const payload = parsePayload(req.body);
    const { filename, buffer } = await fillLessonPlanDocx(payload);
    const token = encodeDownloadToken({ ...payload, filename });
    const downloadUrl = `${getPublicBaseUrl(
      req
    )}/api/download-lesson-plan-docx?t=${encodeURIComponent(token)}`;

    res.status(200).json({
      file_name: filename,
      size: buffer.length,
      download_url: downloadUrl,
      markdown_link: `[下载 ${filename}](${downloadUrl})`
    });
  } catch (error) {
    res.status(400).json({
      error: "DOCX_LINK_FAILED",
      message: error.message
    });
  }
};
