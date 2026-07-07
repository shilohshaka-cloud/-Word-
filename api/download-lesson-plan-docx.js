const {
  contentDisposition,
  decodeDownloadToken,
  fillLessonPlanDocx
} = require("../lib/fill-docx");

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") {
    res.status(204).end();
    return;
  }

  if (req.method !== "GET") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  try {
    const payload = decodeDownloadToken(req.query.t);
    const { buffer, filename, mimeType } = await fillLessonPlanDocx(payload);

    res.setHeader("Content-Type", mimeType);
    res.setHeader("Content-Disposition", contentDisposition(filename));
    res.setHeader("Content-Length", String(buffer.length));
    res.status(200).send(buffer);
  } catch (error) {
    res.status(400).json({
      error: "DOCX_DOWNLOAD_FAILED",
      message: error.message
    });
  }
};
