const {
  contentDisposition,
  fillLessonPlanDocx,
  parsePayload
} = require("../lib/fill-docx");

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

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
    const { buffer, filename, mimeType } = await fillLessonPlanDocx(payload);

    res.setHeader("Content-Type", mimeType);
    res.setHeader("Content-Disposition", contentDisposition(filename));
    res.setHeader("Content-Length", String(buffer.length));
    res.status(200).send(buffer);
  } catch (error) {
    res.status(400).json({
      error: "DOCX_GENERATION_FAILED",
      message: error.message
    });
  }
};
