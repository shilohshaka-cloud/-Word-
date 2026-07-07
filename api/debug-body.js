function describeBody(body) {
  const type = Buffer.isBuffer(body) ? "buffer" : typeof body;
  let text = "";

  if (body === undefined) {
    text = "";
  } else if (body === null) {
    text = "null";
  } else if (Buffer.isBuffer(body)) {
    text = body.toString("utf8");
  } else if (typeof body === "string") {
    text = body;
  } else {
    text = JSON.stringify(body);
  }

  const chars = Array.from(text);
  const visiblePreview = text
    .slice(0, 300)
    .replace(/\s+/g, " ")
    .replace(/[\u200B-\u200D\uFEFF]/g, "");

  return {
    type,
    is_empty_value: body === undefined || body === null || body === "",
    length: text.length,
    trimmed_length: text.trim().length,
    visible_preview: visiblePreview || "<no visible characters>",
    first_40_codepoints: chars
      .slice(0, 40)
      .map((char) => `U+${char.codePointAt(0).toString(16).toUpperCase().padStart(4, "0")}`)
  };
}

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Accept");

  if (req.method === "OPTIONS") {
    res.statusCode = 204;
    res.end();
    return;
  }

  res.statusCode = 200;
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.end(
    JSON.stringify(
      {
        method: req.method,
        content_type: req.headers["content-type"] || "",
        body: describeBody(req.body)
      },
      null,
      2
    )
  );
};
