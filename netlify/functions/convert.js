const busboy = require("busboy");
const MsgReader = require("msgreader").default || require("msgreader");
const { simpleParser } = require("mailparser");
const mammoth = require("mammoth");
const PDFDocument = require("pdfkit");
const docx = require("docx");

// ─────────────────────────────────────────
// HANDLER
// ─────────────────────────────────────────
exports.handler = async (event) => {
  const headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
  };

  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 200, headers, body: "" };
  }

  if (event.httpMethod !== "POST") {
    return { statusCode: 405, headers, body: JSON.stringify({ error: "Method not allowed" }) };
  }

  try {
    const { fileBuffer, filename, targetFormat } = await parseForm(event);
    if (!fileBuffer || !targetFormat) {
      return { statusCode: 400, headers, body: JSON.stringify({ error: "Missing file or targetFormat" }) };
    }

    const ext = filename.split(".").pop().toLowerCase();
    const baseName = filename.replace(/\.[^.]+$/, "");

    // Step 1 — Extract structured content from source
    const content = await extract(fileBuffer, ext);

    // Step 2 — Convert to target
    const result = await convert(content, targetFormat, baseName);

    return {
      statusCode: 200,
      headers: {
        ...headers,
        "Content-Type": result.mime,
        "Content-Disposition": `attachment; filename="${result.name}"`,
        "X-Filename": result.name,
      },
      body: result.base64,
      isBase64Encoded: true,
    };
  } catch (err) {
    console.error("Convert error:", err);
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ error: err.message || "Conversion failed" }),
    };
  }
};

// ─────────────────────────────────────────
// PARSE MULTIPART FORM
// ─────────────────────────────────────────
function parseForm(event) {
  return new Promise((resolve, reject) => {
    const ct = event.headers["content-type"] || event.headers["Content-Type"];
    const bb = busboy({ headers: { "content-type": ct } });

    let fileBuffer = null;
    let filename = "";
    let targetFormat = "";

    bb.on("file", (_field, stream, info) => {
      filename = info.filename || "file";
      const chunks = [];
      stream.on("data", (c) => chunks.push(c));
      stream.on("end", () => { fileBuffer = Buffer.concat(chunks); });
    });

    bb.on("field", (name, val) => {
      if (name === "targetFormat") targetFormat = val;
    });

    bb.on("finish", () => resolve({ fileBuffer, filename, targetFormat }));
    bb.on("error", reject);

    const raw = event.isBase64Encoded
      ? Buffer.from(event.body, "base64")
      : Buffer.from(event.body, "binary");

    bb.end(raw);
  });
}

// ─────────────────────────────────────────
// EXTRACT: Source → { text, html, meta }
// ─────────────────────────────────────────
async function extract(buf, ext) {
  switch (ext) {
    case "msg": return extractMSG(buf);
    case "eml": return extractEML(buf);
    case "docx": return extractDOCX(buf);
    case "html":
    case "htm": {
      const s = buf.toString("utf-8");
      // Strip tags for plain text
      const plain = s.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
      return { text: plain, html: s, meta: {} };
    }
    case "csv": return extractCSV(buf.toString("utf-8"));
    case "json": return extractJSON(buf.toString("utf-8"));
    default: {
      const s = buf.toString("utf-8");
      return { text: s, html: `<pre>${esc(s)}</pre>`, meta: {} };
    }
  }
}

// ── MSG (accurate via msgreader) ──
function extractMSG(buf) {
  try {
    const reader = new MsgReader(buf);
    const d = reader.getFileData();

    const subject = d.subject || "(Tanpa Subjek)";
    const senderName = d.senderName || "";
    const senderEmail = d.senderEmail || "";
    const from = senderName
      ? `${senderName}${senderEmail ? ` <${senderEmail}>` : ""}`
      : senderEmail || "(Tidak diketahui)";
    const recipients = (d.recipients || []).map((r) => r.name || r.email || "").filter(Boolean);
    const to = recipients.join(", ") || "(Tidak diketahui)";
    const date = d.messageDeliveryTime || d.clientSubmitTime || "";
    const body = d.body || "";
    const bodyHtml = d.bodyHtml || "";

    const attachments = (d.attachments || []).map((a) => ({
      name: a.fileName || a.name || "file",
      size: a.content ? a.content.length : 0,
    }));

    const text = buildTextEmail({ subject, from, to, date, body, attachments });
    const html = buildHtmlEmail({ subject, from, to, date, body: bodyHtml || `<div style="white-space:pre-wrap">${esc(body)}</div>`, attachments });

    return { text, html, meta: { subject, from, to, date, attachments, type: "email" } };
  } catch (err) {
    console.warn("msgreader failed:", err.message);
    // Fallback: brute-force text extraction
    const bytes = new Uint8Array(buf);
    let longest = "";
    let cur = "";
    for (let i = 0; i < bytes.length; i++) {
      const b = bytes[i];
      if (b >= 32 && b < 127) { cur += String.fromCharCode(b); }
      else { if (cur.length > longest.length) longest = cur; cur = ""; }
    }
    if (cur.length > longest.length) longest = cur;
    const text = longest || "(Tidak dapat membaca konten MSG)";
    return { text, html: `<pre>${esc(text)}</pre>`, meta: {} };
  }
}

// ── EML (accurate via mailparser) ──
async function extractEML(buf) {
  try {
    const parsed = await simpleParser(buf);

    const subject = parsed.subject || "(Tanpa Subjek)";
    const from = parsed.from?.text || "(Tidak diketahui)";
    const to = parsed.to?.text || "(Tidak diketahui)";
    const cc = parsed.cc?.text || "";
    const date = parsed.date ? parsed.date.toLocaleString("id-ID") : "";
    const body = parsed.text || "";
    const bodyHtml = parsed.html || "";

    const attachments = (parsed.attachments || []).map((a) => ({
      name: a.filename || "file",
      size: a.size || 0,
    }));

    const text = buildTextEmail({ subject, from, to, cc, date, body, attachments });
    const html = buildHtmlEmail({ subject, from, to, cc, date, body: bodyHtml || `<div style="white-space:pre-wrap">${esc(body)}</div>`, attachments });

    return { text, html, meta: { subject, from, to, cc, date, attachments, type: "email" } };
  } catch (err) {
    const raw = buf.toString("utf-8");
    return { text: raw, html: `<pre>${esc(raw)}</pre>`, meta: {} };
  }
}

// ── DOCX (via mammoth) ──
async function extractDOCX(buf) {
  const textRes = await mammoth.extractRawText({ buffer: buf });
  const htmlRes = await mammoth.convertToHtml({ buffer: buf });
  return { text: textRes.value, html: htmlRes.value, meta: {} };
}

// ── CSV ──
function extractCSV(str) {
  const rows = str.trim().split("\n").map((r) => r.split(","));
  let html = '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:sans-serif;">';
  rows.forEach((row, i) => {
    html += "<tr>";
    row.forEach((cell) => {
      const tag = i === 0 ? "th" : "td";
      html += `<${tag} style="${i === 0 ? "background:#f0f0f0;font-weight:bold;" : ""}">${esc(cell.trim())}</${tag}>`;
    });
    html += "</tr>";
  });
  html += "</table>";
  return { text: str, html, meta: {} };
}

// ── JSON ──
function extractJSON(str) {
  try {
    const pretty = JSON.stringify(JSON.parse(str), null, 2);
    return { text: pretty, html: `<pre>${esc(pretty)}</pre>`, meta: {} };
  } catch {
    return { text: str, html: `<pre>${esc(str)}</pre>`, meta: {} };
  }
}

// ─────────────────────────────────────────
// CONVERT: { text, html, meta } → file
// ─────────────────────────────────────────
async function convert(content, format, baseName) {
  switch (format) {
    case "pdf": return generatePDF(content, baseName);
    case "docx": return generateDOCX(content, baseName);
    case "txt": return generateTXT(content, baseName);
    case "html": return generateHTML(content, baseName);
    default: throw new Error(`Format "${format}" tidak didukung`);
  }
}

// ── Generate PDF (pdfkit — accurate, supports unicode) ──
function generatePDF(content, baseName) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: "A4", margin: 50, bufferPages: true });
    const chunks = [];

    doc.on("data", (c) => chunks.push(c));
    doc.on("end", () => {
      const buf = Buffer.concat(chunks);
      resolve({ base64: buf.toString("base64"), mime: "application/pdf", name: `${baseName}.pdf` });
    });
    doc.on("error", reject);

    // If it's an email, render with structure
    if (content.meta?.type === "email") {
      const m = content.meta;

      // Header box
      doc.save();
      doc.roundedRect(50, 50, 495, 90, 6).fill("#f5f6fa");
      doc.restore();

      doc.fontSize(16).fillColor("#1a1a2e").text(m.subject || "Email", 65, 60, { width: 465 });
      doc.fontSize(9).fillColor("#666");
      doc.text(`From: ${m.from || ""}`, 65, doc.y + 6, { width: 465 });
      doc.text(`To: ${m.to || ""}`, 65, doc.y + 2, { width: 465 });
      if (m.cc) doc.text(`CC: ${m.cc}`, 65, doc.y + 2, { width: 465 });
      if (m.date) doc.text(`Date: ${m.date}`, 65, doc.y + 2, { width: 465 });

      if (m.attachments?.length) {
        doc.text(`Attachments: ${m.attachments.map((a) => a.name).join(", ")}`, 65, doc.y + 2, { width: 465 });
      }

      // Divider
      const bodyY = Math.max(doc.y + 20, 160);
      doc.moveTo(50, bodyY - 5).lineTo(545, bodyY - 5).strokeColor("#ddd").stroke();

      // Body
      doc.fontSize(10).fillColor("#222");
      // Use plain text for body (more reliable than HTML parsing)
      const bodyText = content.text.split("\n").slice(content.text.split("\n").findIndex((l) => l.startsWith("─")) + 1).join("\n").trim() || content.text;
      doc.text(bodyText, 50, bodyY, { width: 495, lineGap: 3 });
    } else {
      // Generic text document
      doc.fontSize(10).fillColor("#222").text(content.text, { width: 495, lineGap: 3 });
    }

    doc.end();
  });
}

// ── Generate DOCX (docx library — proper .docx) ──
async function generateDOCX(content, baseName) {
  const children = [];

  if (content.meta?.type === "email") {
    const m = content.meta;

    // Email header as styled paragraphs
    children.push(
      new docx.Paragraph({
        children: [new docx.TextRun({ text: m.subject || "Email", bold: true, size: 32, font: "Calibri" })],
        spacing: { after: 200 },
      })
    );

    const headerFields = [
      m.from ? `From: ${m.from}` : null,
      m.to ? `To: ${m.to}` : null,
      m.cc ? `CC: ${m.cc}` : null,
      m.date ? `Date: ${m.date}` : null,
      m.attachments?.length ? `Attachments: ${m.attachments.map((a) => a.name).join(", ")}` : null,
    ].filter(Boolean);

    for (const field of headerFields) {
      children.push(
        new docx.Paragraph({
          children: [new docx.TextRun({ text: field, size: 18, color: "666666", font: "Calibri" })],
          spacing: { after: 40 },
        })
      );
    }

    // Divider
    children.push(
      new docx.Paragraph({
        border: { bottom: { color: "cccccc", style: docx.BorderStyle.SINGLE, size: 1, space: 1 } },
        spacing: { after: 300 },
      })
    );

    // Body paragraphs
    const bodyLines = content.text.split("\n");
    const bodyStart = bodyLines.findIndex((l) => l.startsWith("─"));
    const bodyText = bodyStart >= 0 ? bodyLines.slice(bodyStart + 1).join("\n").trim() : content.text;

    for (const line of bodyText.split("\n")) {
      children.push(
        new docx.Paragraph({
          children: [new docx.TextRun({ text: line, size: 22, font: "Calibri" })],
          spacing: { after: 80 },
        })
      );
    }
  } else {
    // Generic text
    for (const line of content.text.split("\n")) {
      children.push(
        new docx.Paragraph({
          children: [new docx.TextRun({ text: line, size: 22, font: "Calibri" })],
          spacing: { after: 80 },
        })
      );
    }
  }

  const doc = new docx.Document({
    sections: [{ properties: { page: { size: { width: 12240, height: 15840 } } }, children }],
  });

  const buf = await docx.Packer.toBuffer(doc);
  return {
    base64: buf.toString("base64"),
    mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    name: `${baseName}.docx`,
  };
}

// ── Generate TXT ──
function generateTXT(content, baseName) {
  const buf = Buffer.from(content.text, "utf-8");
  return { base64: buf.toString("base64"), mime: "text/plain; charset=utf-8", name: `${baseName}.txt` };
}

// ── Generate HTML ──
function generateHTML(content, baseName) {
  const full = `<!DOCTYPE html>
<html lang="id">
<head><meta charset="utf-8"><title>${esc(baseName)}</title>
<style>
  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; max-width: 800px; margin: 40px auto; padding: 20px; color: #222; line-height: 1.6; }
  .email-header { background: #f5f6fa; padding: 20px; border-radius: 8px; margin-bottom: 24px; }
  .email-header p { margin: 4px 0; font-size: 14px; color: #555; }
  .email-header .subject { font-size: 20px; font-weight: 600; color: #1a1a2e; margin-bottom: 8px; }
  pre { white-space: pre-wrap; word-wrap: break-word; }
  table { border-collapse: collapse; width: 100%; }
  th, td { border: 1px solid #ddd; padding: 8px 12px; text-align: left; }
  th { background: #f0f0f0; }
</style>
</head>
<body>${content.html}</body>
</html>`;

  const buf = Buffer.from(full, "utf-8");
  return { base64: buf.toString("base64"), mime: "text/html; charset=utf-8", name: `${baseName}.html` };
}

// ─────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────
function esc(s) {
  return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function buildTextEmail({ subject, from, to, cc, date, body, attachments }) {
  const lines = [
    `Subject: ${subject}`,
    `From: ${from}`,
    `To: ${to}`,
  ];
  if (cc) lines.push(`CC: ${cc}`);
  if (date) lines.push(`Date: ${date}`);
  if (attachments?.length) lines.push(`Attachments: ${attachments.map((a) => a.name).join(", ")}`);
  lines.push("", "─".repeat(50), "", body);
  return lines.join("\n");
}

function buildHtmlEmail({ subject, from, to, cc, date, body, attachments }) {
  let html = `<div class="email-header">`;
  html += `<div class="subject">${esc(subject)}</div>`;
  html += `<p><strong>From:</strong> ${esc(from)}</p>`;
  html += `<p><strong>To:</strong> ${esc(to)}</p>`;
  if (cc) html += `<p><strong>CC:</strong> ${esc(cc)}</p>`;
  if (date) html += `<p><strong>Date:</strong> ${esc(date)}</p>`;
  if (attachments?.length) {
    html += `<p><strong>Attachments:</strong> ${attachments.map((a) => `📎 ${esc(a.name)}`).join(", ")}</p>`;
  }
  html += `</div>`;
  html += `<div class="email-body">${body}</div>`;
  return html;
}
