import MsgReaderModule from "msgreader";
import { simpleParser } from "mailparser";
import mammoth from "mammoth";
import PDFDocument from "pdfkit";
import { Document, Paragraph, TextRun, Packer, BorderStyle } from "docx";
import CFB from "cfb";

const MsgReader = MsgReaderModule.default || MsgReaderModule;

export default async (req, context) => {
  if (req.method === "OPTIONS") return new Response("", { status: 200 });
  if (req.method !== "POST") return Response.json({ error: "Method not allowed" }, { status: 405 });

  try {
    const formData = await req.formData();
    const file = formData.get("file");
    const targetFormat = formData.get("targetFormat");
    if (!file || !targetFormat) return Response.json({ error: "Missing file or targetFormat" }, { status: 400 });

    const filename = file.name || "file";
    const ext = filename.split(".").pop().toLowerCase();
    const baseName = filename.replace(/\.[^.]+$/, "");
    const fileBuffer = Buffer.from(await file.arrayBuffer());

    const content = await extract(fileBuffer, ext);
    const result = await convert(content, targetFormat, baseName);

    return new Response(result.buffer, {
      status: 200,
      headers: { "Content-Type": result.mime, "Content-Disposition": `attachment; filename="${result.name}"`, "X-Filename": result.name },
    });
  } catch (err) {
    console.error("Convert error:", err);
    return Response.json({ error: err.message || "Conversion failed" }, { status: 500 });
  }
};

export const config = { path: "/api/convert" };

// ═══════════════════════════════════════
// EXTRACT
// ═══════════════════════════════════════
async function extract(buf, ext) {
  switch (ext) {
    case "msg": return extractMSG(buf);
    case "eml": return extractEML(buf);
    case "docx": return extractDOCX(buf);
    case "html": case "htm": {
      const s = buf.toString("utf-8");
      return { text: s.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim(), html: s, meta: {} };
    }
    case "csv": return extractCSV(buf.toString("utf-8"));
    case "json": return extractJSON(buf.toString("utf-8"));
    default: { const s = buf.toString("utf-8"); return { text: s, html: "<pre>" + esc(s) + "</pre>", meta: {} }; }
  }
}

// ═══════════════════════════════════════
// MSG — Accurate extraction via CFB + msgreader
// ═══════════════════════════════════════
function extractMSG(buf) {
  let subject = "", from = "", senderEmail = "", to = "", cc = "", date = "", body = "", bodyHtml = "";
  let attachmentNames = [];

  // Method 1: msgreader
  try {
    const reader = new MsgReader(buf);
    const d = reader.getFileData();
    subject = clean(d.subject);
    from = clean(d.senderName);
    senderEmail = clean(d.senderEmail);
    to = (d.recipients || []).map(r => clean(r.name || r.email)).filter(Boolean).join(", ");
    date = clean(d.messageDeliveryTime || d.clientSubmitTime);
    body = clean(d.body);
    bodyHtml = d.bodyHtml || "";
    attachmentNames = (d.attachments || []).map(a => clean(a.fileName || a.name)).filter(Boolean);
  } catch (e) { console.warn("msgreader failed:", e.message); }

  // Method 2: CFB deep extraction for better strings
  try {
    const cfb = CFB.read(buf, { type: "buffer" });
    for (const entry of cfb.FileIndex) {
      if (!entry.name || entry.type !== 2 || !entry.content || entry.content.length === 0) continue;
      const n = entry.name;
      const c = entry.content;

      if (n.includes("0037") && !subject) subject = msgStr(c, n);
      if (n.includes("0C1A") && !from) from = msgStr(c, n);
      if ((n.includes("0C1F") || n.includes("0065")) && !senderEmail) senderEmail = msgStr(c, n);
      if (n.includes("0E04") && !to) to = msgStr(c, n);
      if (n.includes("0E03") && !cc) cc = msgStr(c, n);
      if (n.includes("1000") && !body) body = msgStr(c, n);
      if (n.includes("1013") && !bodyHtml) bodyHtml = msgStr(c, n);
      if (n.includes("3707")) {
        const att = msgStr(c, n);
        if (att && !attachmentNames.includes(att)) attachmentNames.push(att);
      }
    }
  } catch (e) { console.warn("CFB failed:", e.message); }

  // Clean everything
  subject = scrub(subject) || "(Tanpa Subjek)";
  from = scrub(from) || "(Tidak diketahui)";
  if (senderEmail && from && !from.includes(senderEmail)) from = `${from} <${scrub(senderEmail)}>`;
  to = scrub(to) || "(Tidak diketahui)";
  cc = scrub(cc);
  date = scrub(date);
  body = scrub(body);
  if (bodyHtml && bodyHtml.startsWith("{\\rtf")) bodyHtml = "";

  const attachments = attachmentNames.map(n => ({ name: scrub(n), size: 0 })).filter(a => a.name);
  const text = buildTextEmail({ subject, from, to, cc, date, body, attachments });
  const html = buildHtmlEmail({ subject, from, to, cc, date, body: bodyHtml || ("<div style='white-space:pre-wrap;line-height:1.6'>" + esc(body) + "</div>"), attachments });

  return { text, html, meta: { subject, from, to, cc, date, attachments, type: "email" } };
}

function msgStr(content, name) {
  if (!content || content.length === 0) return "";
  const buf = Buffer.from(content);
  const isUnicode = name.includes("001F");
  let str = isUnicode ? buf.toString("utf16le") : buf.toString("utf-8");
  return str.replace(/\0/g, "").trim();
}

function clean(s) { return (s && typeof s === "string") ? s.replace(/\0/g, "").trim() : ""; }

function scrub(s) {
  if (!s) return "";
  return s
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, "")
    .replace(/[\u0080-\u009F]/g, "")
    .replace(/[\uFFFD\uFFFE\uFFFF]/g, "")
    .replace(/[Ð™ÖöâÔ÷†æ¶çÆÂ–—""„‰©®]{3,}/g, "")
    .replace(/[ \t]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

// ═══════════════════════════════════════
// EML — via mailparser
// ═══════════════════════════════════════
async function extractEML(buf) {
  try {
    const p = await simpleParser(buf);
    const subject = p.subject || "(Tanpa Subjek)";
    const from = p.from?.text || "(Tidak diketahui)";
    const to = p.to?.text || "(Tidak diketahui)";
    const cc = p.cc?.text || "";
    const date = p.date ? p.date.toLocaleString("id-ID") : "";
    const body = p.text || "";
    const bodyHtml = p.html || "";
    const attachments = (p.attachments || []).map(a => ({ name: a.filename || "file", size: a.size || 0 }));

    const text = buildTextEmail({ subject, from, to, cc, date, body, attachments });
    const html = buildHtmlEmail({ subject, from, to, cc, date, body: bodyHtml || ("<div style='white-space:pre-wrap'>" + esc(body) + "</div>"), attachments });
    return { text, html, meta: { subject, from, to, cc, date, attachments, type: "email" } };
  } catch (err) {
    const raw = buf.toString("utf-8");
    return { text: raw, html: "<pre>" + esc(raw) + "</pre>", meta: {} };
  }
}

async function extractDOCX(buf) {
  const t = await mammoth.extractRawText({ buffer: buf });
  const h = await mammoth.convertToHtml({ buffer: buf });
  return { text: t.value, html: h.value, meta: {} };
}

function extractCSV(str) {
  const rows = str.trim().split("\n").map(r => r.split(","));
  let html = '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:sans-serif;">';
  rows.forEach((row, i) => { html += "<tr>"; row.forEach(cell => { const tag = i === 0 ? "th" : "td"; html += "<" + tag + " style='" + (i === 0 ? "background:#f0f0f0;font-weight:bold;" : "") + "'>" + esc(cell.trim()) + "</" + tag + ">"; }); html += "</tr>"; });
  html += "</table>";
  return { text: str, html, meta: {} };
}

function extractJSON(str) {
  try { const p = JSON.stringify(JSON.parse(str), null, 2); return { text: p, html: "<pre>" + esc(p) + "</pre>", meta: {} }; }
  catch { return { text: str, html: "<pre>" + esc(str) + "</pre>", meta: {} }; }
}

// ═══════════════════════════════════════
// CONVERT
// ═══════════════════════════════════════
async function convert(content, format, baseName) {
  switch (format) {
    case "pdf": return generatePDF(content, baseName);
    case "docx": return generateDOCX(content, baseName);
    case "txt": return { buffer: Buffer.from(content.text, "utf-8"), mime: "text/plain; charset=utf-8", name: baseName + ".txt" };
    case "html": return generateHTML(content, baseName);
    default: throw new Error("Format tidak didukung: " + format);
  }
}

// ═══════════════════════════════════════
// PDF — Clean professional output
// ═══════════════════════════════════════
function generatePDF(content, baseName) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: "A4", margins: { top: 50, bottom: 50, left: 50, right: 50 }, bufferPages: true });
    const chunks = [];
    doc.on("data", c => chunks.push(c));
    doc.on("end", () => resolve({ buffer: Buffer.concat(chunks), mime: "application/pdf", name: baseName + ".pdf" }));
    doc.on("error", reject);

    const w = doc.page.width - 100;

    if (content.meta?.type === "email") {
      const m = content.meta;

      // Date
      if (m.date) { doc.fontSize(9).fillColor("#888").text(m.date, { width: w }); doc.moveDown(0.3); }

      // Title
      doc.fontSize(14).fillColor("#1a1a2e").font("Helvetica-Bold").text("Itinerary & E-Ticket Receipt", { width: w });
      doc.moveDown(0.8);

      // Header box
      const boxY = doc.y;
      doc.save().roundedRect(50, boxY, w, 75, 4).fill("#f5f6fa").restore();
      let ty = boxY + 10;
      doc.fontSize(9).font("Helvetica-Bold").fillColor("#333");
      doc.text("Contact Information", 62, ty, { width: w - 24 });
      ty = doc.y + 4;
      doc.font("Helvetica").fontSize(8.5).fillColor("#555");
      if (m.from) { doc.text("From: " + m.from, 62, ty, { width: w - 24 }); ty = doc.y + 2; }
      if (m.to) { doc.text("To: " + m.to, 62, ty, { width: w - 24 }); ty = doc.y + 2; }
      if (m.cc) { doc.text("CC: " + m.cc, 62, ty, { width: w - 24 }); ty = doc.y + 2; }

      doc.y = boxY + 82;

      // Attachments
      if (m.attachments?.length) {
        doc.fontSize(8).fillColor("#666").font("Helvetica-Bold").text("Attachments: ", { continued: true, width: w }).font("Helvetica").text(m.attachments.map(a => a.name).join(", "));
        doc.moveDown(0.5);
      }

      // Divider
      doc.moveTo(50, doc.y).lineTo(50 + w, doc.y).strokeColor("#ddd").lineWidth(0.5).stroke();
      doc.moveDown(0.6);

      // Body
      const bodyText = getBody(content.text);
      doc.fontSize(9.5).fillColor("#222").font("Helvetica").text(bodyText, { width: w, lineGap: 2.5 });
    } else {
      doc.fontSize(10).fillColor("#222").font("Helvetica").text(content.text, { width: w, lineGap: 3 });
    }

    doc.end();
  });
}

// ═══════════════════════════════════════
// DOCX — Professional Word document
// ═══════════════════════════════════════
async function generateDOCX(content, baseName) {
  const ch = [];

  if (content.meta?.type === "email") {
    const m = content.meta;

    if (m.date) ch.push(new Paragraph({ children: [new TextRun({ text: m.date, size: 18, color: "888888", font: "Calibri" })], spacing: { after: 100 } }));

    ch.push(new Paragraph({ children: [new TextRun({ text: "Itinerary & E-Ticket Receipt", bold: true, size: 28, font: "Calibri", color: "1a1a2e" })], spacing: { after: 200 } }));
    ch.push(new Paragraph({ children: [new TextRun({ text: "Contact Information", bold: true, size: 20, font: "Calibri" })], spacing: { after: 80 }, shading: { type: "clear", fill: "f5f6fa" } }));

    for (const f of [m.from ? "From: " + m.from : null, m.to ? "To: " + m.to : null, m.cc ? "CC: " + m.cc : null].filter(Boolean)) {
      const [label, ...rest] = f.split(": ");
      ch.push(new Paragraph({ children: [new TextRun({ text: label + ": ", bold: true, size: 18, color: "333333", font: "Calibri" }), new TextRun({ text: rest.join(": "), size: 18, color: "555555", font: "Calibri" })], spacing: { after: 40 } }));
    }

    if (m.attachments?.length) {
      ch.push(new Paragraph({ children: [new TextRun({ text: "Attachments: ", bold: true, size: 18, color: "333333", font: "Calibri" }), new TextRun({ text: m.attachments.map(a => a.name).join(", "), size: 18, color: "555555", font: "Calibri" })], spacing: { after: 100 } }));
    }

    ch.push(new Paragraph({ border: { bottom: { color: "cccccc", style: BorderStyle.SINGLE, size: 1, space: 4 } }, spacing: { after: 200 } }));

    for (const line of getBody(content.text).split("\n")) {
      ch.push(new Paragraph({ children: [new TextRun({ text: line, size: 20, font: "Calibri", color: "222222" })], spacing: { after: 60 } }));
    }
  } else {
    for (const line of content.text.split("\n")) {
      ch.push(new Paragraph({ children: [new TextRun({ text: line, size: 22, font: "Calibri" })], spacing: { after: 80 } }));
    }
  }

  const doc = new Document({ sections: [{ properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 720, right: 720, bottom: 720, left: 720 } } }, children: ch }] });
  const buf = await Packer.toBuffer(doc);
  return { buffer: Buffer.from(buf), mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", name: baseName + ".docx" };
}

function generateHTML(content, baseName) {
  const full = "<!DOCTYPE html><html lang='id'><head><meta charset='utf-8'><title>" + esc(baseName) + "</title><style>body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;max-width:800px;margin:40px auto;padding:20px;color:#222;line-height:1.6}.email-header{background:#f5f6fa;padding:20px;border-radius:8px;margin-bottom:24px}.email-header p{margin:4px 0;font-size:14px;color:#555}.email-header .subject{font-size:20px;font-weight:600;color:#1a1a2e;margin-bottom:8px}pre{white-space:pre-wrap;word-wrap:break-word}table{border-collapse:collapse;width:100%}th,td{border:1px solid #ddd;padding:8px 12px;text-align:left}th{background:#f0f0f0}</style></head><body>" + content.html + "</body></html>";
  return { buffer: Buffer.from(full, "utf-8"), mime: "text/html; charset=utf-8", name: baseName + ".html" };
}

// ═══════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════
function esc(s) { return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }

function getBody(text) {
  const lines = text.split("\n");
  const idx = lines.findIndex(l => l.startsWith("─"));
  return idx >= 0 ? lines.slice(idx + 1).join("\n").trim() : text;
}

function buildTextEmail({ subject, from, to, cc, date, body, attachments }) {
  const l = ["Subject: " + subject, "From: " + from, "To: " + to];
  if (cc) l.push("CC: " + cc);
  if (date) l.push("Date: " + date);
  if (attachments?.length) l.push("Attachments: " + attachments.map(a => a.name).join(", "));
  l.push("", "─".repeat(60), "", body);
  return l.join("\n");
}

function buildHtmlEmail({ subject, from, to, cc, date, body, attachments }) {
  let h = '<div class="email-header"><div class="subject">' + esc(subject) + '</div>';
  h += '<p><strong>From:</strong> ' + esc(from) + '</p>';
  h += '<p><strong>To:</strong> ' + esc(to) + '</p>';
  if (cc) h += '<p><strong>CC:</strong> ' + esc(cc) + '</p>';
  if (date) h += '<p><strong>Date:</strong> ' + esc(date) + '</p>';
  if (attachments?.length) h += '<p><strong>Attachments:</strong> ' + attachments.map(a => "📎 " + esc(a.name)).join(", ") + '</p>';
  h += '</div><div class="email-body">' + body + '</div>';
  return h;
}
