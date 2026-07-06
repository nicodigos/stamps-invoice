import * as pdfjsLib from "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.min.mjs";

const PDF_EXTS = new Set([".pdf"]);
const ZIP_EXTS = new Set([".zip"]);
const IMAGE_EXTS = new Set([".jpg", ".jpeg", ".png", ".webp"]);
const EML_EXTS = new Set([".eml"]);
const MSG_EXTS = new Set([".msg"]);
const MESSAGE_EXTS = new Set([...EML_EXTS, ...MSG_EXTS]);
const POSTAL_MIME_URL = "https://esm.sh/postal-mime@2.4.3";
const MSG_READER_URL = "https://esm.sh/@kenjiuno/msgreader@1.28.0";
const PDFJS_WORKER_URL = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.worker.min.mjs";

export async function uploadedFileToPdfBytes(file) {
  const pdfs = await uploadedFileToPdfByteList(file);
  return mergePdfBytes(pdfs);
}

export async function uploadedFileToPdfByteList(file) {
  const ext = extensionOf(file.name);
  if (PDF_EXTS.has(ext)) {
    return [new Uint8Array(await file.arrayBuffer())];
  }
  if (ZIP_EXTS.has(ext)) {
    return zipToPdfByteList(file);
  }
  if (IMAGE_EXTS.has(ext)) {
    return [await imageToPdfBytes(await file.arrayBuffer(), ext)];
  }
  if (MESSAGE_EXTS.has(ext)) {
    const parsed = await parseMessageFile(file);
    return Promise.all(parsed.attachments.map((attachment) => fileContentToPdfBytes(attachment.bytes, attachment.name)));
  }
  throw new Error(`Unsupported file type: ${ext || file.type || "no extension"}.`);
}

export async function fileToAnalysisPayload(file) {
  const ext = extensionOf(file.name);
  if (MESSAGE_EXTS.has(ext)) {
    const parsed = await parseMessageFile(file);
    const images = [];
    for (const attachment of parsed.attachments) {
      images.push(...await contentToAnalysisImages(attachment.bytes, attachment.name, attachment.name));
    }
    return {
      sourceName: file.name,
      email: parsed.email,
      images,
    };
  }

  return {
    sourceName: file.name,
    email: null,
    images: await contentToAnalysisImages(new Uint8Array(await file.arrayBuffer()), file.name, file.name),
  };
}

export async function emailFileToPreview(file) {
  const ext = extensionOf(file.name);
  if (!MESSAGE_EXTS.has(ext)) return null;
  const parsed = await parseMessageFile(file, { includeAllAttachments: true });
  return {
    sourceName: file.name,
    email: parsed.email,
    attachments: await Promise.all(parsed.attachments.map(async (attachment) => ({
      name: attachment.name,
      mimeType: attachment.mimeType,
      size: attachment.bytes.byteLength,
      thumbnailDataUrl: await attachmentToThumbnailDataUrl(attachment.bytes, attachment.name),
      bytes: attachment.bytes,
    }))),
  };
}

export async function stampPdfBytes(inputBytes, { date, paymentCode, clientInvoice }) {
  const { PDFDocument, StandardFonts, rgb } = window.PDFLib;
  const doc = await PDFDocument.load(inputBytes);
  const font = await doc.embedFont(StandardFonts.HelveticaBold);
  const color = rgb(0.75, 0, 0);

  for (const page of doc.getPages()) {
    const x = 40;
    const y = 105;
    drawBoldText(page, "PAID", x, y + 92, { size: 22, font, color });
    [
      `Payment Code: ${paymentCode}`,
      `Client Invoice: ${clientInvoice}`,
      `Payment Date: ${date}`,
      "Paid by Paola Pongo",
    ].forEach((line, index) => {
      drawBoldText(page, line, x, y + 62 - index * 18, { size: 14, font, color });
    });
  }

  return doc.save();
}

export async function mergePdfBytes(pdfFilesBytes) {
  const { PDFDocument } = window.PDFLib;
  const output = await PDFDocument.create();

  for (const bytes of pdfFilesBytes) {
    const src = await PDFDocument.load(bytes);
    const pages = await output.copyPages(src, src.getPageIndices());
    pages.forEach((page) => output.addPage(page));
  }

  if (output.getPageCount() === 0) {
    throw new Error("There are no pages to merge.");
  }

  return output.save();
}

async function zipToPdfBytes(file) {
  const pdfs = await zipToPdfByteList(file);
  return mergePdfBytes(pdfs);
}

async function zipToPdfByteList(file) {
  const { PDFDocument } = window.PDFLib;
  const zip = await window.JSZip.loadAsync(await file.arrayBuffer());
  const out = [];
  const entries = Object.values(zip.files)
    .filter((entry) => !entry.dir)
    .sort((left, right) => left.name.localeCompare(right.name));

  for (const entry of entries) {
    const ext = extensionOf(entry.name);
    if (!PDF_EXTS.has(ext) && !IMAGE_EXTS.has(ext)) continue;

    const bytes = await entry.async("uint8array");
    if (PDF_EXTS.has(ext)) {
      await PDFDocument.load(bytes);
      out.push(bytes);
    } else {
      out.push(await imageToPdfBytes(bytes, ext));
    }
  }

  if (out.length === 0) {
    throw new Error("The ZIP does not contain supported PDFs or images.");
  }

  return out;
}

async function parseMessageFile(file, options = {}) {
  const ext = extensionOf(file.name);
  if (MSG_EXTS.has(ext)) return parseMsgFile(file, options);
  return parseEmlFile(file, options);
}

async function parseEmlFile(file, options = {}) {
  const { default: PostalMime } = await import(POSTAL_MIME_URL);
  const email = await PostalMime.parse(await file.arrayBuffer());
  let attachments = (email.attachments || [])
    .map((attachment) => ({
      name: attachment.filename || attachment.name || "attachment",
      mimeType: attachment.mimeType || attachment.contentType || "",
      disposition: attachment.disposition || "",
      contentId: attachment.contentId || "",
      related: Boolean(attachment.related),
      bytes: toUint8Array(attachment.content),
    }))
    .filter((attachment) => isRealAttachment(attachment));

  if (!options.includeAllAttachments) {
    attachments = attachments.filter((attachment) => isProcessableAttachment(attachment));
  }

  if (!attachments.length && !options.includeAllAttachments) {
    throw new Error(`The email "${file.name}" does not have supported PDF or image attachments.`);
  }

  return {
    email: {
      subject: email.subject || "",
      from: email.from?.address || email.from?.name || "",
      date: email.date || "",
      text: String(email.text || "").slice(0, 6000),
    },
    attachments,
  };
}

async function parseMsgFile(file, options = {}) {
  const { default: MsgReader } = await import(MSG_READER_URL);
  const reader = new MsgReader(await file.arrayBuffer());
  const message = reader.getFileData();
  let attachments = (message.attachments || [])
    .map((attachment) => {
      const content = reader.getAttachment(attachment);
      const name = content?.fileName || attachment.fileName || attachment.fileNameShort || "attachment";
      return {
        name,
        mimeType: mimeTypeFromName(name),
        disposition: "attachment",
        contentId: attachment.contentId || "",
        related: false,
        bytes: toUint8Array(content?.content),
      };
    })
    .filter((attachment) => attachment.bytes.byteLength > 0);

  if (!options.includeAllAttachments) {
    attachments = attachments.filter((attachment) => isProcessableAttachment(attachment));
  }

  if (!attachments.length && !options.includeAllAttachments) {
    throw new Error(`The Outlook message "${file.name}" does not have supported PDF or image attachments.`);
  }

  return {
    email: {
      subject: message.subject || "",
      from: message.senderEmail || message.senderName || headerValue(message.headers, "From"),
      date: message.messageDeliveryTime || message.clientSubmitTime || message.creationTime || headerValue(message.headers, "Date"),
      text: String(message.body || message.bodyHTML || "").slice(0, 6000),
    },
    attachments,
  };
}

function isRealAttachment(attachment) {
  return attachment.disposition === "attachment" || (!attachment.related && !attachment.contentId);
}

function isProcessableAttachment(attachment) {
  const ext = extensionOf(attachment.name);
  return isRealAttachment(attachment) && (PDF_EXTS.has(ext) || IMAGE_EXTS.has(ext));
}

async function fileContentToPdfBytes(bytes, name) {
  const ext = extensionOf(name);
  if (PDF_EXTS.has(ext)) return bytes;
  if (IMAGE_EXTS.has(ext)) return imageToPdfBytes(bytes, ext);
  throw new Error(`Unsupported attachment: ${name}`);
}

function headerValue(headers, name) {
  const pattern = new RegExp(`^${escapeRegExp(name)}:\\s*(.+)$`, "im");
  return String(headers || "").match(pattern)?.[1]?.trim() || "";
}

function mimeTypeFromName(name) {
  const ext = extensionOf(name);
  if (PDF_EXTS.has(ext)) return "application/pdf";
  if (ext === ".jpg" || ext === ".jpeg") return "image/jpeg";
  if (ext === ".png") return "image/png";
  if (ext === ".webp") return "image/webp";
  if (ext === ".zip") return "application/zip";
  if (ext === ".eml") return "message/rfc822";
  if (ext === ".msg") return "application/vnd.ms-outlook";
  if (ext === ".xlsx") return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  if (ext === ".xls") return "application/vnd.ms-excel";
  if (ext === ".docx") return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  if (ext === ".doc") return "application/msword";
  return "application/octet-stream";
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

async function attachmentToThumbnailDataUrl(bytes, name) {
  const ext = extensionOf(name);
  try {
    if (PDF_EXTS.has(ext)) return pdfBytesToThumbnailDataUrl(bytes);
    if (IMAGE_EXTS.has(ext)) return imageBytesToThumbnailDataUrl(bytes, ext);
  } catch {
    return "";
  }
  return "";
}

async function pdfBytesToThumbnailDataUrl(pdfBytes) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER_URL;
  const loadingTask = pdfjsLib.getDocument({ data: pdfBytes.slice ? pdfBytes.slice() : pdfBytes });
  const pdf = await loadingTask.promise;
  const page = await pdf.getPage(1);
  const baseViewport = page.getViewport({ scale: 1 });
  const scale = Math.min(180 / baseViewport.width, 132 / baseViewport.height);
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d");
  canvas.width = Math.max(1, Math.floor(viewport.width));
  canvas.height = Math.max(1, Math.floor(viewport.height));
  await page.render({ canvasContext: context, viewport }).promise;
  return canvas.toDataURL("image/png");
}

async function imageBytesToThumbnailDataUrl(bytes, ext) {
  const blobType = ext === ".webp" ? "image/webp" : ext === ".png" ? "image/png" : "image/jpeg";
  const bitmap = await createImageBitmap(new Blob([bytes], { type: blobType }));
  const maxWidth = 180;
  const maxHeight = 132;
  const scale = Math.min(maxWidth / bitmap.width, maxHeight / bitmap.height, 1);
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(bitmap.width * scale));
  canvas.height = Math.max(1, Math.round(bitmap.height * scale));
  canvas.getContext("2d").drawImage(bitmap, 0, 0, canvas.width, canvas.height);
  return canvas.toDataURL("image/png");
}

async function contentToAnalysisImages(bytes, name, label) {
  const ext = extensionOf(name);
  if (PDF_EXTS.has(ext)) {
    return renderPdfBytesToImages(bytes, label);
  }
  if (ZIP_EXTS.has(ext)) {
    const zip = await window.JSZip.loadAsync(bytes);
    const images = [];
    const entries = Object.values(zip.files)
      .filter((entry) => !entry.dir)
      .sort((left, right) => left.name.localeCompare(right.name));
    for (const entry of entries) {
      const entryExt = extensionOf(entry.name);
      if (!PDF_EXTS.has(entryExt) && !IMAGE_EXTS.has(entryExt)) continue;
      images.push(...await contentToAnalysisImages(await entry.async("uint8array"), entry.name, entry.name));
    }
    return images;
  }
  if (IMAGE_EXTS.has(ext)) {
    const image = await imageBytesToPngPayload(bytes, ext);
    return [{
      sourceName: label,
      pageNumber: 1,
      colorHint: image.colorHint,
      imageBase64: image.imageBase64,
    }];
  }
  return [];
}

async function renderPdfBytesToImages(pdfBytes, sourceName) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER_URL;
  const loadingTask = pdfjsLib.getDocument({ data: pdfBytes.slice ? pdfBytes.slice() : pdfBytes });
  const pdf = await loadingTask.promise;
  const images = [];
  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const viewport = page.getViewport({ scale: 2.2 });
    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d");
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    await page.render({ canvasContext: context, viewport }).promise;
    const colorHint = pageNumber === 1 ? detectBankColorHint(canvas) : null;
    images.push({
      sourceName,
      pageNumber,
      colorHint,
      imageBase64: canvas.toDataURL("image/png").replace(/^data:image\/png;base64,/, ""),
    });
  }
  return images;
}

async function imageBytesToPngPayload(bytes, ext) {
  if (ext === ".png") {
    const bitmap = await createImageBitmap(new Blob([bytes], { type: "image/png" }));
    const canvas = document.createElement("canvas");
    canvas.width = bitmap.width;
    canvas.height = bitmap.height;
    canvas.getContext("2d").drawImage(bitmap, 0, 0);
    return {
      colorHint: detectBankColorHint(canvas),
      imageBase64: bytesToBase64(bytes),
    };
  }
  const blobType = ext === ".webp" ? "image/webp" : "image/jpeg";
  const bitmap = await createImageBitmap(new Blob([bytes], { type: blobType }));
  const canvas = document.createElement("canvas");
  canvas.width = bitmap.width;
  canvas.height = bitmap.height;
  canvas.getContext("2d").drawImage(bitmap, 0, 0);
  const pngBlob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png"));
  if (!pngBlob) throw new Error("Could not prepare the image for AI.");
  return {
    colorHint: detectBankColorHint(canvas),
    imageBase64: bytesToBase64(new Uint8Array(await pngBlob.arrayBuffer())),
  };
}

function detectBankColorHint(canvas) {
  const context = canvas.getContext("2d", { willReadFrequently: true });
  if (!context || !canvas.width || !canvas.height) return null;

  const imageData = context.getImageData(0, 0, canvas.width, canvas.height);
  const targets = [
    { bank: "Desjardins", color: "light green", rgb: [159, 204, 113] },
    { bank: "National Bank", color: "light blue", rgb: [104, 171, 218] },
    { bank: "Scotia Bank", color: "red", rgb: [207, 67, 65] },
  ];
  const counts = Object.fromEntries(targets.map((target) => [target.bank, 0]));
  const samplePoints = randomColorPoints(imageData, canvas.width, canvas.height, 20);
  if (samplePoints.length < 6) return null;

  for (const point of samplePoints) {
    const nearest = targets
      .map((target) => ({ ...target, distance: colorDistance(point.rgb, target.rgb) }))
      .sort((left, right) => left.distance - right.distance)[0];
    if (nearest && nearest.distance <= 125) counts[nearest.bank] += 1;
  }

  const ranked = Object.entries(counts).sort((left, right) => right[1] - left[1]);
  const [bank, count] = ranked[0] || ["", 0];
  const secondCount = ranked[1]?.[1] || 0;
  if (!bank || count < 4 || count - secondCount < 2) return null;

  const target = targets.find((item) => item.bank === bank);
  return {
    bank,
    color: target?.color || "",
    confidence: Math.round((count / samplePoints.length) * 100) / 100,
    sampleCount: samplePoints.length,
    counts,
  };
}

function randomColorPoints(imageData, width, height, desiredCount) {
  const points = [];
  const maxAttempts = 1400;
  const edgePaddingX = Math.max(1, Math.floor(width * 0.03));
  const edgePaddingY = Math.max(1, Math.floor(height * 0.03));

  for (let attempt = 0; attempt < maxAttempts && points.length < desiredCount; attempt += 1) {
    const x = randomInt(edgePaddingX, Math.max(edgePaddingX + 1, width - edgePaddingX));
    const y = randomInt(edgePaddingY, Math.max(edgePaddingY + 1, height - edgePaddingY));
    const rgb = averagePatchColor(imageData, width, height, x, y, 3);
    if (isUsefulBankColorSample(rgb)) {
      points.push({ rgb });
    }
  }

  return points;
}

function randomInt(min, max) {
  return Math.floor(min + Math.random() * Math.max(1, max - min));
}

function averagePatchColor(imageData, width, height, centerX, centerY, radius) {
  let red = 0;
  let green = 0;
  let blue = 0;
  let count = 0;

  for (let y = Math.max(0, centerY - radius); y <= Math.min(height - 1, centerY + radius); y += 1) {
    for (let x = Math.max(0, centerX - radius); x <= Math.min(width - 1, centerX + radius); x += 1) {
      const index = (y * width + x) * 4;
      red += imageData.data[index];
      green += imageData.data[index + 1];
      blue += imageData.data[index + 2];
      count += 1;
    }
  }

  return [
    Math.round(red / count),
    Math.round(green / count),
    Math.round(blue / count),
  ];
}

function isUsefulBankColorSample(rgb) {
  const brightness = (rgb[0] + rgb[1] + rgb[2]) / 3;
  if (brightness < 85 || brightness > 245) return false;
  return colorfulnessScore(rgb) >= 28;
}

function colorfulnessScore([red, green, blue]) {
  const max = Math.max(red, green, blue);
  const min = Math.min(red, green, blue);
  return max - min;
}

function colorDistance(left, right) {
  return Math.sqrt(
    ((left[0] - right[0]) ** 2)
    + ((left[1] - right[1]) ** 2)
    + ((left[2] - right[2]) ** 2),
  );
}

async function imageToPdfBytes(bytes, ext) {
  const { PDFDocument } = window.PDFLib;
  const doc = await PDFDocument.create();
  const normalizedBytes = ext === ".webp" ? await webpToPngBytes(bytes) : bytes;
  const image = ext === ".jpg" || ext === ".jpeg"
    ? await doc.embedJpg(normalizedBytes)
    : await doc.embedPng(normalizedBytes);
  const maxWidth = 612;
  const maxHeight = 792;
  const scale = Math.min(maxWidth / image.width, maxHeight / image.height, 1);
  const width = image.width * scale;
  const height = image.height * scale;
  const page = doc.addPage([maxWidth, maxHeight]);
  page.drawImage(image, {
    x: (maxWidth - width) / 2,
    y: (maxHeight - height) / 2,
    width,
    height,
  });
  return doc.save();
}

async function webpToPngBytes(bytes) {
  const blob = new Blob([bytes], { type: "image/webp" });
  const bitmap = await createImageBitmap(blob);
  const canvas = document.createElement("canvas");
  canvas.width = bitmap.width;
  canvas.height = bitmap.height;
  canvas.getContext("2d").drawImage(bitmap, 0, 0);
  const pngBlob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png"));
  if (!pngBlob) throw new Error("Could not convert the WebP image.");
  return new Uint8Array(await pngBlob.arrayBuffer());
}

function drawBoldText(page, text, x, y, options) {
  page.drawText(text, { ...options, x, y });
  page.drawText(text, { ...options, x: x + 0.4, y });
}

function extensionOf(name) {
  const match = String(name || "").toLowerCase().match(/\.[^.]+$/);
  return match ? match[0] : "";
}

function toUint8Array(value) {
  if (value instanceof Uint8Array) return value;
  if (value instanceof ArrayBuffer) return new Uint8Array(value);
  if (ArrayBuffer.isView(value)) return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  return new Uint8Array();
}

function bytesToBase64(bytes) {
  let binary = "";
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.length; index += chunkSize) {
    binary += String.fromCharCode(...bytes.slice(index, index + chunkSize));
  }
  return btoa(binary);
}
