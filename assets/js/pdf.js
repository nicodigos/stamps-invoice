import * as pdfjsLib from "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.min.mjs";

const PDF_EXTS = new Set([".pdf"]);
const ZIP_EXTS = new Set([".zip"]);
const IMAGE_EXTS = new Set([".jpg", ".jpeg", ".png", ".webp"]);
const EML_EXTS = new Set([".eml"]);
const POSTAL_MIME_URL = "https://esm.sh/postal-mime@2.4.3";
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
  if (EML_EXTS.has(ext)) {
    const parsed = await parseEmailFile(file);
    return Promise.all(parsed.attachments.map((attachment) => fileContentToPdfBytes(attachment.bytes, attachment.name)));
  }
  throw new Error(`Tipo de archivo no soportado: ${ext || file.type || "sin extension"}.`);
}

export async function fileToAnalysisPayload(file) {
  const ext = extensionOf(file.name);
  if (EML_EXTS.has(ext)) {
    const parsed = await parseEmailFile(file);
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
    throw new Error("No hay paginas para combinar.");
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
    throw new Error("El ZIP no contiene PDFs o imagenes soportadas.");
  }

  return out;
}

async function parseEmailFile(file) {
  const { default: PostalMime } = await import(POSTAL_MIME_URL);
  const email = await PostalMime.parse(await file.arrayBuffer());
  const attachments = (email.attachments || [])
    .map((attachment) => ({
      name: attachment.filename || attachment.name || "attachment",
      mimeType: attachment.mimeType || attachment.contentType || "",
      disposition: attachment.disposition || "",
      contentId: attachment.contentId || "",
      related: Boolean(attachment.related),
      bytes: toUint8Array(attachment.content),
    }))
    .filter((attachment) => isProcessableAttachment(attachment));

  if (!attachments.length) {
    throw new Error(`El email "${file.name}" no tiene adjuntos PDF o imagen soportados.`);
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

function isProcessableAttachment(attachment) {
  const ext = extensionOf(attachment.name);
  const realAttachment = attachment.disposition === "attachment" || (!attachment.related && !attachment.contentId);
  return realAttachment && (PDF_EXTS.has(ext) || IMAGE_EXTS.has(ext));
}

async function fileContentToPdfBytes(bytes, name) {
  const ext = extensionOf(name);
  if (PDF_EXTS.has(ext)) return bytes;
  if (IMAGE_EXTS.has(ext)) return imageToPdfBytes(bytes, ext);
  throw new Error(`Adjunto no soportado: ${name}`);
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
    return [{
      sourceName: label,
      pageNumber: 1,
      imageBase64: await imageBytesToPngBase64(bytes, ext),
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
    images.push({
      sourceName,
      pageNumber,
      imageBase64: canvas.toDataURL("image/png").replace(/^data:image\/png;base64,/, ""),
    });
  }
  return images;
}

async function imageBytesToPngBase64(bytes, ext) {
  if (ext === ".png") {
    return bytesToBase64(bytes);
  }
  const blobType = ext === ".webp" ? "image/webp" : "image/jpeg";
  const bitmap = await createImageBitmap(new Blob([bytes], { type: blobType }));
  const canvas = document.createElement("canvas");
  canvas.width = bitmap.width;
  canvas.height = bitmap.height;
  canvas.getContext("2d").drawImage(bitmap, 0, 0);
  const pngBlob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png"));
  if (!pngBlob) throw new Error("No se pudo preparar la imagen para AI.");
  return bytesToBase64(new Uint8Array(await pngBlob.arrayBuffer()));
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
  if (!pngBlob) throw new Error("No se pudo convertir la imagen WebP.");
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
