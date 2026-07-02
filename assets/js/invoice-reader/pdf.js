import * as pdfjsLib from "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.min.mjs";

const { PDFDocument } = window.PDFLib;

export async function splitPdfToPages(file) {
  const originalBytes = await file.arrayBuffer();
  const sourcePdf = await PDFDocument.load(originalBytes);
  const pages = [];

  for (let index = 0; index < sourcePdf.getPageCount(); index += 1) {
    const singlePdf = await PDFDocument.create();
    const [copiedPage] = await singlePdf.copyPages(sourcePdf, [index]);
    singlePdf.addPage(copiedPage);
    const pdfBytes = await singlePdf.save();
    const imageBase64 = await renderPdfPageToPng(new Uint8Array(pdfBytes), index + 1);
    pages.push({
      pageNumber: index + 1,
      pdfBytes,
      imageBase64,
    });
  }

  return pages;
}

async function renderPdfPageToPng(pdfBytes, pageNumber) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.worker.min.mjs";
  const loadingTask = pdfjsLib.getDocument({ data: pdfBytes });
  const pdf = await loadingTask.promise;
  const page = await pdf.getPage(1);
  const viewport = page.getViewport({ scale: 2.2 });
  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d");
  canvas.width = viewport.width;
  canvas.height = viewport.height;
  await page.render({ canvasContext: context, viewport }).promise;
  return canvas.toDataURL("image/png").replace(/^data:image\/png;base64,/, "");
}

export async function mergePdfBlobs(pdfByteArrays) {
  const merged = await PDFDocument.create();
  for (const bytes of pdfByteArrays) {
    const doc = await PDFDocument.load(bytes);
    const copiedPages = await merged.copyPages(doc, doc.getPageIndices());
    copiedPages.forEach((page) => merged.addPage(page));
  }
  return merged.save();
}
