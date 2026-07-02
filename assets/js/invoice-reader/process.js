import { state } from "./state.js";
import { buildSuggestedFileName, joinSharePointPath, normalizeCardLast4, toFloat } from "./utils.js";
import { deleteSharePointItemById, listChildrenByPath, loadDatabaseRows, saveDatabaseRows, uploadSharePointFile } from "./graph.js";
import { splitPdfToPages } from "./pdf.js";

const TAX_FIELDS = ["gst", "hst", "pst", "qst", "tps", "iva", "vat", "retention"];

export async function processUploadedPdf(elements) {
  const file = elements.pdfInput.files[0];
  const receiptType = elements.receiptTypeInput.value;
  const isReimbursement = receiptType === "reimbursement";
  if (!state.graphToken) {
    throw new Error("Microsoft es obligatorio para guardar PDFs y CSV.");
  }
  if (!file) {
    throw new Error("Selecciona un PDF.");
  }

  const last4 = isReimbursement ? "" : normalizeCardLast4(elements.cardLast4Input.value);
  const description = isReimbursement ? elements.descriptionInput.value.trim() : "";
  if (!isReimbursement && !last4) {
    throw new Error("Card last 4 digits debe tener exactamente 4 numeros.");
  }
  if (isReimbursement && !description) {
    throw new Error("La descripcion es obligatoria para reembolsos.");
  }

  const pages = await splitPdfToPages(file);
  if (!pages.length) {
    throw new Error("No se detectaron paginas en el PDF.");
  }

  const database = await loadDatabaseRows();
  const duplicateScopeRows = [...database.rows];
  const existingNames = new Set((await listChildrenByPath(state.config.receiptsDatabaseDir)).map((item) => item.name));
  const summaryRows = [];
  const rawRows = [];
  const pendingUploads = [];
  const errors = [];

  for (let index = 0; index < pages.length; index += 1) {
    const page = pages[index];
    elements.progressBar.value = Math.round(((index + 1) / pages.length) * 100);
    elements.progressLabel.textContent = `Procesando pagina ${page.pageNumber} de ${pages.length}`;
    const pdfSha256 = await sha256Hex(page.pdfBytes);

    const exactDuplicate = duplicateScopeRows.find((row) => row.pdf_sha256 && row.pdf_sha256 === pdfSha256);
    if (exactDuplicate) {
      errors.push(`Pagina ${page.pageNumber}: omitida porque ya existe como ${exactDuplicate.file_name || "invoice previo"}.`);
      continue;
    }

    let result;
    try {
      result = await classifyPage({
        pageNumber: page.pageNumber,
        imageBase64: page.imageBase64,
        receiptType,
      });
    } catch (error) {
      errors.push(`Pagina ${page.pageNumber}: ${error.message}`);
      continue;
    }

    const baseName = buildSuggestedFileName(
      result.gpt.payment_date || result.compact.date,
      isReimbursement ? "reembolso" : elements.bankInput.value,
      isReimbursement ? "" : elements.cardTypeInput.value,
      result.gpt.merchant_name || result.compact.merchant,
      toFloat(result.gpt.total_amount, result.compact.total),
    );

    const row = {
      status: "Pending",
      processed_at: new Date().toISOString().slice(0, 19),
      source_page_number: page.pageNumber,
      receipt_type: receiptType,
      company: elements.companyInput.value,
      bank: isReimbursement ? "" : elements.bankInput.value,
      card_type: isReimbursement ? "" : elements.cardTypeInput.value,
      card_last4: last4,
      gpt_payment_date: result.gpt.payment_date || result.compact.date || "",
      gpt_total_amount: toFloat(result.gpt.total_amount, result.compact.total),
      gpt_taxes_total: toFloat(result.gpt.taxes_total, result.compact.taxes_total),
      ...buildTaxColumns("gpt", result.gpt),
      gpt_category: result.gpt.category || "Diverse Expenses",
      gpt_merchant_name: result.gpt.merchant_name || result.compact.merchant || "",
      gpt_city: result.gpt.city || result.compact.city || "",
      gpt_province: result.gpt.province || result.compact.province || "",
      gpt_ticket_number: result.gpt.ticket_number || "",
      gpt_description: description,
      gpt_confidence: toFloat(result.gpt.confidence, 0),
      notes: result.gpt.notes || "",
      file_name: "",
      file_path: "",
      pdf_sha256: pdfSha256,
    };
    row.invoice_fingerprint = buildInvoiceFingerprint(row);

    const logicalDuplicate = findDuplicateInvoice(row, duplicateScopeRows);
    if (logicalDuplicate) {
      errors.push(`Pagina ${page.pageNumber}: posible duplicado omitido (${describeDuplicate(logicalDuplicate)}).`);
      continue;
    }

    const uniqueName = makeUniquePdfName(baseName, existingNames);
    const remotePath = joinSharePointPath(state.config.receiptsDatabaseDir, uniqueName);
    row.file_name = uniqueName;
    row.file_path = remotePath;
    pendingUploads.push({
      fileName: uniqueName,
      filePath: remotePath,
      content: page.pdfBytes,
    });
    summaryRows.push(row);
    duplicateScopeRows.push(row);
    rawRows.push({
      ...row,
      raw_google_vision_json: JSON.stringify(result.vision),
      raw_gpt_json: JSON.stringify(result.gpt),
      vision_payment_date: result.compact.date || "",
      vision_total_amount: result.compact.total || 0,
      vision_taxes_total: result.compact.taxes_total || 0,
      ...buildTaxColumns("vision", result.compact),
      vision_merchant_name: result.compact.merchant || "",
      vision_city: result.compact.city || "",
      vision_province: result.compact.province || "",
    });
  }

  state.processed = {
    summaryRows,
    rawRows,
    pendingUploads,
    saved: false,
    errors,
  };
  elements.progressLabel.textContent = errors.length
    ? `Completado con ${errors.length} pagina(s) con error`
    : "Procesamiento completo";
  return state.processed;
}

async function classifyPage(payload) {
  const response = await fetch("/.netlify/functions/process-receipt", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  if (!response.ok) {
    throw new Error(await response.text());
  }
  return response.json();
}

function makeUniquePdfName(baseName, existingNames) {
  const safeBase = String(baseName || "invoice").replace(/[^\w-]+/g, "_");
  let candidate = `${safeBase}.pdf`;
  let counter = 2;
  while (existingNames.has(candidate)) {
    candidate = `${safeBase}__${counter}.pdf`;
    counter += 1;
  }
  existingNames.add(candidate);
  return candidate;
}

export async function keepProcessedResults() {
  const database = await loadDatabaseRows();
  const duplicateScopeRows = [...database.rows];

  for (const pending of state.processed.pendingUploads) {
    if (!pending.content?.byteLength) {
      throw new Error(`No se guardo ningun PDF: ${pending.fileName} esta vacio. Vuelve a procesar el archivo antes de guardar.`);
    }
  }

  for (const row of state.processed.summaryRows) {
    const duplicate = findDuplicateInvoice(row, duplicateScopeRows);
    if (duplicate) {
      throw new Error(`No se guardo ningun PDF: ${row.file_name} parece duplicado de ${describeDuplicate(duplicate)}. Recarga y revisa la base.`);
    }
    duplicateScopeRows.push(row);
  }

  const uploadedItems = [];
  try {
    for (const pending of state.processed.pendingUploads) {
      const uploaded = await uploadSharePointFile(pending.filePath, pending.content, "application/pdf");
      if (uploaded?.id) {
        uploadedItems.push(uploaded);
      }
      if (uploaded?.size === 0) {
        throw new Error(`${pending.fileName} llego a SharePoint con peso 0.`);
      }
    }
  } catch (error) {
    try {
      await rollbackUploadedItems(uploadedItems);
    } catch (rollbackError) {
      throw new Error(`El lote fallo y el CSV no fue actualizado: ${error.message} ${rollbackError.message}`);
    }
    throw new Error(`No se guardo ningun PDF ni se actualizo el CSV: ${error.message}`);
  }

  await saveDatabaseRows([...database.rows, ...state.processed.summaryRows], { expectedEtag: database.eTag });
  state.processed.saved = true;
}

function buildTaxColumns(prefix, source) {
  return Object.fromEntries(
    TAX_FIELDS.map((field) => [`${prefix}_${field}`, toFloat(source?.[field], 0)]),
  );
}

async function rollbackUploadedItems(items) {
  const failures = [];
  for (const item of items.reverse()) {
    try {
      await deleteSharePointItemById(item.id);
    } catch (error) {
      failures.push(item.name || item.id);
    }
  }
  if (failures.length) {
    throw new Error(`No se pudo deshacer la subida de: ${failures.join(", ")}.`);
  }
}

async function sha256Hex(bytes) {
  const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  const digest = await crypto.subtle.digest("SHA-256", buffer);
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, "0")).join("");
}

function findDuplicateInvoice(row, existingRows) {
  return existingRows.find((existing) => invoicesMatch(row, existing));
}

function invoicesMatch(left, right) {
  if (left.pdf_sha256 && right.pdf_sha256 && left.pdf_sha256 === right.pdf_sha256) {
    return true;
  }

  const leftFingerprint = left.invoice_fingerprint || buildInvoiceFingerprint(left);
  const rightFingerprint = right.invoice_fingerprint || buildInvoiceFingerprint(right);
  if (leftFingerprint && rightFingerprint && leftFingerprint === rightFingerprint) {
    return true;
  }

  if (paymentSourceKey(left) !== paymentSourceKey(right)) return false;
  if (amountKey(left.gpt_total_amount) !== amountKey(right.gpt_total_amount)) return false;

  const leftTicket = normalizeIdentifier(left.gpt_ticket_number);
  const rightTicket = normalizeIdentifier(right.gpt_ticket_number);
  if (leftTicket && rightTicket && leftTicket === rightTicket) {
    return extractDateOnly(left.gpt_payment_date) === extractDateOnly(right.gpt_payment_date)
      || merchantsRoughlyMatch(left.gpt_merchant_name, right.gpt_merchant_name);
  }

  return Boolean(
    extractDateOnly(left.gpt_payment_date)
      && extractDateOnly(left.gpt_payment_date) === extractDateOnly(right.gpt_payment_date)
      && merchantsRoughlyMatch(left.gpt_merchant_name, right.gpt_merchant_name),
  );
}

function buildInvoiceFingerprint(row) {
  const source = paymentSourceKey(row);
  const amount = amountKey(row.gpt_total_amount);
  const ticket = normalizeIdentifier(row.gpt_ticket_number);
  const date = extractDateOnly(row.gpt_payment_date);
  const merchant = normalizeMerchant(row.gpt_merchant_name);

  if (!source || !amount) return "";
  if (ticket) return ["ticket", source, ticket, amount, date].join("|");
  if (date && merchant) return ["scan", source, date, amount, merchant].join("|");
  return "";
}

function paymentSourceKey(row) {
  return [
    normalizeIdentifier(row.receipt_type || "bank_transaction"),
    normalizeIdentifier(row.company),
    normalizeIdentifier(row.bank),
    normalizeIdentifier(row.card_type),
    normalizeIdentifier(row.card_last4),
  ].join("|");
}

function amountKey(value) {
  const amount = toFloat(value, NaN);
  return Number.isFinite(amount) ? String(Math.round(amount * 100)) : "";
}

function extractDateOnly(value) {
  return String(value || "").match(/\d{4}-\d{2}-\d{2}/)?.[0] || "";
}

function normalizeIdentifier(value) {
  return String(value || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().replace(/[^A-Z0-9]+/g, "");
}

function normalizeMerchant(value) {
  return normalizeIdentifier(value)
    .replace(/\b(INC|LTD|LLC|CORP|CORPORATION|LIMITED|COMPANY|CO)\b/g, "")
    .replace(/0/g, "O")
    .replace(/1/g, "I");
}

function merchantsRoughlyMatch(left, right) {
  const normalizedLeft = normalizeMerchant(left);
  const normalizedRight = normalizeMerchant(right);
  if (!normalizedLeft || !normalizedRight) return false;
  if (normalizedLeft === normalizedRight) return true;
  if (normalizedLeft.length >= 6 && normalizedRight.length >= 6) {
    return normalizedLeft.includes(normalizedRight) || normalizedRight.includes(normalizedLeft);
  }
  return false;
}

function describeDuplicate(row) {
  return [
    row.file_name,
    extractDateOnly(row.gpt_payment_date),
    row.gpt_merchant_name,
    row.gpt_total_amount,
  ].filter(Boolean).join(" / ") || "invoice previo";
}
