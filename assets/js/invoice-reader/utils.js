export function $(selector) {
  return document.querySelector(selector);
}

export function showFlash(message, type = "info") {
  const flash = $("#flash");
  flash.hidden = false;
  flash.className = `flash ${type}`;
  flash.textContent = message;
}

export function clearFlash() {
  const flash = $("#flash");
  flash.hidden = true;
  flash.textContent = "";
  flash.className = "flash";
}

export function normalizeCardLast4(value) {
  const cleaned = String(value || "").replace(/\D/g, "");
  return cleaned.length === 4 ? `N${cleaned}` : "";
}

export function displayCardLast4(value) {
  const normalized = normalizeCardLast4(value);
  return normalized ? normalized.slice(1) : "";
}

export function toFloat(value, fallback = 0) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

export function sanitizeFilenameComponent(value, fallback) {
  const text = String(value || "").trim();
  if (!text) return fallback;
  return text.replace(/[^\w-]+/g, "_").replace(/_+/g, "_").replace(/^_+|_+$/g, "") || fallback;
}

export function buildSuggestedFileName(paymentDate, bank, cardType, merchantName, totalAmount) {
  const datePart = String(paymentDate || "").match(/\d{4}-\d{2}-\d{2}/)?.[0] || String(paymentDate || "");
  const amountPart = toFloat(totalAmount, 0).toFixed(2);
  return [
    sanitizeFilenameComponent(datePart, "no_date"),
    sanitizeFilenameComponent(bank, "no_bank"),
    sanitizeFilenameComponent(cardType, "no_card"),
    sanitizeFilenameComponent(merchantName, "no_merchant"),
    sanitizeFilenameComponent(amountPart, "0_00"),
  ].join("__");
}

export function formatColumnLabel(columnName) {
  const text = String(columnName || "").replace(/^(gpt|vision)_/, "");
  if (text === "gpt_description" || text === "description") return "Description";
  return text.replaceAll("_", " ").replace(/\b\w/g, (char) => char.toUpperCase());
}

export function csvEscape(value) {
  const text = value == null ? "" : String(value);
  if (/[",\n]/.test(text)) {
    return `"${text.replaceAll('"', '""')}"`;
  }
  return text;
}

export function rowsToCsv(rows) {
  if (!rows.length) return "";
  const columns = collectRowColumns(rows);
  const lines = [
    columns.join(","),
    ...rows.map((row) => columns.map((column) => csvEscape(row[column])).join(",")),
  ];
  return `\ufeff${lines.join("\n")}`;
}

export function collectRowColumns(rows) {
  const columns = [];
  const seen = new Set();
  rows.forEach((row) => {
    Object.keys(row).forEach((column) => {
      if (seen.has(column)) return;
      seen.add(column);
      columns.push(column);
    });
  });
  return columns;
}

export function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

export function joinSharePointPath(...parts) {
  return parts.map((part) => String(part || "").trim().replace(/^\/+|\/+$/g, "")).filter(Boolean).join("/");
}
