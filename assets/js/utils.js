export const $ = (selector) => document.querySelector(selector);

export function showFlash(message, type = "info") {
  const flash = $("#flash");
  flash.textContent = message;
  flash.className = `flash ${type}`;
  flash.hidden = false;
}

export function clearFlash() {
  const flash = $("#flash");
  flash.textContent = "";
  flash.hidden = true;
}

export function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

export function sanitizeFilename(text, fallback = "INV") {
  const cleaned = String(text ?? "").trim().replace(/[^A-Za-z0-9_-]+/g, "_");
  return cleaned || fallback;
}

export function joinSharePointPath(...parts) {
  return parts
    .flatMap((part) => String(part ?? "").split("/"))
    .map((part) => part.trim())
    .filter(Boolean)
    .join("/");
}

export function parseIsoDate(value) {
  const date = new Date(`${value}T00:00:00`);
  if (!value || Number.isNaN(date.getTime())) {
    throw new Error("Fecha invalida. Usa YYYY-MM-DD.");
  }
  return date;
}

export function monthName(date) {
  return date.toLocaleString("en-US", { month: "long" });
}

export function yyyymmdd(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}${month}${day}`;
}

export function downloadBlobUrl(bytes, type) {
  return URL.createObjectURL(new Blob([bytes], { type }));
}
