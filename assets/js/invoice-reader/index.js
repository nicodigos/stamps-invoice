import { state as shellState } from "../state.js";
import { clearFlash, collectRowColumns, formatColumnLabel, showFlash, $ } from "./utils.js";
import { resolveDriveId } from "./graph.js";
import { state, syncInvoiceState } from "./state.js";
import {
  applyFilters,
  downloadFilteredExcel,
  downloadFilteredPdfs,
  goToNextPage,
  goToPreviousPage,
  refreshDatabase,
  saveDatabaseEdits,
} from "./table.js";
import { keepProcessedResults, processUploadedPdf } from "./process.js";

const SUMMARY_PREFERRED_COLUMNS = [
  "source_page_number",
  "receipt_type",
  "company",
  "gpt_payment_date",
  "gpt_total_amount",
  "gpt_taxes_total",
  "gpt_gst",
  "gpt_hst",
  "gpt_pst",
  "gpt_qst",
  "gpt_tps",
  "gpt_iva",
  "gpt_vat",
  "gpt_retention",
  "gpt_category",
  "gpt_merchant_name",
  "gpt_ticket_number",
  "gpt_city",
  "gpt_province",
  "gpt_description",
  "notes",
  "file_name",
];

let initialized = false;

export async function initInvoiceReader() {
  if (initialized) return;
  initialized = true;
  syncInvoiceState(shellState);
  wireButtons();
  populateSelects();
  syncReceiptTypeFields();
  syncHeaderControls();
  if (state.graphToken) {
    await resolveDriveId();
    await refreshDatabase(elements);
  }
}

export function syncInvoiceReaderAuth() {
  syncInvoiceState(shellState);
  syncHeaderControls();
  if (!state.graphToken) {
    state.databaseRows = [];
    state.filteredRows = [];
    elements.dbTableHead.innerHTML = "";
    elements.dbTableBody.innerHTML = "";
    elements.pageIndicator.textContent = "Conecta Microsoft para cargar la base.";
    setFiltersSidebar(false);
  }
}

const elements = {
  appShell: $(".app-shell"),
  refreshDbBtn: $("#invoice-refresh-db-btn"),
  saveDbBtn: $("#invoice-save-db-btn"),
  downloadExcelBtn: $("#invoice-download-excel-btn"),
  downloadPdfsBtn: $("#invoice-download-pdfs-btn"),
  filtersAccordionPanel: $("#invoice-filters-accordion-panel"),
  showFiltersBtn: $("#invoice-show-filters-btn"),
  hideFiltersBtn: $("#invoice-hide-filters-btn"),
  sidebarBackdrop: $("#invoice-sidebar-backdrop"),
  clearFiltersBtn: $("#invoice-clear-filters-btn"),
  databaseLayout: $("#invoice-database-layout"),
  filtersContainer: $("#invoice-filters-container"),
  dbTableHead: $("#invoice-db-table thead"),
  dbTableBody: $("#invoice-db-table tbody"),
  prevPageBtn: $("#invoice-prev-page-btn"),
  nextPageBtn: $("#invoice-next-page-btn"),
  pageIndicator: $("#invoice-page-indicator"),
  pdfInput: $("#invoice-pdf-input"),
  receiptTypeInput: $("#invoice-receipt-type-input"),
  companyInput: $("#invoice-company-input"),
  bankInput: $("#invoice-bank-input"),
  bankField: $("#invoice-bank-field"),
  cardTypeInput: $("#invoice-card-type-input"),
  cardTypeField: $("#invoice-card-type-field"),
  cardLast4Input: $("#invoice-card-last4-input"),
  cardLast4Field: $("#invoice-card-last4-field"),
  descriptionInput: $("#invoice-description-input"),
  descriptionField: $("#invoice-description-field"),
  processBtn: $("#invoice-process-btn"),
  keepResultsBtn: $("#invoice-keep-results-btn"),
  dropResultsBtn: $("#invoice-drop-results-btn"),
  downloadSummaryBtn: $("#invoice-download-summary-btn"),
  progressBar: $("#invoice-progress-bar"),
  progressLabel: $("#invoice-progress-label"),
  uploadCaption: $("#invoice-upload-caption"),
  summaryTableHead: $("#invoice-summary-table thead"),
  summaryTableBody: $("#invoice-summary-table tbody"),
  rawOutput: $("#invoice-raw-output"),
};

function wireButtons() {
  elements.refreshDbBtn.addEventListener("click", async () => runGuarded(() => refreshDatabase(elements)));
  elements.showFiltersBtn.addEventListener("click", () => setFiltersSidebar(true));
  elements.hideFiltersBtn.addEventListener("click", () => setFiltersSidebar(false));
  elements.sidebarBackdrop.addEventListener("click", () => setFiltersSidebar(false));
  elements.saveDbBtn.addEventListener("click", async () => runGuarded(async () => {
    await saveDatabaseEdits();
    showFlash("Cambios guardados en el CSV de SharePoint.");
  }));
  elements.downloadExcelBtn.addEventListener("click", () => runGuarded(() => downloadFilteredExcel()));
  elements.downloadPdfsBtn.addEventListener("click", () => runGuarded(() => downloadFilteredPdfs()));
  elements.prevPageBtn.addEventListener("click", () => goToPreviousPage(elements));
  elements.nextPageBtn.addEventListener("click", () => goToNextPage(elements));
  elements.receiptTypeInput.addEventListener("change", syncReceiptTypeFields);
  elements.clearFiltersBtn.addEventListener("click", () => {
    document.querySelectorAll("#invoice-filters-container input, #invoice-filters-container select").forEach((input) => {
      if (input.type === "range") {
        input.value = input.id.endsWith("-min") ? input.min : input.max;
      } else {
        input.value = input.id === "filter-status" ? "all" : "";
      }
    });
    document.querySelectorAll(".range-filter-values span").forEach((label) => {
      const relatedInput = document.querySelector(`#${label.id.replace("-label", "")}`);
      if (relatedInput) {
        label.textContent = Number(relatedInput.value).toFixed(2).replace(/\.00$/, "");
      }
    });
    applyFilters(elements);
  });
  elements.filtersContainer.addEventListener("input", () => applyFilters(elements));
  elements.filtersContainer.addEventListener("change", () => applyFilters(elements));
  elements.processBtn.addEventListener("click", async () => runGuarded(handleProcess));
  elements.keepResultsBtn.addEventListener("click", async () => runGuarded(handleKeepResults));
  elements.dropResultsBtn.addEventListener("click", handleDropResults);
  elements.downloadSummaryBtn.addEventListener("click", handleDownloadSummary);
}

function setFiltersSidebar(opened) {
  elements.databaseLayout.classList.toggle("is-collapsed", !opened);
  elements.databaseLayout.classList.toggle("is-sidebar-open", opened);
  elements.filtersAccordionPanel.classList.toggle("is-open", opened);
  elements.sidebarBackdrop.hidden = !opened;
  elements.appShell.classList.toggle("sidebar-open", opened);
  elements.appShell.classList.toggle("sidebar-collapsed", !opened);
  syncHeaderControls();
}

function populateSelects() {
  const companies = state.config.receiptCompanyOptions || state.config.companyOptions || [];
  const banks = state.config.receiptBankOptions || state.config.bankOptions || [];
  elements.companyInput.innerHTML = companies.map((item) => `<option value="${item}">${item}</option>`).join("");
  elements.bankInput.innerHTML = banks.map((item) => `<option value="${item}">${item}</option>`).join("");
}

function syncHeaderControls() {
  const connected = Boolean(state.graphToken);
  const databaseTabActive = $("#invoice-database-tab")?.classList.contains("is-active");
  const sidebarOpen = elements.databaseLayout.classList.contains("is-sidebar-open");
  elements.showFiltersBtn.hidden = !connected || !databaseTabActive || sidebarOpen;
}

function syncReceiptTypeFields() {
  const receiptType = elements.receiptTypeInput.value;
  const isReimbursement = receiptType === "reimbursement";
  elements.bankField.hidden = isReimbursement;
  elements.bankInput.disabled = isReimbursement;
  elements.cardTypeField.hidden = isReimbursement;
  elements.cardTypeInput.disabled = isReimbursement;
  elements.cardLast4Field.hidden = isReimbursement;
  elements.descriptionField.hidden = !isReimbursement;
  elements.cardLast4Input.disabled = isReimbursement;
  elements.descriptionInput.disabled = !isReimbursement;
  if (isReimbursement) {
    elements.bankInput.selectedIndex = 0;
    elements.cardTypeInput.selectedIndex = 0;
    elements.cardLast4Input.value = "";
  } else {
    elements.descriptionInput.value = "";
  }
}

async function handleProcess() {
  clearFlash();
  elements.progressBar.value = 0;
  elements.progressLabel.textContent = "Preparando procesamiento";
  const result = await processUploadedPdf(elements);
  renderSummary();
  elements.keepResultsBtn.disabled = !result.summaryRows.length;
  elements.dropResultsBtn.disabled = !result.summaryRows.length;
  elements.downloadSummaryBtn.disabled = !result.summaryRows.length;
  elements.uploadCaption.textContent = result.summaryRows.length
    ? `Se procesaron ${result.summaryRows.length} pagina(s).`
    : "No hubo paginas procesadas.";
  if (result.errors?.length) {
    showFlash(result.errors.join(" | "), "warning");
  } else {
    showFlash("Todas las paginas fueron procesadas.");
  }
}

async function handleKeepResults() {
  await keepProcessedResults();
  elements.keepResultsBtn.disabled = true;
  await refreshDatabase(elements);
  showFlash("Resultados persistidos en SharePoint.");
}

function handleDropResults() {
  state.processed = { summaryRows: [], rawRows: [], pendingUploads: [], saved: false };
  elements.summaryTableHead.innerHTML = "";
  elements.summaryTableBody.innerHTML = "";
  elements.rawOutput.textContent = "";
  elements.keepResultsBtn.disabled = true;
  elements.dropResultsBtn.disabled = true;
  elements.downloadSummaryBtn.disabled = true;
  elements.uploadCaption.textContent = "Resultados descartados.";
  showFlash("Resultados descartados.");
}

function handleDownloadSummary() {
  const rows = state.processed.summaryRows;
  if (!rows.length) return;
  const columns = getSummaryColumns(rows);
  const exportRows = rows.map((row) => Object.fromEntries(
    columns.map((column) => [formatColumnLabel(column), displaySummaryValue(column, row[column] ?? "")]),
  ));
  const worksheet = XLSX.utils.json_to_sheet(exportRows);
  worksheet["!autofilter"] = {
    ref: XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: exportRows.length, c: Math.max(0, columns.length - 1) },
    }),
  };
  worksheet["!cols"] = columns.map((column) => ({ wch: getSummaryColumnWidth(column) }));
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "summary");
  XLSX.writeFile(workbook, "invoice_summary_google_vision.xlsx");
}

function renderSummary() {
  const rows = state.processed.summaryRows;
  if (!rows.length) {
    elements.summaryTableHead.innerHTML = "";
    elements.summaryTableBody.innerHTML = "";
    elements.rawOutput.textContent = JSON.stringify(state.processed.rawRows, null, 2);
    return;
  }

  const columns = getSummaryColumns(rows);
  elements.summaryTableHead.innerHTML = `<tr>${columns.map((column) => `<th>${column}</th>`).join("")}</tr>`;
  elements.summaryTableBody.innerHTML = rows.map((row) => `
    <tr>
      ${columns.map((column) => `<td>${escapeHtml(displaySummaryValue(column, row[column] ?? ""))}</td>`).join("")}
    </tr>
  `).join("");

  elements.rawOutput.textContent = JSON.stringify(state.processed.rawRows, null, 2);
}

async function runGuarded(work) {
  try {
    syncInvoiceState(shellState);
    if (!state.graphToken) throw new Error("Conecta Microsoft antes de usar Invoice App.");
    clearFlash();
    await work();
  } catch (error) {
    showFlash(error.message, "error");
  }
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function displaySummaryValue(column, value) {
  if (column === "receipt_type") {
    return String(value).trim().toLowerCase() === "reimbursement" ? "Reembolso" : "Transaccion bancaria";
  }
  return value;
}

function getSummaryColumns(rows) {
  const availableColumns = new Set(collectRowColumns(rows));
  return SUMMARY_PREFERRED_COLUMNS.filter((column) => availableColumns.has(column));
}

function getSummaryColumnWidth(column) {
  if (["source_page_number", "gpt_gst", "gpt_hst", "gpt_pst", "gpt_qst", "gpt_tps", "gpt_iva", "gpt_vat", "gpt_retention"].includes(column)) {
    return 12;
  }
  if (["gpt_total_amount", "gpt_taxes_total"].includes(column)) {
    return 14;
  }
  if (["receipt_type", "company", "gpt_category", "gpt_city", "gpt_province", "gpt_ticket_number"].includes(column)) {
    return 18;
  }
  if (["gpt_merchant_name", "gpt_description", "notes"].includes(column)) {
    return 26;
  }
  if (column === "file_name") {
    return 40;
  }
  return 18;
}
