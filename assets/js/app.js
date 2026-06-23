import { loadConfig } from "./config.js";
import { state } from "./state.js";
import {
  BANKS,
  BANK_TO_CHECKINGS_DIR,
  CATEGORIES,
  CATEGORY_TO_MONTH_REL_PATH,
  COMPANIES,
  COMPANY_SHAREPOINT_DIRS,
} from "./data.js";
import { uploadSharePointFile } from "./graph.js";
import { emailFileToPreview, fileToAnalysisPayload, mergePdfBytes, stampPdfBytes, uploadedFileToPdfBytes } from "./pdf.js";
import {
  $,
  clearFlash,
  downloadBlobUrl,
  escapeHtml,
  joinSharePointPath,
  monthName,
  parseIsoDate,
  sanitizeFilename,
  showFlash,
  yyyymmdd,
} from "./utils.js";

const elements = {
  microsoftAuthBtn: $("#microsoft-auth-btn"),
  connectCta: $("#connect-cta"),
  connectPanel: $("#connect-panel"),
  mainPanel: $("#main-panel"),
  stampForm: $("#stamp-form"),
  invoiceFile: $("#invoice-file"),
  invoiceDropzone: $("#invoice-dropzone"),
  dropzoneTitle: $("#dropzone-title"),
  dropzoneCaption: $("#dropzone-caption"),
  uploadList: $("#upload-list"),
  emailPreviewPanel: $("#email-preview-panel"),
  emailPreviewCount: $("#email-preview-count"),
  emailPreviewList: $("#email-preview-list"),
  dateInput: $("#date-input"),
  companyInput: $("#company-input"),
  bankInput: $("#bank-input"),
  categoryInput: $("#category-input"),
  paymentCodeInput: $("#payment-code-input"),
  clientInvoiceInput: $("#client-invoice-input"),
  autofillBtn: $("#autofill-btn"),
  processBtn: $("#process-btn"),
  stampResult: $("#stamp-result"),
  bankPathOutput: $("#bank-path-output"),
  categoryPathOutput: $("#category-path-output"),
  downloadLink: $("#download-link"),
  invoiceDownloadBtn: $("#invoice-download-btn"),
  downloadStatusText: $("#download-status-text"),
  downloadCount: $("#download-count"),
  downloadProgressFill: $("#download-progress-fill"),
  downloadTotalCount: $("#download-total-count"),
  downloadSuccessCount: $("#download-success-count"),
  downloadSkippedCount: $("#download-skipped-count"),
  downloadErrors: $("#download-errors"),
  downloadErrorList: $("#download-error-list"),
  refreshFileTreeBtn: $("#refresh-file-tree-btn"),
  fileTreeSearch: $("#file-tree-search"),
  fileTreeTotal: $("#file-tree-total"),
  fileTree: $("#file-tree"),
};

boot().catch((error) => showFlash(error.message, "error"));

let queuedFiles = [];
let currentFileTreeResult = null;
let currentEmailPreviews = [];
let emailPreviewRenderId = 0;

async function boot() {
  wireTabs();
  wireButtons();
  await loadConfig();
  populateSelect(elements.companyInput, COMPANIES);
  populateSelect(elements.bankInput, BANKS);
  populateSelect(elements.categoryInput, CATEGORIES);
  elements.dateInput.value = new Date().toISOString().slice(0, 10);
  await completeMicrosoftRedirectIfNeeded();
  renderAuthState();
  window.lucide?.createIcons();
}

function wireTabs() {
  document.querySelectorAll(".tab-button").forEach((button) => {
    button.addEventListener("click", () => {
      document.querySelectorAll(".tab-button").forEach((item) => item.classList.remove("is-active"));
      document.querySelectorAll(".tab-panel").forEach((item) => item.classList.remove("is-active"));
      button.classList.add("is-active");
      $(`#${button.dataset.tabTarget}`).classList.add("is-active");
    });
  });
}

function wireButtons() {
  elements.microsoftAuthBtn.addEventListener("click", () => runGuarded(toggleMicrosoft));
  elements.connectCta.addEventListener("click", () => runGuarded(connectMicrosoft));
  elements.invoiceFile.addEventListener("change", () => {
    addFilesToQueue(Array.from(elements.invoiceFile.files || []));
    elements.invoiceFile.value = "";
  });
  ["dragenter", "dragover"].forEach((eventName) => {
    elements.invoiceDropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      elements.invoiceDropzone.classList.add("is-dragging");
    });
  });
  ["dragleave", "drop"].forEach((eventName) => {
    elements.invoiceDropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      elements.invoiceDropzone.classList.remove("is-dragging");
    });
  });
  elements.invoiceDropzone.addEventListener("drop", (event) => {
    addFilesToQueue(Array.from(event.dataTransfer?.files || []));
  });
  elements.uploadList.addEventListener("click", (event) => {
    const button = event.target.closest(".upload-remove");
    if (!button) return;
    queuedFiles = queuedFiles.filter((item) => item.id !== button.dataset.fileId);
    renderUploadQueue();
  });
  elements.emailPreviewList.addEventListener("click", (event) => {
    const button = event.target.closest(".email-attachment-download");
    if (!button) return;
    event.preventDefault();
    runGuarded(() => downloadEmailAttachment(button.dataset.emailIndex, button.dataset.attachmentIndex));
  });
  elements.stampForm.addEventListener("submit", (event) => {
    event.preventDefault();
    runGuarded(processAndUploadInvoice);
  });
  elements.autofillBtn.addEventListener("click", () => runGuarded(autoFillFromQueuedFiles));
  elements.invoiceDownloadBtn.addEventListener("click", () => runGuarded(downloadCnetInvoices));
  elements.refreshFileTreeBtn.addEventListener("click", () => runGuarded(refreshFileTree));
  elements.fileTreeSearch.addEventListener("input", () => renderFileTree(currentFileTreeResult));
  elements.fileTree.addEventListener("click", (event) => {
    const fileButton = event.target.closest(".file-download");
    if (fileButton) {
      event.preventDefault();
      event.stopPropagation();
      runGuarded(() => downloadTreeFile(fileButton.dataset.fileId, fileButton.dataset.fileName));
      return;
    }
    const monthButton = event.target.closest(".tree-download");
    if (monthButton) {
      event.preventDefault();
      event.stopPropagation();
      runGuarded(() => downloadTreeMonth(monthButton.dataset.monthKey));
    }
  });
}

function addFilesToQueue(files) {
  for (const file of files) {
    queuedFiles.push({
      id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
      file,
    });
  }
  renderUploadQueue();
}

function renderUploadQueue() {
  if (!queuedFiles.length) {
    elements.dropzoneTitle.textContent = "Drop files here";
    elements.dropzoneCaption.textContent = "or click to select EML, PDFs, ZIPs, JPG, PNG, or WebP";
    elements.uploadList.innerHTML = "";
    renderEmailPreviews();
    return;
  }
  elements.dropzoneTitle.textContent = `${queuedFiles.length} file${queuedFiles.length === 1 ? "" : "s"} queued`;
  elements.dropzoneCaption.textContent = "You can drop more files or remove one with the X";
  elements.uploadList.innerHTML = queuedFiles.map(({ id, file }) => `
    <li class="upload-item">
      <span class="upload-item-text">
        <strong title="${escapeHtml(file.name)}">${escapeHtml(file.name)}</strong>
        <small>${formatFileSize(file.size)}</small>
      </span>
      <button class="upload-remove" type="button" data-file-id="${escapeHtml(id)}" aria-label="Remove ${escapeHtml(file.name)}">
        <i data-lucide="x"></i>
      </button>
    </li>
  `).join("");
  window.lucide?.createIcons();
  renderEmailPreviews();
}

async function renderEmailPreviews() {
  const renderId = ++emailPreviewRenderId;
  const emailFiles = queuedFiles
    .map((item) => item.file)
    .filter((file) => /\.eml$/i.test(file.name || ""));

  currentEmailPreviews = [];
  if (!emailFiles.length) {
    elements.emailPreviewPanel.hidden = true;
    elements.emailPreviewCount.textContent = "";
    elements.emailPreviewList.innerHTML = "";
    return;
  }

  elements.emailPreviewPanel.hidden = false;
  elements.emailPreviewCount.textContent = "Reading...";
  elements.emailPreviewList.innerHTML = "";

  const previews = await Promise.all(emailFiles.map(async (file) => {
    try {
      return await emailFileToPreview(file);
    } catch (error) {
      return {
        sourceName: file.name,
        error: error.message || String(error),
        email: { subject: "", from: "", date: "", text: "" },
        attachments: [],
      };
    }
  }));

  if (renderId !== emailPreviewRenderId) return;
  currentEmailPreviews = previews.filter(Boolean);
  elements.emailPreviewCount.textContent = `${currentEmailPreviews.length} email${currentEmailPreviews.length === 1 ? "" : "s"}`;
  elements.emailPreviewList.innerHTML = currentEmailPreviews.map(renderEmailPreview).join("");
  window.lucide?.createIcons();
}

function renderEmailPreview(preview, emailIndex) {
  const email = preview.email || {};
  const attachments = Array.isArray(preview.attachments) ? preview.attachments : [];
  const attachmentHtml = attachments.length
    ? attachments.map((attachment, attachmentIndex) => `
      <li>
        <span>
          <strong title="${escapeHtml(attachment.name)}">${escapeHtml(attachment.name)}</strong>
          <small>${escapeHtml(attachment.mimeType || "Attachment")} &middot; ${formatFileSize(attachment.size)}</small>
        </span>
        <button class="email-attachment-download" type="button" data-email-index="${emailIndex}" data-attachment-index="${attachmentIndex}" title="Download attachment" aria-label="Download ${escapeHtml(attachment.name)}">
          <i data-lucide="download"></i>
        </button>
      </li>
    `).join("")
    : "<li><span><strong>No attachments found</strong></span></li>";

  return `
    <article class="email-preview-card">
      <div class="email-preview-meta">
        <strong title="${escapeHtml(preview.sourceName)}">${escapeHtml(preview.sourceName)}</strong>
        ${preview.error ? `<small class="email-preview-error">${escapeHtml(preview.error)}</small>` : ""}
        <small>${escapeHtml(email.from || "Unknown sender")}${email.date ? ` &middot; ${escapeHtml(email.date)}` : ""}</small>
      </div>
      <dl class="email-preview-details">
        <dt>Subject</dt>
        <dd>${escapeHtml(email.subject || "(No subject)")}</dd>
        <dt>Text</dt>
        <dd class="email-preview-text">${escapeHtml(email.text || "(No text body)")}</dd>
      </dl>
      <ul class="email-attachment-list">
        ${attachmentHtml}
      </ul>
    </article>
  `;
}

function downloadEmailAttachment(emailIndex, attachmentIndex) {
  const preview = currentEmailPreviews[Number(emailIndex)];
  const attachment = preview?.attachments?.[Number(attachmentIndex)];
  if (!attachment) throw new Error("Attachment not found.");
  const blob = new Blob([attachment.bytes], { type: attachment.mimeType || "application/octet-stream" });
  triggerBrowserDownload(blob, attachment.name || "attachment");
}

async function autoFillFromQueuedFiles() {
  if (!queuedFiles.length) throw new Error("Select or drop at least one file first.");

  elements.autofillBtn.disabled = true;
  elements.autofillBtn.querySelector("span").textContent = "Reading";
  try {
    const documents = [];
    for (let index = 0; index < queuedFiles.length; index += 1) {
      const file = queuedFiles[index].file;
      showFlash(`Preparing AI ${index + 1} of ${queuedFiles.length}: ${file.name}`, "info");
      const payload = await fileToAnalysisPayload(file);
      if (payload.images.length) {
        documents.push(payload);
      }
    }

    if (!documents.length) {
      throw new Error("No supported PDFs or images were found for analysis.");
    }

    showFlash("Reading documents with Vision and GPT...", "info");
    const response = await fetch("/.netlify/functions/analyze-stamp-documents", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        documents,
        options: {
          companies: COMPANIES,
          banks: BANKS,
          categories: CATEGORIES,
        },
      }),
    });
    if (!response.ok) {
      throw new Error(await response.text());
    }

    const result = await response.json();
    applyAutoFillResult(result);
    showFlash(buildAutoFillMessage(result), result.lowConfidence?.length ? "warning" : "success");
  } finally {
    elements.autofillBtn.disabled = false;
    elements.autofillBtn.querySelector("span").textContent = "Auto-fill";
  }
}

function applyAutoFillResult(result) {
  setInputValue(elements.dateInput, normalizeDate(result.date));
  setSelectValue(elements.companyInput, result.company);
  setSelectValue(elements.categoryInput, result.category);
  setSelectValue(elements.bankInput, result.bank);
  setInputValue(elements.paymentCodeInput, result.payment_code);
  setInputValue(elements.clientInvoiceInput, result.client_invoice);
}

function setInputValue(input, value) {
  const nextValue = String(value || "").trim();
  if (nextValue) input.value = nextValue;
}

function setSelectValue(select, value) {
  const nextValue = String(value || "").trim();
  if (!nextValue) return;
  const option = Array.from(select.options).find((item) => item.value === nextValue);
  if (option) select.value = nextValue;
}

function normalizeDate(value) {
  const text = String(value || "").trim();
  if (!text) return "";
  const match = text.match(/\d{4}-\d{2}-\d{2}/);
  return match ? match[0] : "";
}

function buildAutoFillMessage(result) {
  const low = Array.isArray(result.lowConfidence) ? result.lowConfidence : [];
  if (!low.length) {
    return "Fields filled with AI. Review before processing.";
  }
  return `Fields filled with AI. Review before processing. Low confidence: ${low.join(", ")}.`;
}

function formatFileSize(bytes) {
  if (!Number.isFinite(bytes) || bytes <= 0) return "File";
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function populateSelect(select, options) {
  select.innerHTML = options.map((item) => `<option value="${escapeHtml(item)}">${escapeHtml(item)}</option>`).join("");
}

async function toggleMicrosoft() {
  if (state.graphToken) {
    disconnectMicrosoft();
  } else {
    await connectMicrosoft();
  }
}

async function connectMicrosoft() {
  clearFlash();
  const response = await fetch("/.netlify/functions/microsoft-login");
  if (!response.ok) {
    throw new Error(await response.text());
  }
  const payload = await response.json();
  localStorage.setItem("msalState", payload.state);
  window.location.href = payload.authUrl;
}

async function completeMicrosoftRedirectIfNeeded() {
  const params = new URLSearchParams(window.location.search);
  const code = params.get("code");
  if (!code) return;

  const stateParam = params.get("state");
  const expectedState = localStorage.getItem("msalState");
  if (!stateParam || !expectedState || stateParam !== expectedState) {
    throw new Error("Microsoft state does not match.");
  }

  const redirectUri = window.location.origin + "/";
  const response = await fetch(`/.netlify/functions/microsoft-callback?code=${encodeURIComponent(code)}&redirect_uri=${encodeURIComponent(redirectUri)}`);
  if (!response.ok) {
    throw new Error(await response.text());
  }
  const payload = await response.json();
  state.graphToken = payload.accessToken;
  sessionStorage.setItem("graphToken", state.graphToken);
  localStorage.removeItem("msalState");
  history.replaceState({}, document.title, window.location.pathname);
}

function disconnectMicrosoft() {
  state.graphToken = "";
  state.driveId = "";
  sessionStorage.removeItem("graphToken");
  renderAuthState();
  showFlash("Microsoft disconnected.", "info");
}

function renderAuthState() {
  const connected = Boolean(state.graphToken);
  elements.microsoftAuthBtn.querySelector("span").textContent = connected ? "Disconnect" : "Connect";
  elements.mainPanel.hidden = !connected;
  elements.connectPanel.hidden = connected;
  window.lucide?.createIcons();
}

async function processAndUploadInvoice() {
  clearFlash();
  elements.processBtn.disabled = true;
  elements.processBtn.querySelector("span").textContent = "Processing";
  elements.stampResult.hidden = true;

  try {
    if (!queuedFiles.length) throw new Error("Select or drop at least one invoice first.");

    const date = parseIsoDate(elements.dateInput.value);
    const paymentCode = elements.paymentCodeInput.value.trim();
    const clientInvoice = elements.clientInvoiceInput.value.trim();
    if (!paymentCode) throw new Error("Payment Code is required.");
    if (!clientInvoice) throw new Error("Client Invoice is required.");

    const filesToProcess = [...queuedFiles];
    const sourcePdfs = [];
    for (let index = 0; index < filesToProcess.length; index += 1) {
      const file = filesToProcess[index].file;
      showFlash(`Preparing file ${index + 1} of ${filesToProcess.length}: ${file.name}`, "info");
      sourcePdfs.push(await uploadedFileToPdfBytes(file));
    }

    showFlash("Combining files into one PDF...", "info");
    const combinedBytes = await mergePdfBytes(sourcePdfs);
    const stampedBytes = await stampPdfBytes(combinedBytes, {
      date: elements.dateInput.value,
      paymentCode,
      clientInvoice,
    });

    const fileName = buildStampedFileName({ paymentCode, clientInvoice, date });
    const { bankTargetPath, categoryTargetPath } = buildTargetPaths({
      company: elements.companyInput.value,
      bank: elements.bankInput.value,
      category: elements.categoryInput.value,
      date,
      fileName,
    });

    showFlash("Uploading combined PDF to SharePoint...", "info");
    await uploadSharePointFile(bankTargetPath, stampedBytes, "application/pdf");
    await uploadSharePointFile(categoryTargetPath, stampedBytes, "application/pdf");

    if (state.stampedBlobUrl) URL.revokeObjectURL(state.stampedBlobUrl);
    state.stampedBlobUrl = downloadBlobUrl(stampedBytes, "application/pdf");
    elements.downloadLink.href = state.stampedBlobUrl;
    elements.downloadLink.download = fileName;
    elements.bankPathOutput.textContent = bankTargetPath;
    elements.categoryPathOutput.textContent = categoryTargetPath;
    elements.stampResult.hidden = false;
    queuedFiles = [];
    renderUploadQueue();
    showFlash(`${filesToProcess.length} file${filesToProcess.length === 1 ? "" : "s"} combined into one PDF and uploaded successfully.`, "success");
  } finally {
    elements.processBtn.disabled = false;
    elements.processBtn.querySelector("span").textContent = "Process";
  }
}

function buildStampedFileName({ paymentCode, clientInvoice, date }) {
  const baseName = `P_${sanitizeFilename(paymentCode)}_I_${sanitizeFilename(clientInvoice)}_${yyyymmdd(date)}`;
  return `${baseName}.pdf`;
}

function buildTargetPaths({ company, bank, category, date, fileName }) {
  const companyRoot = COMPANY_SHAREPOINT_DIRS[company];
  const bankRel = BANK_TO_CHECKINGS_DIR[bank];
  const categoryRel = CATEGORY_TO_MONTH_REL_PATH[category];
  if (!companyRoot) throw new Error(`Company has no mapping: ${company}`);
  if (!bankRel) throw new Error(`Bank has no mapping: ${bank}`);
  if (!categoryRel) throw new Error(`Category has no mapping: ${category}`);

  const monthRoot = joinSharePointPath(companyRoot, String(date.getFullYear()), `${date.getMonth() + 1} ${monthName(date)}`);
  return {
    bankTargetPath: joinSharePointPath(monthRoot, "2 Bank Transactions", bankRel, "Direct Payments", fileName),
    categoryTargetPath: joinSharePointPath(monthRoot, ...categoryRel, fileName),
  };
}

async function downloadCnetInvoices() {
  elements.invoiceDownloadBtn.disabled = true;
  elements.invoiceDownloadBtn.querySelector("span").textContent = "Downloading";
  resetDownloaderUi();

  try {
    setDownloaderStatus("Downloading CNET CSV...", 0, 0);
    const exportResult = await postJson("/.netlify/functions/cnet-invoices-export", {});
    const rows = Array.isArray(exportResult.rows) ? exportResult.rows : [];
    const rowErrors = Array.isArray(exportResult.errors) ? exportResult.errors : [];
    const validRows = rows.filter((row) => !row.errors?.length);
    const errors = [...rowErrors, ...rows.flatMap((row) => (row.errors || []).map((message) => ({
      invoiceId: row.invoiceId || "-",
      message,
    })))];
    let completed = 0;
    let uploaded = 0;

    elements.downloadTotalCount.textContent = String(rows.length);
    elements.downloadSuccessCount.textContent = "0";
    renderDownloadErrors(errors);

    if (!rows.length) {
      throw new Error("The CNET CSV does not contain invoices.");
    }
    if (!validRows.length) {
      throw new Error("No CSV row has all required fields.");
    }

    setDownloaderStatus("Checking existing PDFs in SharePoint...", 0, validRows.length);
    const existingResult = await postJson("/.netlify/functions/cnet-invoices-existing", {
      accessToken: state.graphToken,
    });
    renderFileTree(existingResult);
    const existingIndex = existingResult.existing || {};
    const rowsToDownload = validRows.filter((row) => !hasSameStatusPdf(existingIndex, row));
    const skipped = validRows.length - rowsToDownload.length;
    const totalWork = rowsToDownload.length;
    elements.downloadSkippedCount.textContent = String(skipped);

    if (!rowsToDownload.length) {
      setDownloaderStatus(`Everything is up to date. Skipped ${skipped} invoices with unchanged status.`, 0, 0);
      showFlash(`Everything is up to date. ${skipped} invoice${skipped === 1 ? "" : "s"} skipped because they already exist with the same status.`, "success");
      return;
    }

    for (let index = 0; index < rowsToDownload.length; index += 1) {
      const row = rowsToDownload[index];
      setDownloaderStatus(`Processing invoice ${index + 1} of ${rowsToDownload.length}: ${row.invoiceId}`, completed, totalWork);
      const batchResult = await postJson("/.netlify/functions/cnet-invoices-download-batch", {
        accessToken: state.graphToken,
        rows: [row],
      });
      const results = Array.isArray(batchResult.results) ? batchResult.results : [];
      for (const item of results) {
        completed += 1;
        if (item.ok) {
          uploaded += 1;
        } else {
          errors.push({
            invoiceId: item.invoiceId || "-",
            message: item.error || "Unknown error.",
          });
        }
      }
      elements.downloadSuccessCount.textContent = String(uploaded);
      renderDownloadErrors(errors);
      setDownloaderStatus(`Uploaded ${uploaded} of ${totalWork} PDFs.`, completed, totalWork);
    }

    const message = errors.length
      ? `Invoice Downloader finished with ${uploaded} PDF${uploaded === 1 ? "" : "s"} uploaded and ${errors.length} error${errors.length === 1 ? "" : "s"}.`
      : `${uploaded} PDF${uploaded === 1 ? "" : "s"} uploaded successfully.`;
    showFlash(message, errors.length ? "warning" : "success");
    await refreshFileTree({ quiet: true });
  } finally {
    elements.invoiceDownloadBtn.disabled = false;
    elements.invoiceDownloadBtn.querySelector("span").textContent = "Download invoices";
  }
}

async function postJson(url, payload) {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  if (!response.ok) {
    throw new Error(await response.text());
  }
  return response.json();
}

function resetDownloaderUi() {
  elements.downloadTotalCount.textContent = "-";
  elements.downloadSuccessCount.textContent = "-";
  elements.downloadSkippedCount.textContent = "-";
  elements.downloadErrorList.innerHTML = "";
  elements.downloadErrors.hidden = true;
  setDownloaderStatus("Preparing Invoice Downloader...", 0, 0);
}

function setDownloaderStatus(message, completed, total) {
  const safeCompleted = Math.max(0, Number(completed) || 0);
  const safeTotal = Math.max(0, Number(total) || 0);
  const percent = safeTotal ? Math.min(100, Math.round((safeCompleted / safeTotal) * 100)) : 0;
  elements.downloadStatusText.textContent = message;
  elements.downloadCount.textContent = `${safeCompleted} / ${safeTotal}`;
  elements.downloadProgressFill.style.width = `${percent}%`;
}

function renderDownloadErrors(errors) {
  const visibleErrors = Array.isArray(errors) ? errors : [];
  elements.downloadErrors.hidden = !visibleErrors.length;
  elements.downloadErrorList.innerHTML = visibleErrors.map((error) => {
    const invoiceId = escapeHtml(error.invoiceId || "-");
    const message = escapeHtml(error.message || String(error));
    return `<li><strong>${invoiceId}</strong>: ${message}</li>`;
  }).join("");
}

async function refreshFileTree(options = {}) {
  elements.refreshFileTreeBtn.disabled = true;
  elements.refreshFileTreeBtn.querySelector("span").textContent = "Refreshing";
  try {
    if (!options.quiet) {
      setDownloaderStatus("Scanning SharePoint file tree...", 0, 0);
    }
    const result = await postJson("/.netlify/functions/cnet-invoices-existing", {
      accessToken: state.graphToken,
    });
    renderFileTree(result);
    if (!options.quiet) {
      showFlash(`File tree loaded: ${result.files || 0} PDF${Number(result.files) === 1 ? "" : "s"}.`, "success");
    }
  } finally {
    elements.refreshFileTreeBtn.disabled = false;
    elements.refreshFileTreeBtn.querySelector("span").textContent = "Refresh tree";
  }
}

function renderFileTree(result) {
  currentFileTreeResult = result || null;
  const totalFiles = Number(result?.files) || 0;
  const companies = filterFileTree(Array.isArray(result?.tree) ? result.tree : [], elements.fileTreeSearch.value);
  elements.fileTreeTotal.textContent = `${totalFiles} PDF${totalFiles === 1 ? "" : "s"} uploaded across ${companies.length} compan${companies.length === 1 ? "y" : "ies"}.`;
  elements.fileTree.innerHTML = companies.map(renderCompanyNode).join("");
  window.lucide?.createIcons();
}

function renderCompanyNode(company) {
  return `
    <details class="tree-node">
      <summary>
        <span class="tree-label">${escapeHtml(company.name)}</span>
        <span class="tree-count">${Number(company.count) || 0}</span>
      </summary>
      <div class="tree-children">
        ${(company.years || []).map(renderYearNode).join("")}
      </div>
    </details>
  `;
}

function renderYearNode(year) {
  return `
    <details class="tree-node">
      <summary>
        <span class="tree-label">${escapeHtml(year.name)}</span>
        <span class="tree-count">${Number(year.count) || 0}</span>
      </summary>
      <div class="tree-children">
        ${(year.months || []).map(renderMonthNode).join("")}
      </div>
    </details>
  `;
}

function renderMonthNode(month) {
  const monthKey = escapeHtml(month.key || `${month.name}-${Math.random().toString(16).slice(2)}`);
  return `
    <details class="tree-node">
      <summary>
        <span class="tree-label">${escapeHtml(month.name)}</span>
        <span class="tree-count">${Number(month.count) || 0}</span>
        <button class="tree-download" type="button" data-month-key="${monthKey}" title="Download month as ZIP" aria-label="Download ${escapeHtml(month.name)} as ZIP">
          <i data-lucide="download"></i>
        </button>
      </summary>
      <div class="tree-children">
        <ul class="file-list">
          ${(month.files || []).map((file) => `
            <li>
              <span>${escapeHtml(file.name)}</span>
              <button class="file-download" type="button" data-file-id="${escapeHtml(file.id)}" data-file-name="${escapeHtml(file.name)}" title="Download PDF" aria-label="Download ${escapeHtml(file.name)}">
                <i data-lucide="download"></i>
              </button>
            </li>
          `).join("")}
        </ul>
      </div>
    </details>
  `;
}

function filterFileTree(tree, query) {
  const needle = String(query || "").trim().toLowerCase();
  const keyedTree = addTreeKeys(tree);
  if (!needle) return keyedTree;

  return keyedTree.map((company) => {
    const companyMatch = company.name.toLowerCase().includes(needle);
    const years = (company.years || []).map((year) => {
      const yearMatch = year.name.toLowerCase().includes(needle);
      const months = (year.months || []).map((month) => {
        const monthMatch = month.name.toLowerCase().includes(needle);
        const files = (month.files || []).filter((file) => (
          companyMatch
          || yearMatch
          || monthMatch
          || file.name.toLowerCase().includes(needle)
        ));
        return files.length ? { ...month, files, count: files.length } : null;
      }).filter(Boolean);
      const count = months.reduce((total, month) => total + month.count, 0);
      return months.length ? { ...year, months, count } : null;
    }).filter(Boolean);
    const count = years.reduce((total, year) => total + year.count, 0);
    return years.length ? { ...company, years, count } : null;
  }).filter(Boolean);
}

function addTreeKeys(tree) {
  return tree.map((company) => ({
    ...company,
    key: company.name,
    years: (company.years || []).map((year) => ({
      ...year,
      key: `${company.name}/${year.name}`,
      months: (year.months || []).map((month) => ({
        ...month,
        key: `${company.name}/${year.name}/${month.name}`,
        files: month.files || [],
      })),
    })),
  }));
}

function findMonthByKey(key) {
  for (const company of addTreeKeys(currentFileTreeResult?.tree || [])) {
    for (const year of company.years || []) {
      for (const month of year.months || []) {
        if (month.key === key) return { company, year, month };
      }
    }
  }
  return null;
}

async function downloadTreeFile(fileId, fileName) {
  if (!fileId) throw new Error("Missing file id.");
  const blob = await downloadGraphFile(fileId);
  triggerBrowserDownload(blob, fileName || "invoice.pdf");
}

async function downloadTreeMonth(monthKey) {
  const found = findMonthByKey(monthKey);
  if (!found) throw new Error("Month folder not found.");
  const files = found.month.files || [];
  if (!files.length) throw new Error("Month folder has no files.");
  if (!window.JSZip) throw new Error("JSZip is not loaded.");

  showFlash(`Preparing ZIP for ${found.month.name} (${files.length} PDFs)...`, "info");
  const zip = new window.JSZip();
  for (let index = 0; index < files.length; index += 1) {
    const file = files[index];
    setDownloaderStatus(`Downloading ${index + 1} of ${files.length} files for ZIP...`, index, files.length);
    zip.file(file.name, await downloadGraphFile(file.id));
  }
  const blob = await zip.generateAsync({ type: "blob" });
  triggerBrowserDownload(blob, `${sanitizeZipName(found.company.name)}_${sanitizeZipName(found.year.name)}_${sanitizeZipName(found.month.name)}.zip`);
  setDownloaderStatus(`ZIP ready: ${files.length} PDFs.`, files.length, files.length);
  showFlash(`Month ZIP downloaded: ${files.length} PDFs.`, "success");
}

async function downloadGraphFile(fileId) {
  const driveId = currentFileTreeResult?.driveId || state.driveId;
  if (!driveId) throw new Error("Missing SharePoint drive id. Refresh the tree first.");
  const response = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${encodeURIComponent(fileId)}/content`, {
    headers: { Authorization: `Bearer ${state.graphToken}` },
  });
  if (!response.ok) throw new Error(await response.text());
  return response.blob();
}

function triggerBrowserDownload(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function sanitizeZipName(value) {
  return String(value || "folder").replace(/[^A-Za-z0-9_-]+/g, "_").replace(/^_+|_+$/g, "") || "folder";
}

function hasSameStatusPdf(existingIndex, row) {
  const invoiceId = String(row.invoiceId || "").trim();
  const status = normalizeStatusKey(row.paymentStatus);
  return Boolean(invoiceId && status && existingIndex[invoiceId]?.includes(status));
}

function normalizeStatusKey(value) {
  return String(value || "")
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[\\/:*?"<>|#%&{}~\x00-\x1f]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/[. ]+$/g, "");
}

async function runGuarded(work) {
  try {
    clearFlash();
    await work();
  } catch (error) {
    showFlash(error.message || String(error), "error");
  }
}
