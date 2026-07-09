import { loadConfig } from "./config.js";
import { initInvoiceReader, syncInvoiceReaderAuth } from "./invoice-reader/index.js";
import { state } from "./state.js";
import {
  BANKS,
  BANK_TO_CHECKINGS_DIR,
  CATEGORIES,
  CATEGORY_TO_MONTH_REL_PATH,
  COMPANIES,
  COMPANY_SHAREPOINT_DIRS,
  EXPENSE_CATEGORIES,
  NEW_STRUCTURE_CATEGORY_TO_MONTH_REL_PATH,
  NEW_STRUCTURE_FOLDER_ROOTS,
  NEW_STRUCTURE_STAMP_CATEGORIES,
  STAMP_BANK_TO_TEMPLATE_BANK,
  TEMPLATE_MONTHS,
} from "./data.js";
import { downloadSharePointFile, ensureSharePointFolderPath, listSharePointFileNames, listSharePointFolders, uploadSharePointFile } from "./graph.js";
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
  newBankField: $("#new-bank-field"),
  newBankInput: $("#new-bank-input"),
  accountTypeField: $("#account-type-field"),
  accountTypeInput: $("#account-type-input"),
  accountNumberField: $("#account-number-field"),
  accountNumberInput: $("#account-number-input"),
  accountNumberOptions: $("#account-number-options"),
  incomeAmountField: $("#income-amount-field"),
  incomeAmountInput: $("#income-amount-input"),
  incomePaymentTypeField: $("#income-payment-type-field"),
  incomePaymentTypeInput: $("#income-payment-type-input"),
  categoryInput: $("#category-input"),
  folderStructureLegacyInput: $("#folder-structure-legacy-input"),
  folderStructureInput: $("#folder-structure-input"),
  newFolderRootField: $("#new-folder-root-field"),
  newFolderRootInputs: Array.from(document.querySelectorAll("input[name='new-folder-root']")),
  newFolderRootNoneInput: $("#new-folder-root-none-input"),
  paymentCodeInput: $("#payment-code-input"),
  paymentCodeLabel: $("#payment-code-label"),
  clientInvoiceInput: $("#client-invoice-input"),
  clientInvoiceLabel: $("#client-invoice-label"),
  autofillBtn: $("#autofill-btn"),
  processBtn: $("#process-btn"),
  stampResult: $("#stamp-result"),
  bankPathLabel: $("#bank-path-label"),
  bankPathOutput: $("#bank-path-output"),
  categoryPathLabel: $("#category-path-label"),
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
  headerNav: $("#header-nav"),
  headerNavPrev: $("#header-nav-prev"),
  headerNavNext: $("#header-nav-next"),
  refreshFileTreeBtn: $("#refresh-file-tree-btn"),
  fileTreeSearch: $("#file-tree-search"),
  fileTreeTotal: $("#file-tree-total"),
  fileTree: $("#file-tree"),
  templateCompanyInput: $("#template-company-input"),
  templateYearInput: $("#template-year-input"),
  templateMonthInput: $("#template-month-input"),
  generateFolderTemplateBtn: $("#generate-folder-template-btn"),
  copyFolderTemplateBtn: $("#copy-folder-template-btn"),
  templateSourceOutput: $("#template-source-output"),
  templateCountOutput: $("#template-count-output"),
  templatePathOutput: $("#template-path-output"),
  folderTemplateTree: $("#folder-template-tree"),
  folderConfirmModal: $("#folder-confirm-modal"),
  folderConfirmExpected: $("#folder-confirm-expected"),
  folderConfirmExisting: $("#folder-confirm-existing"),
  folderConfirmMissing: $("#folder-confirm-missing"),
  folderConfirmCreateList: $("#folder-confirm-create-list"),
  folderConfirmCancel: $("#folder-confirm-cancel"),
  folderConfirmCreate: $("#folder-confirm-create"),
};

boot().catch((error) => showFlash(error.message, "error"));

let queuedFiles = [];
let currentFileTreeResult = null;
let currentEmailPreviews = [];
let emailPreviewRenderId = 0;
let excludedEmailAttachmentKeysByFileId = new Map();
let currentFolderTemplateText = "";
let templateBankAccounts = [];
let folderConfirmResolver = null;
let stampFolderOptionsRequestId = 0;
let discoveredCompanyFolders = false;
let accountNumberFolderOptions = [];
let autoFillInProgress = false;
const BANK_ACCOUNTS_WORKBOOK_PATH = "General/Cuentas Bancarias NO MOVER.xlsx";
const NOT_IDENTIFIED_FOLDER = "Not Identified";
const NEW_STRUCTURE_MIN_UPLOAD_DATE = new Date("2026-07-01T00:00:00");
const TEMPLATE_BANKS = ["Desjardin", "National", "Scotiabank"];
const TEMPLATE_ACCOUNT_TYPES = ["debit", "credit"];
const EXPENSE_FILENAME_BANK_CODES = {
  "scotia bank": "SB",
  scotiabank: "SB",
  "national bank": "NB",
  national: "NB",
  desjardins: "DJ",
  desjardin: "DJ",
};
const EXPENSE_FILENAME_PAYMENT_TYPE_CODES = {
  credit: "C",
  debit: "D",
};

async function boot() {
  wireTabs();
  wireHeaderNavScrolling();
  wireButtons();
  await loadConfig();
  populateSelect(elements.companyInput, COMPANIES);
  populateSelect(elements.bankInput, BANKS);
  populateSelectWithPreservedValue(elements.newBankInput, withCommonFolderOptions([], templateBankOptions()));
  populateSelectWithPreservedValue(elements.accountTypeInput, withCommonFolderOptions([], ["debit", "credit"]));
  populateAccountNumberOptions([NOT_IDENTIFIED_FOLDER]);
  populateSelect(elements.templateCompanyInput, COMPANIES);
  elements.dateInput.value = new Date().toISOString().slice(0, 10);
  restoreStampFolderStructurePreference();
  syncStampCategoryOptions();
  elements.templateYearInput.value = String(new Date().getFullYear());
  elements.templateMonthInput.value = TEMPLATE_MONTHS[new Date().getMonth()];
  renderFolderTemplate();
  await completeMicrosoftRedirectIfNeeded();
  renderAuthState();
  await initInvoiceReader();
  window.lucide?.createIcons();
  revealActiveHeaderNavButton("auto");
  setTimeout(() => revealActiveHeaderNavButton("auto"), 0);
}

function wireTabs() {
  document.querySelectorAll(".tab-button").forEach((button) => {
    button.addEventListener("click", () => {
      document.querySelectorAll(".tab-button").forEach((item) => item.classList.remove("is-active"));
      document.querySelectorAll(".tab-panel").forEach((item) => item.classList.remove("is-active"));
      button.classList.add("is-active");
      $(`#${button.dataset.tabTarget}`).classList.add("is-active");
      scrollHeaderNavButtonIntoView(button);
      syncInvoiceReaderAuth();
      if (button.dataset.tabTarget === "folder-template-tab") {
        runGuarded(loadFolderTemplateCsv);
      }
    });
  });
}

function wireHeaderNavScrolling() {
  elements.headerNavPrev.addEventListener("click", () => scrollHeaderNavByStep(-1));
  elements.headerNavNext.addEventListener("click", () => scrollHeaderNavByStep(1));
  elements.headerNav.addEventListener("scroll", updateHeaderNavScrollButtons, { passive: true });
  window.addEventListener("resize", updateHeaderNavScrollButtons);
  if (window.ResizeObserver) {
    new ResizeObserver(updateHeaderNavScrollButtons).observe(elements.headerNav);
  }
  updateHeaderNavScrollButtons();
  requestAnimationFrame(() => {
    revealActiveHeaderNavButton("auto");
  });
}

function scrollHeaderNavByStep(direction) {
  const nav = elements.headerNav;
  const maxScrollLeft = Math.max(0, nav.scrollWidth - nav.clientWidth);
  const push = Math.max(160, Math.floor(nav.clientWidth * 0.8));
  const targetLeft = Math.max(0, Math.min(nav.scrollLeft + direction * push, maxScrollLeft));
  nav.scrollTo({ left: targetLeft, behavior: "smooth" });
}

function revealActiveHeaderNavButton(behavior = "smooth") {
  syncHeaderNavActiveButton();
  const activeButton = document.querySelector(".tab-button.is-active");
  if (activeButton) {
    scrollHeaderNavButtonIntoView(activeButton, behavior);
  }
}

function syncHeaderNavActiveButton() {
  const activePanel = document.querySelector(".tab-panel.is-active");
  if (!activePanel) return;
  const activeButton = document.querySelector(`.tab-button[data-tab-target="${activePanel.id}"]`);
  if (!activeButton) return;
  document.querySelectorAll(".tab-button").forEach((item) => item.classList.toggle("is-active", item === activeButton));
}

function scrollHeaderNavButtonIntoView(button, behavior = "smooth") {
  const nav = elements.headerNav;
  const maxScrollLeft = Math.max(0, nav.scrollWidth - nav.clientWidth);
  const buttonLeft = button.offsetLeft - nav.offsetLeft;
  const centeredLeft = buttonLeft - (nav.clientWidth - button.offsetWidth) / 2;
  const targetLeft = Math.max(0, Math.min(centeredLeft, maxScrollLeft));
  nav.scrollTo({ left: targetLeft, behavior });
  updateHeaderNavScrollButtons();
}

function updateHeaderNavScrollButtons() {
  const nav = elements.headerNav;
  const hasOverflow = nav.scrollWidth > nav.clientWidth + 1;
  const shell = nav.parentElement;
  const atStart = nav.scrollLeft <= 1;
  const atEnd = nav.scrollLeft + nav.clientWidth >= nav.scrollWidth - 1;
  shell?.classList.toggle("has-overflow", hasOverflow);
  shell?.classList.toggle("is-at-start", hasOverflow && atStart);
  shell?.classList.toggle("is-at-end", hasOverflow && atEnd);
  elements.headerNavPrev.hidden = !hasOverflow;
  elements.headerNavNext.hidden = !hasOverflow;
  if (!hasOverflow) return;
  elements.headerNavPrev.disabled = atStart;
  elements.headerNavNext.disabled = atEnd;
}

function wireButtons() {
  elements.microsoftAuthBtn.addEventListener("click", () => runGuarded(toggleMicrosoft));
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
    excludedEmailAttachmentKeysByFileId.delete(button.dataset.fileId);
    queuedFiles = queuedFiles.filter((item) => item.id !== button.dataset.fileId);
    renderUploadQueue();
  });
  elements.emailPreviewList.addEventListener("click", (event) => {
    const removeButton = event.target.closest(".email-attachment-remove");
    if (removeButton) {
      event.preventDefault();
      removeEmailAttachment(removeButton.dataset.emailIndex, removeButton.dataset.attachmentKey);
      return;
    }

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
  elements.generateFolderTemplateBtn.addEventListener("click", () => runGuarded(generateFolderTemplate));
  elements.copyFolderTemplateBtn.addEventListener("click", () => runGuarded(copyFolderTemplate));
  [elements.folderStructureLegacyInput, elements.folderStructureInput].forEach((input) => input.addEventListener("change", () => {
    localStorage.setItem("stampFolderStructure", stampUsesNewFolderStructure() ? "new" : "legacy");
    syncStampCategoryOptions();
    runGuarded(refreshStampFolderControls);
  }));
  elements.newFolderRootInputs.forEach((input) => input.addEventListener("change", () => {
    syncStampCategoryOptions();
    runGuarded(refreshStampFolderControls);
  }));
  elements.companyInput.addEventListener("change", () => runGuarded(refreshStampFolderControls));
  elements.dateInput.addEventListener("change", () => runGuarded(refreshStampFolderControls));
  elements.categoryInput.addEventListener("change", () => runGuarded(refreshStampFolderControls));
  elements.newBankInput.addEventListener("change", () => runGuarded(refreshStampFolderControls));
  elements.accountTypeInput.addEventListener("change", () => runGuarded(refreshStampFolderControls));
  elements.accountNumberInput.addEventListener("focus", renderAccountNumberOptions);
  elements.accountNumberInput.addEventListener("input", () => {
    elements.accountNumberInput.value = lastFourDigits(elements.accountNumberInput.value);
    renderAccountNumberOptions();
  });
  elements.accountNumberField.addEventListener("focusout", () => {
    window.setTimeout(() => {
      if (!elements.accountNumberField.contains(document.activeElement)) hideAccountNumberOptions();
    }, 0);
  });
  elements.accountNumberOptions.addEventListener("mousedown", (event) => {
    event.preventDefault();
    const button = event.target.closest("[data-account-number]");
    if (!button) return;
    elements.accountNumberInput.value = lastFourDigits(button.dataset.accountNumber);
    hideAccountNumberOptions();
  });
  elements.folderConfirmCancel.addEventListener("click", () => closeFolderCreationDialog(false));
  elements.folderConfirmCreate.addEventListener("click", () => closeFolderCreationDialog(true));
  elements.folderConfirmModal.addEventListener("click", (event) => {
    if (event.target === elements.folderConfirmModal) closeFolderCreationDialog(false);
  });
  window.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && !elements.folderConfirmModal.hidden) closeFolderCreationDialog(false);
    if (event.key === "Escape") hideAccountNumberOptions();
  });
  document.addEventListener("pointerdown", closeAccountNumberOptionsOnOutsidePointer, true);
  elements.templateCompanyInput.addEventListener("change", renderFolderTemplate);
  elements.templateYearInput.addEventListener("change", renderFolderTemplate);
  elements.templateMonthInput.addEventListener("change", renderFolderTemplate);
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
  if (files.length) {
    resetStampFolderRoot();
  }
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
    elements.dropzoneCaption.textContent = "or click to select EML, MSG, PDFs, ZIPs, JPG, PNG, or WebP";
    elements.uploadList.innerHTML = "";
    syncAutoFillButtonState();
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
  syncAutoFillButtonState();
  renderEmailPreviews();
}

async function renderEmailPreviews() {
  const renderId = ++emailPreviewRenderId;
  const emailItems = queuedFiles
    .filter((item) => /\.(eml|msg)$/i.test(item.file.name || ""));

  currentEmailPreviews = [];
  if (!emailItems.length) {
    elements.emailPreviewPanel.hidden = true;
    elements.emailPreviewCount.textContent = "";
    elements.emailPreviewList.innerHTML = "";
    return;
  }

  elements.emailPreviewPanel.hidden = false;
  elements.emailPreviewCount.textContent = "Reading...";
  elements.emailPreviewList.innerHTML = "";

  const previews = await Promise.all(emailItems.map(async (item) => {
    try {
      return {
        ...await emailFileToPreview(item.file),
        fileId: item.id,
      };
    } catch (error) {
      return {
        fileId: item.id,
        sourceName: item.file.name,
        error: error.message || String(error),
        email: { subject: "", from: "", date: "", text: "" },
        attachments: [],
      };
    }
  }));

  if (renderId !== emailPreviewRenderId) return;
  currentEmailPreviews = previews.filter(Boolean);
  elements.emailPreviewCount.textContent = `${currentEmailPreviews.length} message${currentEmailPreviews.length === 1 ? "" : "s"}`;
  elements.emailPreviewList.innerHTML = currentEmailPreviews.map(renderEmailPreview).join("");
  window.lucide?.createIcons();
}

function renderEmailPreview(preview, emailIndex) {
  const email = preview.email || {};
  const excludedKeys = excludedEmailAttachmentKeysByFileId.get(preview.fileId) || new Set();
  const attachments = (Array.isArray(preview.attachments) ? preview.attachments : [])
    .filter((attachment) => !excludedKeys.has(emailAttachmentKey(attachment)));
  const attachmentHtml = attachments.length
    ? attachments.map((attachment, attachmentIndex) => `
      <li>
        ${renderAttachmentThumbnail(attachment)}
        <span>
          <strong title="${escapeHtml(attachment.name)}">${escapeHtml(attachment.name)}</strong>
          <small>${escapeHtml(attachment.mimeType || "Attachment")} &middot; ${formatFileSize(attachment.size)}</small>
        </span>
        <span class="email-attachment-actions">
          <button class="email-attachment-download" type="button" data-email-index="${emailIndex}" data-attachment-index="${attachmentIndex}" title="Download attachment" aria-label="Download ${escapeHtml(attachment.name)}">
            <i data-lucide="download"></i>
          </button>
          <button class="email-attachment-remove" type="button" data-email-index="${emailIndex}" data-attachment-key="${escapeHtml(emailAttachmentKey(attachment))}" title="Remove from upload" aria-label="Remove ${escapeHtml(attachment.name)} from upload">
            <i data-lucide="x"></i>
          </button>
        </span>
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

function renderAttachmentThumbnail(attachment) {
  if (attachment.thumbnailDataUrl) {
    return `
      <span class="email-attachment-thumb">
        <img src="${escapeHtml(attachment.thumbnailDataUrl)}" alt="">
      </span>
    `;
  }
  return `
    <span class="email-attachment-thumb email-attachment-thumb-empty">
      <i data-lucide="file"></i>
    </span>
  `;
}

function downloadEmailAttachment(emailIndex, attachmentIndex) {
  const preview = currentEmailPreviews[Number(emailIndex)];
  const excludedKeys = excludedEmailAttachmentKeysByFileId.get(preview?.fileId) || new Set();
  const attachments = (preview?.attachments || []).filter((attachment) => !excludedKeys.has(emailAttachmentKey(attachment)));
  const attachment = attachments[Number(attachmentIndex)];
  if (!attachment) throw new Error("Attachment not found.");
  const blob = new Blob([attachment.bytes], { type: attachment.mimeType || "application/octet-stream" });
  triggerBrowserDownload(blob, attachment.name || "attachment");
}

function removeEmailAttachment(emailIndex, attachmentKey) {
  const preview = currentEmailPreviews[Number(emailIndex)];
  if (!preview?.fileId || !attachmentKey) return;
  const excludedKeys = excludedEmailAttachmentKeysByFileId.get(preview.fileId) || new Set();
  excludedKeys.add(attachmentKey);
  excludedEmailAttachmentKeysByFileId.set(preview.fileId, excludedKeys);
  elements.emailPreviewList.innerHTML = currentEmailPreviews.map(renderEmailPreview).join("");
  window.lucide?.createIcons();
}

function emailAttachmentKey(attachment) {
  return [
    String(attachment.name || "").trim().toLowerCase(),
    String(attachment.mimeType || "").trim().toLowerCase(),
    String(attachment.size || attachment.bytes?.byteLength || 0),
  ].join("|");
}

async function autoFillFromQueuedFiles() {
  if (!queuedFiles.length) throw new Error("Select or drop at least one file first.");

  autoFillInProgress = true;
  elements.autofillBtn.disabled = true;
  elements.autofillBtn.querySelector("span").textContent = "Reading";
  try {
    const folderRoot = currentStampFolderRoot();
    const documents = [];
    for (let index = 0; index < queuedFiles.length; index += 1) {
      const item = queuedFiles[index];
      const file = item.file;
      showFlash(`Preparing AI ${index + 1} of ${queuedFiles.length}: ${file.name}`, "info");
      const payload = await fileToAnalysisPayload(file, emailProcessingOptions(item));
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
          categories: currentStampCategories(),
          folderStructure: stampUsesNewFolderStructure() ? "new" : "legacy",
          folderRoot,
          newStructureCategoryPaths: NEW_STRUCTURE_CATEGORY_TO_MONTH_REL_PATH,
        },
      }),
    });
    if (!response.ok) {
      throw new Error(await response.text());
    }

    const result = await response.json();
    applyAutoFillResult(result, folderRoot);
    showFlash(buildAutoFillMessage(result), result.lowConfidence?.length ? "warning" : "success");
  } finally {
    autoFillInProgress = false;
    syncAutoFillButtonState();
    elements.autofillBtn.querySelector("span").textContent = "Auto-fill";
  }
}

function applyAutoFillResult(result, folderRoot = currentStampFolderRoot()) {
  if (folderRoot === "Documents") {
    applyDocumentsAutoFillResult(result);
  } else if (folderRoot === "Expenses" || !stampUsesNewFolderStructure()) {
    applyExpensesAutoFillResult(result);
  } else if (folderRoot === "Income") {
    applyIncomeAutoFillResult(result);
  } else if (folderRoot === "Statements") {
    applyStatementsAutoFillResult(result);
  }
  runGuarded(refreshStampFolderControls);
}

function applyDocumentsAutoFillResult(result) {
  setInputValue(elements.dateInput, normalizeDate(result.date));
  setSelectValue(elements.companyInput, result.company);
  setNewStructureRootFromCategory(result.category);
  syncStampCategoryOptions();
  setSelectValue(elements.categoryInput, result.category);
  setInputValue(elements.paymentCodeInput, result.suggested_title);
  setInputValue(elements.clientInvoiceInput, result.description);
}

function applyExpensesAutoFillResult(result) {
  setInputValue(elements.dateInput, normalizeDate(result.date));
  setSelectValue(elements.companyInput, result.company);
  setSelectValue(elements.categoryInput, result.category);
  setSelectValue(elements.bankInput, result.bank);
  applyExpenseFinancialFields(result);
  setInputValue(elements.paymentCodeInput, result.payment_code);
  setInputValue(elements.clientInvoiceInput, result.client_invoice);
}

function applyIncomeAutoFillResult(result) {
  setInputValue(elements.dateInput, normalizeDate(result.date));
  setSelectValue(elements.companyInput, result.company);
  setNewStructureRootFromCategory(result.category);
  syncStampCategoryOptions();
  setSelectValue(elements.categoryInput, result.category);
  clearInputValue(elements.incomeAmountInput);
  setInputValue(elements.incomeAmountInput, result.amount);
  clearInputValue(elements.incomePaymentTypeInput);
  setSelectValue(elements.incomePaymentTypeInput, result.payment_type);
}

function applyStatementsAutoFillResult(result) {
  setInputValue(elements.dateInput, normalizeDate(result.date));
  setSelectValue(elements.companyInput, result.company);
  setInputValue(elements.paymentCodeInput, result.suggested_title);
}

function applyExpenseFinancialFields(result) {
  if (!currentStampRootIsExpenses()) {
    clearInputValue(elements.newBankInput);
    clearInputValue(elements.accountTypeInput);
    clearInputValue(elements.accountNumberInput);
    return;
  }

  const confidence = result.confidence || {};
  clearInputValue(elements.newBankInput);
  clearInputValue(elements.accountTypeInput);
  clearInputValue(elements.accountNumberInput);

  if (confidence.bank >= 0.65) {
    setInputValue(elements.newBankInput, stampBankToTemplateBank(result.bank));
  }

  if (confidence.account_type >= 0.65) {
    setSelectValue(elements.accountTypeInput, result.account_type);
  }

  setInputValue(elements.accountNumberInput, lastFourDigits(result.account_last_four));
}

function setInputValue(input, value) {
  const nextValue = String(value || "").trim();
  if (nextValue) input.value = nextValue;
}

function clearInputValue(input) {
  input.value = "";
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
  select.innerHTML = options.map((item) => {
    const value = String(item ?? "");
    const label = value || NOT_IDENTIFIED_FOLDER;
    return `<option value="${escapeHtml(value)}">${escapeHtml(label)}</option>`;
  }).join("");
}

function populateSelectWithPreservedValue(select, options) {
  const previousValue = select.value;
  const uniqueOptions = Array.from(new Set(options));
  populateSelect(select, uniqueOptions);
  if (uniqueOptions.includes(previousValue)) {
    select.value = previousValue;
  }
}

function populateAccountNumberOptions(options) {
  accountNumberFolderOptions = Array.from(new Set(options.map(lastFourDigits).filter((value) => /^\d{4}$/.test(value))));
  renderAccountNumberOptions();
}

function renderAccountNumberOptions() {
  const input = elements.accountNumberInput;
  const query = lastFourDigits(input.value);
  const visibleOptions = accountNumberFolderOptions.filter((option) => !query || option.includes(query));
  elements.accountNumberOptions.innerHTML = visibleOptions
    .map((option) => `
      <li>
        <button type="button" data-account-number="${escapeHtml(option)}">${escapeHtml(option)}</button>
      </li>
    `)
    .join("");
  elements.accountNumberOptions.hidden = document.activeElement !== input || !visibleOptions.length;
}

function hideAccountNumberOptions() {
  elements.accountNumberOptions.hidden = true;
}

function lastFourDigits(value) {
  return String(value || "").replace(/\D/g, "").slice(0, 4);
}

function normalizeIdentifier(value) {
  return String(value || "").trim().toLowerCase().replace(/[_-]+/g, " ").replace(/\s+/g, " ");
}

function closeAccountNumberOptionsOnOutsidePointer(event) {
  if (!elements.accountNumberField.contains(event.target)) {
    hideAccountNumberOptions();
  }
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
  syncInvoiceReaderAuth();
  showFlash("Microsoft disconnected.", "info");
}

function renderAuthState() {
  const connected = Boolean(state.graphToken);
  elements.microsoftAuthBtn.querySelector("span").textContent = connected ? "Disconnect" : "Connect";
  elements.mainPanel.hidden = !connected;
  elements.connectPanel.hidden = connected;
  syncInvoiceReaderAuth();
  if (connected) {
    void discoverSharePointCompanyFolders().then(refreshStampFolderControls).catch(() => {});
  }
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
    assertAllowedUploadDate(date);
    const category = currentStampEffectiveCategory();
    if (!category) throw new Error("Select a folder and category first.");
    const isDocuments = currentStampRootIsDocuments();
    const isExpenses = currentStampRootIsExpenses();
    const isIncome = currentStampRootIsIncome();
    const isStatements = currentStampRootIsStatements();
    const paymentCode = elements.paymentCodeInput.value.trim();
    const clientInvoice = elements.clientInvoiceInput.value.trim();
    const incomeAmount = elements.incomeAmountInput.value.trim();
    const incomePaymentType = elements.incomePaymentTypeInput.value.trim();
    const suggestedTitle = paymentCode;
    const description = clientInvoice;
    if (isDocuments) {
      if (!suggestedTitle) throw new Error("Suggested title is required.");
      if (!description) throw new Error("Description is required.");
    } else if (isStatements) {
      if (!suggestedTitle) throw new Error("Suggested file name is required.");
    } else if (isExpenses) {
      if (!paymentCode) throw new Error("Payment Code is required.");
      if (!clientInvoice) throw new Error("Client Invoice is required.");
    } else if (isIncome) {
      if (!incomeAmount) throw new Error("Monto is required.");
      if (!incomePaymentType) throw new Error("Payment type is required.");
    }

    const filesToProcess = [...queuedFiles];
    const sourcePdfs = [];
    for (let index = 0; index < filesToProcess.length; index += 1) {
      const item = filesToProcess[index];
      const file = item.file;
      showFlash(`Preparing file ${index + 1} of ${filesToProcess.length}: ${file.name}`, "info");
      sourcePdfs.push(await uploadedFileToPdfBytes(file, emailProcessingOptions(item)));
    }

    showFlash("Combining files into one PDF...", "info");
    const combinedBytes = await mergePdfBytes(sourcePdfs);
    const stampedBytes = await stampPdfBytes(combinedBytes, {
      date: elements.dateInput.value,
      paymentCode,
      clientInvoice,
      stampType: isExpenses ? "paid" : "processed",
      title: isDocuments ? suggestedTitle : "",
      description: isDocuments ? description : "",
    });

    const initialFileName = isDocuments
      ? buildDocumentsFileName({ suggestedTitle })
      : isStatements
        ? buildDocumentsFileName({ suggestedTitle })
        : isExpenses
          ? buildStampedFileName({
              paymentCode,
              clientInvoice,
              date,
              bank: currentStampBankValue(),
              accountType: elements.accountTypeInput.value,
              accountNumber: elements.accountNumberInput.value,
            })
          : buildGeneralStampedFileName({ category, date });
    const useNewFolderStructure = stampUsesNewFolderStructure();
    if (!useNewFolderStructure) {
      throw new Error("Uploads must use the new folder structure.");
    }
    const targetPaths = buildTargetPaths({
      company: elements.companyInput.value,
      bank: currentStampBankValue(),
      accountType: elements.accountTypeInput.value,
      accountNumber: elements.accountNumberInput.value,
      incomePaymentType,
      category,
      date,
      fileName: initialFileName,
      useNewFolderStructure,
    });
    const fileName = isStatements
      ? await uniqueSharePointFileName(targetPaths.folderPaths[0], initialFileName)
      : initialFileName;
    if (fileName !== initialFileName) {
      targetPaths.targetPaths = targetPaths.targetPaths.map((targetPath) => replaceSharePointFileName(targetPath, fileName));
      targetPaths.uploadTargetPath = targetPaths.targetPaths[0];
    }

    // Accounting folders may be created only after the user sees exactly what is missing and approves it.
    // Never create SharePoint accounting folders silently from the stamping flow.
    for (const folderPath of targetPaths.folderPaths) {
      await ensureSharePointFolderPath(folderPath, { confirmCreate: confirmSharePointFolderCreation });
    }

    showFlash("Uploading combined PDF to SharePoint...", "info");
    for (const targetPath of targetPaths.targetPaths) {
      await uploadSharePointFile(targetPath, stampedBytes, "application/pdf");
    }

    if (state.stampedBlobUrl) URL.revokeObjectURL(state.stampedBlobUrl);
    state.stampedBlobUrl = downloadBlobUrl(stampedBytes, "application/pdf");
    elements.downloadLink.href = state.stampedBlobUrl;
    elements.downloadLink.download = fileName;
    renderStampResultPaths(targetPaths);
    elements.stampResult.hidden = false;
    queuedFiles = [];
    excludedEmailAttachmentKeysByFileId = new Map();
    renderUploadQueue();
    showFlash(`${filesToProcess.length} file${filesToProcess.length === 1 ? "" : "s"} combined into one PDF and uploaded successfully.`, "success");
  } finally {
    elements.processBtn.disabled = false;
    elements.processBtn.querySelector("span").textContent = "Process";
  }
}

function emailProcessingOptions(item) {
  return {
    excludedAttachmentKeys: excludedEmailAttachmentKeysByFileId.get(item.id) || new Set(),
  };
}

function assertAllowedUploadDate(date) {
  if (date < NEW_STRUCTURE_MIN_UPLOAD_DATE) {
    throw new Error("Uploads are blocked for dates before July 1, 2026.");
  }
}

function buildStampedFileName({ paymentCode, clientInvoice, date, bank, accountType, accountNumber }) {
  const baseName = `${expenseFileNamePrefix({ bank, accountType, accountNumber })}_P_${sanitizeFilename(paymentCode)}_I_${sanitizeFilename(clientInvoice)}_${yyyymmdd(date)}`;
  return `${baseName}.pdf`;
}

function expenseFileNamePrefix({ bank, accountType, accountNumber }) {
  const bankCode = EXPENSE_FILENAME_BANK_CODES[normalizeIdentifier(bank)];
  const paymentTypeCode = EXPENSE_FILENAME_PAYMENT_TYPE_CODES[normalizeIdentifier(accountType)];
  if (!bankCode) throw new Error("Select a valid bank for the expense file name prefix.");
  if (!paymentTypeCode) throw new Error("Select Credit or debit as the payment type for the expense file name prefix.");
  return `${bankCode}${paymentTypeCode}${lastFourDigits(accountNumber) || "0000"}`;
}

function buildDocumentsFileName({ suggestedTitle }) {
  return `${sanitizeFilename(suggestedTitle, "document")}.pdf`;
}

function buildGeneralStampedFileName({ category, date }) {
  const root = currentStampFolderRoot() || "document";
  const label = stampCategoryLabel(category) || category || root;
  return `${sanitizeFilename(root)}_${sanitizeFilename(label)}_${yyyymmdd(date)}.pdf`;
}

async function uniqueSharePointFileName(folderPath, fileName) {
  const existingNames = new Set((await listSharePointFileNames(folderPath)).map((name) => name.toLowerCase()));
  const extension = fileName.match(/(\.[^.]+)$/)?.[1] || "";
  const baseName = extension ? fileName.slice(0, -extension.length) : fileName;
  let candidate = fileName;
  let counter = 2;
  while (existingNames.has(candidate.toLowerCase())) {
    candidate = `${baseName} (${counter})${extension}`;
    counter += 1;
  }
  return candidate;
}

function replaceSharePointFileName(path, fileName) {
  const parts = String(path || "").split("/");
  parts[parts.length - 1] = fileName;
  return parts.join("/");
}

function renderStampResultPaths(targetPaths) {
  if (targetPaths.mode === "new") {
    elements.bankPathLabel.textContent = "Upload path";
    elements.bankPathOutput.textContent = targetPaths.uploadTargetPath;
    elements.categoryPathLabel.hidden = true;
    elements.categoryPathOutput.hidden = true;
    elements.categoryPathOutput.textContent = "";
    return;
  }

  elements.bankPathLabel.textContent = "Bank folder";
  elements.bankPathOutput.textContent = targetPaths.bankTargetPath;
  elements.categoryPathLabel.textContent = "Category folder";
  elements.categoryPathOutput.textContent = targetPaths.categoryTargetPath;
  elements.categoryPathLabel.hidden = false;
  elements.categoryPathOutput.hidden = false;
}

function confirmSharePointFolderCreation({ fullPath, missingPath, missingParts, lastExistingPath }) {
  return showFolderCreationDialog({ fullPath, missingPath, missingParts, lastExistingPath });
}

function showFolderCreationDialog({ fullPath, missingPath, missingParts, lastExistingPath }) {
  return new Promise((resolve) => {
    folderConfirmResolver = resolve;
    elements.folderConfirmExpected.textContent = fullPath;
    elements.folderConfirmExisting.textContent = lastExistingPath;
    elements.folderConfirmMissing.textContent = missingPath;
    elements.folderConfirmCreateList.innerHTML = missingParts
      .map((part) => `<li>${escapeHtml(part)}</li>`)
      .join("");
    elements.folderConfirmModal.hidden = false;
    elements.folderConfirmModal.classList.remove("is-closing");
    window.lucide?.createIcons();
    elements.folderConfirmCreate.focus();
  });
}

function closeFolderCreationDialog(confirmed) {
  if (!folderConfirmResolver) return;
  const resolve = folderConfirmResolver;
  folderConfirmResolver = null;
  elements.folderConfirmModal.classList.add("is-closing");
  window.setTimeout(() => {
    elements.folderConfirmModal.hidden = true;
    elements.folderConfirmModal.classList.remove("is-closing");
    resolve(confirmed);
  }, 140);
}

function stampUsesNewFolderStructure() {
  return elements.folderStructureInput.checked;
}

function currentStampCategories() {
  if (!stampUsesNewFolderStructure()) return CATEGORIES;
  const root = currentStampFolderRoot();
  if (!root) return [];
  return NEW_STRUCTURE_STAMP_CATEGORIES.filter((category) => newStructureRootForCategory(category) === root);
}

function restoreStampFolderStructurePreference() {
  const stored = localStorage.getItem("stampFolderStructure");
  elements.folderStructureInput.checked = stored !== "legacy";
  elements.folderStructureLegacyInput.checked = stored === "legacy";
}

function syncStampCategoryOptions() {
  const previousValue = elements.categoryInput.value;
  const categories = currentStampCategories();
  populateStampCategorySelect(categories);
  if (currentStampRootIsStatements() && categories.includes("Statements")) {
    elements.categoryInput.value = "Statements";
  } else if (categories.includes(previousValue)) {
    elements.categoryInput.value = previousValue;
  }
  syncStampFolderControlVisibility();
  syncStampMetadataFields();
}

function populateStampCategorySelect(categories) {
  const options = categories
    .map((category) => `<option value="${escapeHtml(category)}">${escapeHtml(stampCategoryLabel(category))}</option>`);
  if (stampUsesNewFolderStructure()) {
    const placeholder = currentStampFolderRoot() ? "Select category" : "Select a folder first";
    options.unshift(`<option value="">${placeholder}</option>`);
  }
  elements.categoryInput.innerHTML = options.join("");
}

function stampCategoryLabel(category) {
  if (!stampUsesNewFolderStructure()) return category;
  const relPath = NEW_STRUCTURE_CATEGORY_TO_MONTH_REL_PATH[category] || [];
  return relPath.slice(1).join(" - ") || category;
}

function currentStampBankValue() {
  return stampUsesNewFolderStructure() ? elements.newBankInput.value : elements.bankInput.value;
}

function currentStampFolderRoot() {
  return elements.newFolderRootInputs.find((input) => input.checked)?.value || "";
}

function currentStampEffectiveCategory() {
  if (currentStampRootIsStatements()) return "Statements";
  return elements.categoryInput.value;
}

function resetStampFolderRoot() {
  elements.newFolderRootNoneInput.checked = true;
  syncStampCategoryOptions();
}

function setNewStructureRootFromCategory(category) {
  const root = newStructureRootForCategory(category);
  if (root && NEW_STRUCTURE_FOLDER_ROOTS.includes(root)) {
    elements.newFolderRootInputs.forEach((input) => {
      input.checked = input.value === root;
    });
  }
}

function syncStampFolderControlVisibility() {
  const newMode = stampUsesNewFolderStructure();
  const category = elements.categoryInput.value;
  const root = currentStampFolderRoot();
  const hasNewRoot = !newMode || Boolean(root);
  const isDocuments = newMode && root === "Documents";
  const isExpenses = newMode && root === "Expenses";
  const isIncome = newMode && root === "Income";
  const isStatements = newMode && root === "Statements";
  const needsBank = newMode && newStructureCategoryNeedsBank(category);
  const needsAccount = newMode && newStructureCategoryNeedsAccount(category);

  elements.stampForm.classList.toggle("is-expenses-root", isExpenses);
  elements.stampForm.classList.toggle("is-income-root", isIncome);
  elements.dateInput.closest(".field").hidden = !hasNewRoot;
  elements.dateInput.required = hasNewRoot;
  elements.companyInput.closest(".field").hidden = !hasNewRoot;
  elements.companyInput.required = hasNewRoot;
  elements.categoryInput.closest(".field").hidden = !hasNewRoot || isStatements;
  elements.categoryInput.required = hasNewRoot && !isStatements;
  elements.bankInput.closest(".field").hidden = newMode;
  elements.bankInput.required = !newMode;
  elements.newFolderRootField.hidden = !newMode;
  elements.newFolderRootInputs.forEach((input) => {
    input.required = newMode;
  });
  elements.newBankField.hidden = !isExpenses || !needsBank;
  elements.accountTypeField.hidden = !isExpenses || !needsAccount;
  elements.accountNumberField.hidden = !isExpenses || !needsAccount;
  elements.incomeAmountField.hidden = !isIncome;
  elements.incomePaymentTypeField.hidden = !isIncome;
  syncAutoFillButtonState();
  syncStampMetadataFields();
}

function syncAutoFillButtonState() {
  const hasReadableFiles = queuedFiles.length > 0;
  const hasRequiredFolder = !stampUsesNewFolderStructure() || Boolean(currentStampFolderRoot());
  elements.autofillBtn.disabled = autoFillInProgress || !hasReadableFiles || !hasRequiredFolder;
}

function syncStampMetadataFields() {
  const isDocuments = currentStampRootIsDocuments();
  const isExpenses = currentStampRootIsExpenses();
  const isIncome = currentStampRootIsIncome();
  const isStatements = currentStampRootIsStatements();
  const showTitleField = isDocuments || isExpenses || isStatements;
  const showDescriptionField = isDocuments || isExpenses;
  elements.stampForm.classList.toggle("is-documents-mode", isDocuments);
  elements.paymentCodeInput.closest(".field").hidden = !showTitleField;
  elements.clientInvoiceInput.closest(".field").hidden = !showDescriptionField;
  elements.paymentCodeInput.required = showTitleField;
  elements.clientInvoiceInput.required = showDescriptionField;
  elements.incomeAmountInput.required = isIncome;
  elements.incomePaymentTypeInput.required = isIncome;
  elements.paymentCodeLabel.textContent = isDocuments || isStatements ? "Suggested file name" : "Payment Code";
  elements.clientInvoiceLabel.textContent = isDocuments ? "Description" : "Client Invoice";
  elements.paymentCodeInput.placeholder = isDocuments || isStatements ? "Final PDF title" : "";
  elements.clientInvoiceInput.placeholder = isDocuments ? "PDF subject metadata" : "";
}

function currentStampRootIsDocuments() {
  return stampUsesNewFolderStructure() && currentStampFolderRoot() === "Documents";
}

function currentStampRootIsExpenses() {
  return !stampUsesNewFolderStructure() || currentStampFolderRoot() === "Expenses";
}

function currentStampRootIsIncome() {
  return stampUsesNewFolderStructure() && currentStampFolderRoot() === "Income";
}

function currentStampRootIsStatements() {
  return stampUsesNewFolderStructure() && currentStampFolderRoot() === "Statements";
}

async function refreshStampFolderControls() {
  syncStampFolderControlVisibility();
  if (!stampUsesNewFolderStructure() || !state.graphToken) return;

  const requestId = ++stampFolderOptionsRequestId;
  const category = currentStampEffectiveCategory();
  if (!category) return;
  const basePath = newStructureCategoryBasePath({
    company: elements.companyInput.value,
    category,
    date: parseIsoDate(elements.dateInput.value),
  });
  if (!basePath) return;

  if (newStructureCategoryNeedsBank(category)) {
    const bankOptions = await listSharePointFolders(basePath);
    if (requestId !== stampFolderOptionsRequestId) return;
    populateSelectWithPreservedValue(elements.newBankInput, withCommonFolderOptions(bankOptions, templateBankOptions()));
  }

  if (!newStructureCategoryNeedsAccount(category)) return;

  const bank = normalizeFolderLevel(elements.newBankInput.value);
  if (bank && bank !== NOT_IDENTIFIED_FOLDER) {
    const typeOptions = await listSharePointFolders(joinSharePointPath(basePath, bank));
    if (requestId !== stampFolderOptionsRequestId) return;
    populateSelectWithPreservedValue(elements.accountTypeInput, withCommonFolderOptions(typeOptions, ["debit", "credit"]));
  } else {
    populateSelectWithPreservedValue(elements.accountTypeInput, withCommonFolderOptions([], ["debit", "credit"]));
  }

  const accountType = normalizeFolderLevel(elements.accountTypeInput.value);
  if (bank && bank !== NOT_IDENTIFIED_FOLDER && accountType && accountType !== NOT_IDENTIFIED_FOLDER) {
    const accountOptions = await listSharePointFolders(joinSharePointPath(basePath, bank, accountType));
    if (requestId !== stampFolderOptionsRequestId) return;
    populateAccountNumberOptions(withCommonFolderOptions(accountOptions, []));
  } else {
    populateAccountNumberOptions([NOT_IDENTIFIED_FOLDER]);
  }
}

function withCommonFolderOptions(existingOptions, commonOptions) {
  return ["", ...existingOptions, ...commonOptions];
}

function templateBankOptions() {
  const workbookBanks = templateBankAccounts.map((account) => account.bank).filter(Boolean);
  return workbookBanks.length ? Array.from(new Set(workbookBanks)) : TEMPLATE_BANKS;
}

async function discoverSharePointCompanyFolders() {
  if (discoveredCompanyFolders || !state.graphToken) return;
  discoveredCompanyFolders = true;

  const generalFolders = await listSharePointFolders("General");
  const dtechFolder = generalFolders.find((name) => /d[\s-]*t[\s-]*e?[\s-]*c[\s-]*h/i.test(name));
  if (!dtechFolder) return;

  const dtechRoot = await resolveCompanyAccountantRoot(dtechFolder);
  addCompanyOption(dtechFolder, dtechRoot);
}

async function resolveCompanyAccountantRoot(companyFolderName) {
  const companyPath = joinSharePointPath("General", companyFolderName);
  const children = await listSharePointFolders(companyPath);
  const accountantFolder = children.find((name) => /documents?\s*account|accountant/i.test(name));
  return accountantFolder ? joinSharePointPath(companyPath, accountantFolder) : companyPath;
}

function addCompanyOption(company, companyRoot) {
  if (!company || !companyRoot) return;
  COMPANY_SHAREPOINT_DIRS[company] = companyRoot;
  if (!COMPANIES.includes(company)) {
    COMPANIES.push(company);
  }
  repopulateCompanySelect(elements.companyInput);
  repopulateCompanySelect(elements.templateCompanyInput);
}

function repopulateCompanySelect(select) {
  const previousValue = select.value;
  populateSelect(select, COMPANIES);
  if (COMPANIES.includes(previousValue)) {
    select.value = previousValue;
  }
}

function buildTargetPaths({ company, bank, accountType, accountNumber, incomePaymentType, category, date, fileName, useNewFolderStructure }) {
  const companyRoot = COMPANY_SHAREPOINT_DIRS[company];
  if (!companyRoot) throw new Error(`Company has no mapping: ${company}`);

  const monthRoot = joinSharePointPath(companyRoot, String(date.getFullYear()), `${date.getMonth() + 1} ${monthName(date)}`);
  if (useNewFolderStructure) {
    return buildNewStructureTargetPaths({ bank, accountType, accountNumber, incomePaymentType, category, monthRoot, fileName });
  }

  const bankRel = BANK_TO_CHECKINGS_DIR[bank];
  const categoryRel = CATEGORY_TO_MONTH_REL_PATH[category];
  if (!bankRel) throw new Error(`Bank has no mapping: ${bank}`);
  if (!categoryRel) throw new Error(`Category has no mapping: ${category}`);

  const bankFolderPath = joinSharePointPath(monthRoot, "2 Bank Transactions", bankRel, "Direct Payments");
  const categoryFolderPath = joinSharePointPath(monthRoot, ...categoryRel);
  const bankTargetPath = joinSharePointPath(bankFolderPath, fileName);
  const categoryTargetPath = joinSharePointPath(categoryFolderPath, fileName);
  return {
    mode: "legacy",
    bankFolderPath,
    categoryFolderPath,
    bankTargetPath,
    categoryTargetPath,
    folderPaths: [bankFolderPath, categoryFolderPath],
    targetPaths: [bankTargetPath, categoryTargetPath],
  };
}

function buildNewStructureTargetPaths({ bank, accountType, accountNumber, incomePaymentType, category, monthRoot, fileName }) {
  const categoryRel = NEW_STRUCTURE_CATEGORY_TO_MONTH_REL_PATH[category];
  if (!categoryRel) throw new Error(`Category has no new folder structure mapping: ${category}`);

  const categoryFolderPath = joinSharePointPath(monthRoot, ...categoryRel);
  if (newStructureCategoryNeedsPaymentType(category)) {
    return singleNewStructureTarget(joinSharePointPath(categoryFolderPath, normalizeFolderLevel(incomePaymentType)), fileName);
  }

  if (!newStructureCategoryNeedsAccount(category)) {
    return singleNewStructureTarget(categoryFolderPath, fileName);
  }

  if (newStructureRootForCategory(category) === "Expenses") {
    return singleNewStructureTarget(categoryFolderPath, fileName);
  }

  return singleNewStructureTarget(buildAccountScopedFolderPath({
    basePath: categoryFolderPath,
    bank,
    accountType,
    accountNumber,
  }), fileName);
}

function singleNewStructureTarget(folderPath, fileName) {
  const targetPath = joinSharePointPath(folderPath, fileName);
  return {
    mode: "new",
    uploadFolderPath: folderPath,
    uploadTargetPath: targetPath,
    folderPaths: [folderPath],
    targetPaths: [targetPath],
  };
}

function newStructureCategoryNeedsAccount(category) {
  const relPath = NEW_STRUCTURE_CATEGORY_TO_MONTH_REL_PATH[category] || [];
  return relPath[0] === "Expenses";
}

function newStructureRootForCategory(category) {
  const relPath = NEW_STRUCTURE_CATEGORY_TO_MONTH_REL_PATH[category] || [];
  return relPath[0] || "";
}

function newStructureCategoryNeedsBank(category) {
  return newStructureCategoryNeedsAccount(category);
}

function newStructureCategoryNeedsPaymentType(category) {
  const relPath = NEW_STRUCTURE_CATEGORY_TO_MONTH_REL_PATH[category] || [];
  return relPath[0] === "Income" && relPath[1] === "Transactions";
}

function newStructureCategoryBasePath({ company, category, date }) {
  const companyRoot = COMPANY_SHAREPOINT_DIRS[company];
  const categoryRel = NEW_STRUCTURE_CATEGORY_TO_MONTH_REL_PATH[category];
  if (!companyRoot || !categoryRel || !date) return "";
  const monthRoot = joinSharePointPath(companyRoot, String(date.getFullYear()), `${date.getMonth() + 1} ${monthName(date)}`);
  return joinSharePointPath(monthRoot, ...categoryRel);
}

function buildAccountScopedFolderPath({ basePath, bank, accountType, accountNumber }) {
  const bankFolder = normalizeFolderLevel(bank);
  if (bankFolder === NOT_IDENTIFIED_FOLDER) {
    return joinSharePointPath(basePath, NOT_IDENTIFIED_FOLDER);
  }

  const typeFolder = normalizeFolderLevel(accountType);
  if (typeFolder === NOT_IDENTIFIED_FOLDER) {
    return joinSharePointPath(basePath, bankFolder, NOT_IDENTIFIED_FOLDER);
  }

  return joinSharePointPath(basePath, bankFolder, typeFolder, normalizeAccountFolder(accountNumber));
}

function stampBankToTemplateBank(bank) {
  return STAMP_BANK_TO_TEMPLATE_BANK[bank] || bank;
}

function normalizeFolderLevel(value) {
  const text = String(value || "").trim().replace(/[\\/:*?"<>|\x00-\x1f]/g, " ").replace(/\s+/g, " ");
  return text || NOT_IDENTIFIED_FOLDER;
}

function normalizeAccountFolder(value) {
  const digits = lastFourDigits(value);
  if (!digits) return NOT_IDENTIFIED_FOLDER;
  if (!/^\d{4}$/.test(digits)) {
    throw new Error("Last four numbers of the card must be exactly 4 digits.");
  }
  return `N${digits}`;
}

function renderFolderTemplate() {
  const company = elements.templateCompanyInput.value;
  const year = Number(elements.templateYearInput.value) || new Date().getFullYear();
  const companyRoot = COMPANY_SHAREPOINT_DIRS[company];
  if (!companyRoot) throw new Error(`Company has no mapping: ${company}`);

  const month = elements.templateMonthInput.value;
  if (!templateBankAccounts.length) {
    currentFolderTemplateText = "";
    elements.templateSourceOutput.textContent = "Load the bank accounts workbook before generating a folder template.";
    elements.templateCountOutput.textContent = "0 folders";
    elements.templatePathOutput.textContent = joinSharePointPath(companyRoot, String(year), month);
    elements.folderTemplateTree.innerHTML = "";
    return;
  }

  const accounts = bankAccountsForCompany(company);
  const tree = buildFolderTemplateTree({ companyRoot, year, month, accounts });
  const folderCount = countFolderNodes(tree);
  currentFolderTemplateText = treeToText(tree);

  elements.templateSourceOutput.textContent = accounts.length
    ? `Bank accounts loaded: ${accounts.length} for this company.`
    : "Workbook loaded. No bank accounts found for this company; generating folders without bank subfolders.";
  elements.templateCountOutput.textContent = `${folderCount} folder${folderCount === 1 ? "" : "s"}`;
  elements.templatePathOutput.textContent = joinSharePointPath(companyRoot, String(year), month);
  elements.folderTemplateTree.innerHTML = renderTemplateNode(tree, true);
  window.lucide?.createIcons();
}

async function generateFolderTemplate() {
  elements.generateFolderTemplateBtn.disabled = true;
  elements.generateFolderTemplateBtn.querySelector("span").textContent = "Downloading";
  try {
    if (!window.JSZip) throw new Error("JSZip is not loaded.");
    await loadFolderTemplateCsv();
    const fileName = buildFolderTemplateFileName();
    const blob = await folderTemplateToZip(buildCurrentMonthFolderTemplateTree());
    triggerBrowserDownload(blob, fileName);
    showFlash(`Folder ZIP downloaded: ${fileName}`, "success");
  } finally {
    elements.generateFolderTemplateBtn.disabled = false;
    elements.generateFolderTemplateBtn.querySelector("span").textContent = "Download";
  }
}

async function loadFolderTemplateCsv() {
  elements.templateSourceOutput.textContent = "Loading bank accounts workbook...";
  elements.templateCountOutput.textContent = "0 folders";
  elements.folderTemplateTree.innerHTML = "";
  currentFolderTemplateText = "";
  await loadCloudTemplateBankWorkbook();
  renderFolderTemplate();
}

function buildFolderTemplateFileName() {
  const company = elements.templateCompanyInput.value;
  const year = Number(elements.templateYearInput.value) || new Date().getFullYear();
  const month = elements.templateMonthInput.value;
  return `${sanitizeFilename(company, "company")}_${year}_${sanitizeFilename(month, "month")}_folders.zip`;
}

function buildCurrentFolderTemplateTree() {
  const company = elements.templateCompanyInput.value;
  const year = Number(elements.templateYearInput.value) || new Date().getFullYear();
  const companyRoot = COMPANY_SHAREPOINT_DIRS[company];
  if (!companyRoot) throw new Error(`Company has no mapping: ${company}`);
  return buildFolderTemplateTree({
    companyRoot,
    year,
    month: elements.templateMonthInput.value,
    accounts: bankAccountsForCompany(company),
  });
}

function buildCurrentMonthFolderTemplateTree() {
  const company = elements.templateCompanyInput.value;
  const month = elements.templateMonthInput.value;
  const accounts = bankAccountsForCompany(company);
  return {
    name: month,
    children: [
      buildExpensesTemplate(accounts),
      buildIncomeTemplate(accounts),
      buildStatementsTemplate(accounts),
      buildDocumentsTemplate(),
    ],
  };
}

async function folderTemplateToZip(tree) {
  const zip = new window.JSZip();
  addFolderTreeToZip(zip, tree);
  return zip.generateAsync({ type: "blob" });
}

function addFolderTreeToZip(zip, node, path = "") {
  const name = typeof node === "string" ? node : node.name;
  const nextPath = joinSharePointPath(path, sanitizeZipPathPart(name));
  zip.folder(nextPath);
  for (const child of (typeof node === "string" ? [] : node.children || [])) {
    addFolderTreeToZip(zip, child, nextPath);
  }
}

function sanitizeZipPathPart(value) {
  return String(value || "folder")
    .replace(/[\\/:*?"<>|\x00-\x1f]/g, " ")
    .replace(/\s+/g, " ")
    .trim() || "folder";
}

function buildFolderTemplateTree({ companyRoot, year, month, accounts }) {
  return {
    name: joinSharePointPath(companyRoot, String(year)),
    children: [{
      name: month,
      children: [
        buildExpensesTemplate(accounts),
        buildIncomeTemplate(accounts),
        buildStatementsTemplate(accounts),
        buildDocumentsTemplate(),
      ],
    }],
  };
}

function buildDocumentsTemplate() {
  return {
    name: "Documents",
    children: [
      {
        name: "Tax Remittances",
        children: ["GST HST", "Payroll Remittances", "QST"],
      },
      {
        name: "Work Safety",
        children: ["CNESST QC", "WCB NB", "WCB NS", "WCB PEI", "WSIB ON"],
      },
      "Comite Paritario",
      "Investment Land",
      "Letters Received",
      "Union",
      "Xoom Investment",
    ],
  };
}

function buildExpensesTemplate(accounts) {
  return {
    name: "Expenses",
    children: [
      ...EXPENSE_CATEGORIES.map((category) => ({
        name: category,
        children: accountTree(accounts),
      })),
      {
        name: "Reimbursements",
        children: ["OP", "Reimbursements"],
      },
    ],
  };
}

function buildIncomeTemplate(accounts) {
  return {
    name: "Income",
    children: [
      "Invoices",
      {
        name: "Transactions",
        children: ["Cheque", "Cash"],
      },
    ],
  };
}

function buildStatementsTemplate(accounts) {
  return {
    name: "Statements",
    children: accountTree(accounts),
  };
}

function accountTree(accounts) {
  const grouped = new Map();
  for (const account of accounts) {
    if (!grouped.has(account.bank)) grouped.set(account.bank, new Map());
    const typeMap = grouped.get(account.bank);
    if (!typeMap.has(account.type)) typeMap.set(account.type, []);
    typeMap.get(account.type).push(account.last4);
  }

  return Array.from(grouped.entries())
    .sort(([left], [right]) => left.localeCompare(right, undefined, { numeric: true }))
    .map(([bank, typeMap]) => ({
      name: bank,
      children: Array.from(typeMap.entries())
        .sort(([left], [right]) => left.localeCompare(right, undefined, { numeric: true }))
        .map(([type, last4List]) => ({
          name: type,
          children: Array.from(new Set(last4List.filter(Boolean))).sort().map((last4) => ({ name: last4, children: [] })),
        })),
    }));
}

function renderTemplateNode(node, isRoot = false) {
  if (typeof node === "string") {
    return `<div class="template-leaf"><i data-lucide="folder"></i><span>${escapeHtml(node)}</span></div>`;
  }
  const children = node.children || [];
  if (!children.length) {
    return `<div class="template-leaf"><i data-lucide="folder"></i><span>${escapeHtml(node.name)}</span></div>`;
  }
  return `
    <details class="tree-node template-node" ${isRoot ? "open" : ""}>
      <summary>
        <span class="tree-label">${escapeHtml(node.name)}</span>
        <span class="tree-count">${countFolderNodes(node)}</span>
      </summary>
      <div class="tree-children">
        ${children.map((child) => renderTemplateNode(child)).join("")}
      </div>
    </details>
  `;
}

function countFolderNodes(node) {
  if (typeof node === "string") return 1;
  return 1 + (node.children || []).reduce((total, child) => total + countFolderNodes(child), 0);
}

function treeToText(node, depth = 0) {
  if (typeof node === "string") {
    const leafPrefix = depth ? `${"  ".repeat(depth - 1)}- ` : "";
    return `${leafPrefix}${node}`;
  }
  const prefix = depth ? `${"  ".repeat(depth - 1)}- ` : "";
  const lines = [`${prefix}${node.name}`];
  for (const child of node.children || []) {
    lines.push(treeToText(child, depth + 1));
  }
  return lines.join("\n");
}

async function loadCloudTemplateBankWorkbook() {
  try {
    const bytes = await downloadSharePointFile(BANK_ACCOUNTS_WORKBOOK_PATH);
    templateBankAccounts = parseBankAccountsWorkbook(bytes);
  } catch (error) {
    if (!isCsvNotFoundError(error)) throw error;
    templateBankAccounts = [];
    renderFolderTemplate();
    throw new Error(`Bank accounts workbook not found: ${BANK_ACCOUNTS_WORKBOOK_PATH}`);
  }

  if (!templateBankAccounts.length) {
    renderFolderTemplate();
    throw new Error(`Bank accounts workbook is empty or has no valid rows: ${BANK_ACCOUNTS_WORKBOOK_PATH}`);
  }
}

async function ensureFolderTemplateBankAccountsLoaded() {
  if (templateBankAccounts.length) return;
  await loadCloudTemplateBankWorkbook();
}

function isCsvNotFoundError(error) {
  return /itemNotFound|could not be found|not found/i.test(error?.message || String(error));
}

function parseBankAccountsWorkbook(bytes) {
  if (!window.XLSX) throw new Error("XLSX is not loaded.");
  const workbook = window.XLSX.read(bytes, { type: "array" });
  const sheet = workbook.Sheets["Cuentas Bancarias"] || workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) return [];

  const rows = window.XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
  if (rows.length < 2) return [];
  const headers = rows[0].map(normalizeCsvHeader);
  const indexes = {
    company: firstHeaderIndex(headers, ["company", "empresa"]),
    bank: firstHeaderIndex(headers, ["bank", "banco"]),
    type: firstHeaderIndex(headers, ["type", "tipo"]),
    last4: firstHeaderIndex(headers, ["last 4 digits", "ultimos 4 digitos"]),
  };
  if (indexes.company < 0 || indexes.bank < 0 || indexes.type < 0 || indexes.last4 < 0) {
    throw new Error("Workbook must include Company, Bank, Type, and Last 4 digits columns.");
  }

  return rows.slice(1).map((row, index) => {
    const company = String(row[indexes.company] || "").trim();
    const bank = String(row[indexes.bank] || "").trim();
    const type = String(row[indexes.type] || "").trim();
    const last4 = String(row[indexes.last4] || "").trim().toUpperCase();
    if (company && !COMPANIES.includes(company)) {
      throw new Error(`Invalid Company at workbook row ${index + 2}. Choose a company from the allowed list.`);
    }
    if (bank && !TEMPLATE_BANKS.includes(bank)) {
      throw new Error(`Invalid Bank at workbook row ${index + 2}. Choose a bank from the allowed list.`);
    }
    if (type && !TEMPLATE_ACCOUNT_TYPES.includes(type)) {
      throw new Error(`Invalid Type at workbook row ${index + 2}. Choose debit or credit.`);
    }
    if ((company || bank || type || last4) && !/^N\d{4}$/.test(last4)) {
      throw new Error(`Invalid Last 4 digits at workbook row ${index + 2}. Expected NXXXX.`);
    }
    return {
      company,
      bank,
      type,
      last4,
    };
  }).filter((account) => account.bank && account.type && account.last4);
}

function firstHeaderIndex(headers, names) {
  return names.map((name) => headers.indexOf(name)).find((index) => index >= 0) ?? -1;
}

function bankAccountsForCompany(company) {
  if (!templateBankAccounts.length) return [];
  const normalizedCompany = normalizeCsvHeader(company);
  const exactMatches = templateBankAccounts.filter((account) => normalizeCsvHeader(account.company) === normalizedCompany);
  if (exactMatches.length) return exactMatches;
  const genericAccounts = templateBankAccounts.filter((account) => !account.company);
  return genericAccounts;
}

function normalizeCsvHeader(value) {
  return String(value || "")
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

async function copyFolderTemplate() {
  if (!currentFolderTemplateText) renderFolderTemplate();
  if (!currentFolderTemplateText) {
    throw new Error("Load the bank accounts workbook before copying a folder template.");
  }
  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(currentFolderTemplateText);
  } else {
    const textArea = document.createElement("textarea");
    textArea.value = currentFolderTemplateText;
    textArea.setAttribute("readonly", "");
    textArea.style.position = "fixed";
    textArea.style.opacity = "0";
    document.body.appendChild(textArea);
    textArea.select();
    document.execCommand("copy");
    textArea.remove();
  }
  showFlash("Folder template copied to clipboard.", "success");
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
    throw new Error(cleanErrorText(await response.text()));
  }
  return response.json();
}

function cleanErrorText(value) {
  const text = String(value || "");
  const stripped = text
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/\s+/g, " ")
    .trim();
  if (/Inactivity Timeout|Too much time has passed without sending any data/i.test(stripped)) {
    return "CNET session timed out while exporting invoices. Try the download again.";
  }
  return stripped || "Request failed.";
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
