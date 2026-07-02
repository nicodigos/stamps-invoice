import { state } from "./state.js";
import { joinSharePointPath, normalizeCardLast4, rowsToCsv, sanitizeFilenameComponent } from "./utils.js";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const BACKUP_FOLDER_NAME = "data_versions";
const BACKUP_MARKER_FILE = "lastupdated.txt";
const BACKUP_INTERVAL_MS = 24 * 60 * 60 * 1000;
const MAX_BACKUP_VERSIONS = 60;

let backupCheckPromise = null;

class GraphRequestError extends Error {
  constructor(message, status) {
    super(message);
    this.name = "GraphRequestError";
    this.status = status;
  }
}

function authHeaders() {
  if (!state.graphToken) {
    throw new Error("Microsoft no esta conectado.");
  }
  return { Authorization: `Bearer ${state.graphToken}` };
}

async function graphFetch(url, options = {}) {
  const response = await fetch(url, {
    ...options,
    headers: {
      ...authHeaders(),
      ...(options.headers || {}),
    },
  });
  if (!response.ok) {
    throw new GraphRequestError(await response.text(), response.status);
  }
  return response;
}

async function graphJson(url, options = {}) {
  const response = await graphFetch(url, options);
  return response.status === 204 ? {} : response.json();
}

async function graphBytes(url, options = {}) {
  const response = await graphFetch(url, options);
  return {
    bytes: await response.arrayBuffer(),
    eTag: response.headers.get("etag"),
  };
}

async function downloadSharePointFileWithMeta(path) {
  const driveId = await resolveDriveId();
  const encodedPath = encodeURIComponent(path).replaceAll("%2F", "/");
  return graphBytes(`${GRAPH_BASE}/drives/${driveId}/root:/${encodedPath}:/content`);
}

export async function deleteSharePointItemById(itemId) {
  const driveId = await resolveDriveId();
  await graphFetch(`${GRAPH_BASE}/drives/${driveId}/items/${itemId}`, { method: "DELETE" });
}

async function createSharePointFolder(parentPath, folderName) {
  const driveId = await resolveDriveId();
  const encodedParentPath = encodeURIComponent(parentPath).replaceAll("%2F", "/");
  return graphJson(`${GRAPH_BASE}/drives/${driveId}/root:/${encodedParentPath}:/children`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      name: folderName,
      folder: {},
      "@microsoft.graph.conflictBehavior": "replace",
    }),
  });
}

export async function resolveDriveId() {
  if (state.driveId) return state.driveId;
  const { spHostname, spSitePath, spDriveName } = state.config;
  const site = await graphJson(`${GRAPH_BASE}/sites/${spHostname}:${spSitePath}`);
  const drives = (await graphJson(`${GRAPH_BASE}/sites/${site.id}/drives`)).value || [];
  const drive = drives.find((item) => item.name === spDriveName) || drives[0];
  if (!drive) {
    throw new Error("No se pudo resolver el drive de SharePoint.");
  }
  state.driveId = drive.id;
  return drive.id;
}

export async function listChildrenByPath(path) {
  const driveId = await resolveDriveId();
  const encodedPath = encodeURIComponent(path).replaceAll("%2F", "/");
  try {
    const url = `${GRAPH_BASE}/drives/${driveId}/root:/${encodedPath}:/children?$top=200&$select=id,name,folder,file,webUrl,lastModifiedDateTime`;
    const data = await graphJson(url);
    return data.value || [];
  } catch (error) {
    if (error instanceof GraphRequestError && error.status === 404) {
      return [];
    }
    throw error;
  }
}

export async function downloadSharePointFile(path) {
  const { bytes } = await downloadSharePointFileWithMeta(path);
  return new Uint8Array(bytes);
}

export async function uploadSharePointFile(path, content, contentType = "application/octet-stream", options = {}) {
  const driveId = await resolveDriveId();
  const encodedPath = encodeURIComponent(path).replaceAll("%2F", "/");
  return graphJson(`${GRAPH_BASE}/drives/${driveId}/root:/${encodedPath}:/content`, {
    method: "PUT",
    headers: {
      "Content-Type": contentType,
      ...(options.expectedEtag === null ? { "If-None-Match": "*" } : {}),
      ...(options.expectedEtag ? { "If-Match": options.expectedEtag } : {}),
    },
    body: content,
  });
}

export async function loadDatabaseRows() {
  await ensureDailyCsvBackup();
  const csvPath = joinSharePointPath(state.config.receiptsDatabaseDir, state.config.receiptsDatabaseCsv);
  try {
    const { bytes, eTag } = await downloadSharePointFileWithMeta(csvPath);
    const text = new TextDecoder("utf-8").decode(bytes);
    return { rows: parseCsv(text), eTag };
  } catch (error) {
    if (error instanceof GraphRequestError && error.status === 404) {
      return { rows: [], eTag: null };
    }
    if (error instanceof GraphRequestError) {
      throw new Error("No se pudo leer el CSV de SharePoint. Recarga antes de volver a guardar.");
    }
    throw new Error("El CSV de SharePoint no se pudo parsear. Revisa el archivo antes de volver a guardar.");
  }
}

export async function saveDatabaseRows(rows, options = {}) {
  await ensureDailyCsvBackup();
  const csvPath = joinSharePointPath(state.config.receiptsDatabaseDir, state.config.receiptsDatabaseCsv);
  const csvText = rowsToCsv(rows);
  try {
    const item = await uploadSharePointFile(
      csvPath,
      new TextEncoder().encode(csvText),
      "text/csv;charset=utf-8",
      { expectedEtag: options.expectedEtag },
    );
    return item.eTag || null;
  } catch (error) {
    if (error instanceof GraphRequestError && (error.status === 409 || error.status === 412)) {
      throw new Error("El CSV cambio en SharePoint desde tu ultima carga. Recarga la base antes de guardar para evitar sobreescribir cambios ajenos.");
    }
    throw error;
  }
}

async function ensureDailyCsvBackup() {
  if (!backupCheckPromise) {
    backupCheckPromise = runDailyCsvBackupCheck().finally(() => {
      backupCheckPromise = null;
    });
  }
  return backupCheckPromise;
}

async function runDailyCsvBackupCheck() {
  const { csvPath, backupDirPath, markerPath } = getBackupPaths();
  await ensureBackupFolderExists(backupDirPath);
  const lastUpdatedAt = await readBackupMarker(markerPath);
  if (!shouldCreateBackup(lastUpdatedAt)) {
    return;
  }

  let csvBytes;
  try {
    const downloaded = await downloadSharePointFileWithMeta(csvPath);
    csvBytes = downloaded.bytes;
  } catch (error) {
    if (error instanceof GraphRequestError && error.status === 404) {
      await writeBackupMarker(markerPath, new Date());
      return;
    }
    throw new Error("No se pudo leer el CSV para generar la version diaria.");
  }

  const backupFilePath = joinSharePointPath(backupDirPath, buildBackupFileName());
  await uploadSharePointFile(backupFilePath, new Uint8Array(csvBytes), "text/csv;charset=utf-8");
  await trimBackupVersions(backupDirPath);
  await writeBackupMarker(markerPath, new Date());
}

function getBackupPaths() {
  const csvPath = joinSharePointPath(state.config.receiptsDatabaseDir, state.config.receiptsDatabaseCsv);
  const backupDirPath = joinSharePointPath(state.config.receiptsDatabaseDir, BACKUP_FOLDER_NAME);
  const markerPath = joinSharePointPath(backupDirPath, BACKUP_MARKER_FILE);
  return { csvPath, backupDirPath, markerPath };
}

async function ensureBackupFolderExists(backupDirPath) {
  const parentPath = state.config.receiptsDatabaseDir;
  const folderName = BACKUP_FOLDER_NAME;
  const children = await listChildrenByPath(parentPath);
  const folder = children.find((item) => item.folder && item.name === folderName);
  if (folder) return;
  await createSharePointFolder(parentPath, folderName);
}

async function readBackupMarker(path) {
  try {
    const { bytes } = await downloadSharePointFileWithMeta(path);
    const text = new TextDecoder("utf-8").decode(bytes).trim();
    const match = text.match(/lastupdated at\s*[:=]\s*(.+)$/i);
    const value = match ? match[1].trim() : text;
    const parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  } catch (error) {
    if (error instanceof GraphRequestError && error.status === 404) {
      return null;
    }
    throw new Error("No se pudo leer el archivo de control de versiones del CSV.");
  }
}

function shouldCreateBackup(lastUpdatedAt) {
  if (!(lastUpdatedAt instanceof Date)) return true;
  return (Date.now() - lastUpdatedAt.getTime()) > BACKUP_INTERVAL_MS;
}

function buildBackupFileName() {
  const baseName = sanitizeFilenameComponent(
    state.config.receiptsDatabaseCsv.replace(/\.csv$/i, ""),
    "receipts_database",
  );
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
  return `${baseName}__${timestamp}.csv`;
}

async function trimBackupVersions(backupDirPath) {
  const children = await listChildrenByPath(backupDirPath);
  const csvFiles = children
    .filter((item) => item.file && /\.csv$/i.test(item.name))
    .sort((left, right) => new Date(left.lastModifiedDateTime || 0) - new Date(right.lastModifiedDateTime || 0));

  while (csvFiles.length > MAX_BACKUP_VERSIONS) {
    const oldest = csvFiles.shift();
    if (!oldest?.id) break;
    await deleteSharePointItemById(oldest.id);
  }
}

async function writeBackupMarker(path, date) {
  const content = `lastupdated at: ${date.toISOString()}\n`;
  await uploadSharePointFile(path, new TextEncoder().encode(content), "text/plain;charset=utf-8");
}

function parseCsv(text) {
  const workbook = XLSX.read(text, { type: "string" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  return rows.map((row) => {
    const legacyChecked = String(row.checked).toLowerCase() === "true" || row.checked === true;
    const normalizedStatus = normalizeStatus(row.status || (legacyChecked ? "Paid" : "Pending"));
    const normalizedReceiptType = normalizeReceiptType(row.receipt_type);
    const { checked, ...rest } = row;
    return {
      ...rest,
      card_last4: normalizeStoredCardLast4(rest.card_last4),
      receipt_type: normalizedReceiptType,
      status: normalizedStatus,
    };
  });
}

function normalizeStoredCardLast4(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return "";

  const digits = raw.replace(/\D/g, "");
  if (digits.length >= 1 && digits.length <= 4) {
    return `N${digits.padStart(4, "0")}`;
  }

  const normalized = normalizeCardLast4(raw);
  if (normalized) return normalized;

  return raw;
}

function normalizeStatus(value) {
  return String(value).trim().toLowerCase() === "paid" ? "Paid" : "Pending";
}

function normalizeReceiptType(value) {
  return String(value).trim().toLowerCase() === "reimbursement" ? "reimbursement" : "bank_transaction";
}
