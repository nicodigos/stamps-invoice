const CNET_BASE_URL = "https://app.master.cnetfranchise.com";
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const INVOICES_BASE_PATH = "General/Invoices";
const REQUIRED_COLUMNS = [
  "Invoice ID",
  "Vendor Company Name",
  "Creation Date",
  "Buyer Company Name",
  "Payment Status",
  "Total Amount With Taxes",
];
const MONTH_NAMES = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];

class CookieJar {
  constructor() {
    this.cookies = new Map();
  }

  addFromHeaders(headers) {
    const setCookies = typeof headers.getSetCookie === "function"
      ? headers.getSetCookie()
      : splitSetCookie(headers.get("set-cookie") || "");
    for (const source of setCookies) {
      const first = String(source || "").split(";")[0];
      const separatorIndex = first.indexOf("=");
      if (separatorIndex <= 0) continue;
      this.cookies.set(first.slice(0, separatorIndex).trim(), first.slice(separatorIndex + 1).trim());
    }
  }

  header() {
    return Array.from(this.cookies.entries()).map(([key, value]) => `${key}=${value}`).join("; ");
  }
}

async function cnetRequest(jar, pathname, options = {}) {
  const url = absoluteCnetUrl(pathname);
  const response = await fetch(url, {
    redirect: "manual",
    ...options,
    headers: {
      "User-Agent": "Mozilla/5.0",
      ...(jar.header() ? { Cookie: jar.header() } : {}),
      ...(options.headers || {}),
    },
  });
  jar.addFromHeaders(response.headers);
  return response;
}

async function loginToCnet() {
  const username = process.env.CNET_INVOICES_USER || "";
  const password = process.env.CNET_INVOICES_PASS || "";
  if (!username || !password) {
    throw new Error("Missing CNET_INVOICES_USER or CNET_INVOICES_PASS.");
  }

  const jar = new CookieJar();
  const loginResponse = await cnetRequest(jar, "/login");
  const loginHtml = await loginResponse.text();
  if (!loginResponse.ok) {
    throw new Error(`CNET login page failed: HTTP ${loginResponse.status}.`);
  }

  const csrfToken = extractCsrfToken(loginHtml);
  const postResponse = await cnetRequest(jar, "/login_check", {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
      Referer: `${CNET_BASE_URL}/login`,
    },
    body: new URLSearchParams({
      _csrf_token: csrfToken,
      _username: username,
      _password: password,
      _remember_me: "on",
      _submit: "Login",
    }),
  });

  if (isRedirect(postResponse.status)) {
    const location = postResponse.headers.get("location") || "/manager";
    const follow = await followCnetRedirects(jar, location);
    const followText = await follow.text();
    if (!follow.ok || /name=["']_username["']/i.test(followText)) {
      throw new Error("CNET did not accept the login.");
    }
    return jar;
  }

  const postText = await postResponse.text();
  if (!postResponse.ok || /name=["']_username["']/i.test(postText)) {
    throw new Error("CNET did not accept the login.");
  }
  return jar;
}

async function followCnetRedirects(jar, location) {
  let response = await cnetRequest(jar, location);
  for (let count = 0; count < 8 && isRedirect(response.status); count += 1) {
    const nextLocation = response.headers.get("location");
    if (!nextLocation) break;
    response = await cnetRequest(jar, nextLocation);
  }
  return response;
}

async function exportInvoiceRows() {
  const jar = await loginToCnet();
  const response = await cnetRequest(jar, "/manager/invoices/export", {
    headers: { Accept: "text/csv,application/csv,text/plain,*/*" },
  });
  const content = await response.text();
  if (!response.ok) {
    throw new Error(`CNET export failed: HTTP ${response.status}. ${content.slice(0, 300)}`);
  }
  if (/name=["']_username["']/i.test(content)) {
    throw new Error("CNET returned to login while exporting invoices.");
  }
  const fileName = contentDispositionFilename(response.headers.get("content-disposition") || "");
  return normalizeCsvRows(content, fileName);
}

async function downloadInvoicePdf(row, jar) {
  validateRow(row);
  const showResponse = await cnetRequest(jar, `/manager/invoices/${encodeURIComponent(row.invoiceId)}/show`);
  const showHtml = await showResponse.text();
  if (!showResponse.ok) {
    throw new Error(`Could not open invoice ${row.invoiceId}: HTTP ${showResponse.status}.`);
  }
  const pdfHref = extractPdfHref(showHtml);
  const pdfResponse = await cnetRequest(jar, pdfHref, {
    headers: { Accept: "application/pdf,*/*" },
  });
  const contentType = pdfResponse.headers.get("content-type") || "";
  const bytes = Buffer.from(await pdfResponse.arrayBuffer());
  if (!pdfResponse.ok) {
    throw new Error(`Could not download PDF for invoice ${row.invoiceId}: HTTP ${pdfResponse.status}.`);
  }
  if (!contentType.includes("pdf") && bytes.slice(0, 4).toString("latin1") !== "%PDF") {
    throw new Error(`CNET did not return a PDF for invoice ${row.invoiceId}.`);
  }
  return bytes;
}

async function uploadInvoicePdf(row, pdfBytes, accessToken) {
  const driveId = await resolveDriveId(accessToken);
  const targetPath = buildInvoiceTargetPath(row);
  const encodedPath = encodeURIComponent(targetPath).replaceAll("%2F", "/");
  const response = await fetch(`${GRAPH_BASE}/drives/${driveId}/root:/${encodedPath}:/content`, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/pdf",
    },
    body: pdfBytes,
  });
  if (!response.ok) {
    throw new Error(await response.text());
  }
  return { targetPath, graphItem: await response.json() };
}

async function listExistingInvoiceIndex(accessToken) {
  return (await listExistingInvoiceTree(accessToken)).existing;
}

async function listExistingInvoiceTree(accessToken) {
  const driveId = await resolveDriveId(accessToken);
  const encodedPath = encodeURIComponent(INVOICES_BASE_PATH).replaceAll("%2F", "/");
  const rootUrl = `${GRAPH_BASE}/drives/${driveId}/root:/${encodedPath}:/children?$select=id,name,folder,file`;
  const items = [];

  try {
    await collectGraphChildren(rootUrl, accessToken, items, driveId, []);
  } catch (error) {
    if (String(error.message || "").includes("\"code\":\"itemNotFound\"") || String(error.message || "").includes("itemNotFound")) {
      return { existing: {}, tree: [], totalFiles: 0 };
    }
    throw error;
  }

  const index = {};
  const treeMap = new Map();
  for (const item of items) {
    if (!item.file || !/\.pdf$/i.test(item.name || "")) continue;
    const parsed = parseExistingInvoiceFileName(item.name);
    if (parsed) {
      if (!index[parsed.invoiceId]) index[parsed.invoiceId] = [];
      if (!index[parsed.invoiceId].includes(parsed.paymentStatus)) {
        index[parsed.invoiceId].push(parsed.paymentStatus);
      }
    }
    addFileToTree(treeMap, item.pathSegments || [], item);
  }
  return {
    existing: index,
    tree: treeFromMap(treeMap),
    totalFiles: items.filter((item) => item.file && /\.pdf$/i.test(item.name || "")).length,
    driveId,
  };
}

async function collectGraphChildren(url, token, items, driveId, pathSegments) {
  let nextUrl = url;
  while (nextUrl) {
    const page = await graphJson(nextUrl, token);
    for (const item of page.value || []) {
      const nextPathSegments = [...pathSegments, item.name];
      if (item.folder) {
        await collectGraphChildren(`${GRAPH_BASE}/drives/${driveId}/items/${item.id}/children?$select=id,name,folder,file`, token, items, driveId, nextPathSegments);
      } else {
        items.push({ ...item, pathSegments: nextPathSegments });
      }
    }
    nextUrl = page["@odata.nextLink"] || "";
  }
}

function addFileToTree(treeMap, pathSegments, fileItem) {
  const [company = "Unknown", year = "Unknown", month = "Unknown"] = pathSegments;
  if (!treeMap.has(company)) {
    treeMap.set(company, { name: company, count: 0, years: new Map() });
  }
  const companyNode = treeMap.get(company);
  if (!companyNode.years.has(year)) {
    companyNode.years.set(year, { name: year, count: 0, months: new Map() });
  }
  const yearNode = companyNode.years.get(year);
  if (!yearNode.months.has(month)) {
    yearNode.months.set(month, { name: month, count: 0, files: [] });
  }
  const monthNode = yearNode.months.get(month);
  companyNode.count += 1;
  yearNode.count += 1;
  monthNode.count += 1;
  monthNode.files.push({
    id: fileItem.id || "",
    name: fileItem.name || "",
  });
}

function treeFromMap(treeMap) {
  return Array.from(treeMap.values())
    .sort(compareName)
    .map((company) => ({
      name: company.name,
      count: company.count,
      years: Array.from(company.years.values())
        .sort((left, right) => String(right.name).localeCompare(String(left.name), undefined, { numeric: true }))
        .map((year) => ({
          name: year.name,
          count: year.count,
          months: Array.from(year.months.values())
            .sort(compareName)
            .map((month) => ({
              name: month.name,
              count: month.count,
              files: month.files.sort((left, right) => left.name.localeCompare(right.name, undefined, { numeric: true })),
            })),
        })),
    }));
}

function compareName(left, right) {
  return String(left.name).localeCompare(String(right.name), undefined, { numeric: true });
}

async function resolveDriveId(token) {
  const spHostname = process.env.SP_HOSTNAME || "";
  const spSitePath = process.env.SP_SITE_PATH || "";
  const spDriveName = process.env.SP_DRIVE_NAME || "Documents";
  if (!spHostname || !spSitePath) {
    throw new Error("Missing SP_HOSTNAME or SP_SITE_PATH.");
  }

  const site = await graphJson(`${GRAPH_BASE}/sites/${spHostname}:${spSitePath}`, token);
  const drives = (await graphJson(`${GRAPH_BASE}/sites/${site.id}/drives`, token)).value || [];
  const drive = drives.find((item) => item.name === spDriveName) || drives[0];
  if (!drive) {
    throw new Error("Could not resolve the SharePoint drive.");
  }
  return drive.id;
}

async function graphJson(url, token) {
  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!response.ok) {
    throw new Error(await response.text());
  }
  return response.json();
}

function normalizeCsvRows(csvText, fileName) {
  const parsed = parseCsv(csvText);
  if (!parsed.length) {
    return { fileName, rows: [], errors: [{ invoiceId: "-", message: "CSV is empty." }] };
  }
  const headers = parsed[0].map((header) => String(header || "").trim());
  const headerMap = new Map(headers.map((header, index) => [normalizeHeader(header), index]));
  const missingColumns = REQUIRED_COLUMNS.filter((column) => !headerMap.has(normalizeHeader(column)));
  if (missingColumns.length) {
    throw new Error(`The CSV is missing required columns: ${missingColumns.join(", ")}.`);
  }

  const rows = [];
  const errors = [];
  for (let index = 1; index < parsed.length; index += 1) {
    const source = parsed[index];
    if (!source.some((cell) => String(cell || "").trim())) continue;
    const row = {
      sourceRowNumber: index + 1,
      invoiceId: csvValue(source, headerMap, "Invoice ID"),
      vendorCompanyName: csvValue(source, headerMap, "Vendor Company Name"),
      creationDate: csvValue(source, headerMap, "Creation Date"),
      buyerCompanyName: csvValue(source, headerMap, "Buyer Company Name"),
      paymentStatus: csvValue(source, headerMap, "Payment Status"),
      totalAmountWithTaxes: csvValue(source, headerMap, "Total Amount With Taxes"),
      errors: [],
    };
    for (const field of ["invoiceId", "vendorCompanyName", "creationDate", "buyerCompanyName", "paymentStatus", "totalAmountWithTaxes"]) {
      if (!row[field]) row.errors.push(`Row ${row.sourceRowNumber}: missing ${field}.`);
    }
    try {
      parseCreationDate(row.creationDate);
    } catch (error) {
      row.errors.push(`Fila ${row.sourceRowNumber}: ${error.message}`);
    }
    rows.push(row);
  }
  return { fileName, rows, errors };
}

function parseCsv(source) {
  const rows = [];
  let row = [];
  let cell = "";
  let quoted = false;

  for (let index = 0; index < source.length; index += 1) {
    const char = source[index];
    const next = source[index + 1];
    if (quoted) {
      if (char === "\"" && next === "\"") {
        cell += "\"";
        index += 1;
      } else if (char === "\"") {
        quoted = false;
      } else {
        cell += char;
      }
      continue;
    }
    if (char === "\"") {
      quoted = true;
    } else if (char === ",") {
      row.push(cell);
      cell = "";
    } else if (char === "\n") {
      row.push(cell.replace(/\r$/, ""));
      rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += char;
    }
  }
  row.push(cell.replace(/\r$/, ""));
  if (row.length > 1 || row[0]) rows.push(row);
  return rows;
}

function csvValue(row, headerMap, headerName) {
  const index = headerMap.get(normalizeHeader(headerName));
  return String(row[index] || "").trim();
}

function normalizeHeader(value) {
  return String(value || "").trim().toLowerCase().replace(/\s+/g, " ");
}

function validateRow(row) {
  if (!row || typeof row !== "object") throw new Error("Invalid row.");
  for (const field of ["invoiceId", "vendorCompanyName", "creationDate", "buyerCompanyName", "paymentStatus", "totalAmountWithTaxes"]) {
    if (!String(row[field] || "").trim()) {
      throw new Error(`Invoice ${row.invoiceId || "-"} is missing required field: ${field}.`);
    }
  }
}

function buildInvoiceTargetPath(row) {
  const date = parseCreationDate(row.creationDate);
  const year = String(date.getFullYear());
  const month = `${String(date.getMonth() + 1).padStart(2, "0")} - ${MONTH_NAMES[date.getMonth()]}`;
  const fileName = [
    row.invoiceId,
    row.buyerCompanyName,
    row.paymentStatus,
    row.totalAmountWithTaxes,
  ].map((part) => sanitizePathPart(part)).join("_");

  return [
    INVOICES_BASE_PATH,
    sanitizePathPart(row.vendorCompanyName),
    year,
    month,
    `${fileName}.pdf`,
  ].join("/");
}

function parseExistingInvoiceFileName(fileName) {
  const stem = String(fileName || "").replace(/\.pdf$/i, "");
  const parts = stem.split("_");
  if (parts.length < 4) return null;
  const invoiceId = parts[0]?.trim();
  const paymentStatus = parts[parts.length - 2]?.trim();
  if (!invoiceId || !paymentStatus) return null;
  return { invoiceId, paymentStatus };
}

function parseCreationDate(value) {
  const text = String(value || "").trim();
  if (!text) throw new Error("Creation Date is empty.");
  const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (iso) return validDate(Number(iso[1]), Number(iso[2]), Number(iso[3]), text);
  const slash = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (slash) return validDate(Number(slash[3]), Number(slash[1]), Number(slash[2]), text);
  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) return parsed;
  throw new Error(`Invalid Creation Date: ${text}.`);
}

function validDate(year, month, day, source) {
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
    throw new Error(`Invalid Creation Date: ${source}.`);
  }
  return date;
}

function sanitizePathPart(value) {
  const sanitized = String(value || "")
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[\\/:*?"<>|#%&{}~\x00-\x1f]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/[. ]+$/g, "");
  return (sanitized || "Unknown").slice(0, 120);
}

function extractCsrfToken(html) {
  const match = html.match(/name=["']_csrf_token["'][^>]*value=["']([^"']+)["']/i)
    || html.match(/value=["']([^"']+)["'][^>]*name=["']_csrf_token["']/i);
  if (!match) {
    throw new Error("Could not find _csrf_token in CNET.");
  }
  return decodeHtml(match[1]);
}

function extractPdfHref(html) {
  const anchorMatches = html.match(/<a\b[\s\S]*?<\/a>/gi) || [];
  for (const anchor of anchorMatches) {
    if (!/Download\s+PDF/i.test(stripTags(anchor)) && !/print\/pdf/i.test(anchor)) continue;
    const href = anchor.match(/\bhref=["']([^"']+)["']/i)?.[1];
    if (href) return decodeHtml(href);
  }
  throw new Error("Could not find the Download PDF button in CNET.");
}

function absoluteCnetUrl(pathname) {
  if (/^https?:\/\//i.test(pathname)) return pathname;
  return new URL(pathname, CNET_BASE_URL).toString();
}

function contentDispositionFilename(header) {
  const match = header.match(/filename\*=UTF-8''([^;]+)/i) || header.match(/filename="?([^";]+)"?/i);
  return match ? decodeURIComponent(match[1]) : "";
}

function splitSetCookie(header) {
  if (!header) return [];
  return header.split(/,(?=\s*[^;,]+=)/g);
}

function isRedirect(status) {
  return [301, 302, 303, 307, 308].includes(status);
}

function stripTags(value) {
  return String(value || "").replace(/<[^>]+>/g, " ");
}

function decodeHtml(value) {
  return String(value || "")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, "\"")
    .replace(/&#039;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">");
}

module.exports = {
  loginToCnet,
  exportInvoiceRows,
  downloadInvoicePdf,
  uploadInvoicePdf,
  buildInvoiceTargetPath,
  listExistingInvoiceIndex,
  listExistingInvoiceTree,
};
