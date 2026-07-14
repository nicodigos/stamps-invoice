import { mkdir, readFile, readdir, stat } from "node:fs/promises";
import path from "node:path";
import process from "node:process";

const GENERAL_ROOT = String.raw`C:\Users\Melissa Martinez\OneDrive - Group Castillo (12433087 Canada Inc)\Group Castillo Team Site - General`;
const APPLY = process.argv.includes("--apply");

const APPROVED_ROOTS = [
  ["1001298527 ONTARIO INC", "Documents Accountant"],
  ["10342548 CANADA INC", "02-Documents accountant 10342548"],
  ["10377180 CANADA INC-standby", "02-Documents Accountant"],
  ["10696480 CANADA LTD", "02-Documents accountant 10696480"],
  ["12433087 CANADA INC-MASTER", "01-Documents Accountant 12433087"],
  ["13037622 CANADA INC", "01-Documents Accountant 13037622"],
  ["16021166 Canada Inc-Diego-Chacho", "Documents Accountant"],
  ["9359-6633 QUEBEC INC", "02-Document Accountant"],
  ["9390-9216 QUEBEC INC", "02-Documents accountant 9390-9216"],
  ["TAYANTI-CANADA", "01-Documents Accountant Tayanti"],
];

const MONTHS = [
  ...["8 August", "9 September", "10 October", "11 November", "12 December"].map((month) => [2026, month]),
  ...["1 January", "2 February", "3 March", "4 April", "5 May", "6 June", "7 July", "8 August", "9 September", "10 October", "11 November", "12 December"].map((month) => [2027, month]),
];

const dataSource = await readFile(new URL("../assets/js/data.js", import.meta.url), "utf8");
const expenseMatch = dataSource.match(/export const EXPENSE_CATEGORIES\s*=\s*(\[[\s\S]*?\]);/);
if (!expenseMatch) throw new Error("Could not read EXPENSE_CATEGORIES from assets/js/data.js.");
const EXPENSE_CATEGORIES = JSON.parse(expenseMatch[1].replace(/,\s*]/g, "]"));

const MONTH_TEMPLATE = [
  ...EXPENSE_CATEGORIES.map((name) => ["Expenses", name]),
  ["Expenses", "Card Purchases"],
  ["Expenses", "Tax Payments"],
  ["Expenses", "Reimbursements", "OP"],
  ["Expenses", "Reimbursements", "Reimbursements"],
  ["Income", "Invoices"],
  ["Income", "Transactions", "Cheque"],
  ["Income", "Transactions", "Cash"],
  ["Statements"],
  ["Documents", "Tax Remittances", "GST HST"],
  ["Documents", "Tax Remittances", "Payroll Remittances"],
  ["Documents", "Tax Remittances", "QST"],
  ["Documents", "Work Safety", "CNESST QC"],
  ["Documents", "Work Safety", "WCB NB"],
  ["Documents", "Work Safety", "WCB NS"],
  ["Documents", "Work Safety", "WCB PEI"],
  ["Documents", "Work Safety", "WSIB ON"],
  ["Documents", "Comite Paritario"],
  ["Documents", "Investment Land"],
  ["Documents", "Letters Received"],
  ["Documents", "Union"],
  ["Documents", "Xoom Investment"],
];

function normalizedName(value) {
  return String(value).normalize("NFKD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

async function directoryEntries(parent) {
  return readdir(parent, { withFileTypes: true });
}

function classifyName(entries, expected) {
  const exact = entries.find((entry) => entry.name.toLowerCase() === expected.toLowerCase());
  if (exact) return { status: exact.isDirectory() ? "EXACT" : "CONFLICT", actual: exact.name };
  const normalized = normalizedName(expected);
  const similar = entries.find((entry) => normalizedName(entry.name) === normalized);
  return similar ? { status: "CONFLICT", actual: similar.name } : { status: "MISSING", actual: "" };
}

async function assertDirectory(folderPath, label) {
  const info = await stat(folderPath);
  if (!info.isDirectory()) throw new Error(`${label} is not a directory: ${folderPath}`);
}

const results = [];
let createdYears = 0;
let createdMonths = 0;

await assertDirectory(GENERAL_ROOT, "General root");

for (const [company, accountantFolder] of APPROVED_ROOTS) {
  const companyPath = path.join(GENERAL_ROOT, company);
  const accountantPath = path.join(companyPath, accountantFolder);
  await assertDirectory(companyPath, `Company ${company}`);
  await assertDirectory(accountantPath, `Accountant root ${company}`);

  for (const [year, month] of MONTHS) {
    const yearName = String(year);
    let rootEntries = await directoryEntries(accountantPath);
    let yearState = classifyName(rootEntries, yearName);
    let yearPath = path.join(accountantPath, yearName);

    if (yearState.status === "CONFLICT") {
      results.push({ company, year, month, status: "SKIP_CONFLICT", reason: `Year collides with ${yearState.actual}` });
      continue;
    }

    if (yearState.status === "MISSING") {
      if (!APPLY) {
        results.push({ company, year, month, status: "CREATE", reason: "Year and month are absent" });
        continue;
      }
      rootEntries = await directoryEntries(accountantPath);
      yearState = classifyName(rootEntries, yearName);
      if (yearState.status === "CONFLICT") {
        results.push({ company, year, month, status: "SKIP_CONFLICT", reason: `Year appeared as ${yearState.actual}` });
        continue;
      }
      if (yearState.status === "MISSING") {
        try {
          await mkdir(yearPath);
          createdYears += 1;
        } catch (error) {
          if (error.code !== "EEXIST") throw error;
        }
      }
    }

    const yearEntries = await directoryEntries(yearPath);
    const monthState = classifyName(yearEntries, month);
    if (monthState.status === "EXACT") {
      results.push({ company, year, month, status: "SKIP_EXISTING", reason: monthState.actual });
      continue;
    }
    if (monthState.status === "CONFLICT") {
      results.push({ company, year, month, status: "SKIP_CONFLICT", reason: `Month collides with ${monthState.actual}` });
      continue;
    }
    if (!APPLY) {
      results.push({ company, year, month, status: "CREATE", reason: "Month is absent" });
      continue;
    }

    const freshEntries = await directoryEntries(yearPath);
    const freshState = classifyName(freshEntries, month);
    if (freshState.status !== "MISSING") {
      results.push({
        company,
        year,
        month,
        status: freshState.status === "EXACT" ? "SKIP_EXISTING" : "SKIP_CONFLICT",
        reason: `Month appeared during execution as ${freshState.actual}`,
      });
      continue;
    }

    const monthPath = path.join(yearPath, month);
    try {
      await mkdir(monthPath);
    } catch (error) {
      if (error.code === "EEXIST") {
        results.push({ company, year, month, status: "SKIP_CONFLICT", reason: "Month appeared during creation" });
        continue;
      }
      throw error;
    }

    for (const relativeParts of MONTH_TEMPLATE) {
      await mkdir(path.join(monthPath, ...relativeParts), { recursive: true });
    }
    createdMonths += 1;
    results.push({ company, year, month, status: "CREATED", reason: "Created and populated" });
  }
}

if (APPLY) {
  for (const result of results.filter((item) => item.status === "CREATED")) {
    const accountantFolder = APPROVED_ROOTS.find(([company]) => company === result.company)[1];
    const monthPath = path.join(GENERAL_ROOT, result.company, accountantFolder, String(result.year), result.month);
    await assertDirectory(monthPath, "Created month");
    for (const relativeParts of MONTH_TEMPLATE) {
      await assertDirectory(path.join(monthPath, ...relativeParts), `Template folder ${relativeParts.join("/")}`);
    }
  }
}

const summary = results.reduce((counts, item) => {
  counts[item.status] = (counts[item.status] || 0) + 1;
  return counts;
}, {});

console.log(JSON.stringify({ mode: APPLY ? "APPLY" : "DRY_RUN", approvedCompanies: APPROVED_ROOTS.length, createdYears, createdMonths, summary, results }, null, 2));
