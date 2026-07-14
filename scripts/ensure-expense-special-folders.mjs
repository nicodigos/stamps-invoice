import { mkdir, readdir, stat } from "node:fs/promises";
import path from "node:path";
import process from "node:process";

const GENERAL_ROOT = String.raw`C:\Users\Melissa Martinez\OneDrive - Group Castillo (12433087 Canada Inc)\Group Castillo Team Site - General`;
const APPLY = process.argv.includes("--apply");
const SPECIAL_FOLDERS = ["Card Purchases", "Tax Payments"];

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
  ...["7 July", "8 August", "9 September", "10 October", "11 November", "12 December"].map((month) => [2026, month]),
  ...["1 January", "2 February", "3 March", "4 April", "5 May", "6 June", "7 July", "8 August", "9 September", "10 October", "11 November", "12 December"].map((month) => [2027, month]),
];

function normalizedName(value) {
  return String(value).normalize("NFKD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function classify(entries, expected) {
  const exact = entries.find((entry) => entry.name.toLowerCase() === expected.toLowerCase());
  if (exact) return { status: exact.isDirectory() ? "EXACT" : "CONFLICT", actual: exact.name };
  const similar = entries.find((entry) => normalizedName(entry.name) === normalizedName(expected));
  return similar ? { status: "CONFLICT", actual: similar.name } : { status: "MISSING", actual: "" };
}

async function assertDirectory(folderPath, label) {
  const info = await stat(folderPath);
  if (!info.isDirectory()) throw new Error(`${label} is not a directory: ${folderPath}`);
}

async function inspectChild(parentPath, expected) {
  return classify(await readdir(parentPath, { withFileTypes: true }), expected);
}

async function ensureChild(parentPath, expected, createdPaths) {
  let state = await inspectChild(parentPath, expected);
  if (state.status === "EXACT") return { ok: true, path: path.join(parentPath, state.actual), created: false };
  if (state.status === "CONFLICT") return { ok: false, reason: `${expected} collides with ${state.actual}` };

  try {
    await mkdir(path.join(parentPath, expected));
    createdPaths.push(path.join(parentPath, expected));
    return { ok: true, path: path.join(parentPath, expected), created: true };
  } catch (error) {
    if (error.code !== "EEXIST") throw error;
  }

  state = await inspectChild(parentPath, expected);
  if (state.status === "EXACT") return { ok: true, path: path.join(parentPath, state.actual), created: false };
  return { ok: false, reason: `${expected} appeared with conflicting name ${state.actual || "unknown"}` };
}

const results = [];
const createdPaths = [];
await assertDirectory(GENERAL_ROOT, "General root");

for (const [company, accountantFolder] of APPROVED_ROOTS) {
  const accountantPath = path.join(GENERAL_ROOT, company, accountantFolder);
  await assertDirectory(accountantPath, `Accountant root ${company}`);

  for (const [year, month] of MONTHS) {
    const segments = [String(year), month, "Expenses"];
    let parentPath = accountantPath;
    let parentMissing = false;
    let conflict = "";

    for (const segment of segments) {
      if (parentMissing) continue;
      const state = await inspectChild(parentPath, segment);
      if (state.status === "CONFLICT") {
        conflict = `${segment} collides with ${state.actual}`;
        break;
      }
      if (state.status === "MISSING") {
        parentMissing = true;
      } else {
        parentPath = path.join(parentPath, state.actual);
      }
    }

    if (conflict) {
      for (const folder of SPECIAL_FOLDERS) results.push({ company, year, month, folder, status: "SKIP_CONFLICT", reason: conflict });
      continue;
    }

    if (!APPLY) {
      if (parentMissing) {
        for (const folder of SPECIAL_FOLDERS) results.push({ company, year, month, folder, status: "CREATE", reason: "Required parent path is absent" });
        continue;
      }
      for (const folder of SPECIAL_FOLDERS) {
        const state = await inspectChild(parentPath, folder);
        results.push({
          company,
          year,
          month,
          folder,
          status: state.status === "EXACT" ? "EXISTING" : state.status === "MISSING" ? "CREATE" : "SKIP_CONFLICT",
          reason: state.status === "CONFLICT" ? `${folder} collides with ${state.actual}` : state.actual,
        });
      }
      continue;
    }

    parentPath = accountantPath;
    let failedReason = "";
    for (const segment of segments) {
      const ensured = await ensureChild(parentPath, segment, createdPaths);
      if (!ensured.ok) {
        failedReason = ensured.reason;
        break;
      }
      parentPath = ensured.path;
    }
    if (failedReason) {
      for (const folder of SPECIAL_FOLDERS) results.push({ company, year, month, folder, status: "SKIP_CONFLICT", reason: failedReason });
      continue;
    }

    for (const folder of SPECIAL_FOLDERS) {
      const ensured = await ensureChild(parentPath, folder, createdPaths);
      results.push({
        company,
        year,
        month,
        folder,
        status: ensured.ok ? (ensured.created ? "CREATED" : "EXISTING") : "SKIP_CONFLICT",
        reason: ensured.ok ? "" : ensured.reason,
      });
    }
  }
}

if (APPLY) {
  for (const [company, accountantFolder] of APPROVED_ROOTS) {
    for (const [year, month] of MONTHS) {
      for (const folder of SPECIAL_FOLDERS) {
        await assertDirectory(path.join(GENERAL_ROOT, company, accountantFolder, String(year), month, "Expenses", folder), `${company} ${year} ${month} ${folder}`);
      }
    }
  }
}

const summary = results.reduce((counts, result) => {
  counts[result.status] = (counts[result.status] || 0) + 1;
  return counts;
}, {});

console.log(JSON.stringify({
  mode: APPLY ? "APPLY" : "DRY_RUN",
  approvedCompanies: APPROVED_ROOTS.length,
  expectedMonths: APPROVED_ROOTS.length * MONTHS.length,
  createdPathCount: createdPaths.length,
  summary,
  results,
}, null, 2));
