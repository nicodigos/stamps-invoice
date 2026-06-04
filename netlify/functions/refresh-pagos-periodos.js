const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const mysql = require("mysql2/promise");

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const DEFAULT_WORKBOOK_PATH = "General/12433087 CANADA INC-MASTER/09-Pagos Periodos/2025/Building & Contractor Pay List/Building Address & Contractor Pay List.xlsx";
const CONTRACTOR_SHEET_NAME = "contractor list";
const CONTRACTOR_HEADER_NAME = "NAME";
const BUILDING_SHEET_NAME = "building list";
const BUILDING_HEADER_NAME = "BUILDING LIST";

exports.handler = async function handler(event) {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: "Metodo no permitido." };
  }

  try {
    const payload = JSON.parse(event.body || "{}");
    const accessToken = String(payload.accessToken || "").trim();
    if (!accessToken) {
      return { statusCode: 401, body: "Falta el token de Microsoft." };
    }

    const driveId = await resolveDriveId(accessToken);
    const workbookBytes = await downloadSharePointFile(workbookPath(), accessToken, driveId);
    const workbookData = parseWorkbook(workbookBytes);
    await replaceLookupTables(workbookData);

    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        contractors: workbookData.contractors.length,
        buildings: workbookData.buildings.length,
      }),
    };
  } catch (error) {
    return { statusCode: 500, body: error.message || String(error) };
  }
};

function workbookPath() {
  return (process.env.PAGOS_PERIODOS_WORKBOOK_PATH || DEFAULT_WORKBOOK_PATH).trim().replace(/^\/+|\/+$/g, "");
}

async function graphJson(url, token, options = {}) {
  const response = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      ...(options.headers || {}),
    },
  });
  if (!response.ok) {
    throw new Error(await response.text());
  }
  return response.json();
}

async function graphBytes(url, token) {
  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!response.ok) {
    throw new Error(await response.text());
  }
  return Buffer.from(await response.arrayBuffer());
}

async function resolveDriveId(token) {
  const spHostname = process.env.SP_HOSTNAME || "";
  const spSitePath = process.env.SP_SITE_PATH || "";
  const spDriveName = process.env.SP_DRIVE_NAME || "Documents";
  if (!spHostname || !spSitePath) {
    throw new Error("Falta SP_HOSTNAME o SP_SITE_PATH.");
  }

  const site = await graphJson(`${GRAPH_BASE}/sites/${spHostname}:${spSitePath}`, token);
  const drives = (await graphJson(`${GRAPH_BASE}/sites/${site.id}/drives`, token)).value || [];
  const drive = drives.find((item) => item.name === spDriveName) || drives[0];
  if (!drive) {
    throw new Error("No se pudo resolver el drive de SharePoint.");
  }
  return drive.id;
}

async function downloadSharePointFile(relativePath, token, driveId) {
  const encodedPath = encodeURIComponent(relativePath).replaceAll("%2F", "/");
  return graphBytes(`${GRAPH_BASE}/drives/${driveId}/root:/${encodedPath}:/content`, token);
}

function parseWorkbook(content) {
  const workbook = global.XLSX
    ? global.XLSX.read(content, { type: "buffer" })
    : requireXlsx().read(content, { type: "buffer" });

  return {
    contractors: columnValues(workbook, CONTRACTOR_SHEET_NAME, CONTRACTOR_HEADER_NAME),
    buildings: columnValues(workbook, BUILDING_SHEET_NAME, BUILDING_HEADER_NAME),
  };
}

function requireXlsx() {
  try {
    return require("xlsx");
  } catch {
    throw new Error("Falta la dependencia xlsx. Ejecuta npm install xlsx o agrega xlsx a package.json.");
  }
}

function columnValues(workbook, sheetName, headerName) {
  const sheet = findSheet(workbook, sheetName);
  const rows = requireXlsx().utils.sheet_to_json(sheet, { header: 1, defval: "" });
  let headerColumn = -1;
  let headerRow = -1;

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex];
    for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
      if (normalize(row[colIndex]) === normalize(headerName)) {
        headerColumn = colIndex;
        headerRow = rowIndex;
        break;
      }
    }
    if (headerColumn !== -1) break;
  }

  if (headerColumn === -1) {
    throw new Error(`No se encontro el header "${headerName}".`);
  }

  const seen = new Set();
  const values = [];
  for (let rowIndex = headerRow + 1; rowIndex < rows.length; rowIndex += 1) {
    const value = String(rows[rowIndex][headerColumn] ?? "").trim();
    if (!value || seen.has(value)) continue;
    seen.add(value);
    values.push(value);
  }
  return values.sort((left, right) => left.localeCompare(right));
}

function findSheet(workbook, sheetName) {
  const actualName = workbook.SheetNames.find((name) => normalize(name) === normalize(sheetName));
  if (!actualName) {
    throw new Error(`No se encontro la hoja "${sheetName}".`);
  }
  return workbook.Sheets[actualName];
}

function normalize(value) {
  return String(value ?? "").trim().toLowerCase().replace(/\s+/g, " ");
}

async function replaceLookupTables(data) {
  const connection = await mysql.createConnection(parseMysqlConfig());
  try {
    await connection.beginTransaction();
    await replaceLookupTable(connection, "pagos_periodos_contractors", data.contractors);
    await replaceLookupTable(connection, "pagos_periodos_buildings", data.buildings);
    await connection.commit();
  } catch (error) {
    await connection.rollback();
    throw error;
  } finally {
    await connection.end();
  }
}

async function replaceLookupTable(connection, tableName, values) {
  await connection.query(`DELETE FROM ${tableName}`);
  for (const value of values) {
    await connection.query(`INSERT INTO ${tableName} (name) VALUES (?)`, [value]);
  }
}

function parseMysqlConfig() {
  const localDsn = (process.env.LOCAL_MYSQL_DSN || "").trim();
  const env = (process.env.ENV || "development").trim();
  if (env === "development" && localDsn) {
    return parseLocalMysqlDsn(localDsn);
  }

  const host = (process.env.DB_HOST || "").trim();
  const port = Number(process.env.DB_PORT || "3306") || 3306;
  const user = (process.env.DB_USER || "").trim();
  const password = process.env.DB_PASS || "";
  const database = (process.env.DB_NAME || "").trim();
  if (!host || !user || !password || !database) {
    throw new Error("Faltan DB_HOST / DB_PORT / DB_USER / DB_PASS / DB_NAME.");
  }

  const config = {
    host,
    port,
    user,
    password,
    database,
    charset: "utf8mb4",
    multipleStatements: false,
  };

  const caPem = process.env.DO_MYSQL_CA_PEM || "";
  if (caPem.trim()) {
    const caPath = path.join(os.tmpdir(), "cnet_mysql_ca.pem");
    fs.writeFileSync(caPath, caPem, "utf8");
    config.ssl = { ca: fs.readFileSync(caPath, "utf8") };
  }

  return config;
}

function parseLocalMysqlDsn(dsn) {
  const [left, right] = dsn.split("@");
  if (!left || !right) throw new Error("LOCAL_MYSQL_DSN invalido.");

  const [user, password = ""] = left.split(":");
  const [protocolPart, dbPart = ""] = right.split("/");
  const database = dbPart.split("?")[0].trim();
  if (!user || !database) throw new Error("LOCAL_MYSQL_DSN invalido.");

  const config = {
    user: user.trim(),
    password,
    database,
    charset: "utf8mb4",
    multipleStatements: false,
  };

  const protocol = protocolPart.trim();
  const match = protocol.match(/^tcp\(([^:)]+)(?::(\d+))?\)$/);
  if (match) {
    config.host = match[1] || "127.0.0.1";
    config.port = Number(match[2] || "3306");
  } else {
    config.host = "127.0.0.1";
    config.port = 3306;
  }

  return config;
}
