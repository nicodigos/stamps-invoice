const fs = require("node:fs");
const path = require("node:path");
const http = require("node:http");

const ROOT_DIR = process.cwd();
const FUNCTIONS_DIR = path.join(ROOT_DIR, "netlify", "functions");
const DEFAULT_PORT = Number(process.env.PORT) || 8888;

loadEnvFile(path.join(ROOT_DIR, ".env"));

const MIME_TYPES = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
};

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url || "/", `http://${req.headers.host || `localhost:${DEFAULT_PORT}`}`);
    const pathname = decodeURIComponent(url.pathname);

    if (pathname.startsWith("/.netlify/functions/") || pathname.startsWith("/api/")) {
      await handleFunctionRequest(req, res, url, pathname);
      return;
    }

    await handleStaticRequest(res, pathname);
  } catch (error) {
    res.writeHead(500, { "Content-Type": "text/plain; charset=utf-8" });
    res.end(error.stack || error.message || "Internal server error");
  }
});

server.listen(DEFAULT_PORT, () => {
  console.log(`Dev server ready on http://localhost:${DEFAULT_PORT}`);
});

async function handleFunctionRequest(req, res, url, pathname) {
  const functionName = getFunctionName(pathname);
  const functionPath = path.join(FUNCTIONS_DIR, `${functionName}.js`);

  if (!isSafePath(functionPath, FUNCTIONS_DIR) || !fs.existsSync(functionPath)) {
    res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    res.end(`Function not found: ${functionName}`);
    return;
  }

  const requestBody = await readRequestBody(req);
  clearFunctionCache();
  const loaded = require(functionPath);
  const handler = loaded.handler || loaded.default || loaded;

  const event = {
    path: pathname,
    httpMethod: req.method || "GET",
    headers: normalizeHeaders(req.headers),
    queryStringParameters: Object.fromEntries(url.searchParams.entries()),
    rawUrl: url.toString(),
    body: requestBody,
    isBase64Encoded: false,
  };

  const result = await handler(event, {});
  const headers = result?.headers || {};
  const statusCode = Number(result?.statusCode) || 200;
  const body = typeof result?.body === "string" ? result.body : JSON.stringify(result?.body ?? "");

  res.writeHead(statusCode, headers);
  res.end(body);
}

function clearFunctionCache() {
  for (const cachedPath of Object.keys(require.cache)) {
    if (isSafePath(cachedPath, FUNCTIONS_DIR)) {
      delete require.cache[cachedPath];
    }
  }
}

async function handleStaticRequest(res, pathname) {
  const candidatePath = pathname === "/" ? "/index.html" : pathname;
  const filePath = path.join(ROOT_DIR, candidatePath);

  if (!isSafePath(filePath, ROOT_DIR) || !fs.existsSync(filePath) || fs.statSync(filePath).isDirectory()) {
    res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    res.end("Not found");
    return;
  }

  const extension = path.extname(filePath).toLowerCase();
  res.writeHead(200, { "Content-Type": MIME_TYPES[extension] || "application/octet-stream" });
  fs.createReadStream(filePath).pipe(res);
}

function getFunctionName(pathname) {
  if (pathname.startsWith("/.netlify/functions/")) {
    return pathname.slice("/.netlify/functions/".length).split("/")[0];
  }
  return pathname.slice("/api/".length).split("/")[0];
}

function normalizeHeaders(headers) {
  return Object.fromEntries(
    Object.entries(headers).map(([key, value]) => [key.toLowerCase(), Array.isArray(value) ? value.join(",") : String(value || "")]),
  );
}

function isSafePath(candidatePath, rootDir) {
  const resolvedCandidate = path.resolve(candidatePath);
  const resolvedRoot = path.resolve(rootDir);
  return resolvedCandidate === resolvedRoot || resolvedCandidate.startsWith(`${resolvedRoot}${path.sep}`);
}

function readRequestBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on("data", (chunk) => chunks.push(chunk));
    req.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    req.on("error", reject);
  });
}

function loadEnvFile(filePath) {
  if (!fs.existsSync(filePath)) return;

  const source = fs.readFileSync(filePath, "utf8");
  for (const line of source.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;

    const separatorIndex = trimmed.indexOf("=");
    if (separatorIndex === -1) continue;

    const key = trimmed.slice(0, separatorIndex).trim();
    let value = trimmed.slice(separatorIndex + 1).trim();

    if ((value.startsWith("\"") && value.endsWith("\"")) || (value.startsWith("'") && value.endsWith("'"))) {
      value = value.slice(1, -1);
    }

    value = value.replace(/\\n/g, "\n");
    process.env[key] = value;
  }
}
