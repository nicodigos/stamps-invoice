const { webcrypto } = require("node:crypto");

exports.handler = async function handler(event) {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: "Method not allowed" };
  }

  const openAiKey = process.env.OPENAI_API_KEY || "";
  const googleJson = process.env.GOOGLE_SERVICE_ACCOUNT_JSON || "";
  if (!openAiKey || !googleJson) {
    return { statusCode: 500, body: "Missing OPENAI_API_KEY or GOOGLE_SERVICE_ACCOUNT_JSON." };
  }

  try {
    const payload = JSON.parse(event.body || "{}");
    const documents = Array.isArray(payload.documents) ? payload.documents : [];
    const options = normalizeOptions(payload.options);
    if (!documents.length) {
      return { statusCode: 400, body: "No documents to analyze." };
    }

    const accessToken = await getGoogleAccessToken(parseJsonInput(googleJson, "GOOGLE_SERVICE_ACCOUNT_JSON"));
    const visionPages = [];
    for (const document of documents) {
      for (const image of document.images || []) {
        const vision = await extractTextWithVision(image.imageBase64, accessToken);
        visionPages.push({
          sourceName: image.sourceName || document.sourceName || "",
          pageNumber: image.pageNumber || 1,
          colorHint: image.colorHint || null,
          fullText: vision.full_text || "",
        });
      }
    }

    if (!visionPages.length) {
      return { statusCode: 400, body: "No rendered images found for Vision." };
    }

    const emailContexts = documents
      .filter((document) => document.email)
      .map((document) => ({
        sourceName: document.sourceName || "",
        ...document.email,
      }));
    const gpt = await classifyStampFields({ emailContexts, visionPages, options, apiKey: openAiKey });
    const normalized = normalizeClassification(gpt, options);

    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(normalized),
    };
  } catch (error) {
    return { statusCode: 500, body: error.message || String(error) };
  }
};

function normalizeOptions(options) {
  return {
    companies: Array.isArray(options?.companies) ? options.companies.map(String) : [],
    banks: Array.isArray(options?.banks) ? options.banks.map(String) : [],
    categories: Array.isArray(options?.categories) ? options.categories.map(String) : [],
    folderStructure: options?.folderStructure === "legacy" ? "legacy" : "new",
    folderRoot: normalizeFolderRoot(options?.folderRoot, options?.folderStructure),
    newStructureCategoryPaths: normalizeCategoryPaths(options?.newStructureCategoryPaths),
  };
}

function normalizeFolderRoot(value, folderStructure) {
  if (folderStructure === "legacy") return "Expenses";
  const root = String(value || "").trim();
  return ["Documents", "Expenses", "Income", "Statements"].includes(root) ? root : "";
}

function normalizeCategoryPaths(value) {
  if (!value || typeof value !== "object") return {};
  return Object.fromEntries(Object.entries(value).map(([key, path]) => [
    String(key),
    Array.isArray(path) ? path.map(String) : [],
  ]));
}

async function classifyStampFields({ emailContexts, visionPages, options, apiKey }) {
  const workflowInstructions = workflowPrompt(options.folderRoot);
  const activeFields = workflowFields(options.folderRoot);
  const systemPrompt = [
    "You fill an invoice stamping form from email context and OCR.",
    "Return JSON only. Choose the closest available dropdown option even when text is not exact.",
    "Human will validate after you fill the form, so provide best suggestions rather than blanks when there is useful evidence.",
    "Only fill fields for the selected folder workflow. Use empty string and 0 confidence for every field outside the selected workflow.",
    "Description must be at most 300 characters. Suggested title must be short and filename-friendly.",
    "Use visual color evidence only after deciding a page/document is bank-related, such as a payment confirmation, transfer, deposit, bank transaction, or account document.",
    "For bank-related pages, colorHint maps as: light green means Desjardins, light blue means National Bank, red means Scotia Bank.",
    "Ignore colorHint for invoices, bills, receipts, supplier documents, or any page that does not appear to be from a bank.",
    workflowInstructions,
  ].join(" ");

  const userPrompt = [
    `Selected workflow fields to fill: ${activeFields.join(", ")}.`,
    "All other schema fields must be empty strings with confidence 0.",
    `Company options: ${options.companies.join(" | ")}`,
    `Category options: ${options.categories.join(" | ")}`,
    `Bank options: ${options.banks.join(" | ")}`,
    "Account type options: debit | credit",
    "Income payment_type options: Cheque | Cash",
    `Folder structure mode: ${options.folderStructure}.`,
    `Selected folder root: ${options.folderRoot || "none"}.`,
    options.folderStructure === "new"
      ? `Validate category against these generated-folder destinations: ${JSON.stringify(options.newStructureCategoryPaths).slice(0, 5000)}`
      : "Legacy folder structure is active; use the legacy category names directly.",
    "Important mappings: T4A or PAGO T4A should map to 5 T4A Payments. 1001298527 ONTARIO should map to 1001298527 ONTARIO INC. Scotia should map to Scotia Bank.",
    "date must be YYYY-MM-DD when possible.",
    "First decide from OCR/email whether each page is bank-related or invoice-related. If bank text is weak or absent on a bank-related page, use first-page colorHint to choose the closest bank option. Do not use colorHint from non-bank pages.",
    "Return confidence 0-1 for each field.",
    `Email context:\n${JSON.stringify(emailContexts).slice(0, 10000)}`,
    `Vision OCR pages:\n${JSON.stringify(visionPages.map((page) => ({ ...page, fullText: String(page.fullText || "").slice(0, 9000) }))).slice(0, 26000)}`,
  ].join("\n\n");

  const response = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: "gpt-4o-mini",
      temperature: 0,
      response_format: {
        type: "json_schema",
        json_schema: {
          name: "stamp_autofill",
          strict: true,
          schema: {
            type: "object",
            additionalProperties: false,
            properties: {
              date: { type: "string" },
              company: { type: "string" },
              category: { type: "string" },
              bank: { type: "string" },
              account_type: { type: "string" },
              account_last_four: { type: "string" },
              payment_code: { type: "string" },
              client_invoice: { type: "string" },
              suggested_title: { type: "string" },
              description: { type: "string" },
              amount: { type: "string" },
              payment_type: { type: "string" },
              notes: { type: "string" },
              confidence: {
                type: "object",
                additionalProperties: false,
                properties: {
                  date: { type: "number" },
                  company: { type: "number" },
                  category: { type: "number" },
                  bank: { type: "number" },
                  account_type: { type: "number" },
                  account_last_four: { type: "number" },
                  payment_code: { type: "number" },
                  client_invoice: { type: "number" },
                  suggested_title: { type: "number" },
                  description: { type: "number" },
                  amount: { type: "number" },
                  payment_type: { type: "number" },
                },
                required: ["date", "company", "category", "bank", "account_type", "account_last_four", "payment_code", "client_invoice", "suggested_title", "description", "amount", "payment_type"],
              },
            },
            required: ["date", "company", "category", "bank", "account_type", "account_last_four", "payment_code", "client_invoice", "suggested_title", "description", "amount", "payment_type", "notes", "confidence"],
          },
        },
      },
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: userPrompt },
      ],
    }),
  });
  const payload = await response.json();
  if (!response.ok) {
    throw new Error(`OpenAI error: ${JSON.stringify(payload)}`);
  }
  return JSON.parse(payload.choices?.[0]?.message?.content || "{}");
}

function workflowFields(folderRoot) {
  if (folderRoot === "Documents") return ["date", "company", "category", "suggested_title", "description"];
  if (folderRoot === "Income") return ["date", "company", "category", "amount", "payment_type"];
  if (folderRoot === "Statements") return ["date", "company", "suggested_title"];
  return ["date", "company", "category", "bank", "account_type", "payment_code", "client_invoice"];
}

function workflowPrompt(folderRoot) {
  if (folderRoot === "Documents") {
    return [
      "Documents workflow:",
      "Detect date, company, and category from document OCR first; use email context only as fallback.",
      "Generate description as a plain summary of what the document is, maximum 300 characters.",
      "Generate suggested_title as a short title including sender/remitter and a quick idea of the document.",
      "Do not fill Expenses, Income, or Statements fields.",
    ].join(" ");
  }

  if (folderRoot === "Income") {
    return [
      "Income workflow:",
      "Only try to identify company, category, amount, and payment_type.",
      "Use the selected Date field for filing; if a clear income document date is visible, return it as date.",
      "Amount should be the best visible received/deposit/invoice total as a short readable string.",
      "Payment_type must be Cheque or Cash. Use Cheque only when cheque/check payment evidence is visible; use Cash only when cash evidence is visible; otherwise leave empty with low confidence.",
      "Do not fill Documents, Expenses, or Statements fields.",
    ].join(" ");
  }

  if (folderRoot === "Statements") {
    return [
      "Statements workflow:",
      "Only try to identify date, company, and suggested_title.",
      "Suggested_title should be a short appropriate filename title for the statement.",
      "Do not fill Documents, Expenses, or Income fields.",
    ].join(" ");
  }

  return [
    "Expenses workflow:",
    "Best-effort identify category and company.",
    "The key fields are payment_code, client_invoice, and company; fill those carefully from visible payment, invoice, supplier, reference, or period evidence.",
    "Detect bank and account_type only with good evidence; otherwise use empty string and low confidence.",
    "Never fill account_last_four. Always return account_last_four as an empty string with confidence 0; the user will enter it manually if needed.",
    "Keep using colorHint to infer bank only when the page appears bank-related.",
    "Do not fill Documents, Income, or Statements fields.",
  ].join(" ");
}

function normalizeClassification(raw, options) {
  const confidence = normalizeConfidence(raw.confidence);
  const out = {
    date: normalizeDate(raw.date),
    company: closestOption(raw.company, options.companies),
    category: closestOption(raw.category, options.categories),
    bank: confidence.bank >= 0.65 ? closestOption(raw.bank, options.banks, { allowEmpty: true }) : "",
    account_type: confidence.account_type >= 0.65 ? closestOption(raw.account_type, ["debit", "credit"], { allowEmpty: true }) : "",
    account_last_four: "",
    payment_code: sanitizeField(raw.payment_code),
    client_invoice: sanitizeField(raw.client_invoice),
    suggested_title: sanitizeField(raw.suggested_title),
    description: sanitizeDescription(raw.description),
    amount: sanitizeAmount(raw.amount),
    payment_type: confidence.payment_type >= 0.65 ? closestOption(raw.payment_type, ["Cheque", "Cash"], { allowEmpty: true }) : "",
    notes: String(raw.notes || "").trim().split(/\s+/).slice(0, 30).join(" "),
    confidence,
  };

  out.lowConfidence = Object.entries(confidence)
    .filter(([, value]) => value > 0 && value < 0.65)
    .map(([key]) => key);
  return out;
}

function normalizeDate(value) {
  const text = String(value || "").trim();
  const direct = text.match(/\b(20\d{2})[-/](0[1-9]|1[0-2])[-/](0[1-9]|[12]\d|3[01])\b/);
  if (direct) return `${direct[1]}-${direct[2]}-${direct[3]}`;
  return "";
}

function sanitizeField(value) {
  return String(value || "").trim().replace(/\s+/g, " ").slice(0, 80);
}

function sanitizeDescription(value) {
  return String(value || "").trim().replace(/\s+/g, " ").slice(0, 300);
}

function sanitizeAmount(value) {
  return String(value || "").trim().replace(/\s+/g, " ").slice(0, 40);
}

function normalizeConfidence(value) {
  const fields = ["date", "company", "category", "bank", "account_type", "account_last_four", "payment_code", "client_invoice", "suggested_title", "description", "amount", "payment_type"];
  return Object.fromEntries(fields.map((field) => [field, clamp01(Number(value?.[field]) || 0)]));
}

function clamp01(value) {
  return Math.max(0, Math.min(1, value));
}

function closestOption(value, options, settings = {}) {
  const raw = String(value || "").trim();
  if (!raw && settings.allowEmpty) return "";
  if (!options.length) return raw;
  const exact = options.find((option) => option === raw);
  if (exact) return exact;

  const normalizedRaw = normalizeText(raw);
  if (!normalizedRaw && settings.allowEmpty) return "";

  let best = options[0];
  let bestScore = -Infinity;
  for (const option of options) {
    const score = similarityScore(normalizedRaw, normalizeText(option));
    if (score > bestScore) {
      best = option;
      bestScore = score;
    }
  }
  return best;
}

function similarityScore(left, right) {
  if (!left || !right) return 0;
  if (left === right) return 100;
  let score = 0;
  if (left.includes(right) || right.includes(left)) score += 60;
  const leftTokens = new Set(left.split(" ").filter(Boolean));
  const rightTokens = new Set(right.split(" ").filter(Boolean));
  for (const token of leftTokens) {
    if (rightTokens.has(token)) score += 16;
    else if ([...rightTokens].some((rightToken) => rightToken.includes(token) || token.includes(rightToken))) score += 8;
  }
  score -= levenshtein(left, right) / Math.max(left.length, right.length);
  return score;
}

function normalizeText(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/&/g, " and ")
    .replace(/[^\w\s]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function levenshtein(left, right) {
  const dp = Array.from({ length: left.length + 1 }, () => Array(right.length + 1).fill(0));
  for (let i = 0; i <= left.length; i += 1) dp[i][0] = i;
  for (let j = 0; j <= right.length; j += 1) dp[0][j] = j;
  for (let i = 1; i <= left.length; i += 1) {
    for (let j = 1; j <= right.length; j += 1) {
      dp[i][j] = Math.min(
        dp[i - 1][j] + 1,
        dp[i][j - 1] + 1,
        dp[i - 1][j - 1] + (left[i - 1] === right[j - 1] ? 0 : 1),
      );
    }
  }
  return dp[left.length][right.length];
}

async function getGoogleAccessToken(serviceAccount) {
  const now = Math.floor(Date.now() / 1000);
  const jwtHeader = base64UrlEncode(JSON.stringify({ alg: "RS256", typ: "JWT" }));
  const jwtClaim = base64UrlEncode(JSON.stringify({
    iss: serviceAccount.client_email,
    scope: "https://www.googleapis.com/auth/cloud-platform",
    aud: "https://oauth2.googleapis.com/token",
    exp: now + 3600,
    iat: now,
  }));
  const unsignedJwt = `${jwtHeader}.${jwtClaim}`;
  const key = await webcrypto.subtle.importKey(
    "pkcs8",
    pemToArrayBuffer(serviceAccount.private_key),
    { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" },
    false,
    ["sign"],
  );
  const signature = await webcrypto.subtle.sign("RSASSA-PKCS1-v1_5", key, new TextEncoder().encode(unsignedJwt));
  const jwt = `${unsignedJwt}.${base64UrlEncode(signature)}`;

  const response = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: jwt,
    }),
  });
  const payload = await response.json();
  if (!response.ok || !payload.access_token) {
    throw new Error(`Google token error: ${JSON.stringify(payload)}`);
  }
  return payload.access_token;
}

async function extractTextWithVision(imageBase64, accessToken) {
  const response = await fetch("https://vision.googleapis.com/v1/images:annotate", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      requests: [
        {
          image: { content: imageBase64 },
          features: [{ type: "DOCUMENT_TEXT_DETECTION" }],
        },
      ],
    }),
  });
  const payload = await response.json();
  if (!response.ok) {
    throw new Error(`Google Vision error: ${JSON.stringify(payload)}`);
  }
  return { full_text: payload.responses?.[0]?.fullTextAnnotation?.text || "" };
}

function parseJsonInput(text, label) {
  const rawText = String(text || "");
  try {
    const parsed = JSON.parse(rawText);
    return normalizeParsedJson(parsed, label);
  } catch (error) {
    const sanitizedText = sanitizeJsonText(rawText);
    if (sanitizedText !== rawText) {
      try {
        return normalizeParsedJson(JSON.parse(sanitizedText), label);
      } catch {
        // Fall through to the original parse error.
      }
    }
    throw new Error(`${label} no contiene JSON valido: ${error.message}`);
  }
}

function normalizeParsedJson(parsed, label) {
  if (label === "GOOGLE_SERVICE_ACCOUNT_JSON" && parsed && typeof parsed === "object") {
    return { ...parsed, private_key: String(parsed.private_key || "").replace(/\\n/g, "\n") };
  }
  return parsed;
}

function sanitizeJsonText(text) {
  let result = "";
  let inString = false;
  let escaped = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (inString) {
      if (escaped) {
        result += char;
        escaped = false;
        continue;
      }
      if (char === "\\") {
        result += char;
        escaped = true;
        continue;
      }
      if (char === "\"") {
        result += char;
        inString = false;
        continue;
      }
      if (char === "\n") {
        result += "\\n";
        continue;
      }
      if (char === "\r") {
        result += "\\r";
        continue;
      }
      if (char === "\t") {
        result += "\\t";
        continue;
      }
      result += char;
      continue;
    }

    result += char;
    if (char === "\"") {
      inString = true;
    }
  }

  return result;
}

function pemToArrayBuffer(pem) {
  const base64 = String(pem || "").replace(/-----[^-]+-----/g, "").replace(/\s+/g, "");
  const bytes = Buffer.from(base64, "base64");
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function base64UrlEncode(input) {
  const bytes = typeof input === "string" ? Buffer.from(input, "utf8") : Buffer.from(input);
  return bytes.toString("base64url");
}
