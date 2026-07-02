const ALLOWED_CATEGORIES = [
  "Labor T4A",
  "Supplies",
  "Equipment",
  "Repairs",
  "Warehouse Rent",
  "Office Rent (Colonnade 219)",
  "Office utilities",
  "DN Fees",
  "DN Supplies",
  "DN Early Payment",
  "DN Others",
  "Early Discount",
  "Loan Installment",
  "Franchise Loan",
  "Tax Payments",
  "Master Franchise Fees",
  "Diverse Expenses",
  "Labor Franchise",
  "Labor Payroll",
  "Management fees",
  "Reimbousement",
  "Professional fees",
  "Car Allowance",
  "Credit Card Payment",
  "Labor Subs Invoice",
  "Rental Equipment",
  "Insurance",
  "Leasing",
  "Inmatriculation",
  "Automobile Repair & Maintenance",
  "Autmobile Fuel & Oil",
  "Cleaning Supplies",
  "Office Supplies",
  "Business Parking",
  "Meals",
  "Tools & Equipment",
  "Gifts",
  "Pago TC Visa Desjardins",
];
const ALLOWED_CATEGORY_SET = new Set(ALLOWED_CATEGORIES);
const PROVINCE_CODES = new Set(["AB", "BC", "MB", "NB", "NL", "NS", "NT", "NU", "ON", "PE", "QC", "SK", "YT"]);
const TAX_FIELD_TOKENS = {
  gst: ["GST"],
  hst: ["HST"],
  pst: ["PST", "RST"],
  qst: ["QST", "TVQ"],
  tps: ["TPS"],
  iva: ["IVA"],
  vat: ["VAT"],
  retention: ["RETENTION", "RETENCION", "RETENCIÓN", "WITHHOLDING", "WH TAX"],
};

exports.handler = async function handler(event) {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: "Method not allowed" };
  }

  const openAiKey = process.env.OPENAI_API_KEY || "";
  const googleJson = process.env.GOOGLE_SERVICE_ACCOUNT_JSON || "";
  if (!openAiKey || !googleJson) {
    return { statusCode: 500, body: "Falta OPENAI_API_KEY o GOOGLE_SERVICE_ACCOUNT_JSON." };
  }

  let payload;
  try {
    payload = parseJsonInput(event.body || "{}", "event.body");
  } catch (error) {
    return { statusCode: 400, body: error.message };
  }
  const imageBase64 = payload.imageBase64 || "";
  const receiptType = String(payload.receiptType || "bank_transaction").trim().toLowerCase();
  if (!imageBase64) {
    return { statusCode: 400, body: "Falta imageBase64." };
  }

  try {
    const accessToken = await getGoogleAccessToken(parseJsonInput(googleJson, "GOOGLE_SERVICE_ACCOUNT_JSON"));
    const vision = await extractTextWithVision(imageBase64, accessToken);
    const compact = parseReceiptFromText(vision.full_text || "");
    const gpt = await classifyWithGpt(vision, compact, receiptType, openAiKey);
    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ vision, compact, gpt }),
    };
  } catch (error) {
    return { statusCode: 500, body: error.message };
  }
};

function parseJsonInput(text, label) {
  const rawText = String(text || "");
  try {
    const parsed = JSON.parse(rawText);
    return normalizeParsedJson(parsed, label);
  } catch (error) {
    const sanitizedText = sanitizeJsonText(rawText);
    if (sanitizedText !== rawText) {
      try {
        const parsed = JSON.parse(sanitizedText);
        return normalizeParsedJson(parsed, label);
      } catch {
        // Fall through to the original error message below.
      }
    }
    throw new Error(`${label} no contiene JSON valido: ${error.message}`);
  }
}

function normalizeParsedJson(parsed, label) {
  if (label === "GOOGLE_SERVICE_ACCOUNT_JSON" && parsed && typeof parsed === "object") {
    return {
      ...parsed,
      private_key: String(parsed.private_key || "").replace(/\\n/g, "\n"),
    };
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
  const fullText = payload.responses?.[0]?.fullTextAnnotation?.text || "";
  return { full_text: fullText };
}

function parseReceiptFromText(fullText) {
  const lines = String(fullText || "").split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  const merchant = lines[0] || "";

  let date = "";
  const dateMatch = String(fullText || "").match(/\b(20\d{2}[-/](0[1-9]|1[0-2])[-/](0[1-9]|[12]\d|3[01]))(?:[ T]([01]\d|2[0-3]):[0-5]\d(?::[0-5]\d)?)?\b/);
  if (dateMatch) {
    const base = dateMatch[1].replaceAll("/", "-");
    const timePart = dateMatch[4] ? `${dateMatch[4]}${dateMatch[4].length === 5 ? ":00" : ""}` : "00:00:00";
    date = `${base}T${timePart}`;
  }

  let total = 0;
  let taxesTotal = 0;
  const taxBreakdown = createEmptyTaxBreakdown();
  for (const line of lines) {
    const upper = line.toUpperCase();
    const amount = extractAmountFromLine(line);
    if (amount <= 0) continue;
    if (upper.includes("TOTAL") && !upper.includes("SUBTOTAL")) total = Math.max(total, amount);
    const matchedTaxFields = detectTaxFields(upper);
    if (matchedTaxFields.length) {
      taxesTotal += amount;
      matchedTaxFields.forEach((field) => {
        taxBreakdown[field] += amount;
      });
    }
  }

  if (total === 0) {
    for (const line of lines) {
      total = Math.max(total, extractAmountFromLine(line));
    }
  }

  let city = "";
  let province = "";
  for (const line of lines) {
    const match = line.toUpperCase().match(/\b([A-Z]{2})\b/);
    if (match && PROVINCE_CODES.has(match[1])) {
      province = match[1];
      city = line.slice(0, match.index).trim().replace(/[ ,.-]+$/g, "").split(",").pop()?.trim() || "";
      break;
    }
  }

  return {
    merchant,
    date,
    total,
    taxes_total: taxesTotal,
    ...taxBreakdown,
    city,
    province,
    address: "",
    items: [],
  };
}

function extractAmountFromLine(line) {
  const matches = String(line || "").match(/(?<!\d)(\d{1,3}(?:,\d{3})*(?:\.\d{2})|\d+\.\d{2})(?!\d)/g) || [];
  return Math.max(0, ...matches.map((value) => Number(value.replaceAll(",", ""))).filter(Number.isFinite));
}

async function classifyWithGpt(vision, compact, receiptType, apiKey) {
  const today = new Date();
  const currentYear = today.getFullYear();
  const previousYear = currentYear - 1;
  const systemPrompt = [
    "You are extracting fields from an invoice OCR result.",
    "Return JSON only with keys: payment_date, total_amount, taxes_total, gst, hst, pst, qst, tps, iva, vat, retention, category, confidence, merchant_name, city, province, ticket_number, notes.",
    `Category must be one of: ${ALLOWED_CATEGORIES.join(", ")}.`,
    "payment_date must be ISO datetime when possible. Monetary fields must be numbers.",
    "Use 0 for any tax field not present. taxes_total should equal the sum of detected taxes when possible.",
    "Always choose the closest allowed category based on merchant, line items, and receipt context, not just one keyword.",
    "Do not default to Diverse Expenses when another allowed category is even moderately more specific and plausible.",
    "Use Diverse Expenses only as a last resort when no other allowed category reasonably fits the receipt.",
    "Distinguish carefully between Supplies, Office Supplies, Cleaning Supplies, Equipment, Tools & Equipment, Repairs, Rental Equipment, Meals, Business Parking, Insurance, Leasing, Professional fees, Tax Payments, and Credit Card Payment.",
  ].join(" ");

  const baseUserPrompt = [
    "Esto viene de un invoice. Quiero que llenes estos campos: payment_date, total_amount, taxes_total, gst, hst, pst, qst, tps, iva, vat, retention, category, merchant_name, city, province, ticket_number.",
    `El tipo de recibo elegido en el formulario es: ${receiptType}.`,
    "Prefer OCR evidence. Even if uncertain, choose the best category from the allowed list instead of defaulting generically.",
    "Use Diverse Expenses only if the receipt genuinely does not fit any more specific allowed category.",
    "Map each specific tax into its own field: GST, HST, PST, QST/TVQ, TPS, IVA, VAT, retention/withholding. Leave missing fields as 0.",
    "ticket_number should be the receipt number, ticket number, invoice number, folio, reference, transaction id, or similar document identifier when visible. If none is visible, return an empty string.",
    "Choose the best category even if several seem plausible. Use the merchant name, line item words, and overall receipt purpose together.",
    "Use Credit Card Payment only for a card statement or payment receipt, not for ordinary purchases paid by card.",
    "Use Tax Payments for remittances and government tax payments, not normal purchases that merely include taxes.",
    "Use Office utilities only for utilities/services tied to the office. Use Warehouse Rent or Office Rent (Colonnade 219) only when the receipt clearly corresponds to rent.",
    "Use Equipment for purchased equipment, Tools & Equipment for tools and small operational equipment, and Rental Equipment only when the document is a rental.",
    "Use Automobile Repair & Maintenance for vehicle service/repair and Autmobile Fuel & Oil for fuel, gas, oil, or charging-like fuel usage.",
    "Notes <=20 words.",
    `For payment_date, the invoice year is very likely ${currentYear}; depending on how close the receipt is to the start of the year, it may be ${previousYear} instead.`,
    `Google Vision OCR JSON:\n${JSON.stringify({ ...vision, full_text: String(vision.full_text || "").slice(0, 12000) })}`,
    `Parsed helper fields:\n${JSON.stringify(compact)}`,
  ].join("\n\n");
  const passInstructions = [
    "Pass 1 of 3. Produce your best full extraction and category selection from the OCR.",
    "Pass 2 of 3. Re-check every field critically. Challenge the category choice and switch it if another allowed category fits better.",
    "Pass 3 of 3. Audit the whole receipt again. Prioritize the most defensible final category and totals, even if they differ from an earlier instinct.",
  ];
  const candidates = [];
  for (const instruction of passInstructions) {
    candidates.push(await runGptReceiptPass(systemPrompt, `${baseUserPrompt}\n\n${instruction}`, apiKey));
  }
  return selectBestGptCandidate(candidates, compact, receiptType);
}

function safeParseModelJson(content) {
  const text = String(content || "").trim();
  const direct = tryParseJson(text);
  if (direct) return direct;

  const withoutFences = text.replace(/^```(?:json)?\s*/i, "").replace(/\s*```$/i, "").trim();
  const fenced = tryParseJson(withoutFences);
  if (fenced) return fenced;

  const objectSlice = extractFirstJsonObject(withoutFences);
  const sliced = tryParseJson(objectSlice);
  if (sliced) return sliced;

  throw new Error(`No se pudo parsear JSON del modelo: ${text.slice(0, 160)}`);
}

function tryParseJson(text) {
  if (!text) return null;
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

function extractFirstJsonObject(text) {
  const start = text.indexOf("{");
  if (start === -1) return "";
  let depth = 0;
  let inString = false;
  let escaped = false;
  for (let index = start; index < text.length; index += 1) {
    const char = text[index];
    if (inString) {
      if (escaped) {
        escaped = false;
      } else if (char === "\\") {
        escaped = true;
      } else if (char === "\"") {
        inString = false;
      }
      continue;
    }
    if (char === "\"") {
      inString = true;
      continue;
    }
    if (char === "{") depth += 1;
    if (char === "}") {
      depth -= 1;
      if (depth === 0) {
        return text.slice(start, index + 1);
      }
    }
  }
  return "";
}

function createEmptyTaxBreakdown() {
  return {
    gst: 0,
    hst: 0,
    pst: 0,
    qst: 0,
    tps: 0,
    iva: 0,
    vat: 0,
    retention: 0,
  };
}

function detectTaxFields(upperLine) {
  return Object.entries(TAX_FIELD_TOKENS)
    .filter(([, tokens]) => tokens.some((token) => upperLine.includes(token)))
    .map(([field]) => field);
}

function normalizeTaxBreakdown(source) {
  const empty = createEmptyTaxBreakdown();
  for (const key of Object.keys(empty)) {
    empty[key] = Number(source?.[key]) || 0;
  }
  return empty;
}

async function runGptReceiptPass(systemPrompt, userPrompt, apiKey) {
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
          name: "receipt_extraction",
          strict: true,
          schema: {
            type: "object",
            additionalProperties: false,
            properties: {
              payment_date: { type: "string" },
              total_amount: { type: "number" },
              taxes_total: { type: "number" },
              gst: { type: "number" },
              hst: { type: "number" },
              pst: { type: "number" },
              qst: { type: "number" },
              tps: { type: "number" },
              iva: { type: "number" },
              vat: { type: "number" },
              retention: { type: "number" },
              category: { type: "string" },
              confidence: { type: "number" },
              merchant_name: { type: "string" },
              city: { type: "string" },
              province: { type: "string" },
              ticket_number: { type: "string" },
              notes: { type: "string" },
            },
            required: ["payment_date", "total_amount", "taxes_total", "gst", "hst", "pst", "qst", "tps", "iva", "vat", "retention", "category", "confidence", "merchant_name", "city", "province", "ticket_number", "notes"],
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

  const content = payload.choices?.[0]?.message?.content || "{}";
  return normalizeModelCandidate(safeParseModelJson(content));
}

function normalizeModelCandidate(parsed) {
  const category = normalizeCategoryGuess(parsed.category);
  const taxBreakdown = normalizeTaxBreakdown(parsed);
  const normalizedTaxesTotal = Number(parsed.taxes_total) || Object.values(taxBreakdown).reduce((sum, value) => sum + value, 0);
  return {
    payment_date: String(parsed.payment_date || "").trim(),
    total_amount: Number(parsed.total_amount) || 0,
    taxes_total: normalizedTaxesTotal,
    ...taxBreakdown,
    category,
    confidence: Math.max(0, Math.min(1, Number(parsed.confidence) || 0)),
    merchant_name: String(parsed.merchant_name || "").trim(),
    city: String(parsed.city || "").trim(),
    province: String(parsed.province || "").trim(),
    ticket_number: String(parsed.ticket_number || "").trim(),
    notes: String(parsed.notes || "").trim().split(/\s+/).slice(0, 20).join(" "),
  };
}

function selectBestGptCandidate(candidates, compact, receiptType) {
  const categoryVotes = new Map();
  candidates.forEach((candidate) => {
    const current = categoryVotes.get(candidate.category) || { count: 0, confidence: 0 };
    current.count += 1;
    current.confidence += candidate.confidence;
    categoryVotes.set(candidate.category, current);
  });

  const rankedCandidates = candidates
    .map((candidate, index) => ({
      candidate,
      index,
      score: scoreGptCandidate(candidate, compact, categoryVotes.get(candidate.category)),
    }))
    .sort((left, right) => right.score - left.score || right.candidate.confidence - left.candidate.confidence || left.index - right.index);

  const best = rankedCandidates[0]?.candidate || normalizeModelCandidate({});
  return {
    receipt_type: receiptType,
    ...best,
  };
}

function scoreGptCandidate(candidate, compact, categoryVote) {
  const taxFields = Object.keys(createEmptyTaxBreakdown());
  let score = candidate.confidence * 100;
  if (candidate.payment_date) score += 8;
  if (candidate.merchant_name) score += 8;
  if (candidate.city) score += 3;
  if (candidate.province) score += 3;
  if (candidate.ticket_number) score += 4;
  if (candidate.total_amount > 0) score += 12;
  if (candidate.taxes_total >= 0) score += 4;
  if (candidate.notes) score += 2;
  if (candidate.category !== "Diverse Expenses") {
    score += 10;
  } else {
    score -= 12;
  }
  if (categoryVote) {
    score += categoryVote.count * 10;
    score += categoryVote.confidence * 5;
  }

  if (compact?.merchant && candidate.merchant_name && stringsRoughlyMatch(candidate.merchant_name, compact.merchant)) {
    score += 8;
  }
  if (compact?.province && candidate.province && candidate.province.toUpperCase() === String(compact.province).toUpperCase()) {
    score += 5;
  }
  if (compact?.city && candidate.city && stringsRoughlyMatch(candidate.city, compact.city)) {
    score += 4;
  }

  const compactTotal = Number(compact?.total) || 0;
  if (compactTotal > 0) {
    const totalDiff = Math.abs(candidate.total_amount - compactTotal);
    score += Math.max(0, 12 - totalDiff);
  }

  const taxSum = taxFields.reduce((sum, field) => sum + (Number(candidate[field]) || 0), 0);
  if (Math.abs(candidate.taxes_total - taxSum) < 0.011) {
    score += 6;
  }

  return score;
}

function stringsRoughlyMatch(left, right) {
  const normalize = (value) => String(value || "").toLowerCase().replace(/[^\w]+/g, "");
  const leftText = normalize(left);
  const rightText = normalize(right);
  return Boolean(leftText && rightText) && (leftText.includes(rightText) || rightText.includes(leftText));
}

function normalizeCategoryGuess(categoryGuess) {
  const raw = String(categoryGuess || "").trim();
  if (!raw) return "Diverse Expenses";
  if (ALLOWED_CATEGORY_SET.has(raw)) return raw;

  const normalizedGuess = normalizeCategoryText(raw);
  let bestCategory = "Diverse Expenses";
  let bestScore = 0;

  for (const category of ALLOWED_CATEGORIES) {
    const normalizedCategory = normalizeCategoryText(category);
    let score = 0;

    if (normalizedGuess === normalizedCategory) score += 100;
    if (normalizedGuess.includes(normalizedCategory) || normalizedCategory.includes(normalizedGuess)) score += 60;

    const guessTokens = normalizedGuess.split(" ").filter(Boolean);
    const categoryTokens = normalizedCategory.split(" ").filter(Boolean);
    guessTokens.forEach((token) => {
      if (categoryTokens.includes(token)) score += 12;
      else if (categoryTokens.some((categoryToken) => categoryToken.includes(token) || token.includes(categoryToken))) score += 6;
    });

    if (score > bestScore) {
      bestScore = score;
      bestCategory = category;
    }
  }

  return bestScore >= 12 ? bestCategory : "Diverse Expenses";
}

function normalizeCategoryText(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/&/g, " and ")
    .replace(/[^\w\s]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
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
const { webcrypto } = require("node:crypto");
