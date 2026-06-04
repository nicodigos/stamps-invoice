const {
  downloadInvoicePdf,
  loginToCnet,
  uploadInvoicePdf,
} = require("./cnet-invoices-lib");

exports.handler = async function handler(event) {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: "Method not allowed." };
  }

  try {
    const payload = JSON.parse(event.body || "{}");
    const accessToken = String(payload.accessToken || "").trim();
    const rows = Array.isArray(payload.rows) ? payload.rows : [];
    if (!accessToken) {
      return { statusCode: 401, body: "Missing Microsoft token." };
    }
    if (!rows.length) {
      return { statusCode: 400, body: "No invoices to process." };
    }

    const jar = await loginToCnet();
    const results = [];
    for (const row of rows) {
      const invoiceId = String(row?.invoiceId || "").trim();
      try {
        const pdfBytes = await downloadInvoicePdf(row, jar);
        const upload = await uploadInvoicePdf(row, pdfBytes, accessToken);
        results.push({
          ok: true,
          invoiceId,
          targetPath: upload.targetPath,
        });
      } catch (error) {
        results.push({
          ok: false,
          invoiceId,
          error: error.message || String(error),
        });
      }
    }

    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ results }),
    };
  } catch (error) {
    return { statusCode: 500, body: error.message || String(error) };
  }
};
