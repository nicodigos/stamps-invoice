const { listExistingInvoiceTree } = require("./cnet-invoices-lib");

exports.handler = async function handler(event) {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: "Method not allowed." };
  }

  try {
    const payload = JSON.parse(event.body || "{}");
    const accessToken = String(payload.accessToken || "").trim();
    if (!accessToken) {
      return { statusCode: 401, body: "Missing Microsoft token." };
    }

    const result = await listExistingInvoiceTree(accessToken);
    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        existing: result.existing,
        tree: result.tree,
        driveId: result.driveId,
        invoices: Object.keys(result.existing).length,
        files: result.totalFiles,
      }),
    };
  } catch (error) {
    return { statusCode: 500, body: error.message || String(error) };
  }
};
