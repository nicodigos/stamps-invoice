const { exportInvoiceRows } = require("./cnet-invoices-lib");

exports.handler = async function handler(event) {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: "Method not allowed." };
  }

  try {
    const result = await exportInvoiceRows();
    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(result),
    };
  } catch (error) {
    return { statusCode: 500, body: error.message || String(error) };
  }
};
