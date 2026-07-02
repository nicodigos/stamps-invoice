const COMPANIES = [
  "1001298527 ONTARIO INC",
  "10342548 CANADA INC",
  "10377180 CANADA INC-standby",
  "10696480 CANADA LTD",
  "12433087 CANADA INC-MASTER",
  "13037622 CANADA INC",
  "14661796 CANADA LTD- not active",
  "16021166 Canada Inc-Diego-Chacho",
  "9359-6633 QUEBEC INC",
  "9390-9216 QUEBEC INC",
  "TAYANTI-CANADA",
];

const BANKS = ["National Bank", "Scotia Bank", "Desjardins"];

const RECEIPT_COMPANIES = [
  "1001298527 ONTARIO INC",
  "10342548 CANADA INC",
  "10696480 CANADA LTD",
  "12433087 CANADA INC-MASTER",
  "13037622 CANADA INC",
  "9359-6633 QUEBEC INC",
  "9390-9216 QUEBEC INC",
  "D-TECH CONSTRUCTION",
  "TAYANTI-CANADA",
];

const RECEIPT_BANKS = ["Scotiabank", "Desjardins", "National Bank"];

const CATEGORIES = [
  "4 Subs Invoices With Tax",
  "5 T4A Payments",
  "6 Payroll - Paystubs",
  "8 Expenses - Advertising",
  "8 Expenses - Automobile",
  "8 Expenses - Cellphone",
  "8 Expenses - Cleaning",
  "8 Expenses - Equipment Parts And Maintenance",
  "8 Expenses - Insurance - Company",
  "8 Expenses - Insurance - Vehicles",
  "8 Expenses - Office",
  "8 Expenses - Professional",
  "9 Reimbursements - OP",
  "9 Reimbursements - Reimbursements",
];

const DEFAULT_WORKBOOK_PATH = "General/12433087 CANADA INC-MASTER/09-Pagos Periodos/2025/Building & Contractor Pay List/Building Address & Contractor Pay List.xlsx";

exports.handler = async function handler() {
  return {
    statusCode: 200,
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      companyOptions: COMPANIES,
      bankOptions: BANKS,
      categoryOptions: CATEGORIES,
      receiptCompanyOptions: RECEIPT_COMPANIES,
      receiptBankOptions: RECEIPT_BANKS,
      spHostname: process.env.SP_HOSTNAME || "",
      spSitePath: process.env.SP_SITE_PATH || "",
      spDriveName: process.env.SP_DRIVE_NAME || "Documents",
      pagosPeriodosWorkbookPath: (process.env.PAGOS_PERIODOS_WORKBOOK_PATH || DEFAULT_WORKBOOK_PATH).trim().replace(/^\/+|\/+$/g, ""),
      receiptsDatabaseDir: process.env.RECEIPTS_DATABASE_DIR || "General/Sales receipts database",
      receiptsDatabaseCsv: "sales_receipts_database.csv",
    }),
  };
};
