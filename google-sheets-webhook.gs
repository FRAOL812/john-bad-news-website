const VERIFIED_SHEET_NAME = "Received News";
const REJECTED_SHEET_NAME = "Rejected Payments";
const BASE_PRICE_BIRR = 500;
const PAYMENT_WINDOW_MINUTES = 20;
const EXPECTED_RECEIVER_NAME = "Fraol Eshetu Hailu";
const EXPECTED_RECEIVER_ACCOUNT_SUFFIX = "8583";
const CBE_RECEIPT_BASE_URL = "https://mbreciept.cbe.com.et/";

function setupAuthorization() {
  UrlFetchApp.fetch(CBE_RECEIPT_BASE_URL, {
    muteHttpExceptions: true,
    followRedirects: true,
  });
  SpreadsheetApp.getActiveSpreadsheet();
}

function doPost(event) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const data = JSON.parse(event.postData.contents || "{}");
  const receivedAt = new Date();
  const verification = verifyCbeReceipt(data.cbeReceiptUrl, receivedAt);

  if (verification.ok && isDuplicateReference(spreadsheet, verification.reference)) {
    verification.ok = false;
    verification.errors.push("Reference number has already been used");
  }

  if (!verification.ok) {
    appendRejected(spreadsheet, receivedAt, data, verification);
    return jsonResponse({ ok: false, errors: verification.errors });
  }

  const sheet = getSheet(spreadsheet, VERIFIED_SHEET_NAME);
  ensureVerifiedHeader(sheet);
  sheet.appendRow([
    receivedAt.toLocaleString(),
    verification.paymentDate.toLocaleString(),
    data.name || "",
    data.phone || "",
    data.badNews || "",
    data.receiver || "",
    verification.payer || "",
    verification.receiver || "",
    verification.receiverAccount || "",
    verification.amount,
    verification.reference || "",
    `=HYPERLINK("${data.cbeReceiptUrl}", "Open CBE receipt")`,
    "Verified",
    data.audioFile || "",
    data.receiptFile || "",
    data.language || "",
    data.id || "",
  ]);

  return jsonResponse({ ok: true, status: "Verified" });
}

function verifyCbeReceipt(receiptUrl, receivedAt) {
  const errors = [];
  const url = String(receiptUrl || "").trim();

  if (!url.toLowerCase().startsWith(CBE_RECEIPT_BASE_URL)) {
    return { ok: false, errors: ["Missing or invalid CBE receipt URL"] };
  }

  const response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
  });
  const statusCode = response.getResponseCode();

  if (statusCode < 200 || statusCode >= 300) {
    return { ok: false, errors: [`CBE receipt page returned HTTP ${statusCode}`] };
  }

  const text = normalizeReceiptText(response.getContentText());
  const amount = extractAmount(text);
  const paymentDate = extractPaymentDate(text);
  const receiver = extractField(text, /Receiver:\s*\n?\s*([^\n]+)/i);
  const receiverAccount = extractField(text, /Receiver:[\s\S]*?Account:\s*\n?\s*([^\n]+)/i);
  const payer = extractField(text, /Payer:\s*\n?\s*([^\n]+)/i);
  const reference = extractField(text, /Reference No\.?\s*(?:\(VAT Invoice No\.\))?:\s*\n?\s*([A-Za-z0-9-]+)/i);

  if (!amount || amount < BASE_PRICE_BIRR) {
    errors.push(`Transferred amount is below ${BASE_PRICE_BIRR} ETB`);
  }

  if (!receiver || !receiver.toLowerCase().includes(EXPECTED_RECEIVER_NAME.toLowerCase())) {
    errors.push(`Receiver is not ${EXPECTED_RECEIVER_NAME}`);
  }

  if (!receiverAccount || !receiverAccount.replace(/\D/g, "").endsWith(EXPECTED_RECEIVER_ACCOUNT_SUFFIX)) {
    errors.push(`Receiver account does not end with ${EXPECTED_RECEIVER_ACCOUNT_SUFFIX}`);
  }

  if (!paymentDate) {
    errors.push("Payment date/time not found");
  } else {
    const ageMs = receivedAt.getTime() - paymentDate.getTime();
    if (ageMs < 0) {
      errors.push("Payment date/time is in the future");
    }
    if (ageMs > PAYMENT_WINDOW_MINUTES * 60 * 1000) {
      errors.push(`Payment is older than ${PAYMENT_WINDOW_MINUTES} minutes`);
    }
  }

  if (!reference) {
    errors.push("Reference number not found");
  }

  return {
    ok: errors.length === 0,
    errors,
    amount,
    paymentDate,
    receiver,
    receiverAccount,
    payer,
    reference,
  };
}

function normalizeReceiptText(html) {
  return html
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<[^>]+>/g, "\n")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/\s+\n/g, "\n")
    .replace(/\n\s+/g, "\n")
    .replace(/[ \t]+/g, " ")
    .trim();
}

function extractField(text, pattern) {
  const match = text.match(pattern);
  return match ? match[1].trim() : "";
}

function extractAmount(text) {
  const amountText = extractField(text, /Transferred Amount:\s*\n?\s*([0-9,]+(?:\.\d{1,2})?)\s*ETB/i);
  return amountText ? Number(amountText.replace(/,/g, "")) : 0;
}

function extractPaymentDate(text) {
  const dateText = extractField(text, /Payment Date\s*&?\s*Time:\s*\n?\s*([^\n]+)/i);
  if (!dateText) {
    return null;
  }

  const parsed = new Date(dateText);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function isDuplicateReference(spreadsheet, reference) {
  if (!reference) {
    return false;
  }

  const sheet = spreadsheet.getSheetByName(VERIFIED_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    return false;
  }

  const referenceColumn = 11;
  const references = sheet.getRange(2, referenceColumn, sheet.getLastRow() - 1, 1).getValues().flat();
  return references.some((value) => String(value).trim() === reference);
}

function appendRejected(spreadsheet, receivedAt, data, verification) {
  const sheet = getSheet(spreadsheet, REJECTED_SHEET_NAME);
  ensureRejectedHeader(sheet);
  sheet.appendRow([
    receivedAt.toLocaleString(),
    data.name || "",
    data.phone || "",
    data.receiver || "",
    data.paymentAmount || "",
    data.cbeReceiptUrl ? `=HYPERLINK("${data.cbeReceiptUrl}", "Open CBE receipt")` : "",
    verification.errors.join(" | "),
    data.receiptFile || "",
    data.id || "",
  ]);
}

function getSheet(spreadsheet, sheetName) {
  return spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
}

function ensureVerifiedHeader(sheet) {
  if (sheet.getLastRow() > 0) {
    return;
  }

  sheet.appendRow([
    "Received At",
    "Payment Date",
    "Sender Name",
    "Sender Phone",
    "Bad News",
    "Requested Receiver",
    "CBE Payer",
    "CBE Receiver",
    "CBE Receiver Account",
    "Verified Amount",
    "Reference No",
    "CBE Receipt Link",
    "Payment Status",
    "Audio File",
    "Receipt File",
    "Language",
    "Submission ID",
  ]);
}

function ensureRejectedHeader(sheet) {
  if (sheet.getLastRow() > 0) {
    return;
  }

  sheet.appendRow([
    "Received At",
    "Sender Name",
    "Sender Phone",
    "Requested Receiver",
    "Claimed Amount",
    "CBE Receipt Link",
    "Reject Reason",
    "Receipt File",
    "Submission ID",
  ]);
}

function jsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}
