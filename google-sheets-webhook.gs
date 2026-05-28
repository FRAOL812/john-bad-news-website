const VERIFIED_SHEET_NAME = "Received News";
const REJECTED_SHEET_NAME = "Rejected Payments";
const INCOMING_SHEET_NAME = "Incoming Requests";
const ERROR_SHEET_NAME = "Webhook Errors";
const BASE_PRICE_BIRR = 500;
const PAYMENT_WINDOW_MINUTES = 20;
const EXPECTED_RECEIVER_NAME = "Fraol Eshetu Hailu";
const EXPECTED_RECEIVER_ACCOUNT_SUFFIX = "8583";
const CBE_RECEIPT_BASE_URL = "https://mbreciept.cbe.com.et/";
const REQUEST_HEADERS = [
  "Received At",
  "Submission ID",
  "Payment Status",
  "Sender Name",
  "Sender Phone",
  "Bad News",
  "Requested Receiver",
  "Claimed Amount",
  "Verified Amount",
  "CBE Payer",
  "CBE Receiver",
  "CBE Receiver Account",
  "Reference No",
  "CBE Receipt Link",
  "Audio File",
  "Receipt File",
  "Language",
  "Reject Reason",
];

function setupAuthorization() {
  UrlFetchApp.fetch(CBE_RECEIPT_BASE_URL, {
    muteHttpExceptions: true,
    followRedirects: true,
  });
  SpreadsheetApp.getActiveSpreadsheet();
}

function doPost(event) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const receivedAt = new Date();
  const rawBody = event && event.postData ? event.postData.contents || "{}" : "{}";

  try {
    const data = JSON.parse(rawBody);
    appendIncoming(spreadsheet, receivedAt, data, rawBody);
    const verification = verifyCbeReceipt(data.cbeReceiptUrl, receivedAt);

    if (verification.ok && isDuplicateReference(spreadsheet, verification.reference)) {
      verification.ok = false;
      verification.errors.push("Reference number has already been used");
    }

    if (!verification.ok) {
      appendRejected(spreadsheet, receivedAt, data, verification);
      return jsonResponse({ ok: false, errors: verification.errors });
    }

    appendRequestRow(spreadsheet, VERIFIED_SHEET_NAME, [
      receivedAt.toLocaleString(),
      cleanCell(data.id),
      "Verified",
      cleanCell(data.name),
      cleanCell(data.phone),
      cleanCell(data.badNews),
      cleanCell(data.receiver),
      cleanCell(data.paymentAmount),
      verification.amount || "",
      cleanCell(verification.payer),
      cleanCell(verification.receiver),
      cleanCell(verification.receiverAccount),
      cleanCell(verification.reference),
      receiptLinkFormula(data.cbeReceiptUrl),
      cleanCell(data.audioFile),
      cleanCell(data.receiptFile),
      cleanCell(data.language),
      "",
    ]);

    return jsonResponse({ ok: true, status: "Verified" });
  } catch (error) {
    appendWebhookError(spreadsheet, receivedAt, rawBody, error);
    return jsonResponse({ ok: false, errors: [String(error && error.message ? error.message : error)] });
  }
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
  appendRequestRow(spreadsheet, REJECTED_SHEET_NAME, [
    receivedAt.toLocaleString(),
    cleanCell(data.id),
    "Rejected",
    cleanCell(data.name),
    cleanCell(data.phone),
    cleanCell(data.badNews),
    cleanCell(data.receiver),
    cleanCell(data.paymentAmount),
    verification.amount || "",
    cleanCell(verification.payer),
    cleanCell(verification.receiver),
    cleanCell(verification.receiverAccount),
    cleanCell(verification.reference),
    receiptLinkFormula(data.cbeReceiptUrl),
    cleanCell(data.audioFile),
    cleanCell(data.receiptFile),
    cleanCell(data.language),
    verification.errors.join(" | "),
  ]);
}

function appendIncoming(spreadsheet, receivedAt, data, rawBody) {
  const sheet = getSheet(spreadsheet, INCOMING_SHEET_NAME);
  ensureIncomingHeader(sheet);
  sheet.appendRow([
    receivedAt.toLocaleString(),
    cleanCell(data.id),
    "Incoming",
    cleanCell(data.name),
    cleanCell(data.phone),
    cleanCell(data.badNews),
    cleanCell(data.receiver),
    cleanCell(data.paymentAmount),
    "",
    "",
    "",
    "",
    "",
    receiptLinkFormula(data.cbeReceiptUrl),
    cleanCell(data.audioFile),
    cleanCell(data.receiptFile),
    cleanCell(data.language),
    "",
  ]);
}

function appendWebhookError(spreadsheet, receivedAt, rawBody, error) {
  const sheet = getSheet(spreadsheet, ERROR_SHEET_NAME);
  ensureErrorHeader(sheet);
  sheet.appendRow([
    receivedAt.toLocaleString(),
    String(error && error.message ? error.message : error),
    String(error && error.stack ? error.stack : ""),
    rawBody,
  ]);
}

function getSheet(spreadsheet, sheetName) {
  return spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
}

function appendRequestRow(spreadsheet, sheetName, row) {
  const sheet = getSheet(spreadsheet, sheetName);
  ensureRequestHeader(sheet);
  sheet.appendRow(row);
}

function ensureIncomingHeader(sheet) {
  ensureHeader(sheet, REQUEST_HEADERS);
}

function ensureRequestHeader(sheet) {
  ensureHeader(sheet, REQUEST_HEADERS);
}

function ensureErrorHeader(sheet) {
  ensureHeader(sheet, [
    "Received At",
    "Error Message",
    "Stack",
    "Raw Body",
  ]);
}

function ensureHeader(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), headers.length).setVerticalAlignment("top");
  sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), headers.length).setWrap(true);
  sheet.autoResizeColumns(1, headers.length);
}

function cleanCell(value) {
  return String(value || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").trim();
}

function receiptLinkFormula(url) {
  const cleanUrl = cleanCell(url);
  return cleanUrl ? `=HYPERLINK("${cleanUrl}", "Open CBE receipt")` : "";
}

function jsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}
