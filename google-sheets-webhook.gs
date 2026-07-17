const VERIFIED_SHEET_NAME = "Received News";
const INCOMING_SHEET_NAME = "Incoming Requests";
const ERROR_SHEET_NAME = "Webhook Errors";
const WEBHOOK_VERSION = "2026-07-17-telebirr-paypal-only";
const BASE_PRICE_BIRR = 500;
const URGENT_PRICE_BIRR = 2000;
const PAYPAL_BASE_PRICE_USD = 25;
const PAYPAL_URGENT_PRICE_USD = 100;
const TELEBIRR_NAME = "Fraol Eshetu Hailu";
const TELEBIRR_NUMBER = "0913885322";
const PAYPAL_NAME = "Yonatan Woldegiorgis";
const PAYPAL_USERNAME = "@YonatanWoldegiorgis9";
const AUDIO_FOLDER_NAME = "John Bad News Audio";
const FINAL_REQUEST_SHEETS = [VERIFIED_SHEET_NAME];
const REQUEST_HEADERS = [
  "Received At", "Submission ID", "Payment Status", "Done News", "Sender Name",
  "Sender Phone", "Bad News", "Requested Receiver", "Claimed Amount", "Verified Amount",
  "Payment Payer", "Payment Receiver", "Payment Account", "Reference No", "Receipt Type",
  "Audio Link", "Receipt File", "Language", "Reject Reason",
];

function setupAuthorization() {
  getOrCreateDriveFolder(AUDIO_FOLDER_NAME);
  SpreadsheetApp.getActiveSpreadsheet();
}

function doGet() {
  return jsonResponse({ ok: true, version: WEBHOOK_VERSION });
}

function doPost(event) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const receivedAt = new Date();
  const rawBody = event && event.postData ? event.postData.contents || "{}" : "{}";
  try {
    const data = JSON.parse(rawBody);
    const verification = verifyPaymentReceipt(data, receivedAt);
    if (verification.reference && isDuplicateReference(spreadsheet, verification.reference)) {
      verification.ok = false;
      verification.errors.push("Transaction reference has already been used");
    }
    if (!verification.ok) return jsonResponse({ ok: false, errors: verification.errors });

    const audioLink = saveAudioFile(data);
    appendSubmission(spreadsheet, INCOMING_SHEET_NAME, receivedAt, data, verification, audioLink, "Incoming");
    appendSubmission(spreadsheet, VERIFIED_SHEET_NAME, receivedAt, data, verification, audioLink, "Verified");
    return jsonResponse({ ok: true, status: "Verified" });
  } catch (error) {
    appendWebhookError(spreadsheet, receivedAt, rawBody, error);
    return jsonResponse({ ok: false, errors: [String(error && error.message ? error.message : error)] });
  }
}

function verifyPaymentReceipt(data, receivedAt) {
  const method = cleanCell(data.paymentMethod).toLowerCase();
  if (method === "telebirr") return verifyTelebirrReceipt(data, receivedAt);
  if (method === "paypal") return verifyPaypalReceipt(data);
  return { ok: false, errors: ["Unsupported payment method. Select Telebirr or PayPal."] };
}

function verifyTelebirrReceipt(data, receivedAt) {
  const text = cleanCell(data.receiptOcrText);
  const compact = normalizeComparable(text);
  const digits = text.replace(/\D/g, "");
  const expected = getExpectedAmount(data);
  const amount = extractEtbAmount(text);
  const reference = extractReference(text);
  const paymentDate = extractPaymentDate(text);
  const errors = [];

  if (!data.telebirrReceiptVerified || !data.receiptFile || !text) errors.push("Telebirr receipt screenshot was not scanned");
  if (compact.indexOf("telebirr") === -1) errors.push("The screenshot is not a Telebirr receipt");
  if (compact.indexOf(normalizeComparable(TELEBIRR_NAME)) === -1 && digits.indexOf(TELEBIRR_NUMBER) === -1) {
    errors.push(`Telebirr recipient must be ${TELEBIRR_NAME} or ${TELEBIRR_NUMBER}`);
  }
  appendAmountErrors(errors, amount, expected, "ETB");
  if (!reference) errors.push("Telebirr transaction reference was not found");
  if (paymentDate && receivedAt.getTime() - paymentDate.getTime() < -300000) errors.push("Telebirr payment date/time is in the future");
  return {
    ok: errors.length === 0, errors, amount, paymentDate, reference,
    payer: extractPayer(text) || "Telebirr customer", receiver: TELEBIRR_NAME, receiverAccount: TELEBIRR_NUMBER,
  };
}

function verifyPaypalReceipt(data) {
  const text = cleanCell(data.receiptOcrText);
  const compact = normalizeComparable(text);
  const expected = getExpectedAmount(data);
  const amount = extractUsdAmount(text);
  const reference = extractReference(text);
  const errors = [];
  if (!data.paypalReceiptVerified || !data.receiptFile || !text) errors.push("PayPal receipt screenshot was not scanned");
  if (compact.indexOf("paypal") === -1) errors.push("The screenshot is not a PayPal receipt");
  if (compact.indexOf(normalizeComparable(PAYPAL_NAME)) === -1 && compact.indexOf(normalizeComparable(PAYPAL_USERNAME)) === -1) {
    errors.push(`PayPal recipient must be ${PAYPAL_NAME} or ${PAYPAL_USERNAME}`);
  }
  appendAmountErrors(errors, amount, expected, "USD");
  if (!reference) errors.push("PayPal transaction reference was not found");
  return {
    ok: errors.length === 0, errors, amount, reference,
    payer: extractPayer(text) || "PayPal customer", receiver: PAYPAL_NAME, receiverAccount: PAYPAL_USERNAME,
  };
}

function getExpectedAmount(data) {
  const method = cleanCell(data.paymentMethod).toLowerCase();
  const urgent = cleanCell(data.serviceTier).toLowerCase().indexOf("urgent") !== -1;
  const special = parseMoney(data.specialRequestAmount);
  const base = method === "paypal" ? (urgent ? PAYPAL_URGENT_PRICE_USD : PAYPAL_BASE_PRICE_USD) : (urgent ? URGENT_PRICE_BIRR : BASE_PRICE_BIRR);
  return base + special;
}

function appendAmountErrors(errors, receiptAmount, expectedAmount, currency) {
  if (!receiptAmount) errors.push(`Receipt amount was not found`);
  else if (Math.round(receiptAmount * 100) !== Math.round(expectedAmount * 100)) errors.push(`Receipt amount must be ${formatMoney(expectedAmount)} ${currency}`);
}

function extractEtbAmount(text) {
  return extractMoney(text, [/(?:ETB|Birr|Br)\s*([0-9,]+(?:\.\d{1,2})?)/i, /([0-9,]+(?:\.\d{1,2})?)\s*(?:ETB|Birr|Br)\b/i]);
}

function extractUsdAmount(text) {
  return extractMoney(text, [/(?:USD|\$)\s*([0-9,]+(?:\.\d{1,2})?)/i, /([0-9,]+(?:\.\d{1,2})?)\s*USD\b/i]);
}

function extractMoney(text, patterns) {
  for (let i = 0; i < patterns.length; i += 1) {
    const match = String(text || "").match(patterns[i]);
    if (match) return parseMoney(match[1]);
  }
  return 0;
}

function extractReference(text) {
  const patterns = [/(?:Transaction|Payment)\s*(?:No\.?|Number|ID|Reference|Ref\.?)\s*:?\s*([A-Za-z0-9-]{6,64})/i, /(?:Receipt|Reference|Ref\.?)\s*(?:No\.?|Number|ID)?\s*:?\s*([A-Za-z0-9-]{6,64})/i];
  for (let i = 0; i < patterns.length; i += 1) {
    const match = String(text || "").match(patterns[i]);
    if (match) return cleanReceiptField(match[1]);
  }
  return "";
}

function extractPayer(text) {
  const match = String(text || "").match(/(?:Payer|Sender|From|Paid by)(?:\s+Name)?\s*:?\s*([A-Za-z]+(?:\s+[A-Za-z]+){1,4})/i);
  return match ? cleanReceiptField(match[1]) : "";
}

function extractPaymentDate(text) {
  const match = String(text || "").match(/(?:Payment|Transaction)?\s*Date(?:\s*&\s*Time)?\s*:?\s*([A-Za-z0-9,/: -]{8,40})/i);
  if (!match) return null;
  return parseReceiptDate(match[1]);
}

function parseReceiptDate(value) {
  const text = cleanCell(value);
  const direct = new Date(text);
  if (!isNaN(direct.getTime())) return direct;
  const match = text.match(/\b(\d{1,2})[/-](\d{1,2})[/-](\d{4})(?:[,\s]+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?)?/i);
  if (!match) return null;
  let day = Number(match[1]);
  let month = Number(match[2]);
  let hour = Number(match[4] || 0);
  const meridiem = String(match[7] || "").toUpperCase();
  if (month > 12 && day <= 12) { const swap = day; day = month; month = swap; }
  if (meridiem === "PM" && hour < 12) hour += 12;
  if (meridiem === "AM" && hour === 12) hour = 0;
  const parsed = new Date(Number(match[3]), month - 1, day, hour, Number(match[5] || 0), Number(match[6] || 0));
  return isNaN(parsed.getTime()) ? null : parsed;
}

function isDuplicateReference(spreadsheet, reference) {
  const expected = normalizeComparable(reference);
  return FINAL_REQUEST_SHEETS.some(function(sheetName) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return false;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(cleanCell);
    const column = headers.indexOf("Reference No") + 1;
    if (!column) return false;
    return sheet.getRange(2, column, sheet.getLastRow() - 1, 1).getValues().flat().some(function(value) { return normalizeComparable(value) === expected; });
  });
}

function appendSubmission(spreadsheet, sheetName, receivedAt, data, verification, audioLink, status) {
  const sheet = getSheet(spreadsheet, sheetName);
  ensureHeader(sheet, REQUEST_HEADERS);
  const row = [receivedAt.toLocaleString(), cleanCell(data.id), status, false, cleanCell(data.name), getSenderPhone(data), cleanCell(data.badNews), cleanCell(data.receiver), formatClaimedPayment(data), verification.amount || "", cleanReceiptField(verification.payer), cleanReceiptField(verification.receiver), cleanReceiptField(verification.receiverAccount), cleanReceiptField(verification.reference), cleanCell(data.paymentMethod).toUpperCase(), audioLink, cleanCell(data.receiptFile), cleanCell(data.language), ""];
  sheet.appendRow(row);
  sheet.getRange(sheet.getLastRow(), 4).insertCheckboxes();
}

function formatClaimedPayment(data) {
  const currency = cleanCell(data.paymentCurrency || (data.paymentMethod === "paypal" ? "USD" : "ETB"));
  const parts = [`${cleanCell(data.paymentAmount)} ${currency}`, cleanCell(data.paymentMethod).toUpperCase()];
  if (data.serviceTier) parts.push(cleanCell(data.serviceTier));
  if (data.specialRequestAmount) parts.push(`Special: ${cleanCell(data.specialRequestAmount)} ${currency}`);
  if (data.contactChannel) parts.push(`Contact: ${cleanCell(data.contactChannel)}`);
  return parts.join(" | ");
}

function getSenderPhone(data) { return cleanCell(data.phone || data.phoneNumber || data.senderPhone); }
function parseMoney(value) { return Number(String(value || "").replace(/[^0-9.]/g, "")) || 0; }
function formatMoney(value) { return Number(value || 0).toLocaleString("en-US", { maximumFractionDigits: 2 }); }
function normalizeComparable(value) { return String(value || "").toLowerCase().replace(/[^a-z0-9]+/g, ""); }
function cleanCell(value) { return String(value || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").trim(); }
function cleanReceiptField(value) {
  const cleaned = cleanCell(value).replace(/\s+/g, " ");
  return cleaned && cleaned.length <= 140 ? cleaned : "";
}

function saveAudioFile(data) {
  if (!data.audioDataBase64 || !data.audioFile || data.audioFile === "-") return cleanCell(data.audioFile);
  try {
    const folder = getOrCreateDriveFolder(AUDIO_FOLDER_NAME);
    const fileName = `${cleanCell(data.id) || Date.now()}-${cleanCell(data.audioFile).replace(/[\\/:*?"<>|]/g, "-") || "audio"}`;
    const blob = Utilities.newBlob(Utilities.base64Decode(data.audioDataBase64), cleanCell(data.audioMimeType) || "audio/mpeg", fileName);
    const file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (ignored) {}
    return `=HYPERLINK("${file.getUrl()}", "${fileName}")`;
  } catch (error) { return `Audio save failed: ${String(error && error.message ? error.message : error)}`; }
}

function getOrCreateDriveFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function getSheet(spreadsheet, name) { return spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name); }

function ensureHeader(sheet, headers) {
  if (sheet.getLastRow() === 0) sheet.appendRow(headers);
  else sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  const lastRow = Math.max(sheet.getLastRow(), 1);
  sheet.getRange(1, 1, lastRow, headers.length).setVerticalAlignment("top").setWrap(true);
  sheet.autoResizeColumns(1, headers.length);
}

function appendWebhookError(spreadsheet, receivedAt, rawBody, error) {
  const sheet = getSheet(spreadsheet, ERROR_SHEET_NAME);
  ensureHeader(sheet, ["Received At", "Error Message", "Stack", "Raw Body"]);
  sheet.appendRow([receivedAt.toLocaleString(), String(error && error.message ? error.message : error), String(error && error.stack ? error.stack : ""), rawBody]);
}

function jsonResponse(payload) { return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON); }
