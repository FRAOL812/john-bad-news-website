const VERIFIED_SHEET_NAME = "Received News";
const REJECTED_SHEET_NAME = "Rejected Payments";
const INCOMING_SHEET_NAME = "Incoming Requests";
const ERROR_SHEET_NAME = "Webhook Errors";
const WEBHOOK_VERSION = "2026-05-31-cbe-user-web-token";
const BASE_PRICE_BIRR = 500;
const PAYMENT_WINDOW_MINUTES = 60;
const EXPECTED_RECEIVER_NAME = "Fraol Eshetu Hailu";
const EXPECTED_RECEIVER_ACCOUNT_SUFFIX = "8583";
const CBE_RECEIPT_BASE_URL = "https://mbreciept.cbe.com.et/";
const CBE_LEGACY_RECEIPT_BASE_URL = "https://apps.cbe.com.et:100/";
const CBE_TRANSACTION_API_BASE_URL = "https://Mb.cbe.com.et/api/v1/transactions/public/transaction-detail/";
const AUDIO_FOLDER_NAME = "John Bad News Audio";
const MAX_RECEIPT_FIELD_LENGTH = 140;
const FINAL_REQUEST_SHEETS = [VERIFIED_SHEET_NAME, REJECTED_SHEET_NAME];
const CBE_TRANSACTION_API_HEADERS = {
  "X-App-ID": "d1292e42-7400-49de-a2d3-9731caa4c819",
  "X-App-Version": "0a01980b-9859-1369-8198-59f403820000",
};
const REQUEST_HEADERS = [
  "Received At",
  "Submission ID",
  "Payment Status",
  "Done News",
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
  "Audio Link",
  "Receipt File",
  "Language",
  "Reject Reason",
];

function setupAuthorization() {
  UrlFetchApp.fetch(CBE_RECEIPT_BASE_URL, {
    muteHttpExceptions: true,
    followRedirects: true,
  });
  UrlFetchApp.fetch(CBE_LEGACY_RECEIPT_BASE_URL, {
    muteHttpExceptions: true,
    followRedirects: true,
  });
  UrlFetchApp.fetch(`${CBE_TRANSACTION_API_BASE_URL}authorization-check`, {
    headers: CBE_TRANSACTION_API_HEADERS,
    muteHttpExceptions: true,
    followRedirects: true,
  });
  getOrCreateDriveFolder(AUDIO_FOLDER_NAME);
  SpreadsheetApp.getActiveSpreadsheet();
}

function doPost(event) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const receivedAt = new Date();
  const rawBody = event && event.postData ? event.postData.contents || "{}" : "{}";

  try {
    const data = JSON.parse(rawBody);
    data.cbeReceiptUrl = normalizeCbeReceiptUrl(data.cbeReceiptUrl);
    const verification = verifyPaymentReceipt(data, receivedAt);

    if (data.paymentMethod !== "paypal" && isDuplicateReceiptUrl(spreadsheet, data.cbeReceiptUrl)) {
      verification.ok = false;
      verification.errors.push("CBE receipt link has already been submitted");
    }

    if (verification.reference && isDuplicateReference(spreadsheet, verification.reference)) {
      verification.ok = false;
      verification.errors.push("Reference number has already been used");
    }

    if (!verification.ok) {
      return jsonResponse({ ok: false, errors: verification.errors });
    }

    const audioLink = saveAudioFile(data);
    appendIncoming(spreadsheet, receivedAt, data, audioLink);

    appendRequestRow(spreadsheet, VERIFIED_SHEET_NAME, [
      receivedAt.toLocaleString(),
      cleanCell(data.id),
      "Verified",
      false,
      cleanCell(data.name),
      cleanCell(data.phone),
      cleanCell(data.badNews),
      cleanCell(data.receiver),
      formatClaimedPayment(data),
      verification.amount || "",
      cleanCell(verification.payer),
      cleanCell(verification.receiver),
      cleanCell(verification.receiverAccount),
      cleanCell(verification.reference),
      receiptLinkFormula(data.cbeReceiptUrl),
      audioLink,
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

function verifyCbeReceipt(data, receivedAt) {
  const errors = [];
  const receiptUrl = data.cbeReceiptUrl || "";

  if (receiptUrl.toLowerCase().startsWith("ocr:")) {
    const reference = receiptUrl.substring(4).trim();
    if (!reference || reference.length < 5) {
      errors.push("Invalid OCR Reference Number");
    }
    const claimedAmount = Number(data.paymentAmount) || BASE_PRICE_BIRR;
    return {
      ok: errors.length === 0,
      errors,
      amount: claimedAmount,
      paymentDate: receivedAt,
      receiver: EXPECTED_RECEIVER_NAME,
      receiverAccount: `*${EXPECTED_RECEIVER_ACCOUNT_SUFFIX}`,
      payer: "CBE Screenshot (OCR Verified)",
      reference: reference,
    };
  }

  const url = normalizeCbeReceiptUrl(receiptUrl);

  if (!isCbeReceiptUrl(url)) {
    return { ok: false, errors: ["Missing or invalid CBE receipt URL"] };
  }

  const apiVerification = verifyCbeReceiptFromApi(url, receivedAt);
  if (!apiVerification.ok) {
    return apiVerification;
  }

  return apiVerification;
}

function verifyCbeReceiptFromApi(receiptUrl, receivedAt) {
  const receiptToken = extractCbeReceiptToken(receiptUrl);

  if (!receiptToken) {
    return {
      ok: false,
      errors: [`CBE receipt link is missing the receipt token (${WEBHOOK_VERSION})`],
    };
  }

  const apiUrl = buildCbeTransactionApiUrl(receiptUrl, receiptToken);
  let response;
  try {
    response = UrlFetchApp.fetch(apiUrl, {
      method: "get",
      headers: buildCbeTransactionApiHeaders(receiptUrl),
      muteHttpExceptions: true,
      followRedirects: true,
    });
  } catch (error) {
    return {
      ok: false,
      errors: [`Could not reach the CBE transaction API. Try again in a moment. (${WEBHOOK_VERSION})`],
    };
  }
  const statusCode = response.getResponseCode();
  const responseText = response.getContentText() || "";

  if (statusCode < 200 || statusCode >= 300) {
    const apiErrorMessage = extractCbeApiErrorMessage(responseText);
    const invalidTokenMessage = apiErrorMessage.toLowerCase().indexOf("invalid or tampered token") !== -1;

    return {
      ok: false,
      errors: [
        invalidTokenMessage
          ? `CBE rejected this receipt link as invalid or tampered. Please upload the original, clear CBE receipt screenshot again. (${WEBHOOK_VERSION})`
          : `CBE transaction API returned HTTP ${statusCode}${apiErrorMessage ? `: ${apiErrorMessage}` : ""}. Check that the receipt link is valid. (${WEBHOOK_VERSION})`,
      ],
    };
  }

  let transaction;
  try {
    transaction = JSON.parse(responseText || "{}");
  } catch (error) {
    return {
      ok: false,
      errors: [`CBE transaction API returned unreadable data. Try again in a moment. (${WEBHOOK_VERSION})`],
    };
  }

  if (!transaction || typeof transaction !== "object" || !transaction.id) {
    return {
      ok: false,
      errors: [`CBE transaction details were not found for this receipt link. Check that the receipt link is valid. (${WEBHOOK_VERSION})`],
    };
  }

  return buildCbeVerificationFromTransaction(transaction, receivedAt);
}

function extractCbeApiErrorMessage(responseText) {
  try {
    const payload = JSON.parse(responseText || "{}");
    return cleanCell(payload.detail || payload.message || payload.title || "");
  } catch (error) {
    return "";
  }
}

function buildCbeVerificationFromTransaction(transaction, receivedAt) {
  const amount = parseReceiptNumber(transaction.amountCredited || transaction.creditAmount || transaction.debitAmount);
  const paymentDate = parseCbeApiDate(transaction.dateTimes && transaction.dateTimes[0]) || parseCbeApiDate(transaction.processingDate);
  const receiver = cleanReceiptField(transaction.creditAccountHolder);
  const receiverAccount = cleanReceiptField(transaction.creditAccountNo);
  const payer = cleanReceiptField(transaction.debitAccountHolder);
  const reference = cleanReceiptField(transaction.id);
  const errors = [];

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
    if (ageMs < -5 * 60 * 1000) {
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

function extractCbeReceiptToken(receiptUrl) {
  const normalizedUrl = normalizeCbeReceiptUrl(receiptUrl);
  const queryToken = extractCbeReceiptQueryToken(normalizedUrl);
  if (queryToken) {
    return queryToken;
  }

  const match = normalizedUrl.match(/^https?:\/\/(?:mbreciept\.cbe\.com\.et|apps\.cbe\.com\.et:100)\/?([^?#/]+)/i);
  return match ? safeDecodeURIComponent(match[1]) : "";
}

function extractCbeReceiptQueryToken(receiptUrl) {
  const queryString = extractCbeReceiptQueryString(receiptUrl);
  if (!queryString) {
    return "";
  }

  const tokenKeys = ["token", "webToken", "web_token", "accessToken", "access_token", "id", "receiptId", "receipt_id"];
  const pairs = queryString.split("&");
  for (let keyIndex = 0; keyIndex < tokenKeys.length; keyIndex += 1) {
    const expectedKey = tokenKeys[keyIndex].toLowerCase();
    for (let pairIndex = 0; pairIndex < pairs.length; pairIndex += 1) {
      const pair = pairs[pairIndex];
      const equalsIndex = pair.indexOf("=");
      const rawKey = equalsIndex === -1 ? pair : pair.substring(0, equalsIndex);
      const rawValue = equalsIndex === -1 ? "" : pair.substring(equalsIndex + 1);
      if (safeDecodeURIComponent(rawKey).toLowerCase() === expectedKey && rawValue) {
        return safeDecodeURIComponent(rawValue);
      }
    }
  }

  return "";
}

function buildCbeTransactionApiUrl(receiptUrl, receiptToken) {
  const queryString = extractCbeReceiptQueryString(receiptUrl);
  const baseApiUrl = `${CBE_TRANSACTION_API_BASE_URL}${encodeURIComponent(receiptToken)}`;
  return queryString ? `${baseApiUrl}?${queryString}` : baseApiUrl;
}

function buildCbeTransactionApiHeaders(receiptUrl) {
  const origin = getCbeReceiptOrigin(receiptUrl);
  return Object.assign({}, CBE_TRANSACTION_API_HEADERS, {
    Origin: origin,
    Referer: normalizeCbeReceiptUrl(receiptUrl),
  });
}

function extractCbeReceiptQueryString(receiptUrl) {
  const match = String(receiptUrl || "").trim().match(/^https?:\/\/(?:mbreciept\.cbe\.com\.et|apps\.cbe\.com\.et:100)\/?[^?#]*(?:\?([^#]+))?/i);
  return match && match[1] ? match[1] : "";
}

function getCbeReceiptOrigin(receiptUrl) {
  return String(receiptUrl || "").toLowerCase().indexOf("apps.cbe.com.et:100") !== -1
    ? "https://apps.cbe.com.et:100"
    : "https://mbreciept.cbe.com.et";
}

function isCbeReceiptUrl(receiptUrl) {
  const lowerUrl = String(receiptUrl || "").toLowerCase();
  return lowerUrl.startsWith(CBE_RECEIPT_BASE_URL) || lowerUrl.startsWith(CBE_LEGACY_RECEIPT_BASE_URL);
}

function safeDecodeURIComponent(value) {
  try {
    return decodeURIComponent(String(value || "").replace(/\+/g, " "));
  } catch (error) {
    return String(value || "");
  }
}

function parseReceiptNumber(value) {
  const match = String(value || "").match(/[0-9,]+(?:\.\d+)?/);
  return match ? Number(match[0].replace(/,/g, "")) : 0;
}

function parseCbeApiDate(value) {
  const dateText = String(value || "").trim();

  if (!dateText) {
    return null;
  }

  const parsed = new Date(dateText);
  if (!Number.isNaN(parsed.getTime())) {
    return parsed;
  }

  const compactDate = dateText.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (compactDate) {
    return new Date(Number(compactDate[1]), Number(compactDate[2]) - 1, Number(compactDate[3]));
  }

  return null;
}

function verifyPaymentReceipt(data, receivedAt) {
  if (data.paymentMethod === "paypal") {
    return verifyPaypalReceipt(data);
  }

  return verifyCbeReceipt(data, receivedAt);
}

function verifyPaypalReceipt(data) {
  const amount = Number(data.paymentAmount);
  const minimumAmount = String(data.serviceTier || "").toLowerCase().indexOf("urgent") !== -1 ? 100 : 25;

  if (!data.paypalReceiptVerified) {
    return { ok: false, errors: ["PayPal receipt screenshot was not verified"] };
  }

  if (!Number.isFinite(amount) || amount < minimumAmount) {
    return { ok: false, errors: [`PayPal amount is below ${minimumAmount} USD`] };
  }

  return {
    ok: true,
    errors: [],
    amount,
    payer: "PayPal screenshot",
    receiver: "Yonatan Woldegiorgis",
    receiverAccount: "@YonatanWoldegiorgis9",
    reference: cleanCell(data.receiptVerificationValue || data.cbeReceiptUrl || ""),
  };
}

function normalizeCbeReceiptUrl(receiptUrl) {
  let url = String(receiptUrl || "").trim();
  // Strip hashes, but preserve CBE query tokens because the public API validates them.
  url = url.split("#")[0];
  // Strip trailing slashes
  url = url.replace(/\/+$/, "");

  const match = url.match(/^https?:\/\/(mbreciept\.cbe\.com\.et|apps\.cbe\.com\.et:100)\/?([^?#/]*)(\?[^#]*)?$/i);
  if (!match) {
    return url;
  }

  const baseUrl = match[1].toLowerCase() === "apps.cbe.com.et:100" ? CBE_LEGACY_RECEIPT_BASE_URL : CBE_RECEIPT_BASE_URL;
  return `${baseUrl}${match[2]}${match[3] || ""}`;
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

function cleanReceiptField(value) {
  const cleanValue = cleanCell(value).replace(/\s+/g, " ");

  if (!cleanValue || cleanValue.length > MAX_RECEIPT_FIELD_LENGTH) {
    return "";
  }

  if (looksLikeSubmissionJson(cleanValue)) {
    return "";
  }

  return cleanValue;
}

function looksLikeSubmissionJson(value) {
  const trimmedValue = String(value || "").trim();
  return (
    trimmedValue.charAt(0) === "{" &&
    trimmedValue.indexOf('"cbeReceiptUrl"') !== -1 &&
    trimmedValue.indexOf('"paymentAmount"') !== -1
  );
}

function extractAmount(text) {
  const patterns = [
    /Transferred Amount:\s*\n?\s*([0-9,]+(?:\.\d{1,2})?)\s*ETB/i,
    /Amount:\s*\n?\s*([0-9,]+(?:\.\d{1,2})?)\s*ETB/i,
    /Transaction Amount:\s*\n?\s*([0-9,]+(?:\.\d{1,2})?)/i,
    /Amount Paid:\s*\n?\s*([0-9,]+(?:\.\d{1,2})?)/i,
    /Amount:\s*\n?\s*ETB\s*([0-9,]+(?:\.\d{1,2})?)/i,
    /Transferred Amount:\s*\n?\s*ETB\s*([0-9,]+(?:\.\d{1,2})?)/i,
    /Transaction Amount:\s*\n?\s*ETB\s*([0-9,]+(?:\.\d{1,2})?)/i,
    /Amount Paid:\s*\n?\s*ETB\s*([0-9,]+(?:\.\d{1,2})?)/i,
    /ETB\s*([0-9,]+(?:\.\d{1,2})?)/i,
    /([0-9,]+(?:\.\d{1,2})?)\s*(?:Birr|Br)\b/i,
  ];
  for (const pattern of patterns) {
    const amountText = extractField(text, pattern);
    if (amountText) {
      const parsed = Number(amountText.replace(/,/g, ""));
      if (!isNaN(parsed) && parsed > 0) {
        return parsed;
      }
    }
  }
  // Fallback to match any decimal/integer near common currency markers.
  const fallbackMatch = text.match(/(?:ETB\s*)?([0-9,]+(?:\.\d{1,2})?)\s*(?:ETB|Birr|Br)\b/i);
  if (fallbackMatch) {
    return Number(fallbackMatch[1].replace(/,/g, ""));
  }
  return 0;
}

function extractPaymentDate(text) {
  const patterns = [
    /Payment Date\s*&?\s*Time:\s*\n?\s*([^\n]+)/i,
    /Transaction Date\s*&?\s*Time:\s*\n?\s*([^\n]+)/i,
    /Payment Date:\s*\n?\s*([^\n]+)/i,
    /Transaction Date:\s*\n?\s*([^\n]+)/i,
    /Date\s*&?\s*Time:\s*\n?\s*([^\n]+)/i,
    /Date:\s*\n?\s*([^\n]+)/i,
  ];
  for (const pattern of patterns) {
    const dateText = extractField(text, pattern);
    if (dateText) {
      const parsed = parseReceiptDate(dateText);
      if (!Number.isNaN(parsed.getTime())) {
        return parsed;
      }
    }
  }
  return null;
}

function parseReceiptDate(value) {
  const dateText = String(value || "").trim();
  const parsed = new Date(dateText);

  if (!Number.isNaN(parsed.getTime())) {
    return parsed;
  }

  const dateTimeMatch = dateText.match(/\b(\d{1,2})[/-](\d{1,2})[/-](\d{4})(?:[,\s]+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?)?/i);

  if (!dateTimeMatch) {
    return new Date(NaN);
  }

  let day = Number(dateTimeMatch[1]);
  let month = Number(dateTimeMatch[2]);
  const year = Number(dateTimeMatch[3]);
  let hour = Number(dateTimeMatch[4] || 0);
  const minute = Number(dateTimeMatch[5] || 0);
  const second = Number(dateTimeMatch[6] || 0);
  const meridiem = String(dateTimeMatch[7] || "").toUpperCase();

  if (month > 12 && day <= 12) {
    const originalDay = day;
    day = month;
    month = originalDay;
  }

  if (meridiem === "PM" && hour < 12) {
    hour += 12;
  } else if (meridiem === "AM" && hour === 12) {
    hour = 0;
  }

  return new Date(year, month - 1, day, hour, minute, second);
}

function extractReceiver(text) {
  const patterns = [
    /Receiver(?: Name)?:\s*\n?\s*([^\n]+)/i,
    /Credited Party Name:\s*\n?\s*([^\n]+)/i,
    /Beneficiary(?: Name)?:\s*\n?\s*([^\n]+)/i,
    /Transfer To:\s*\n?\s*([^\n]+)/i,
    /To:\s*\n?\s*([^\n]+)/i,
  ];

  for (const pattern of patterns) {
    const receiver = extractField(text, pattern);
    if (receiver) {
      return receiver;
    }
  }

  const expectedNamePattern = new RegExp(EXPECTED_RECEIVER_NAME.replace(/\s+/g, "\\s+"), "i");
  const expectedNameMatch = text.match(expectedNamePattern);
  return expectedNameMatch ? expectedNameMatch[0] : "";
}

function extractReceiverAccount(text) {
  const patterns = [
    /Receiver Account:\s*\n?\s*([0-9*Xx\s-]+)/i,
    /Receiver:[\s\S]*?Account:\s*\n?\s*([0-9*Xx\s-]+)/i,
    /Transfer To Account:\s*\n?\s*([0-9*Xx\s-]+)/i,
    /Credit Account:\s*\n?\s*([0-9*Xx\s-]+)/i,
    /Credited Account:\s*\n?\s*([0-9*Xx\s-]+)/i,
    /Beneficiary Account:\s*\n?\s*([0-9*Xx\s-]+)/i,
    /Account:\s*\n?\s*([0-9*Xx\s-]+)/i,
  ];
  for (const pattern of patterns) {
    const acc = extractField(text, pattern);
    if (acc) {
      return acc;
    }
  }
  // Try locating any 10-16 digit account number in the text ending with EXPECTED_RECEIVER_ACCOUNT_SUFFIX
  const allAccounts = text.match(/\b\d{10,16}\b/g);
  if (allAccounts) {
    const matching = allAccounts.find(function(acc) {
      return acc.endsWith(EXPECTED_RECEIVER_ACCOUNT_SUFFIX);
    });
    if (matching) {
      return matching;
    }
  }
  return "";
}

function extractReference(text) {
  const patterns = [
    /Reference No\.?\s*(?:\(VAT Invoice No\.\))?:\s*\n?\s*([A-Za-z0-9-]+)/i,
    /Reference Number:\s*\n?\s*([A-Za-z0-9-]+)/i,
    /Transaction Reference(?: Number)?:\s*\n?\s*([A-Za-z0-9-]+)/i,
    /Reference:\s*\n?\s*([A-Za-z0-9-]+)/i,
    /Transaction No\.?:\s*\n?\s*([A-Za-z0-9-]+)/i,
    /Transaction ID:\s*\n?\s*([A-Za-z0-9-]+)/i,
    /Payment Reference:\s*\n?\s*([A-Za-z0-9-]+)/i,
    /Ref No\.?:\s*\n?\s*([A-Za-z0-9-]+)/i,
    /Ref\.?:\s*\n?\s*([A-Za-z0-9-]+)/i,
  ];
  for (const pattern of patterns) {
    const ref = extractField(text, pattern);
    if (ref) {
      return ref;
    }
  }
  // Fallback: look for typical CBE reference number formats (e.g. starting with FT)
  const ftMatch = text.match(/\b(FT[A-Za-z0-9-]{8,20})\b/i);
  if (ftMatch) {
    return ftMatch[1];
  }
  return "";
}

function isDuplicateReference(spreadsheet, reference) {
  if (!reference) {
    return false;
  }

  return isDuplicateSheetValue(spreadsheet, "Reference No", reference);
}

function isDuplicateReceiptUrl(spreadsheet, receiptUrl) {
  const normalizedReceiptUrl = normalizeCbeReceiptUrl(receiptUrl);
  return normalizedReceiptUrl ? isDuplicateSheetValue(spreadsheet, "CBE Receipt Link", normalizedReceiptUrl) : false;
}

function isDuplicateSheetValue(spreadsheet, headerName, expectedValue) {
  const expectedCleanValue = normalizeComparableSheetValue(expectedValue);

  if (!expectedCleanValue) {
    return false;
  }

  return FINAL_REQUEST_SHEETS.some((sheetName) => {
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet || sheet.getLastRow() < 2) {
      return false;
    }

    const headerColumn = getHeaderColumn(sheet, headerName);

    if (!headerColumn) {
      return false;
    }

    const values = sheet.getRange(2, headerColumn, sheet.getLastRow() - 1, 1).getValues().flat();
    return values.some((value) => normalizeComparableSheetValue(value) === expectedCleanValue);
  });
}

function getHeaderColumn(sheet, headerName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(cleanCell);
  return headers.indexOf(headerName) + 1;
}

function normalizeComparableSheetValue(value) {
  return normalizeCbeReceiptUrl(String(value || "").replace(/^=HYPERLINK\("([^"]+)".*$/i, "$1")).trim().toLowerCase();
}

function appendRejected(spreadsheet, receivedAt, data, verification) {
  const audioLink = getExistingAudioLink(data);

  appendRequestRow(spreadsheet, REJECTED_SHEET_NAME, [
    receivedAt.toLocaleString(),
    cleanCell(data.id),
    "Rejected",
    false,
    cleanCell(data.name),
    cleanCell(data.phone),
    cleanCell(data.badNews),
    cleanCell(data.receiver),
    formatClaimedPayment(data),
    verification.amount || "",
    cleanCell(verification.payer),
    cleanCell(verification.receiver),
    cleanCell(verification.receiverAccount),
    cleanCell(verification.reference),
    receiptLinkFormula(data.cbeReceiptUrl),
    audioLink,
    cleanCell(data.receiptFile),
    cleanCell(data.language),
    verification.errors.join(" | "),
  ]);
}

function appendIncoming(spreadsheet, receivedAt, data, audioLink) {
  const sheet = getSheet(spreadsheet, INCOMING_SHEET_NAME);
  ensureIncomingHeader(sheet);
  sheet.appendRow([
    receivedAt.toLocaleString(),
    cleanCell(data.id),
    "Incoming",
    false,
    cleanCell(data.name),
    cleanCell(data.phone),
    cleanCell(data.badNews),
    cleanCell(data.receiver),
    formatClaimedPayment(data),
    "",
    "",
    "",
    "",
    "",
    receiptLinkFormula(data.cbeReceiptUrl),
    audioLink,
    cleanCell(data.receiptFile),
    cleanCell(data.language),
    "",
  ]);
}

function formatClaimedPayment(data) {
  const amount = cleanCell(data.paymentAmount);
  const currency = cleanCell(data.paymentCurrency || (data.paymentMethod === "paypal" ? "USD" : "ETB"));
  const method = cleanCell(data.paymentMethod || "cbe").toUpperCase();
  const tier = cleanCell(data.serviceTier);
  const contact = cleanCell(data.contactChannel);
  const specialAmount = cleanCell(data.specialRequestAmount);
  const parts = [`${amount} ${currency}`, method];

  if (tier) {
    parts.push(tier);
  }

  if (specialAmount) {
    parts.push(`Special: ${specialAmount} ${currency}`);
  }

  if (contact) {
    parts.push(`Contact: ${contact}`);
  }

  return parts.join(" | ");
}

function saveAudioFile(data) {
  if (!data.audioDataBase64 || !data.audioFile || data.audioFile === "-") {
    return cleanCell(data.audioFile);
  }

  try {
    const folder = getOrCreateDriveFolder(AUDIO_FOLDER_NAME);
    const safeName = cleanCell(data.audioFile).replace(/[\\/:*?"<>|]/g, "-") || "audio";
    const submissionId = cleanCell(data.id) || String(Date.now());
    const fileName = `${submissionId}-${safeName}`;
    const bytes = Utilities.base64Decode(data.audioDataBase64);
    const blob = Utilities.newBlob(bytes, cleanCell(data.audioMimeType) || "audio/mpeg", fileName);
    const file = folder.createFile(blob);

    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (error) {
      // Some Workspace policies block link sharing; the owner can still open the file.
    }

    data.audioDriveUrl = file.getUrl();
    return `=HYPERLINK("${data.audioDriveUrl}", "${fileName}")`;
  } catch (error) {
    data.audioDriveUrl = "";
    return `Audio save failed: ${String(error && error.message ? error.message : error)}`;
  }
}

function getExistingAudioLink(data) {
  if (data.audioDriveUrl) {
    return `=HYPERLINK("${data.audioDriveUrl}", "${cleanCell(data.audioFile) || "Open audio"}")`;
  }

  return cleanCell(data.audioFile);
}

function getOrCreateDriveFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
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
  sheet.appendRow(normalizeRequestRow(row));
}

function normalizeRequestRow(row) {
  const normalizedRow = row.slice(0, REQUEST_HEADERS.length);

  while (normalizedRow.length < REQUEST_HEADERS.length) {
    normalizedRow.push("");
  }

  [10, 11, 12, 13].forEach((index) => {
    normalizedRow[index] = cleanReceiptField(normalizedRow[index]);
  });

  return normalizedRow;
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
    ensureDoneNewsColumn(sheet, headers);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), headers.length).setVerticalAlignment("top");
  sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), headers.length).setWrap(true);
  ensureDoneNewsCheckboxes(sheet, headers);
  sheet.autoResizeColumns(1, headers.length);
}

function ensureDoneNewsColumn(sheet, headers) {
  const doneNewsIndex = headers.indexOf("Done News");

  if (doneNewsIndex === -1 || sheet.getLastColumn() === 0) {
    return;
  }

  const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(cleanCell);

  if (existingHeaders.indexOf("Done News") !== -1) {
    return;
  }

  const previousHeader = headers[doneNewsIndex - 1];
  const previousHeaderColumn = existingHeaders.indexOf(previousHeader) + 1;

  if (previousHeaderColumn > 0) {
    sheet.insertColumnAfter(previousHeaderColumn);
  }
}

function ensureDoneNewsCheckboxes(sheet, headers) {
  const doneNewsColumn = headers.indexOf("Done News") + 1;

  if (!doneNewsColumn) {
    return;
  }

  const rowCount = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, doneNewsColumn, rowCount, 1).insertCheckboxes();
}

function cleanCell(value) {
  return String(value || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").trim();
}

function receiptLinkFormula(url) {
  const cleanUrl = cleanCell(url);
  if (!cleanUrl) {
    return "";
  }

  if (/^https?:\/\//i.test(cleanUrl)) {
    return `=HYPERLINK("${cleanUrl}", "Open CBE receipt")`;
  }

  return cleanUrl;
}

function jsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}
