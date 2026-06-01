const VERIFIED_SHEET_NAME = "Received News";
const REJECTED_SHEET_NAME = "Rejected Payments";
const INCOMING_SHEET_NAME = "Incoming Requests";
const ERROR_SHEET_NAME = "Webhook Errors";
const WEBHOOK_VERSION = "2026-06-01-cbe-link-testing";
const BASE_PRICE_BIRR = 500;
const URGENT_PRICE_BIRR = 2000;
const PAYPAL_BASE_PRICE_USD = 25;
const PAYPAL_URGENT_PRICE_USD = 100;
const PAYMENT_WINDOW_MINUTES = 60;
const ALLOW_CBE_LINK_TESTING_FALLBACK = false;
const EXPECTED_RECEIVER_NAME = "Fraol Eshetu Hailu";
const EXPECTED_RECEIVER_ACCOUNT_NUMBER = "1000239878583";
const EXPECTED_RECEIVER_ACCOUNT_SUFFIX = "8583";
const EXPECTED_RECEIVER_ACCOUNT_LOOKUP_SUFFIX = EXPECTED_RECEIVER_ACCOUNT_NUMBER.substring(5);
const CBE_RECEIPT_BASE_URL = "https://mbreciept.cbe.com.et/";
const CBE_LEGACY_RECEIPT_BASE_URL = "https://apps.cbe.com.et:100/";
const CBE_PDF_VERIFICATION_BASE_URL = "https://apps.cbe.com.et:100";
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

function doGet() {
  return jsonResponse({ ok: true, version: WEBHOOK_VERSION });
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
  const expectedAmount = getExpectedCbeAmount(data);

  if (receiptUrl.toLowerCase().startsWith("ocr:")) {
    const reference = cleanCbeReference(receiptUrl.substring(4).trim());
    if (!reference) {
      errors.push("Invalid OCR Reference Number");
    }

    return errors.length ? { ok: false, errors } : verifyCbeReceiptReferenceFromPdf(reference, receivedAt, expectedAmount);
  }

  const url = normalizeCbeReceiptUrl(receiptUrl);

  if (!isCbeReceiptUrl(url)) {
    return { ok: false, errors: ["Missing or invalid CBE receipt URL"] };
  }

  const pdfVerification = verifyCbeReceiptFromPdf(url, receivedAt, expectedAmount);
  if (pdfVerification.ok) {
    return pdfVerification;
  }

  const apiVerification = verifyCbeReceiptFromApi(url, receivedAt, expectedAmount);
  if (!apiVerification.ok && data.receiptOcrText) {
    const smsVerification = verifyCbeSmsReceiptText(data.receiptOcrText, url, receivedAt, expectedAmount);
    if (smsVerification.ok) {
      return smsVerification;
    }

    const testingVerification = verifyCbeReceiptLinkForTesting(data, url, receivedAt);
    if (testingVerification.ok) {
      return testingVerification;
    }

    return mergeVerificationFailures(
      mergeVerificationFailures(mergeVerificationFailures(pdfVerification, apiVerification), smsVerification),
      testingVerification,
    );
  }

  if (!apiVerification.ok) {
    const testingVerification = verifyCbeReceiptLinkForTesting(data, url, receivedAt);
    if (testingVerification.ok) {
      return testingVerification;
    }

    return mergeVerificationFailures(mergeVerificationFailures(pdfVerification, apiVerification), testingVerification);
  }

  return apiVerification;
}

function verifyCbeReceiptFromApi(receiptUrl, receivedAt, expectedAmount) {
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

  return buildCbeVerificationFromTransaction(transaction, receivedAt, expectedAmount);
}

function verifyCbeReceiptFromPdf(receiptUrl, receivedAt, expectedAmount) {
  try {
    const reference = extractCbeReferenceFromReceiptUrl(receiptUrl);

    if (!reference) {
      return {
        ok: false,
        errors: [`CBE receipt link does not include an FT reference for PDF verification (${WEBHOOK_VERSION})`],
      };
    }

    return verifyCbeReceiptReferenceFromPdf(reference, receivedAt, expectedAmount);
  } catch (error) {
    return {
      ok: false,
      errors: [`CBE PDF verification failed: ${String(error && error.message ? error.message : error)} (${WEBHOOK_VERSION})`],
    };
  }
}

function verifyCbeReceiptReferenceFromPdf(reference, receivedAt, expectedAmount) {
  const verificationId = buildCbePdfVerificationId(reference);
  const fetchResult = fetchCbePdfVerificationText(verificationId);

  if (!fetchResult.ok) {
    return fetchResult;
  }

  return buildCbeVerificationFromReceiptText(fetchResult.text, receivedAt, reference, expectedAmount);
}

function buildCbePdfVerificationId(reference) {
  const cleanReference = cleanCbeReference(reference);
  return `${cleanReference}${EXPECTED_RECEIVER_ACCOUNT_LOOKUP_SUFFIX}`;
}

function fetchCbePdfVerificationText(verificationId) {
  const urls = [
    `${CBE_PDF_VERIFICATION_BASE_URL}/${encodeURIComponent(verificationId)}`,
    `${CBE_PDF_VERIFICATION_BASE_URL}/?id=${encodeURIComponent(verificationId)}`,
  ];
  const errors = [];

  for (let index = 0; index < urls.length; index += 1) {
    const url = urls[index];
    let response;
    try {
      response = UrlFetchApp.fetch(url, {
        method: "get",
        muteHttpExceptions: true,
        followRedirects: true,
      });
    } catch (error) {
      errors.push(`Could not reach ${url}: ${String(error && error.message ? error.message : error)}`);
      continue;
    }

    const statusCode = response.getResponseCode();
    if (statusCode < 200 || statusCode >= 300) {
      errors.push(`CBE PDF verification returned HTTP ${statusCode} for ${url}`);
      continue;
    }

    const blob = response.getBlob();
    const text = extractTextFromCbeVerificationResponse(blob, response.getContentText() || "", verificationId);
    if (text) {
      return { ok: true, errors: [], text };
    }

    errors.push(`CBE PDF verification returned unreadable receipt text for ${url}`);
  }

  return {
    ok: false,
    errors: errors.length ? errors : [`CBE PDF verification failed for ${verificationId} (${WEBHOOK_VERSION})`],
  };
}

function extractTextFromCbeVerificationResponse(blob, responseText, verificationId) {
  const contentType = String(blob.getContentType() || "").toLowerCase();
  const textPreview = String(responseText || "").trim();

  if (contentType.indexOf("pdf") === -1 && textPreview && textPreview.indexOf("%PDF") !== 0) {
    return normalizeReceiptText(textPreview);
  }

  return extractTextFromPdfBlob(blob, verificationId);
}

function extractTextFromPdfBlob(blob, verificationId) {
  if (typeof Drive === "undefined" || !Drive.Files || !Drive.Files.insert) {
    throw new Error("Advanced Drive service is required for CBE PDF receipt OCR. Enable Drive API in Apps Script.");
  }

  const pdfName = `cbe-verification-${verificationId}.pdf`;
  const pdfBlob = blob.setName(pdfName);
  let convertedFile;

  try {
    convertedFile = Drive.Files.insert(
      {
        title: `cbe-verification-${verificationId}`,
        mimeType: MimeType.GOOGLE_DOCS,
      },
      pdfBlob,
      {
        convert: true,
        ocr: true,
        ocrLanguage: "en",
      },
    );

    const document = DocumentApp.openById(convertedFile.id);
    return normalizeReceiptText(document.getBody().getText());
  } finally {
    if (convertedFile && convertedFile.id) {
      try {
        DriveApp.getFileById(convertedFile.id).setTrashed(true);
      } catch (error) {
        // Best effort cleanup only.
      }
    }
  }
}

function buildCbeVerificationFromReceiptText(text, receivedAt, expectedReference, expectedAmount) {
  const amount = extractAmount(text);
  const paymentDate = extractPaymentDate(text);
  const receiver = cleanReceiptField(extractReceiver(text));
  const receiverAccount = cleanReceiptField(extractReceiverAccount(text));
  const payer = cleanReceiptField(extractPayer(text));
  const reference = cleanCbeReference(extractReference(text)) || cleanCbeReference(expectedReference);
  const errors = [];

  appendCbeAmountErrors(errors, amount, expectedAmount);

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

function verifyCbeSmsReceiptText(text, receiptUrl, receivedAt, expectedAmount) {
  const receiptText = normalizeReceiptText(text);
  const amount = extractAmount(receiptText);
  const paymentDate = extractPaymentDate(receiptText) || extractSmsPaymentDate(receiptText, receivedAt);
  const receiver = cleanReceiptField(extractReceiver(receiptText));
  const receiverAccount = cleanReceiptField(extractReceiverAccount(receiptText));
  const payer = cleanReceiptField(extractSmsPayer(receiptText) || extractPayer(receiptText));
  const reference = cleanReceiptField(extractCbeReceiptToken(receiptUrl) || extractCbeReceiptToken(receiptText) || receiptUrl);
  const errors = [];

  if (!isCbeReceiptUrl(receiptUrl) || receiptText.toLowerCase().indexOf("cbe") === -1) {
    errors.push("CBE SMS receipt text was not found");
  }

  appendCbeAmountErrors(errors, amount, expectedAmount);

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

function verifyCbeReceiptLinkForTesting(data, receiptUrl, receivedAt) {
  const amount = Number(data.paymentAmount) || parseReceiptNumber(data.paymentAmount);
  const reference = cleanReceiptField(extractCbeReceiptToken(receiptUrl));
  const errors = [];

  if (!ALLOW_CBE_LINK_TESTING_FALLBACK) {
    errors.push("CBE receipt link testing fallback is disabled");
  }

  if (!isCbeReceiptUrl(receiptUrl)) {
    errors.push("Missing or invalid CBE receipt URL");
  }

  if (!reference) {
    errors.push("CBE receipt link is missing the receipt token");
  }

  appendCbeAmountErrors(errors, amount, getExpectedCbeAmount(data));

  return {
    ok: errors.length === 0,
    errors,
    amount,
    paymentDate: receivedAt,
    receiver: EXPECTED_RECEIVER_NAME,
    receiverAccount: `*${EXPECTED_RECEIVER_ACCOUNT_SUFFIX}`,
    payer: "CBE receipt link (testing fallback)",
    reference,
  };
}

function mergeVerificationFailures(primary, secondary) {
  return {
    ok: false,
    errors: []
      .concat(primary && primary.errors ? primary.errors : [])
      .concat(secondary && secondary.errors ? secondary.errors : []),
    amount: primary.amount || secondary.amount || "",
    paymentDate: primary.paymentDate || secondary.paymentDate || null,
    receiver: primary.receiver || secondary.receiver || "",
    receiverAccount: primary.receiverAccount || secondary.receiverAccount || "",
    payer: primary.payer || secondary.payer || "",
    reference: primary.reference || secondary.reference || "",
  };
}

function extractCbeApiErrorMessage(responseText) {
  try {
    const payload = JSON.parse(responseText || "{}");
    return cleanCell(payload.detail || payload.message || payload.title || "");
  } catch (error) {
    return "";
  }
}

function buildCbeVerificationFromTransaction(transaction, receivedAt, expectedAmount) {
  const amount = parseReceiptNumber(transaction.amountCredited || transaction.creditAmount || transaction.debitAmount);
  const paymentDate = parseCbeApiDate(transaction.dateTimes && transaction.dateTimes[0]) || parseCbeApiDate(transaction.processingDate);
  const receiver = cleanReceiptField(transaction.creditAccountHolder);
  const receiverAccount = cleanReceiptField(transaction.creditAccountNo);
  const payer = cleanReceiptField(transaction.debitAccountHolder);
  const reference = cleanReceiptField(transaction.id);
  const errors = [];

  appendCbeAmountErrors(errors, amount, expectedAmount);

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

function getExpectedCbeAmount(data) {
  return getExpectedPaymentAmount(data);
}

function getExpectedPaymentAmount(data) {
  const paymentMethod = String(data && data.paymentMethod ? data.paymentMethod : "cbe").toLowerCase();
  const serviceTier = String(data && data.serviceTier ? data.serviceTier : "basic").toLowerCase();
  const specialAmount = parseReceiptNumber(data && data.specialRequestAmount);
  let baseAmount;

  if (paymentMethod === "paypal") {
    baseAmount = serviceTier.indexOf("urgent") !== -1 ? PAYPAL_URGENT_PRICE_USD : PAYPAL_BASE_PRICE_USD;
  } else {
    baseAmount = serviceTier.indexOf("urgent") !== -1 ? URGENT_PRICE_BIRR : BASE_PRICE_BIRR;
  }

  return baseAmount + specialAmount;
}

function appendClaimedAmountErrors(errors, claimedAmount, expectedAmount, currency) {
  const claimedCents = Math.round(Number(claimedAmount || 0) * 100);
  const expectedCents = Math.round(Number(expectedAmount || 0) * 100);

  if (!claimedCents) {
    errors.push("Claimed payment amount not found");
    return;
  }

  if (claimedCents !== expectedCents) {
    errors.push(`Claimed payment amount must match the selected payment option (${formatMoneyAmount(expectedAmount)} ${currency})`);
  }
}

function appendCbeAmountErrors(errors, amount, expectedAmount) {
  const minimumCents = Math.round(BASE_PRICE_BIRR * 100);
  const expectedCents = Math.round(Number(expectedAmount || 0) * 100);
  const amountCents = Math.round(Number(amount || 0) * 100);

  if (!amountCents) {
    errors.push("Transferred amount not found");
    return;
  }

  if (amountCents < minimumCents) {
    errors.push(`Transferred amount is below ${BASE_PRICE_BIRR} ETB`);
    return;
  }

  if (expectedCents && amountCents !== expectedCents) {
    errors.push(`Transferred amount must match the selected CBE payment option (${formatMoneyAmount(expectedAmount)} ETB)`);
  }
}

function formatMoneyAmount(amount) {
  return Number(amount || 0).toLocaleString("en-US", {
    maximumFractionDigits: 2,
    minimumFractionDigits: Number(amount || 0) % 1 === 0 ? 0 : 2,
  });
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

function extractCbeReferenceFromReceiptUrl(receiptUrl) {
  return cleanCbeReference(extractCbeReceiptToken(receiptUrl)) || cleanCbeReference(receiptUrl);
}

function cleanCbeReference(value) {
  const match = String(value || "").match(/\b(FT[A-Za-z0-9]{10})\b/i);
  return match ? match[1].toUpperCase() : "";
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
  const expectedAmount = getExpectedPaymentAmount(data);
  const claimedAmount = parseReceiptNumber(data.paymentAmount);
  const receiptAmount = extractPaypalAmount(data.receiptOcrText);
  const errors = [];

  if (!data.paypalReceiptVerified) {
    errors.push("PayPal receipt screenshot was not verified");
  }

  appendClaimedAmountErrors(errors, claimedAmount, expectedAmount, "USD");

  if (!receiptAmount) {
    errors.push("PayPal receipt amount not found");
  } else if (Math.round(receiptAmount * 100) !== Math.round(expectedAmount * 100)) {
    errors.push(`PayPal receipt amount must match the selected payment option (${formatMoneyAmount(expectedAmount)} USD)`);
  }

  if (errors.length) {
    return { ok: false, errors };
  }

  return {
    ok: true,
    errors: [],
    amount: receiptAmount,
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
    /ETB\s*([0-9,]+(?:\.\d{1,2})?)/i,
    /Amount\s+([0-9,]+(?:\.\d{1,2})?)\s*ETB/i,
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

function extractPaypalAmount(text) {
  const receiptText = normalizeReceiptText(text || "");
  const patterns = [
    /(?:USD|\$)\s*([0-9,]+(?:\.\d{1,2})?)/i,
    /([0-9,]+(?:\.\d{1,2})?)\s*USD\b/i,
    /(?:Amount|Total|Paid|Payment):?\s*\n?\s*(?:USD|\$)?\s*([0-9,]+(?:\.\d{1,2})?)/i,
  ];

  for (const pattern of patterns) {
    const amountText = extractField(receiptText, pattern);
    if (amountText) {
      const parsed = Number(amountText.replace(/,/g, ""));
      if (!isNaN(parsed) && parsed > 0) {
        return parsed;
      }
    }
  }

  return 0;
}

function extractPaymentDate(text) {
  const patterns = [
    /Payment Date\s+\w+\s+([A-Za-z]{3}\s+\d{1,2}\s+\d{4})/i,
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

  const monthNameMatch = dateText.match(/\b([A-Za-z]{3,9})\s+(\d{1,2})\s+(\d{4})(?:[,\s]+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?)?/i);
  if (monthNameMatch) {
    const monthNames = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
    const month = monthNames.indexOf(monthNameMatch[1].substring(0, 3).toLowerCase());
    if (month !== -1) {
      let hour = Number(monthNameMatch[4] || 0);
      const minute = Number(monthNameMatch[5] || 0);
      const second = Number(monthNameMatch[6] || 0);
      const meridiem = String(monthNameMatch[7] || "").toUpperCase();
      if (meridiem === "PM" && hour < 12) {
        hour += 12;
      } else if (meridiem === "AM" && hour === 12) {
        hour = 0;
      }
      return new Date(Number(monthNameMatch[3]), month, Number(monthNameMatch[2]), hour, minute, second);
    }
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
    /Receiver\s+([A-Za-z]+(?:\s+[A-Za-z]+){1,3})/i,
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
    /Receiver\s+[A-Za-z]+(?:\s+[A-Za-z]+){1,3}\s*Account\s*([0-9*Xx\s-]+)/i,
    /Receiver Account:\s*\n?\s*([0-9*Xx\s-]+)/i,
    /Receiver:[\s\S]*?Account:\s*\n?\s*([0-9*Xx\s-]+)/i,
    /to account\s*([0-9*Xx\s-]+)/i,
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

function extractSmsPayer(text) {
  return extractField(text, /Dear\s+([A-Za-z]+(?:\s+[A-Za-z]+){1,5})\s+You\s+have/i);
}

function extractSmsPaymentDate(text, receivedAt) {
  const match = String(text || "").match(/\b(\d{1,2}):(\d{2})\s*(AM|PM)\s*,\s*([A-Za-z]{3,9})\s+(\d{1,2})\b/i);
  if (!match) {
    return null;
  }

  const monthNames = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
  const month = monthNames.indexOf(match[4].substring(0, 3).toLowerCase());
  if (month === -1) {
    return null;
  }

  let hour = Number(match[1]);
  const minute = Number(match[2]);
  const meridiem = match[3].toUpperCase();

  if (meridiem === "PM" && hour < 12) {
    hour += 12;
  } else if (meridiem === "AM" && hour === 12) {
    hour = 0;
  }

  return new Date(receivedAt.getFullYear(), month, Number(match[5]), hour, minute, 0);
}

function extractPayer(text) {
  const patterns = [
    /Payer(?: Name)?:\s*\n?\s*([^\n]+)/i,
    /Payer\s+([A-Za-z]+(?:\s+[A-Za-z]+){1,3})/i,
    /Debited Party Name:\s*\n?\s*([^\n]+)/i,
    /Sender(?: Name)?:\s*\n?\s*([^\n]+)/i,
    /Transfer From:\s*\n?\s*([^\n]+)/i,
    /From:\s*\n?\s*([^\n]+)/i,
  ];

  for (const pattern of patterns) {
    const payer = extractField(text, pattern);
    if (payer) {
      return payer;
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
  writeRequestRow(sheet, [
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
  writeRequestRow(sheet, row);
}

function writeRequestRow(sheet, row) {
  const rowNumber = getLastRequestDataRow(sheet, REQUEST_HEADERS) + 1;
  sheet.getRange(rowNumber, 1, 1, REQUEST_HEADERS.length).setValues([normalizeRequestRow(row)]);
  ensureDoneNewsCheckboxForRow(sheet, REQUEST_HEADERS, rowNumber);
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
  const lastDataRow = getLastRequestDataRow(sheet, headers);
  sheet.getRange(1, 1, Math.max(lastDataRow, 1), headers.length).setVerticalAlignment("top");
  sheet.getRange(1, 1, Math.max(lastDataRow, 1), headers.length).setWrap(true);
  ensureDoneNewsCheckboxes(sheet, headers, lastDataRow);
  clearUnusedDoneNewsCheckboxes(sheet, headers, lastDataRow);
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

function ensureDoneNewsCheckboxes(sheet, headers, lastDataRow) {
  const doneNewsColumn = headers.indexOf("Done News") + 1;

  if (!doneNewsColumn || lastDataRow < 2) {
    return;
  }

  sheet.getRange(2, doneNewsColumn, lastDataRow - 1, 1).insertCheckboxes();
}

function ensureDoneNewsCheckboxForRow(sheet, headers, rowNumber) {
  const doneNewsColumn = headers.indexOf("Done News") + 1;

  if (!doneNewsColumn || rowNumber < 2) {
    return;
  }

  sheet.getRange(rowNumber, doneNewsColumn, 1, 1).insertCheckboxes();
}

function clearUnusedDoneNewsCheckboxes(sheet, headers, lastDataRow) {
  const doneNewsColumn = headers.indexOf("Done News") + 1;
  const firstUnusedRow = Math.max(lastDataRow + 1, 2);
  const rowCount = sheet.getMaxRows() - firstUnusedRow + 1;

  if (!doneNewsColumn || rowCount < 1) {
    return;
  }

  sheet.getRange(firstUnusedRow, doneNewsColumn, rowCount, 1).clearContent().clearDataValidations();
}

function getLastRequestDataRow(sheet, headers) {
  const rowCount = Math.max(sheet.getMaxRows() - 1, 0);

  if (!rowCount) {
    return 1;
  }

  const doneNewsIndex = headers.indexOf("Done News");
  const values = sheet.getRange(2, 1, rowCount, headers.length).getValues();

  for (let rowIndex = values.length - 1; rowIndex >= 0; rowIndex -= 1) {
    const hasData = values[rowIndex].some((value, columnIndex) => {
      return columnIndex !== doneNewsIndex && cleanCell(value) !== "";
    });

    if (hasData) {
      return rowIndex + 2;
    }
  }

  return 1;
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
