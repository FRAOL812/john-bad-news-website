const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

const webhookSource = fs.readFileSync(path.join(__dirname, "..", "google-sheets-webhook.gs"), "utf8");

function receiptHtml({
  reference = "DGP17W7401",
  settledAmount = 50,
  totalPaidAmount = settledAmount + 2,
  receiver = "Fraol Eshetu Hailu",
  receiverAccount = "2519****5322",
  status = "Completed",
} = {}) {
  return `<!doctype html>
    <html><head><title>telebirr receipt</title></head><body>
      <table>
        <tr><td>telebirr Transaction information</td></tr>
        <tr><td>Payer Name</td><td>Test Customer</td></tr>
        <tr><td>Credited Party name</td><td>${receiver}</td></tr>
        <tr><td>Credited party account no</td><td>${receiverAccount}</td></tr>
        <tr><td>transaction status<td>${status}</td></tr>
        <tr><td>Invoice details</td></tr>
        <tr><td>Invoice No.</td><td>Payment date</td><td>Settled Amount</td></tr>
        <tr><td>${reference}</td><td>25-07-2024 01:24:29</td><td>${settledAmount} Birr</td></tr>
        <tr><td>Service fee</td><td>2 Birr</td></tr>
        <tr><td>Total Paid Amount</td><td>${totalPaidAmount} Birr</td></tr>
      </table>
    </body></html>`;
}

function loadWebhook({ html = receiptHtml(), statusCode = 200, failHostnameFetch = false, failAllFetches = false } = {}) {
  const fetchCalls = [];
  const context = {
    console,
    UrlFetchApp: {
      fetch(url, options) {
        fetchCalls.push({ url, options });
        if (failAllFetches || (failHostnameFetch && url.startsWith("https://transactioninfo.ethiotelecom.et/"))) {
          throw new Error(`Address unavailable: ${url}`);
        }
        return {
          getResponseCode: () => statusCode,
          getContentText: () => html,
        };
      },
    },
  };
  vm.createContext(context);
  vm.runInContext(`${webhookSource}\n;globalThis.__testApi = { verifyTelebirrReceipt, verifyPaypalReceipt, queueTelebirrForManualVerification };`, context);
  return { verify: context.__testApi.verifyTelebirrReceipt, verifyPaypal: context.__testApi.verifyPaypalReceipt, queuePending: context.__testApi.queueTelebirrForManualVerification, fetchCalls };
}

function submission({ tier = "basic", special = 0, reference = "DGP17W7401" } = {}) {
  return {
    paymentMethod: "telebirr",
    serviceTier: `${tier} - test`,
    specialRequestAmount: special ? String(special) : "",
    receiptVerificationValue: `https://transactioninfo.ethiotelecom.et/receipt/${reference}`,
    receiptOcrText: `telebirr Payer Name: Test Customer Transaction Number: ${reference}`,
    telebirrReceiptVerified: true,
    receiptFile: "receipt.jpg",
  };
}

function verify(options = {}) {
  const runtime = loadWebhook({ html: options.html, statusCode: options.statusCode, failHostnameFetch: options.failHostnameFetch, failAllFetches: options.failAllFetches });
  const result = runtime.verify(options.data || submission(), new Date("2026-08-16T12:00:00Z"));
  return { result, fetchCalls: runtime.fetchCalls };
}

test("basic plan verifies the settled 50 ETB, not the fee-inclusive total", () => {
  const { result, fetchCalls } = verify();
  assert.equal(result.ok, true);
  assert.equal(result.amount, 50);
  assert.equal(result.reference, "DGP17W7401");
  assert.equal(result.payer, "Test Customer");
  assert.equal(result.receiver, "Fraol Eshetu Hailu");
  assert.equal(fetchCalls.length, 1);
  assert.equal(fetchCalls[0].url, "https://transactioninfo.ethiotelecom.et/receipt/DGP17W7401");
});

for (const paypalPlan of [
  { name: "basic", tier: "basic", amount: 15 },
  { name: "urgent", tier: "urgent", amount: 25 },
]) {
  test(`PayPal ${paypalPlan.name} confirmation requires ${paypalPlan.amount} USD`, () => {
    const runtime = loadWebhook();
    const result = runtime.verifyPaypal({
      paymentMethod: "paypal",
      serviceTier: paypalPlan.tier,
      receiptOcrText: `PayPal Payment to Yonatan Woldegiorgis USD ${paypalPlan.amount}.00 Transaction ID: PAYPAL12345`,
      paypalReceiptVerified: true,
      receiptFile: "paypal-receipt.png",
    });
    assert.equal(result.ok, true);
    assert.equal(result.amount, paypalPlan.amount);
  });
}

test("PayPal mobile details screenshot verifies without a printed PayPal heading", () => {
  const runtime = loadWebhook();
  const result = runtime.verifyPaypal({
    paymentMethod: "paypal",
    serviceTier: "basic",
    receiptOcrText: "Total $15.00 Shipping info Yonatan Woldegiorgis 4813 N Borthwick Ave Portland, OR 97217 United States Transaction ID 0NW80256AE058180W Need help? Get answers Send again",
    paypalReceiptVerified: true,
    receiptFile: "paypal-mobile-details.png",
  });
  assert.equal(result.ok, true);
  assert.equal(result.amount, 15);
  assert.equal(result.reference, "0NW80256AE058180W");
});

test("PayPal-like mobile details still rejects the wrong recipient", () => {
  const runtime = loadWebhook();
  const result = runtime.verifyPaypal({
    paymentMethod: "paypal",
    serviceTier: "basic",
    receiptOcrText: "Total $15.00 Shipping info Somebody Else Transaction ID 0NW80256AE058180W",
    paypalReceiptVerified: true,
    receiptFile: "wrong-recipient.png",
  });
  assert.equal(result.ok, false);
  assert.match(result.errors.join(" "), /PayPal recipient must be/);
});

for (const paymentPlan of [
  { name: "basic", tier: "basic", special: 0, amount: 50 },
  { name: "urgent", tier: "urgent", special: 0, amount: 200 },
  { name: "basic plus special request", tier: "basic", special: 50, amount: 100 },
  { name: "urgent plus special request", tier: "urgent", special: 125, amount: 325 },
]) {
  test(`${paymentPlan.name} plan verifies against the official settled amount`, () => {
    const data = submission({ tier: paymentPlan.tier, special: paymentPlan.special });
    const html = receiptHtml({ settledAmount: paymentPlan.amount });
    const { result } = verify({ data, html });
    assert.equal(result.ok, true);
    assert.equal(result.amount, paymentPlan.amount);
  });
}

test("rejects an official settled amount that does not match the selected plan", () => {
  const { result } = verify({ html: receiptHtml({ settledAmount: 49 }) });
  assert.equal(result.ok, false);
  assert.match(result.errors.join(" "), /Receipt amount must be 50 ETB/);
});

test("rejects a receipt for a different recipient", () => {
  const html = receiptHtml({ receiver: "Different Person", receiverAccount: "2519****0000" });
  const { result } = verify({ html });
  assert.equal(result.ok, false);
  assert.match(result.errors.join(" "), /Telebirr recipient must be/);
});

test("rejects a transaction that is not completed", () => {
  const { result } = verify({ html: receiptHtml({ status: "Pending" }) });
  assert.equal(result.ok, false);
  assert.match(result.errors.join(" "), /transaction is not completed/);
});

test("rejects when the official reference differs from the receipt link", () => {
  const { result } = verify({ html: receiptHtml({ reference: "OTHER12345" }) });
  assert.equal(result.ok, false);
  assert.match(result.errors.join(" "), /reference does not match/);
});

test("rejects invalid and unavailable official receipt pages", () => {
  const invalidPage = verify({ html: "<html><body>Not a receipt</body></html>" }).result;
  assert.equal(invalidPage.ok, false);
  assert.match(invalidPage.errors.join(" "), /invalid receipt page/);

  const unavailablePage = verify({ statusCode: 503 }).result;
  assert.equal(unavailablePage.ok, false);
  assert.match(unavailablePage.errors.join(" "), /HTTP 503/);
});

test("marks the receipt pending when Apps Script cannot resolve the official receipt hostname", () => {
  const { result, fetchCalls } = verify({ failHostnameFetch: true });
  assert.equal(result.ok, false);
  assert.equal(result.pending, true);
  assert.equal(result.errors.length, 0);
  assert.equal(result.reference, "DGP17W7401");
  assert.equal(fetchCalls.length, 3);
});

test("queues a valid Telebirr submission immediately without a remote request", () => {
  const runtime = loadWebhook();
  const result = runtime.queuePending(submission());
  assert.equal(result.ok, false);
  assert.equal(result.pending, true);
  assert.equal(result.reference, "DGP17W7401");
  assert.equal(result.payer, "Test Customer");
  assert.equal(result.receiver, "Fraol Eshetu Hailu (pending verification)");
  assert.equal(result.receiverAccount, "0913885322 (pending verification)");
  assert.equal(result.errors.length, 0);
  assert.equal(runtime.fetchCalls.length, 0);
});
