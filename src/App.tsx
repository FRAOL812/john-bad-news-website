import { FormEvent, useState } from "react";

type IconName =
  | "shield"
  | "bolt"
  | "users"
  | "hat"
  | "check"
  | "card"
  | "send"
  | "phone"
  | "user"
  | "message"
  | "receipt"
  | "upload"
  | "bank"
  | "clock";

type Feature = {
  icon: IconName;
  title: string;
  copy: string;
};

type Step = {
  icon: IconName;
  title: string;
  copy: string;
};

type PlatformName = "facebook" | "telegram" | "whatsapp" | "instagram";

type SubmissionRecord = {
  id: string;
  receivedAt: string;
  name: string;
  phone: string;
  badNews: string;
  receiver: string;
  paymentAmount: string;
  paymentStatus: string;
  cbeReceiptUrl: string;
  audioFile: string;
  audioMimeType?: string;
  audioSize?: number;
  audioDataBase64?: string;
  receiptFile: string;
  language: "en" | "am";
};

type RuntimeConfig = {
  API_BASE?: string;
  SPREADSHEET_WEBHOOK_URL?: string;
};

declare global {
  interface Window {
    BarcodeDetector?: {
      new (options?: { formats?: string[] }): {
        detect(source: ImageBitmapSource): Promise<Array<{ rawValue: string }>>;
      };
    };
    __RUNTIME_CONFIG__?: RuntimeConfig;
  }
}

const basePriceBirr = 500;
const maxReceiptSize = 5 * 1024 * 1024;
const maxAudioSize = 3 * 1024 * 1024;
const cbeReceiptBaseUrl = "https://mbreciept.cbe.com.et/";
const acceptedReceiptTypes = ["image/jpeg", "image/png"];
const acceptedReceiptExtensions = [".jpg", ".jpeg", ".png"];

const countryCodes = [
  { code: "+251", name: "Ethiopia", flag: "🇪🇹", iso: "et", placeholder: "e.g. 911 234 567" },
  { code: "+1", name: "USA/Canada", flag: "🇺🇸", iso: "us", placeholder: "e.g. 555 123 4567" },
  { code: "+44", name: "UK", flag: "🇬🇧", iso: "gb", placeholder: "e.g. 7700 900077" },
  { code: "+91", name: "India", flag: "🇮🇳", iso: "in", placeholder: "e.g. 98765 43210" },
  { code: "+86", name: "China", flag: "🇨🇳", iso: "cn", placeholder: "e.g. 131 2345 6789" },
  { code: "+81", name: "Japan", flag: "🇯🇵", iso: "jp", placeholder: "e.g. 090 1234 5678" },
  { code: "+49", name: "Germany", flag: "🇩🇪", iso: "de", placeholder: "e.g. 151 23456789" },
  { code: "+33", name: "France", flag: "🇫🇷", iso: "fr", placeholder: "e.g. 6 12 34 56 78" },
  { code: "+39", name: "Italy", flag: "🇮🇹", iso: "it", placeholder: "e.g. 312 345 6789" },
  { code: "+34", name: "Spain", flag: "🇪🇸", iso: "es", placeholder: "e.g. 612 345 678" },
  { code: "+55", name: "Brazil", flag: "🇧🇷", iso: "br", placeholder: "e.g. 11 91234 5678" },
  { code: "+61", name: "Australia", flag: "🇦🇺", iso: "au", placeholder: "e.g. 412 345 678" },
  { code: "+971", name: "UAE", flag: "🇦🇪", iso: "ae", placeholder: "e.g. 50 123 4567" },
  { code: "+966", name: "Saudi Arabia", flag: "🇸🇦", iso: "sa", placeholder: "e.g. 51 234 5678" },
  { code: "+27", name: "South Africa", flag: "🇿🇦", iso: "za", placeholder: "e.g. 71 234 5678" },
  { code: "+234", name: "Nigeria", flag: "🇳🇬", iso: "ng", placeholder: "e.g. 801 234 5678" },
  { code: "+254", name: "Kenya", flag: "🇰🇪", iso: "ke", placeholder: "e.g. 712 345678" },
  { code: "+256", name: "Uganda", flag: "🇺🇬", iso: "ug", placeholder: "e.g. 701 234567" },
  { code: "+250", name: "Rwanda", flag: "🇷🇼", iso: "rw", placeholder: "e.g. 781 234567" },
  { code: "+20", name: "Egypt", flag: "🇪🇬", iso: "eg", placeholder: "e.g. 101 234 5678" },
];

const socialLinks: { name: PlatformName; label: string; href: string }[] = [
  { name: "facebook", label: "Facebook", href: "#contact" },
  { name: "telegram", label: "Telegram", href: "#contact" },
  { name: "whatsapp", label: "WhatsApp", href: "#contact" },
  { name: "instagram", label: "Instagram", href: "#contact" },
];

function getSpreadsheetWebhookUrl() {
  return window.__RUNTIME_CONFIG__?.SPREADSHEET_WEBHOOK_URL?.trim() ?? "";
}

function getCbeReceiptUrl(value: string) {
  const normalizedValue = value.replace(/\s*\.\s*/g, ".").replace(/\s+/g, "");
  const match = normalizedValue.match(/https?:\/\/mbreciept\.cbe\.com\.et\/[A-Za-z0-9-]+/i);
  return match?.[0] ?? "";
}

function readFileAsBase64(file: File) {
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = String(reader.result || "");
      resolve(result.includes(",") ? result.split(",")[1] : result);
    };
    reader.onerror = () => reject(reader.error || new Error("File could not be read."));
    reader.readAsDataURL(file);
  });
}

function hasCanvasQrFallback() {
  return Boolean(document.createElement("canvas").getContext("2d"));
}

function canScanReceiptQr() {
  return Boolean((window.BarcodeDetector && "createImageBitmap" in window) || hasCanvasQrFallback());
}

function getReceiptScanDiagnostics() {
  return {
    isSecureContext: window.isSecureContext,
    hasBarcodeDetector: Boolean(window.BarcodeDetector),
    hasCreateImageBitmap: "createImageBitmap" in window,
    hasCanvas2d: hasCanvasQrFallback(),
    userAgent: window.navigator.userAgent,
  };
}

async function readReceiptUrlWithBarcodeDetector(file: File) {
  const BarcodeDetector = window.BarcodeDetector;
  if (!BarcodeDetector) {
    throw new Error("Native BarcodeDetector is not available.");
  }

  const detector = new BarcodeDetector({ formats: ["qr_code"] });
  const bitmap = await createImageBitmap(file);

  try {
    const codes = await detector.detect(bitmap);
    console.info(
      "[receipt] Native QR scan completed",
      codes.map((code) => ({ rawValue: code.rawValue, cbeReceiptUrl: getCbeReceiptUrl(code.rawValue) })),
    );
    return codes.map((code) => getCbeReceiptUrl(code.rawValue)).find(Boolean) ?? "";
  } finally {
    bitmap.close();
  }
}

async function readReceiptUrlWithJsQr(file: File) {
  const { default: jsQR } = await import("jsqr");
  const imageUrl = URL.createObjectURL(file);
  const image = new Image();

  try {
    await new Promise<void>((resolve, reject) => {
      image.onload = () => resolve();
      image.onerror = () => reject(new Error("Receipt image could not be loaded for QR scanning."));
      image.src = imageUrl;
    });

    const imageWidth = image.naturalWidth || image.width;
    const imageHeight = image.naturalHeight || image.height;
    const rotations = [0, 90, 180, 270];

    for (const rotation of rotations) {
      const canvas = document.createElement("canvas");
      const isSideways = rotation === 90 || rotation === 270;
      canvas.width = isSideways ? imageHeight : imageWidth;
      canvas.height = isSideways ? imageWidth : imageHeight;

      const context = canvas.getContext("2d", { willReadFrequently: true });
      if (!context || !canvas.width || !canvas.height) {
        throw new Error("Canvas QR scanner is not available.");
      }

      context.translate(canvas.width / 2, canvas.height / 2);
      context.rotate((rotation * Math.PI) / 180);
      context.drawImage(image, -imageWidth / 2, -imageHeight / 2, imageWidth, imageHeight);

      const imageData = context.getImageData(0, 0, canvas.width, canvas.height);
      const code = jsQR(imageData.data, imageData.width, imageData.height);
      const rawValue = code?.data ?? "";
      const cbeReceiptUrl = getCbeReceiptUrl(rawValue);

      console.info("[receipt] jsQR scan attempt completed", {
        rotation,
        rawValue,
        cbeReceiptUrl,
        imageWidth: canvas.width,
        imageHeight: canvas.height,
      });

      if (cbeReceiptUrl) {
        return cbeReceiptUrl;
      }
    }

    return "";
  } finally {
    URL.revokeObjectURL(imageUrl);
  }
}

async function readReceiptUrlWithOcr(file: File) {
  console.info("[receipt] Using OCR fallback scanner");
  const { recognize } = await import("tesseract.js");
  const result = await recognize(file, "eng");
  const text = result.data.text;
  const cbeReceiptUrl = getCbeReceiptUrl(text);

  console.info("[receipt] OCR scan completed", {
    cbeReceiptUrl,
    textPreview: text.replace(/\s+/g, " ").slice(0, 240),
  });

  return cbeReceiptUrl;
}

async function readCbeReceiptUrlFromImage(file: File) {
  if (!canScanReceiptQr()) {
    console.warn("[receipt] QR scanning unavailable", getReceiptScanDiagnostics());
    throw new Error("Receipt QR scanning is not available in this browser/context.");
  }

  let cbeReceiptUrl = "";

  if (window.BarcodeDetector && "createImageBitmap" in window) {
    console.info("[receipt] Using native BarcodeDetector scanner");
    cbeReceiptUrl = await readReceiptUrlWithBarcodeDetector(file);
  } else {
    console.info("[receipt] Using jsQR fallback scanner");
    cbeReceiptUrl = await readReceiptUrlWithJsQr(file);
  }

  if (cbeReceiptUrl) {
    return cbeReceiptUrl;
  }

  return readReceiptUrlWithOcr(file);
}

async function postSubmissionToSpreadsheet(record: SubmissionRecord) {
  const webhookUrl = getSpreadsheetWebhookUrl();

  if (!webhookUrl) {
    console.error("[receipt] Spreadsheet webhook URL is missing");
    throw new Error("Spreadsheet webhook URL is not configured.");
  }

  console.info("[receipt] Sending submission to spreadsheet webhook", {
    webhookConfigured: Boolean(webhookUrl),
    receiptFile: record.receiptFile,
    cbeReceiptUrl: record.cbeReceiptUrl,
    paymentAmount: record.paymentAmount,
    language: record.language,
  });

  const response = await fetch(webhookUrl, {
    method: "POST",
    headers: {
      "Content-Type": "text/plain;charset=utf-8",
    },
    body: JSON.stringify(record),
  });

  if (!response.ok) {
    throw new Error(`Server responded with ${response.status}`);
  }

  console.info("[receipt] Submission request sent and acknowledged by server.");
}

const features: Feature[] = [
  { icon: "hat", title: "100% Confidential", copy: "Your privacy is our priority." },
  { icon: "bolt", title: "Fast & Reliable", copy: "We deliver on time, every time." },
  { icon: "shield", title: "Professional Delivery", copy: "Handled with care and respect." },
  { icon: "users", title: "All Channels", copy: "Call, voice, message or in-person." },
];

const steps: Step[] = [
  { icon: "message", title: "Submit", copy: "Fill out the form with your bad news and delivery details." },
  { icon: "card", title: "Confirm", copy: "Make payment and upload your receipt for verification." },
  { icon: "hat", title: "We Deliver", copy: "We deliver your message with care, respect, and discretion." },
  { icon: "check", title: "You Relax", copy: "Leave the hard part to us and stay stress-free." },
];

const trustItems: Feature[] = [
  { icon: "shield", title: "Your Privacy Guaranteed", copy: "We never reveal your identity unless you approve." },
  { icon: "users", title: "Trusted by Hundreds", copy: "People trust us to deliver what matters." },
  { icon: "clock", title: "On-Time Delivery", copy: "We respect your time and their time." },
];

const trustTranslations = {
  en: {
    title: "Built on Trust",
    intro: "Professional delivery for sensitive messages, handled with privacy and care.",
    items: trustItems,
  },
  am: {
    title: "በእምነት ላይ የተገነባ",
    intro: "ለስሜታዊ መልእክቶች ሙያዊ ማድረስ፣ በግላዊነት እና በጥንቃቄ።",
    items: [
      { icon: "shield" as const, title: "ግላዊነትዎ የተጠበቀ ነው", copy: "ካላፀደቁ በቀር ማንነትዎን ፈጽሞ አንገልጽም።" },
      { icon: "users" as const, title: "በብዙዎች የታመነ", copy: "ሰዎች አስፈላጊውን መልእክት እንድናደርስ ያምኑናል።" },
      { icon: "clock" as const, title: "በጊዜው ማድረስ", copy: "የእርስዎንም ጊዜ የተቀባዩንም ጊዜ እናከብራለን።" },
    ],
  },
};

const reviewTranslations = {
  en: {
    title: "What People Say",
    intro: "Short notes from people who trusted us with difficult conversations.",
    items: [
      { quote: "They handled a very sensitive message calmly and professionally.", author: "Private client" },
      { quote: "Fast confirmation, clear process, and respectful delivery.", author: "Verified sender" },
      { quote: "The privacy gave me peace of mind from start to finish.", author: "Anonymous customer" },
    ],
  },
  am: {
    title: "ሰዎች ምን ይላሉ",
    intro: "አስቸጋሪ ውይይቶችን ለእኛ ካመኑ ሰዎች አጭር አስተያየቶች።",
    items: [
      { quote: "በጣም ስሜታዊ መልእክትን በሰላም እና በሙያዊነት አደረሱ።", author: "የግል ደንበኛ" },
      { quote: "ፈጣን ማረጋገጫ፣ ግልጽ ሂደት እና በክብር ማድረስ።", author: "የተረጋገጠ ላኪ" },
      { quote: "ግላዊነቱ ከመጀመሪያ እስከ መጨረሻ እርጋታ ሰጠኝ።", author: "ስም-አልባ ደንበኛ" },
    ],
  },
};

const featureTranslations = {
  en: [
    { title: "100% Confidential", copy: "Your privacy is our priority." },
    { title: "Fast & Reliable", copy: "We deliver on time, every time." },
    { title: "Professional Delivery", copy: "Handled with care and respect." },
    { title: "All Channels", copy: "Call, voice, message or in-person." },
  ],
  am: [
    { title: "100% ምስጢራዊ", copy: "ግላዊነትዎ ዋና ቅድሚያችን ነው።" },
    { title: "ፈጣን እና አስተማማኝ", copy: "ሁልጊዜ በጊዜው እናደርሳለን።" },
    { title: "ሙያዊ ማድረስ", copy: "በጥንቃቄ እና በክብር ይከናወናል።" },
    { title: "ሁሉም መንገዶች", copy: "በስልክ፣ በድምጽ፣ በመልእክት ወይም በአካል።" },
  ],
};

const stepTranslations = {
  en: [
    { title: "Submit", copy: "Fill out the form with your bad news and delivery details." },
    { title: "Confirm", copy: "Make payment and upload your receipt for verification." },
    { title: "We Deliver", copy: "We deliver your message with care, respect, and discretion." },
    { title: "You Relax", copy: "Leave the hard part to us and stay stress-free." },
  ],
  am: [
    { title: "ያስገቡ", copy: "መጥፎ ዜናዎን እና የማድረሻ ዝርዝሮችን በቅጹ ይሙሉ።" },
    { title: "ያረጋግጡ", copy: "ክፍያ ይፈጽሙ እና ደረሰኝዎን ለማረጋገጥ ይስቀሉ።" },
    { title: "እኛ እናደርሳለን", copy: "መልእክትዎን በጥንቃቄ፣ በክብር እና በምስጢር እናደርሳለን።" },
    { title: "እርስዎ ይረፉ", copy: "ከባዱን ክፍል ለእኛ ይተዉት እና ከጭንቀት ነጻ ይሁኑ።" },
  ],
};

const processTranslations = {
  en: {
    title: "How It Works",
    intro: "Four clear steps from submission to discreet delivery.",
  },
  am: {
    title: "እንዴት ይሰራል",
    intro: "ከማስገባት እስከ በምስጢር ማድረስ ድረስ አራት ግልጽ ደረጃዎች።",
  },
};

const formTranslations = {
  en: {
    title: "Contact Form",
    intro: "Send your bad news. We will deliver it for you.",
    name: "Name",
    namePlaceholder: "Enter your full name",
    phone: "Phone No",
    message: "Bad News",
    messagePlaceholder: "Write the bad news you want us to deliver...",
    audioUpload: "Upload Audio(optional)",
    audioTooLarge: "Audio file must be 3MB or smaller.",
    receiver: "Receiver",
    receiverPlaceholder: "Enter receiver's name or phone number",
    accountNo: "Account No:- 1000239878583",
    accountName: "Fraol Eshetu Hailu",
    basePrice: "Base price: 500 Birr",
    paymentAmount: "Amount Paid (Birr)",
    paymentAmountPlaceholder: "Enter 500 or more",
    paymentTooLow: "The base price is 500 Birr. Please enter at least 500 Birr.",
    paymentStatusPending: "Pending receipt confirmation",
    receiptUpload: "Click to upload receipt",
    receiptHelp: "Upload a CBE screenshot with the QR code or receipt link visible. JPG or PNG only.",
    receiptReady: "CBE receipt link found",
    receiptChecking: "Checking screenshot for CBE receipt link...",
    receiptCheckingWait: "Please wait while we check your receipt screenshot.",
    receiptMissing: "Please upload your payment receipt before submitting.",
    receiptInvalidType: "Upload a JPG or PNG CBE receipt screenshot only.",
    receiptPdfUnsupported: "PDF receipts are not accepted for automatic verification. Upload a CBE screenshot with the QR/payment link visible.",
    receiptTooLarge: "Receipt must be 5MB or smaller.",
    receiptQrMissing: "No CBE receipt link was found. Upload a clear screenshot that shows the QR code or receipt link.",
    receiptQrUnsupported: "Receipt scanning is not supported in this browser. Please use Chrome or Edge and upload the CBE screenshot again.",
    success: "Your receipt was sent for automatic CBE verification. Verified requests appear in the live sheet.",
    spreadsheetMissing: "Live spreadsheet is not connected yet. Add the webhook URL in runtime-config.js.",
    spreadsheetError: "We could not record this request live. Please try again.",
    recording: "Recording request...",
    security: "Your private details stay confidential. Never submit passwords.",
  },
  am: {
    title: "የመገናኛ ቅጽ",
    intro: "መልእክትዎን ይላኩ። እኛ እናደርሰዋለን።",
    name: "ስም",
    namePlaceholder: "ሙሉ ስምዎን ያስገቡ",
    phone: "ስልክ ቁጥር",
    message: "መጥፎ ዜና",
    messagePlaceholder: "እንድናደርስልዎ የሚፈልጉትን መጥፎ ዜና ይጻፉ...",
    audioUpload: "ድምጽ ይስቀሉ (አማራጭ)",
    audioTooLarge: "የድምጽ ፋይሉ 3MB ወይም ከዚያ በታች መሆን አለበት።",
    receiver: "ተቀባይ",
    receiverPlaceholder: "የተቀባዩን ስም ወይም ስልክ ቁጥር ያስገቡ",
    accountNo: "የሂሳብ ቁጥር:- 1000239878583",
    accountName: "ፍራኦል እሸቱ ኃይሉ",
    basePrice: "መነሻ ዋጋ፡ 500 ብር",
    paymentAmount: "የተከፈለ መጠን (ብር)",
    paymentAmountPlaceholder: "500 ወይም ከዚያ በላይ ያስገቡ",
    paymentTooLow: "መነሻ ዋጋው 500 ብር ነው። እባክዎ ቢያንስ 500 ብር ያስገቡ።",
    paymentStatusPending: "ደረሰኝ ማረጋገጫ በመጠባበቅ ላይ",
    receiptUpload: "ደረሰኝ ለመስቀል ይጫኑ",
    receiptHelp: "QR code ወይም receipt link የሚታይበት የCBE ስክሪንሾት ይስቀሉ። JPG ወይም PNG ብቻ።",
    receiptReady: "የCBE ደረሰኝ ሊንክ ተገኝቷል",
    receiptChecking: "የCBE ደረሰኝ ሊንክ በስክሪንሾቱ ውስጥ እየተፈለገ ነው...",
    receiptCheckingWait: "እባክዎ ደረሰኙን እስክንመረምር ይጠብቁ።",
    receiptMissing: "ከመላክዎ በፊት የክፍያ ደረሰኝዎን ይስቀሉ።",
    receiptInvalidType: "የCBE ደረሰኝ JPG ወይም PNG ስክሪንሾት ብቻ ይስቀሉ።",
    receiptPdfUnsupported: "PDF ደረሰኞች ለራስ-ሰር ማረጋገጫ አይቀበሉም። QR/payment link የሚታይበት የCBE ስክሪንሾት ይስቀሉ።",
    receiptTooLarge: "ደረሰኙ 5MB ወይም ከዚያ በታች መሆን አለበት።",
    receiptQrMissing: "የCBE ደረሰኝ ሊንክ አልተገኘም። QR code ወይም receipt link የሚታይበት ግልጽ ስክሪንሾት ይስቀሉ።",
    receiptQrUnsupported: "የደረሰኝ መቃኘት በዚህ browser አይደገፍም። እባክዎ Chrome ወይም Edge ይጠቀሙና የCBE ስክሪንሾቱን እንደገና ይስቀሉ።",
    success: "ደረሰኝዎ ለCBE ራስ-ሰር ማረጋገጫ ተልኳል። የተረጋገጡ ጥያቄዎች በቀጥታ ሰንጠረዡ ውስጥ ይታያሉ።",
    spreadsheetMissing: "የቀጥታ ሰንጠረዥ ገና አልተገናኘም። webhook URL በ runtime-config.js ውስጥ ያክሉ።",
    spreadsheetError: "ጥያቄውን በቀጥታ መመዝገብ አልቻልንም። እባክዎ ደግመው ይሞክሩ።",
    recording: "ጥያቄው እየተመዘገበ ነው...",
    security: "የግል መረጃዎ በምስጢር ይጠበቃል። የይለፍ ቃሎችን በፍጹም አያስገቡ።",
  },
};

const navTargets = ["home", "how-it-works", "about-us", "reviews", "faq", "contact"];

const faqItems = [
  {
    question: "Will my identity be private?",
    answer: "Yes. We only share your identity if you explicitly ask us to include it in the message.",
  },
  {
    question: "Can you deliver in Amharic?",
    answer: "Yes. You can write the message in English or Amharic, and we can deliver it in the tone you choose.",
  },
  {
    question: "What happens after I submit?",
    answer: "We review your request, confirm payment and delivery details, then contact you before delivery if anything is unclear.",
  },
];

const pageTranslations = {
  en: {
    eyebrow: "BAD NEWS",
    nav: ["Home", "How It Works", "About Us", "Reviews", "FAQ", "Contact"],
    cta: "Send Bad News",
    headline1: "We Deliver",
    headline2: "The Bad News",
    headline3: "You Don't Want To Say.",
    intro: "John Bad News delivers your difficult messages with discretion, professionalism, and impact.",
    submit: "Submit",
    statOne: "Same-day delivery",
    statTwo: "Private by default",
    statThree: "Any channel",
    faqTitle: "Questions Before You Send?",
    faqIntro: "A few details people usually want to know before handing us the hard conversation.",
    footerCopy: "We deliver the bad news you don't want to say, so you can focus on what matters.",
    contactTitle: "Contact Us",
    hours: "Mon - Sun: 8:00 AM - 10:00 PM",
    followTitle: "Follow Us",
    tagline: "Discretion. Delivery. Done.",
  },
  am: {
    eyebrow: "መጥፎ ዜና",
    nav: ["መነሻ", "እንዴት ይሰራል", "ስለ እኛ", "አስተያየቶች", "ጥያቄ", "አግኙን"],
    cta: "መልእክት ላክ",
    headline1: "ለመናገር የሚከብድ መጥፎ ዜናዎትን",
    headline2: "እናደርሳለን",
    headline3: "እርስዎ ማለት የማይፈልጉትን።",
    intro: "ጆን ባድ ኒውስ አስቸጋሪ መልእክቶችን በምስጢር፣ በሙያዊነት እና በጥንቃቄ ያደርሳል።",
    submit: "ላክ",
    statOne: "በተመሳሳይ ቀን ማድረስ",
    statTwo: "ምስጢር በመጀመሪያ",
    statThree: "በማንኛውም መንገድ",
    faqTitle: "ከመላክዎ በፊት ጥያቄ አለዎት?",
    faqIntro: "አስቸጋሪ መልእክትን ከመላክ በፊት ብዙ ሰዎች የሚጠይቁት መረጃ።",
    footerCopy: "እርስዎ ማለት የማይፈልጉትን መጥፎ ዜና እናደርሳለን፣ እርስዎም አስፈላጊው ላይ ያተኩሩ።",
    contactTitle: "አግኙን",
    hours: "ሰኞ - እሁድ፡ 8:00 AM - 10:00 PM",
    followTitle: "ይከተሉን",
    tagline: "ምስጢር። ማድረስ። ተጠናቋል።",
  },
};

const faqTranslations = {
  en: faqItems,
  am: [
    {
      question: "ማንነቴ በምስጢር ይጠበቃል?",
      answer: "አዎ። ማንነትዎን በመልእክቱ ውስጥ እንድናካትት በግልጽ ካልጠየቁን በቀር አንጋራውም።",
    },
    {
      question: "በአማርኛ ማድረስ ትችላላችሁ?",
      answer: "አዎ። መልእክቱን በእንግሊዝኛ ወይም በአማርኛ መጻፍ ይችላሉ፣ በመረጡትም ቃና እናደርሰዋለን።",
    },
    {
      question: "ካስገባሁ በኋላ ምን ይሆናል?",
      answer: "ጥያቄዎን እንመረምራለን፣ ክፍያን እና የማድረሻ ዝርዝሮችን እናረጋግጣለን፣ ከማድረሳችንም በፊት ግልጽ ያልሆነ ነገር ካለ እናገኝዎታለን።",
    },
  ],
};

function Icon({ name }: { name: IconName }) {
  const common = {
    fill: "none",
    stroke: "currentColor",
    strokeWidth: 2,
    strokeLinecap: "round" as const,
    strokeLinejoin: "round" as const,
  };

  switch (name) {
    case "shield":
      return (
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <path {...common} d="M12 3 5 6v5c0 4.6 2.8 8.5 7 10 4.2-1.5 7-5.4 7-10V6l-7-3Z" />
          <path {...common} d="m8.7 12 2.2 2.2 4.6-5" />
        </svg>
      );
    case "bolt":
      return (
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <path {...common} d="M13 2 4 14h7l-1 8 10-13h-7l0-7Z" />
        </svg>
      );
    case "users":
      return (
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <path {...common} d="M16 21v-2a4 4 0 0 0-4-4H7a4 4 0 0 0-4 4v2" />
          <circle {...common} cx="9.5" cy="7" r="4" />
          <path {...common} d="M22 21v-2a4 4 0 0 0-3-3.9" />
          <path {...common} d="M16 3.2a4 4 0 0 1 0 7.6" />
        </svg>
      );
    case "hat":
      return (
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <path {...common} d="M5 14h14" />
          <path {...common} d="M7 14 9 5h6l2 9" />
          <path {...common} d="M4 18c2.2 1 5.1 1.5 8 1.5s5.8-.5 8-1.5" />
        </svg>
      );
    case "check":
      return (
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <circle {...common} cx="12" cy="12" r="9" />
          <path {...common} d="m8 12.5 2.7 2.7L16.5 9" />
        </svg>
      );
    case "card":
      return (
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <rect {...common} x="3" y="5" width="18" height="14" rx="2" />
          <path {...common} d="M3 10h18M7 15h4" />
        </svg>
      );
    case "send":
      return (
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <path {...common} d="m22 2-7 20-4-9-9-4 20-7Z" />
          <path {...common} d="M22 2 11 13" />
        </svg>
      );
    case "phone":
      return (
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <path {...common} d="M22 16.9v3a2 2 0 0 1-2.2 2 19.8 19.8 0 0 1-8.6-3.1 19.5 19.5 0 0 1-6-6A19.8 19.8 0 0 1 2.1 4.2 2 2 0 0 1 4.1 2h3a2 2 0 0 1 2 1.7c.1 1 .4 2 .7 2.8a2 2 0 0 1-.4 2.1L8.1 9.9a16 16 0 0 0 6 6l1.3-1.3a2 2 0 0 1 2.1-.4c.9.3 1.8.6 2.8.7A2 2 0 0 1 22 16.9Z" />
        </svg>
      );
    case "user":
      return (
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <circle {...common} cx="12" cy="8" r="4" />
          <path {...common} d="M4 21a8 8 0 0 1 16 0" />
        </svg>
      );
    case "message":
      return (
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <path {...common} d="M21 15a4 4 0 0 1-4 4H8l-5 3V7a4 4 0 0 1 4-4h10a4 4 0 0 1 4 4v8Z" />
        </svg>
      );
    case "receipt":
      return (
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <path {...common} d="M6 2h12v20l-3-2-3 2-3-2-3 2V2Z" />
          <path {...common} d="M9 7h6M9 11h6M9 15h4" />
        </svg>
      );
    case "upload":
      return (
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <path {...common} d="M12 3v12" />
          <path {...common} d="m7 8 5-5 5 5" />
          <path {...common} d="M5 15v4a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-4" />
        </svg>
      );
    case "bank":
      return (
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <path {...common} d="M3 10h18L12 4 3 10Z" />
          <path {...common} d="M5 10v8M9 10v8M15 10v8M19 10v8M3 20h18" />
        </svg>
      );
    case "clock":
      return (
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <circle {...common} cx="12" cy="12" r="9" />
          <path {...common} d="M12 7v5l3 2" />
        </svg>
      );
  }
}

function PlatformIcon({ name }: { name: PlatformName }) {
  const common = {
    fill: "none",
    stroke: "currentColor",
    strokeWidth: 2,
    strokeLinecap: "round" as const,
    strokeLinejoin: "round" as const,
  };

  switch (name) {
    case "facebook":
      return (
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <path {...common} d="M14 8h3V4h-3c-3 0-5 2-5 5v3H6v4h3v5h4v-5h3l1-4h-4V9c0-.6.4-1 1-1Z" />
        </svg>
      );
    case "telegram":
      return (
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <path {...common} d="M21 4 3 11l7 3 3 6 8-16Z" />
          <path {...common} d="m10 14 4-4" />
        </svg>
      );
    case "whatsapp":
      return (
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <path {...common} d="M4 20 5.3 16.2A8 8 0 1 1 8 18.7L4 20Z" />
          <path {...common} d="M9.5 8.5c.4 3 2.2 4.8 5 6l1.4-1.4-2.1-1-1 1c-.9-.4-1.8-1.2-2.2-2.2l1-1-1-2.1-1.1.7Z" />
        </svg>
      );
    case "instagram":
      return (
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <rect {...common} x="4" y="4" width="16" height="16" rx="5" />
          <circle {...common} cx="12" cy="12" r="3.5" />
          <path {...common} d="M17.5 6.8h.01" />
        </svg>
      );
  }
}

function CountrySelect({ value, onChange }: { value: string; onChange: (val: string) => void }) {
  const [isOpen, setIsOpen] = useState(false);
  const selected = countryCodes.find((c) => c.code === value) || countryCodes[0];

  return (
    <div className="custom-country-select" onBlur={(e) => { if (!e.currentTarget.contains(e.relatedTarget)) setIsOpen(false); }}>
      <button type="button" className="country-select-btn" onClick={() => setIsOpen(!isOpen)} aria-haspopup="listbox" aria-expanded={isOpen} title={`${selected.name} (${selected.code})`}>
        <img src={`https://flagcdn.com/w20/${selected.iso}.png`} alt={selected.name} width="20" />
        <span className="country-select-code">{selected.code}</span>
      </button>
      {isOpen && (
        <div className="country-select-dropdown" role="listbox">
          {countryCodes.map((c) => (
            <button
              type="button"
              key={c.code}
              role="option"
              aria-selected={value === c.code}
              className={`country-select-option ${value === c.code ? 'is-selected' : ''}`}
              onClick={() => { onChange(c.code); setIsOpen(false); }}
            >
              <img src={`https://flagcdn.com/w20/${c.iso}.png`} alt={c.name} width="20" />
              <span>{c.name} ({c.code})</span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

export default function App() {
  const [receiptName, setReceiptName] = useState("");
  const [receiptError, setReceiptError] = useState("");
  const [cbeReceiptUrl, setCbeReceiptUrl] = useState("");
  const [isReceiptChecking, setIsReceiptChecking] = useState(false);
  const [audioName, setAudioName] = useState("");
  const [audioMimeType, setAudioMimeType] = useState("");
  const [audioSize, setAudioSize] = useState(0);
  const [audioDataBase64, setAudioDataBase64] = useState("");
  const [submitted, setSubmitted] = useState(false);
  const [submissionError, setSubmissionError] = useState("");
  const [audioError, setAudioError] = useState("");
  const [isRecording, setIsRecording] = useState(false);
  const [language, setLanguage] = useState<"en" | "am">("en");
  const [selectedCountryCode, setSelectedCountryCode] = useState("+251");
  const [receiverCountryCode, setReceiverCountryCode] = useState("+251");
  const text = pageTranslations[language];
  const featureText = featureTranslations[language];
  const stepText = stepTranslations[language];
  const processText = processTranslations[language];
  const formText = formTranslations[language];
  const faqText = faqTranslations[language];
  const trustText = trustTranslations[language];
  const reviewText = reviewTranslations[language];

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const formElement = event.currentTarget;
    const form = new FormData(formElement);
    setSubmissionError("");

    if (isReceiptChecking) {
      setReceiptError(formText.receiptCheckingWait);
      setSubmitted(false);
      return;
    }

    if (!receiptName) {
      setReceiptError(formText.receiptMissing);
      setSubmitted(false);
      return;
    }

    if (!cbeReceiptUrl) {
      setReceiptError(formText.receiptQrMissing);
      setSubmitted(false);
      return;
    }

    if (receiptError) {
      setSubmissionError(receiptError);
      setSubmitted(false);
      return;
    }

    if (audioError) {
      setSubmissionError(audioError);
      setSubmitted(false);
      return;
    }

    if (!getSpreadsheetWebhookUrl()) {
      setSubmissionError(formText.spreadsheetMissing);
      setSubmitted(false);
      return;
    }

    const paymentAmount = Number(form.get("paymentAmount"));

    if (!Number.isFinite(paymentAmount) || paymentAmount < basePriceBirr) {
      setSubmissionError(formText.paymentTooLow);
      setSubmitted(false);
      return;
    }

    const nextRecord: SubmissionRecord = {
      id: `${Date.now()}`,
      receivedAt: new Date().toLocaleString(),
      name: String(form.get("name") ?? ""),
      phone: `${selectedCountryCode}${String(form.get("phone") ?? "")}`,
      badNews: String(form.get("badNews") ?? ""),
      receiver: `${receiverCountryCode}${String(form.get("receiver") ?? "")}`,
      paymentAmount: `${paymentAmount}`,
      paymentStatus: formText.paymentStatusPending,
      cbeReceiptUrl,
      audioFile: audioName || "-",
      audioMimeType,
      audioSize,
      audioDataBase64,
      receiptFile: receiptName,
      language,
    };

    try {
      setIsRecording(true);
      await postSubmissionToSpreadsheet(nextRecord);

      setSubmitted(true);
      formElement.reset();
      setReceiptName("");
      setCbeReceiptUrl("");
      setAudioName("");
      setAudioMimeType("");
      setAudioSize(0);
      setAudioDataBase64("");
      setAudioError("");
    } catch (error) {
      console.error("[receipt] Submission failed", {
        error,
        receiptFile: nextRecord.receiptFile,
        cbeReceiptUrl: nextRecord.cbeReceiptUrl,
      });
      setSubmitted(false);
      setSubmissionError(formText.spreadsheetError);
    } finally {
      setIsRecording(false);
    }
  }

  async function handleReceiptChange(event: FormEvent<HTMLInputElement>) {
    const receiptInput = event.currentTarget;
    const file = receiptInput.files?.[0];
    setSubmitted(false);
    setSubmissionError("");
    setCbeReceiptUrl("");

    if (isReceiptChecking) {
      console.warn("[receipt] Receipt change ignored while scanning is in progress");
      return;
    }

    setIsReceiptChecking(false);

    if (!file) {
      console.info("[receipt] Receipt selection cleared");
      setReceiptName("");
      setReceiptError("");
      return;
    }

    console.info("[receipt] Receipt file selected", {
      name: file.name,
      type: file.type || "(empty MIME type)",
      size: file.size,
      diagnostics: getReceiptScanDiagnostics(),
    });

    if (!canScanReceiptQr()) {
      console.warn("[receipt] Receipt scan blocked by browser capability/context", getReceiptScanDiagnostics());
      receiptInput.value = "";
      setReceiptName("");
      setCbeReceiptUrl("");
      setReceiptError(formText.receiptQrUnsupported);
      return;
    }

    const lowerName = file.name.toLowerCase();
    const isPdf = file.type === "application/pdf" || lowerName.endsWith(".pdf");
    const hasAcceptedExtension = acceptedReceiptExtensions.some((extension) => lowerName.endsWith(extension));
    const hasAcceptedType = acceptedReceiptTypes.includes(file.type);

    if (isPdf) {
      console.warn("[receipt] Rejected PDF receipt", { name: file.name, type: file.type });
      receiptInput.value = "";
      setReceiptName("");
      setCbeReceiptUrl("");
      setReceiptError(formText.receiptPdfUnsupported);
      return;
    }

    if (!hasAcceptedType && !hasAcceptedExtension) {
      console.warn("[receipt] Rejected unsupported receipt file type", {
        name: file.name,
        type: file.type || "(empty MIME type)",
        hasAcceptedExtension,
        hasAcceptedType,
      });
      receiptInput.value = "";
      setReceiptName("");
      setCbeReceiptUrl("");
      setReceiptError(formText.receiptInvalidType);
      return;
    }

    if (file.size > maxReceiptSize) {
      console.warn("[receipt] Rejected oversized receipt file", {
        name: file.name,
        size: file.size,
        maxReceiptSize,
      });
      receiptInput.value = "";
      setReceiptName("");
      setCbeReceiptUrl("");
      setReceiptError(formText.receiptTooLarge);
      return;
    }

    try {
      setReceiptName(file.name);
      setReceiptError("");
      setIsReceiptChecking(true);
      const extractedUrl = await readCbeReceiptUrlFromImage(file);

      if (!extractedUrl.toLowerCase().startsWith(cbeReceiptBaseUrl)) {
        console.warn("[receipt] No CBE receipt URL found in screenshot scan", {
          name: file.name,
          extractedUrl,
          expectedBaseUrl: cbeReceiptBaseUrl,
        });
        setCbeReceiptUrl("");
        setReceiptError(formText.receiptQrMissing);
        return;
      }

      console.info("[receipt] CBE receipt URL accepted", {
        name: file.name,
        cbeReceiptUrl: extractedUrl,
      });
      setReceiptName(file.name);
      setCbeReceiptUrl(extractedUrl);
      setReceiptError("");
    } catch (error) {
      console.error("[receipt] Failed while scanning receipt image", {
        name: file.name,
        error,
        diagnostics: getReceiptScanDiagnostics(),
      });
      setCbeReceiptUrl("");
      setReceiptError(formText.receiptQrMissing);
    } finally {
      setIsReceiptChecking(false);
    }
  }

  async function handleAudioChange(event: FormEvent<HTMLInputElement>) {
    const audioInput = event.currentTarget;
    const file = audioInput.files?.[0];
    setSubmitted(false);
    setSubmissionError("");

    if (!file) {
      setAudioName("");
      setAudioMimeType("");
      setAudioSize(0);
      setAudioDataBase64("");
      setAudioError("");
      return;
    }

    if (file.size > maxAudioSize) {
      console.warn("[audio] Rejected oversized audio file", {
        name: file.name,
        size: file.size,
        maxAudioSize,
      });
      audioInput.value = "";
      setAudioName("");
      setAudioMimeType("");
      setAudioSize(0);
      setAudioDataBase64("");
      setAudioError(formText.audioTooLarge);
      return;
    }

    try {
      const dataBase64 = await readFileAsBase64(file);
      console.info("[audio] Audio file prepared for spreadsheet webhook", {
        name: file.name,
        type: file.type || "audio/*",
        size: file.size,
        base64Length: dataBase64.length,
      });
      setAudioName(file.name);
      setAudioMimeType(file.type || "audio/mpeg");
      setAudioSize(file.size);
      setAudioDataBase64(dataBase64);
      setAudioError("");
    } catch (error) {
      console.error("[audio] Failed to read audio file", { name: file.name, error });
      audioInput.value = "";
      setAudioName("");
      setAudioMimeType("");
      setAudioSize(0);
      setAudioDataBase64("");
      setAudioError(formText.spreadsheetError);
    }
  }

  return (
    <div className="bad-news-app">
      <header className="site-header">
        <a className="brand" href="#home" aria-label="John Bad News home">
          <img src="/brand/johnny-logo.png" alt="" />
          <span>
            <strong>JOHN</strong>
            <small>{text.eyebrow}</small>
          </span>
        </a>
        <nav aria-label="Primary navigation">
          {text.nav.map((item, index) => (
            <a key={navTargets[index]} href={`#${navTargets[index]}`}>
              {item}
            </a>
          ))}
        </nav>
        <div className="header-actions">
          <label className="language-select">
            <span>Language</span>
            <select value={language} onChange={(event) => setLanguage(event.target.value as "en" | "am")} aria-label="Select language">
              <option value="en">English</option>
              <option value="am">አማርኛ</option>
            </select>
          </label>
          <a className="header-cta" href="#contact">
            {text.cta}
          </a>
        </div>
      </header>

      <main id="home" className="page-shell">
        <section className="hero-section" aria-labelledby="hero-title">
          <div className="hero-copy">
            <img className="hero-logo" src="/brand/johnny-logo.png" alt="Bad News logo" />
            <div className="headline-block">
              <h1 id="hero-title">
                {text.headline1} <span>{text.headline2}</span>
                <small>{text.headline3}</small>
              </h1>
              <p>{text.intro}</p>
              <div className="hero-stats" aria-label="Service highlights">
                <span>{text.statOne}</span>
                <span>{text.statTwo}</span>
                <span>{text.statThree}</span>
              </div>
            </div>

            <div className="feature-grid" aria-label="Service promises">
              {features.map((item, index) => (
                <article className="feature-item" key={item.title}>
                  <span className="round-icon">
                    <Icon name={item.icon} />
                  </span>
                  <strong>{featureText[index].title}</strong>
                  <p>{featureText[index].copy}</p>
                </article>
              ))}
            </div>


          </div>

          <div className="mystery-figure" aria-hidden="true">
            <div className="hat-brim" />
            <div className="hat-top" />
            <div className="coat" />
          </div>

          <aside id="contact" className="contact-card" aria-labelledby="contact-title">
            <div className="form-badge">
              <Icon name="receipt" />
            </div>
            <h2 id="contact-title">{formText.title}</h2>
            <p>{formText.intro}</p>

            <form onSubmit={handleSubmit}>
              <label>
                <span>
                  <Icon name="user" /> {formText.name}
                </span>
                <input name="name" type="text" placeholder={formText.namePlaceholder} required />
              </label>

              <label>
                <span>
                  <Icon name="phone" /> {formText.phone}
                </span>
                <div className="phone-input-group">
                  <CountrySelect
                    value={selectedCountryCode}
                    onChange={setSelectedCountryCode}
                  />
                  <input
                    name="phone"
                    type="tel"
                    placeholder={countryCodes.find((c) => c.code === selectedCountryCode)?.placeholder || "e.g. 123 456 7890"}
                    pattern="[0-9\s\-\(\)]{7,15}"
                    required
                    style={{ flex: 1 }}
                  />
                </div>
              </label>

              <label>
                <span>
                  <Icon name="message" /> {formText.message}
                </span>
                <textarea name="badNews" placeholder={formText.messagePlaceholder} required />
              </label>

              <label className="file-inline">
                <input
                  name="audio"
                  type="file"
                  accept="audio/*"
                  onChange={handleAudioChange}
                />
                <span>
                  <Icon name="upload" /> {audioError || audioName || formText.audioUpload}
                </span>
              </label>

              <label>
                <span>
                  <Icon name="user" /> {formText.receiver}
                </span>
                <div className="phone-input-group">
                  <CountrySelect
                    value={receiverCountryCode}
                    onChange={setReceiverCountryCode}
                  />
                  <input
                    name="receiver"
                    type="tel"
                    placeholder={countryCodes.find((c) => c.code === receiverCountryCode)?.placeholder || "e.g. 123 456 7890"}
                    pattern="[0-9\s\-\(\)]{7,15}"
                    required
                    style={{ flex: 1 }}
                  />
                </div>
              </label>

              <div className="account-box">
                <Icon name="bank" />
                <div>
                  <b>{formText.basePrice}</b>
                  <strong>{formText.accountNo}</strong>
                  <span>{formText.accountName}</span>
                </div>
              </div>

              <label>
                <span>
                  <Icon name="receipt" /> {formText.paymentAmount}
                </span>
                <input name="paymentAmount" type="number" min={basePriceBirr} step="1" placeholder={formText.paymentAmountPlaceholder} required />
              </label>

              <label className="receipt-upload">
                <input
                  name="receipt"
                  type="file"
                  accept=".jpg,.jpeg,.png"
                  required
                  aria-describedby="receipt-status"
                  onChange={handleReceiptChange}
                  disabled={isReceiptChecking}
                />
                <span className={`upload-box ${isReceiptChecking ? "is-uploading" : receiptError ? "is-error" : cbeReceiptUrl ? "is-valid" : ""}`}>
                  <Icon name="upload" />
                  <strong>{receiptName || formText.receiptUpload}</strong>
                  <small id="receipt-status">
                    {isReceiptChecking ? formText.receiptChecking : receiptError || (cbeReceiptUrl ? formText.receiptReady : formText.receiptHelp)}
                  </small>
                  {isReceiptChecking && (
                    <span className="upload-progress" role="progressbar" aria-label={formText.receiptChecking}>
                      <span />
                    </span>
                  )}
                </span>
              </label>

              <button className="submit-button" type="submit" disabled={isReceiptChecking || Boolean(receiptError) || Boolean(audioError) || isRecording}>
                <Icon name="send" /> {isRecording ? formText.recording : isReceiptChecking ? formText.receiptChecking : text.submit}
              </button>
              {submitted && <output className="form-success">{formText.success}</output>}
              {submissionError && <output className="form-error">{submissionError}</output>}
              <small className="security-note">{formText.security}</small>
            </form>
          </aside>
        </section>

        <section id="how-it-works" className="steps-section" aria-labelledby="steps-title">
          <div className="section-heading">
            <h2 id="steps-title">{processText.title}</h2>
            <p>{processText.intro}</p>
          </div>
          <div className="steps-grid">
            {steps.map((step, index) => (
              <article className="step-item" key={step.title}>
                <div className="step-top">
                  <span className="step-number">{String(index + 1).padStart(2, "0")}</span>
                  <span className="round-icon">
                    <Icon name={step.icon} />
                  </span>
                </div>
                <h3>{stepText[index].title}</h3>
                <p>{stepText[index].copy}</p>
                {index < steps.length - 1 && <span className="step-arrow" aria-hidden="true" />}
              </article>
            ))}
          </div>
        </section>

        <section id="about-us" className="trust-section" aria-labelledby="trust-title">
          <div className="section-heading">
            <h2 id="trust-title">{trustText.title}</h2>
            <p>{trustText.intro}</p>
          </div>
          <div className="trust-grid">
            {trustText.items.map((item) => (
              <article className="trust-card" key={item.title}>
                <span className="round-icon">
                  <Icon name={item.icon} />
                </span>
                <div>
                  <h3>{item.title}</h3>
                  <p>{item.copy}</p>
                </div>
              </article>
            ))}
          </div>
        </section>

        <section id="reviews" className="reviews-section" aria-labelledby="reviews-title">
          <div className="section-heading">
            <h2 id="reviews-title">{reviewText.title}</h2>
            <p>{reviewText.intro}</p>
          </div>
          <div className="reviews-grid">
            {reviewText.items.map((item) => (
              <article className="review-card" key={item.author}>
                <p>"{item.quote}"</p>
                <strong>{item.author}</strong>
              </article>
            ))}
          </div>
        </section>

        <section id="faq" className="faq-section" aria-labelledby="faq-title">
          <div className="faq-copy">
            <h2 id="faq-title">{text.faqTitle}</h2>
            <p>{text.faqIntro}</p>
          </div>
          <div className="faq-list">
            {faqText.map((item) => (
              <details key={item.question}>
                <summary>{item.question}</summary>
                <p>{item.answer}</p>
              </details>
            ))}
          </div>
        </section>

      </main>

      <footer className="site-footer">
        <div className="footer-brand">
          <img src="/brand/johnny-logo.png" alt="" />
          <p>{text.footerCopy}</p>
        </div>
        <div>
          <strong>{text.contactTitle}</strong>
          <span>+251 9XX XXX XXX</span>
          <span>info@johnbadnews.com</span>
          <span>Addis Ababa, Ethiopia</span>
          <span>{text.hours}</span>
        </div>
        <div>
          <strong>{text.followTitle}</strong>
          <div className="social-row" aria-label="Social links">
            {socialLinks.map((item) => (
              <a key={item.name} href={item.href} aria-label={item.label} title={item.label}>
                <PlatformIcon name={item.name} />
              </a>
            ))}
          </div>
          <span className="tagline">{text.tagline}</span>
        </div>
      </footer>
    </div>
  );
}
