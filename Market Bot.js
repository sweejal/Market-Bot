const PRODUCE_LIST = [
  "romaine", "red oak", "green oak", "red batavia", "green batavia", "red gem", "bibb", "muir",
  "vixenmix", "malabar spinach", "celery", "swiss chard", "bok choi", "baby kale",
  "assorted hot peppers", "picnic peppers",
  "thai basil", "parsley", "thyme", "oregano", "rosemary", "sage",
  "variegated sage", "garlic chives", "chives", "lemongrass"
];

const PRODUCE_PRICES = {
  "romaine": 1, "red oak": 1, "green oak": 1, "red batavia": 1, "green batavia": 1, "red gem": 1,
  "bibb": 1, "muir": 1, "vixenmix": 14, "malabar spinach": 2, "celery": 2, "swiss chard": 2,
  "bok choi": 1, "baby kale": 2, "assorted hot peppers": 1, "picnic peppers": 1,
  "thai basil": 1, "parsley": 1, "thyme": 1, "oregano": 1, "rosemary": 1, "sage": 1,
  "variegated sage": 1, "garlic chives": 1, "chives": 1, "lemongrass": 1
};

const PRODUCE_SYNONYMS = {
  "spinach": "malabar spinach",
  "batavia": "red batavia",
  "oak": "red oak",
  "green batavia": "green batavia",
  "green oak": "green oak"
};

function getTodayDayName() {
  return ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"][new Date().getDay()];
}

function extractProduceOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mondaySheet = ss.getSheetByName("Monday Orders");
  const fridaySheet = ss.getSheetByName("Friday Orders");

  const day = getTodayDayName();
  const shouldProcessMonday = ["Saturday", "Sunday", "Monday"].includes(day);
  const shouldProcessFriday = ["Tuesday", "Wednesday", "Thursday", "Friday"].includes(day);

  if (!shouldProcessMonday && !shouldProcessFriday) return;

  const threads = GmailApp.search('label:greenhouse-orders newer_than:7d');
  const messages = threads.flatMap(thread => thread.getMessages());

  const existingIdsMonday = getExistingMessageIds(mondaySheet);
  const existingIdsFriday = getExistingMessageIds(fridaySheet);

  messages.forEach(msg => {
    const messageId = msg.getId();
    if (existingIdsMonday.has(messageId) || existingIdsFriday.has(messageId)) return;

    const body = msg.getPlainBody();
    if (!isLikelyAnOrder(body)) return;

    const date = msg.getDate();
    const from = msg.getFrom();
    const name = extractName(from);
    const parsed = parseOrderBody(body);
    const finalDay = parsed.day !== "Unknown" ? parsed.day : fallbackDayByEmailDate(date);

    const row = [
      new Date(date),
      name,
      from,
      parsed.notes,
      messageId,
      false
    ];

    let total = 0;
    PRODUCE_LIST.forEach(item => {
      const qty = parsed.quantities[item] || 0;
      row.push(qty || "");
      if (qty) total += qty * (PRODUCE_PRICES[item] || 0);
    });

    row.push(total);

    if (finalDay === "Monday" && shouldProcessMonday) {
      mondaySheet.appendRow(row);
    } else if (finalDay === "Friday" && shouldProcessFriday) {
      fridaySheet.appendRow(row);
    }
  });
}

function parseOrderBody(body) {
  const lowerBody = body.toLowerCase();
  const pickupMatch = lowerBody.match(/pickup.*|delivery.*|on [a-z]+|this [a-z]+/);
  const pickupLine = pickupMatch ? pickupMatch[0] : "";
  const day = detectPickupDay(lowerBody);
  const quantities = extractProduceQuantities(lowerBody);

  return {
    notes: pickupLine,
    day: day,
    quantities: quantities
  };
}

function extractProduceQuantities(text) {
  const quantities = {};
  const normalizedText = text.toLowerCase().replace(/[^a-z0-9\s]/g, " ");
  const textJoined = normalizedText.split(/\s+/).join(" ");

  [...PRODUCE_LIST, ...Object.keys(PRODUCE_SYNONYMS)].forEach(rawTerm => {
    const official = PRODUCE_SYNONYMS[rawTerm] || rawTerm;
    const pattern = new RegExp(`(\\d+)\\s*(?:x\\s*)?(${rawTerm.replace(/\s+/g, "\\s*")})`, "gi");
    let match;
    while ((match = pattern.exec(textJoined)) !== null) {
      const qty = parseInt(match[1]);
      if (!isNaN(qty)) {
        quantities[official] = (quantities[official] || 0) + qty;
      }
    }
  });

  [...PRODUCE_LIST, ...Object.keys(PRODUCE_SYNONYMS)].forEach(rawTerm => {
    const official = PRODUCE_SYNONYMS[rawTerm] || rawTerm;
    const matches = [...textJoined.matchAll(new RegExp(`\\b${rawTerm}\\b`, "gi"))].length;
    if (matches > (quantities[official] || 0)) {
      quantities[official] = matches;
    } else if (matches === 1 && !quantities[official]) {
      quantities[official] = 1;
    }
  });

  return quantities;
}

function detectPickupDay(text) {
  const monday = ["monday", "pickup monday", "this monday"];
  const friday = ["friday", "pickup friday", "this friday"];
  if (monday.some(d => text.includes(d))) return "Monday";
  if (friday.some(d => text.includes(d))) return "Friday";
  return "Unknown";
}

function fallbackDayByEmailDate(date) {
  const day = date.getDay(); // 0 = Sun ... 6 = Sat
  if ([0, 1, 6].includes(day)) return "Monday";
  if ([3, 4, 5].includes(day)) return "Friday";
  return "Unknown";
}

function extractName(emailStr) {
  const match = emailStr.match(/^(.*?)</);
  return match ? match[1].trim() : emailStr;
}

function isLikelyAnOrder(text) {
  return ["order", "produce", "greenhouse", "pick up", "delivery"].some(kw =>
    text.toLowerCase().includes(kw)
  );
}

function getExistingMessageIds(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return new Set();
  const ids = sheet.getRange(3, 5, lastRow - 2).getValues(); // Col E = Message ID
  return new Set(ids.flat());
}

function clearMondayOrders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Monday Orders");
  clearSheetContent(sheet);
}

function clearFridayOrders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Friday Orders");
  clearSheetContent(sheet);
}

function clearSheetContent(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    sheet.getRange(3, 1, lastRow - 2, sheet.getLastColumn()).clearContent();
  }
}

function sendOrderSummaryEmail() {
  const today = getTodayDayName();
  if (today !== "Monday" && today !== "Friday") return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mondaySheet = ss.getSheetByName("Monday Orders");
  const fridaySheet = ss.getSheetByName("Friday Orders");

  const mondayOrders = getTodayOrders(mondaySheet);
  const fridayOrders = getTodayOrders(fridaySheet);

  const emailBody = `
<b> Greenhouse Order Summary for ${new Date().toDateString()}:</b><br><br>
<b>Monday Orders (${mondayOrders.length}):</b><br>
${formatOrdersHTML(mondayOrders)}<br><br>
<b>Friday Orders (${fridayOrders.length}):</b><br>
${formatOrdersHTML(fridayOrders)}<br><br>
– Market Bot 
`;

  MailApp.sendEmail({
    to: "sweejal.kafle@gmail.com",
    subject: "Daily Greenhouse Order Summary",
    htmlBody: emailBody
  });
}

function getTodayOrders(sheet) {
  const today = new Date().toDateString();
  return sheet.getDataRange().getValues().slice(2).filter(
    row => new Date(row[0]).toDateString() === today
  );
}

function formatOrdersHTML(orders) {
  if (orders.length === 0) return "<i>No orders today.</i>";
  return orders.map(order => {
    const name = order[1];
    const items = order.slice(6, -1).map((qty, i) =>
      qty && qty !== "" ? `${qty} × ${PRODUCE_LIST[i]}` : null
    ).filter(x => x).join(", ");
    return `• <b>${name}</b>: ${items}`;
  }).join("<br>");
}
