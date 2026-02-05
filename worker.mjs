import { google } from "googleapis";
import { chromium } from "playwright";

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const SHEET_MAIN = process.env.SHEET_MAIN || "お届け案件管理";
const SHEET_QUEUE = process.env.SHEET_QUEUE || "_Queue";
const SA_JSON = JSON.parse(process.env.SA_JSON);

const jwt = new google.auth.JWT(
  SA_JSON.client_email,
  null,
  SA_JSON.private_key,
  ["https://www.googleapis.com/auth/spreadsheets"]
);
const sheets = google.sheets({ version: "v4", auth: jwt });

const READYCREW_HOST = "tool.readycrew.cloud";

function nowIso() {
  const d = new Date();
  const tz = new Intl.DateTimeFormat("ja-JP", { timeZone: "Asia/Tokyo", hour12: false,
    year: "numeric", month: "2-digit", day: "2-digit",
    hour: "2-digit", minute: "2-digit", second: "2-digit" }).format(d);
  return tz.replace(/\//g, "-").replace(/ /, "T");
}

async function readQueue(limit = 10) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_QUEUE}!A2:E`
  });
  const rows = res.data.values || [];
  const pending = [];
  rows.forEach((r, i) => {
    const [no, url, corp, , status] = r;
    if (!corp && (status || "pending").toLowerCase() === "pending" && url) {
      pending.push({ idx: i + 2, no, url });
    }
  });
  return pending.slice(0, limit);
}

function extractFromText(txt, expectedNo) {
  const colon = /[：:﹕]/;
  if (expectedNo) {
    const m = txt.match(new RegExp(String(expectedNo) + "\\s*" + colon.source + "\\s*(.+)"));
    if (m) return { corp: m[1].trim() };
  }
  const m2 = txt.match(new RegExp("(\\d{7})\\s*" + colon.source + "\\s*(.+)"));
  if (m2) return { corp: m2[2].trim(), no: m2[1] };
  return null;
}

async function fetchCorp(url, no) {
  const u = new URL(url);
  if (!u.hostname.endsWith(READYCREW_HOST)) return null;

  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();
  await page.setExtraHTTPHeaders({ "Accept-Language": "ja,en;q=0.8" });
  await page.goto(url, { waitUntil: "networkidle" });

  let heads = await page.$$eval("h1,h2,h3", els => els.map(e => e.innerText || ""));
  let body = await page.evaluate(() => document.body?.innerText || "");

  let hit = null;
  for (const h of heads) { hit = extractFromText(h, no); if (hit) break; }
  if (!hit) hit = extractFromText(body, no);

  await browser.close();
  return hit?.corp || null;
}

async function writeBack(no, corp, qIdx) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_MAIN}!A${2}:Z`
  });
  const rows = res.data.values || [];
  let rowIdx = null;
  for (let i = 0; i < rows.length; i++) {
    if ((rows[i][5] || "").toString().trim() === String(no)) { rowIdx = i + 2; break; }
  }
  if (rowIdx) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_MAIN}!G${rowIdx}`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [[corp]] }
    });
  }
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_QUEUE}!C${qIdx}:E${qIdx}`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [[corp, nowIso(), "done"]] }
  });
}

async function main() {
  await jwt.authorize();

  const targets = await readQueue(10);
  if (!targets.length) {
    console.log("no pending");
    return;
  }
  for (const t of targets) {
    try {
      const corp = await fetchCorp(t.url, t.no);
      console.log("no:", t.no, "corp:", corp);
      if (corp) await writeBack(t.no, corp, t.idx);
      else await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `${SHEET_QUEUE}!E${t.idx}`,
        valueInputOption: "USER_ENTERED",
        requestBody: { values: [["retry"]] }
      });
    } catch (e) {
      console.error("error", t.no, e.toString());
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `${SHEET_QUEUE}!E${t.idx}`,
        valueInputOption: "USER_ENTERED",
        requestBody: { values: [["error"]] }
      });
    }
  }
}

main().catch(err => { console.error(err); process.exit(1); });
