import { google } from "googleapis";
import { chromium } from "playwright";

const must = (k) => {
  const v = process.env[k];
  if (!v) throw new Error(`Missing env: ${k}`);
  return v;
};

const SA_JSON = must("SA_JSON");
const SPREADSHEET_ID = must("SPREADSHEET_ID");
const SHEET_MAIN = must("SHEET_MAIN");   // 例: お届け案件管理
const SHEET_QUEUE = must("SHEET_QUEUE"); // 例: _Queue

const jstNow = () => {
  const d = new Date();
  // ざっくりJST文字列（スプレッドシートは文字でもOK）
  const s = new Date(d.getTime() + 9 * 60 * 60 * 1000).toISOString().replace("T", " ").replace("Z", "");
  return s.slice(0, 19);
};

function parseCorpFromH2Texts(h2Texts, akno) {
  // 例: "3007608：株式会社淡路島第一次産業振興公社"
  const re = new RegExp(`^${akno}\\s*[：:]\\s*(.+)$`);
  for (const t of h2Texts) {
    const m = t.match(re);
    if (m) return m[1].trim();
  }
  // aknoが違う表示の可能性もあるので汎用も見る
  for (const t of h2Texts) {
    const m = t.match(/^(\d{7})\s*[：:]\s*(.+)$/);
    if (m && m[2]) return m[2].trim();
  }
  return "";
}

async function sheetsClient() {
  const sa = JSON.parse(SA_JSON);
  const auth = new google.auth.JWT({
    email: sa.client_email,
    key: sa.private_key,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  await auth.authorize();
  return google.sheets({ version: "v4", auth });
}

async function getQueueRows(sheets) {
  // A:案件No B:URL C:企業名 D:最終取得 E:状態
  const range = `${SHEET_QUEUE}!A2:E`;
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range });
  const values = res.data.values || [];
  return values.map((r) => ({
    akno: (r[0] || "").trim(),
    url: (r[1] || "").trim(),
    corp: (r[2] || "").trim(),
    last: (r[3] || "").trim(),
    status: (r[4] || "").trim(),
  }));
}

async function updateQueueRow(sheets, rowNumber, patch) {
  // rowNumber はシート上の行番号（A2開始なので 2+index）
  const a = patch.akno ?? "";
  const b = patch.url ?? "";
  const c = patch.corp ?? "";
  const d = patch.last ?? "";
  const e = patch.status ?? "";
  const range = `${SHEET_QUEUE}!A${rowNumber}:E${rowNumber}`;
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range,
    valueInputOption: "RAW",
    requestBody: { values: [[a, b, c, d, e]] },
  });
}

async function updateMainCorpIfEmpty(sheets, akno, corp) {
  if (!akno || !corp) return;

  // メインの案件No列は F、企業名は G（あなたの固定ヘッダ前提）
  // データ開始は3行目
  const colNoRange = `${SHEET_MAIN}!F3:F`;
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: colNoRange });
  const vals = res.data.values || [];
  let targetRow = -1;
  for (let i = 0; i < vals.length; i++) {
    const v = (vals[i]?.[0] || "").toString().trim();
    if (v === akno) {
      targetRow = 3 + i;
      break;
    }
  }
  if (targetRow < 0) return;

  // 企業名が空なら埋める（上書きしたいならここを外す）
  const corpCell = `${SHEET_MAIN}!G${targetRow}`;
  const cur = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: corpCell });
  const curVal = (cur.data.values?.[0]?.[0] || "").toString().trim();
  if (curVal) return;

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: corpCell,
    valueInputOption: "RAW",
    requestBody: { values: [[corp]] },
  });
}

async function fetchCorpByPlaywright(url, akno) {
  const browser = await chromium.launch({ headless: true });
  try {
    const ctx = await browser.newContext({
      userAgent:
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36",
      locale: "ja-JP",
    });
    const page = await ctx.newPage();

    const resp = await page.goto(url, { waitUntil: "domcontentloaded", timeout: 45000 });
    const status = resp ? resp.status() : 0;

    // JS描画待ち（“確認中…”が消えるの待ち）
    await page.waitForTimeout(2000);

    const bodyText = await page.evaluate(() => document.body?.innerText?.slice(0, 5000) || "");
    if (bodyText.includes("認証情報を確認中")) {
      return { ok: false, reason: "auth_check", status };
    }

    // h2が出るまで少し待つ
    try {
      await page.waitForSelector("h2", { timeout: 15000 });
    } catch (_) {}

    const h2Texts = await page.$$eval("h2", (els) =>
      els.map((e) => (e.textContent || "").trim()).filter(Boolean)
    );

    const corp = parseCorpFromH2Texts(h2Texts, akno);
    if (!corp) {
      // 追加デバッグ情報
      const title = await page.title();
      return { ok: false, reason: `no_corp (h2=${h2Texts.length}, title=${title})`.slice(0, 80), status };
    }

    return { ok: true, corp, status };
  } finally {
    await browser.close();
  }
}

async function main() {
  const sheets = await sheetsClient();
  const rows = await getQueueRows(sheets);

  const pending = rows
    .map((r, i) => ({ ...r, idx: i, rowNumber: 2 + i }))
    .filter((r) => r.status === "pending" && r.akno && r.url);

  console.log(`queue total=${rows.length}, pending=${pending.length}`);

  for (const r of pending) {
    const now = jstNow();
    try {
      console.log(`start akno=${r.akno} url=${r.url}`);
      const got = await fetchCorpByPlaywright(r.url, r.akno);

      if (!got.ok) {
        console.log(`fail akno=${r.akno} reason=${got.reason} http=${got.status}`);
        await updateQueueRow(sheets, r.rowNumber, {
          akno: r.akno,
          url: r.url,
          corp: "",
          last: now,
          status: `error:${got.reason}`.slice(0, 80),
        });
        continue;
      }

      console.log(`ok akno=${r.akno} corp=${got.corp} http=${got.status}`);

      await updateQueueRow(sheets, r.rowNumber, {
        akno: r.akno,
        url: r.url,
        corp: got.corp,
        last: now,
        status: "done",
      });

      await updateMainCorpIfEmpty(sheets, r.akno, got.corp);
    } catch (e) {
      console.log(`exception akno=${r.akno}`, e?.message || e);
      await updateQueueRow(sheets, r.rowNumber, {
        akno: r.akno,
        url: r.url,
        corp: "",
        last: now,
        status: `error:exception`.slice(0, 80),
      });
    }
  }
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
