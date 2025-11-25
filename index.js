// ====== 設定 ======
const SLACK_BOT_TOKEN = PropertiesService.getScriptProperties().getProperty("SLACK_BOT_TOKEN");
const CHANNEL_ID      = PropertiesService.getScriptProperties().getProperty("CHANNEL_ID");
const LOG_SHEET       = "受信ログ";
const SPREADSHEET_ID  = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");



// ===== Slackにボタン送信（テスト用） =====
function sendButton() {
  const message = {
    channel: CHANNEL_ID,
    text: "出勤・退勤ボタンを押してください！",
    attachments: [
      {
        text: "選択してください",
        fallback: "ボタンが表示されません",
        callback_id: "attendance",
        color: "#36a64f",
        attachment_type: "default",
        actions: [
          { name: "punch_in",  text: "出勤",     type: "button", style: "primary" },
          { name: "punch_out", text: "退勤",     type: "button", style: "danger"  },
          { name: "oncall",    text: "オンコール", type: "button", style: "primary" }
        ]
      }
    ]
  };

  const response = UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", {
    method: "post",
    contentType: "application/json",
    headers: { "Authorization": "Bearer " + SLACK_BOT_TOKEN },
    payload: JSON.stringify(message)
  });

  Logger.log("Slack response: " + response.getContentText());
}

// ===== Slackにボタン送信（テスト用） =====
function sendButton() {
  const message = {
    channel: CHANNEL_ID,
    text: "出勤・退勤ボタンを押してください！",
    attachments: [
      {
        text: "選択してください",
        fallback: "ボタンが表示されません",
        callback_id: "attendance",
        color: "#36a64f",
        attachment_type: "default",
        actions: [
          { name: "punch_in", text: "出勤", type: "button", style: "primary" },
          { name: "punch_out", text: "退勤", type: "button", style: "danger" }
        ]
      }
    ]
  };

  UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + SLACK_BOT_TOKEN },
    payload: JSON.stringify(message)
  });
}

// ===== SlackからのPOSTを受け取る =====
function doPost(e) {
  Logger.log("🚀 doPost called! raw=%s", e ? e.postData?.contents : "no data");

  try {
    if (!e || !e.postData) {
      Logger.log("⚠ no postData");
      return ContentService.createTextOutput("no data");
    }

    const contentType = e.postData.type || "";
    const raw = e.postData.contents || "";

    // --- URL検証 (Event Subscriptions) ---
    if (contentType.includes("application/json")) {
      const body = JSON.parse(raw);
      if (body.type === "url_verification" && body.challenge) {
        Logger.log("✅ URL verification OK");
        return ContentService.createTextOutput(body.challenge)
          .setMimeType(ContentService.MimeType.TEXT);
      }
    }

    // --- ボタン押下イベント (Interactivity & Shortcuts) ---
    const params = parseFormUrlEncoded(raw);
    if (!params.payload) {
      Logger.log("⚠ payload empty");
      return ContentService.createTextOutput("ok");
    }

    const payload = JSON.parse(params.payload);
    Logger.log("📦 payload=%s", JSON.stringify(payload));

    const action   = payload.actions?.[0]?.action_id || payload.actions?.[0]?.name || "";
    const userName = payload.user?.username || payload.user?.name || payload.user?.id || "unknown";

    Logger.log(`👤 ${userName} - action=${action}`);

    // Slackに即レス（ボタン押し確認）
    const resp = {
      response_type: "in_channel",
      replace_original: false,
      text: `✅ ${userName} さんが「${action === "punch_in" ? "出勤" : "退勤"}」を押しました！`,
    };

    const output = ContentService.createTextOutput(JSON.stringify(resp))
      .setMimeType(ContentService.MimeType.JSON);

    // シート書き込み（受信ログのみ）
    Utilities.sleep(300);
    saveLogOnly(userName, action);

    return output;

  } catch (err) {
    Logger.log("💥 doPost ERROR: %s", err.stack || err);
    return ContentService.createTextOutput("Error: " + err);
  }
}

// ===== URLエンコードされたデータをパース =====
function parseFormUrlEncoded(body) {
  const o = {};
  body.split("&").forEach(kv => {
    const [k, v] = kv.split("=");
    if (k) o[decodeURIComponent(k)] = decodeURIComponent(v || "");
  });
  return o;
}
// ===== ここまでコピー=====


// ===== 受信ログにのみ記録する関数 =====
function saveLogOnly(userName, action) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const logSheet = ss.getSheetByName(LOG_SHEET);

    if (!logSheet) {
      Logger.log("⚠ シート「" + LOG_SHEET + "」が見つかりません");
      return;
    }

    const now = new Date();
    const dateStr = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd");
    const timeStr = Utilities.formatDate(now, "Asia/Tokyo", "HH:mm:ss");

    // ロック取得（同時書き込み防止）
    const lock = LockService.getScriptLock();
    lock.tryLock(3000);

    // スプレッドシートに書き込み
    logSheet.appendRow([now, userName, action, dateStr, timeStr]);
    Logger.log("📝 受信ログに追記: " + userName + " / " + action);

    lock.releaseLock();
  } catch (err) {
    Logger.log("💥 saveLogOnly ERROR: " + err);
  }
}

function saveLogOnly_(userName, action) {
  try {
    Logger.log("🔍 openById 実行前: " + SPREADSHEET_ID);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log("✅ openById OK");
    const logSheet = ss.getSheetByName(LOG_SHEET);
    Logger.log("✅ getSheetByName OK");

    const now = new Date();
    const dateStr = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd");
    const timeStr = Utilities.formatDate(now, "Asia/Tokyo", "HH:mm:ss");

    const lock = LockService.getScriptLock();
    lock.tryLock(3000);

    logSheet.appendRow([now, userName, action, dateStr, timeStr]);
    Logger.log("📝 受信ログに追記しました: " + userName + " / " + action);

    lock.releaseLock();

  } catch (err) {
    Logger.log("💥 saveLogOnly_ ERROR: " + err);
  }
}  // ← ← ← ✨これが抜けてた！
function testAuth() {
  const id = "19V-S--MPEqAGgothYOfCRKNaq9-fuRLc-PYOJqpj6e8"; // ← IDだけ
  const ss = SpreadsheetApp.openById(id);
  const sheet = ss.getSheets()[0];
  Logger.log("✅ 認証成功: " + sheet.getName());
}

// ===== 勤怠記録へ転記（労働時間＋時給計算つき）=====
function updateAttendanceSheet() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const logSheet = ss.getSheetByName("受信ログ");
    const attendanceSheet = ss.getSheetByName("勤怠記録");
    const staffSheet = ss.getSheetByName("スタッフマスタ"); // ← ここ追加！

    if (!logSheet || !attendanceSheet || !staffSheet) {
      Logger.log("⚠ シートが見つかりません");
      return;
    }

    // --- スタッフマスタをマップ化（名前 → 時給）---
    const staffData = staffSheet.getDataRange().getValues();
    staffData.shift(); // ヘッダー除去
    const staffMap = new Map();
    staffData.forEach(([name, wage]) => {
      if (name && wage) staffMap.set(String(name).trim(), Number(wage));
    });

    // --- 受信ログの読み込み ---
    const logs = logSheet.getDataRange().getValues();
    logs.shift();
    const attendanceMap = new Map();

    logs.forEach(row => {
      const [timestamp, name, action, date, time] = row;
      if (!name || !date || !time) return;

      const key = `${date}_${name}`;
      const record = attendanceMap.get(key) || { date, name, in: "", out: "" };

      if (action === "punch_in") record.in = time;
      if (action === "punch_out") record.out = time;

      attendanceMap.set(key, record);
    });

   // --- 勤怠記録クリア＆ヘッダー再作成 ---
attendanceSheet.clearContents();
attendanceSheet.appendRow(["日付", "名前", "出勤", "退勤", "労働時間", "勤務金額", "休憩時間"]);

// --- データ書き込み（ヘッダと列数7つ揃える）---
const rows = [];
attendanceMap.forEach(r => rows.push([r.date, r.name, r.in, r.out, "", "", ""]));
if (rows.length) attendanceSheet.getRange(2, 1, rows.length, 7).setValues(rows);

// === 行数チェック ===
const lastRow = attendanceSheet.getLastRow();
if (lastRow < 2) {
  Logger.log("ℹ 明細0件。終了。");
  return;
}

const n = lastRow - 1;

// 💡 G列にデフォルト休憩時間（1:00）を自動挿入
const restRange = attendanceSheet.getRange(2, 7, n, 1);
const restValues = restRange.getValues().map(r => [r[0] || "1:00"]);
restRange.setValues(restValues);
restRange.setNumberFormat("[h]:mm");

// 💡 E列：労働時間（出勤-退勤-休憩）
attendanceSheet.getRange(2, 5, n, 1).setFormulaR1C1(
  '=IF(AND(RC[-2]<>"",RC[-1]<>""),(RC[-1]-RC[-2]-RC[2]),"")'
);

// 💡 F列：勤務金額
attendanceSheet.getRange(2, 6, n, 1).setFormulaR1C1(
  '=IF(RC[-1]="","",RC[-1]*24*VLOOKUP(RC[-4],\'スタッフマスタ\'!C1:C2,2,false))'
);

// 💡 表示形式
attendanceSheet.getRange(2, 5, n, 1).setNumberFormat("[h]:mm"); // 労働時間
attendanceSheet.getRange(2, 6, n, 1).setNumberFormat("¥#,##0"); // 金額

Logger.log("✅ 勤怠記録＋時給計算 更新OK（休憩時間デフォルト1h対応）");

  } catch (err) {
    Logger.log("💥 updateAttendanceSheet ERROR: " + err);
  }
}


