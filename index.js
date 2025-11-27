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

// ===== 勤怠記録へ転記（出勤丸め・休憩優先・割増計算・分単位の正確計算）=====
function updateAttendanceSheet() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const logSheet = ss.getSheetByName("受信ログ");
    const attendanceSheet = ss.getSheetByName("勤怠記録");
    const staffSheet = ss.getSheetByName("スタッフマスタ");

    if (!logSheet || !attendanceSheet || !staffSheet) {
      Logger.log("⚠ シートが見つかりません");
      return;
    }

    // --- スタッフマスタをマップ化（名前 → {時給, 出勤丸め時間}）---
    const staffData = staffSheet.getDataRange().getValues();
    staffData.shift(); // ヘッダー除去

    const staffMap = new Map();
    staffData.forEach(([name, wage, start]) => {
      if (name && wage) {
        staffMap.set(String(name).trim(), {
          wage: Number(wage),
          start: start ? Utilities.formatDate(start, "Asia/Tokyo", "HH:mm") : ""
        });
      }
    });

    // --- 受信ログ ---
    const logs = logSheet.getDataRange().getValues();
    logs.shift();

    const map = new Map(); // key: "日付_名前" → {in,out,rest}

    logs.forEach(row => {
      const [timestamp, name, action, date, time, restInput] = row;
      if (!name || !date || !time) return;

      const key = `${date}_${name}`;
      const obj = map.get(key) || { date, name, in: "", out: "", rest: "" };

      if (action === "punch_in") obj.in = time;
      if (action === "punch_out") obj.out = time;

      // ★ 休憩時間（優先）
      if (restInput) obj.rest = restInput;

      map.set(key, obj);
    });

    // --- 勤怠記録の初期化 ---
    attendanceSheet.clearContents();
    attendanceSheet.appendRow(["日付", "名前", "出勤", "退勤", "労働時間", "勤務金額", "休憩時間"]);

    const rows = [];

    map.forEach(rec => {
      const staff = staffMap.get(String(rec.name).trim());
      if (!staff) return;

      let start = rec.in;
      let end = rec.out;

      // --- 出勤丸め ---
      if (staff.start) {
        const scheduled = staff.start;      // 例 "08:30"
        const pressed = rec.in;             // 押した時刻 "08:17" など

        if (pressed) {
          if (pressed < scheduled) start = scheduled; // 早すぎ → 丸め上げ
          else start = pressed;                       // 遅刻 → そのまま
        }
      }

      // --- 退勤はそのまま ---
      if (!end) end = "";

      // --- 休憩 ---
      let restStr = rec.rest ? rec.rest : "1:00"; // 受信ログ優先・なければデフォルト

     // === 時刻を「分」に変換（Date型にも対応） ===
function toMinutes(v) {
  try {
    if (v instanceof Date) {
      return v.getHours() * 60 + v.getMinutes();
    }
    if (typeof v === "string") {
      const [h, m] = v.split(":").map(Number);
      return h * 60 + m;
    }
    return 0;
  } catch (e) {
    return 0;
  }
}


      // === 割増計算 ===
      const normalMinutes = Math.min(workMinutes, 480); // 8時間まで
      const overtimeMinutes = Math.max(0, workMinutes - 480);

      const wage = staff.wage;
      const money =
        (normalMinutes / 60 * wage) +
        (overtimeMinutes / 60 * wage * 1.25);

      rows.push([
        rec.date,
        rec.name,
        start,
        end,
        minutesToHHMM(workMinutes),
        money,
        restStr
      ]);
    });

    if (rows.length)
      attendanceSheet.getRange(2, 1, rows.length, 7).setValues(rows);

    // 表示形式
    attendanceSheet.getRange(2, 5, rows.length, 1).setNumberFormat("[h]:mm");
    attendanceSheet.getRange(2, 6, rows.length, 1).setNumberFormat("¥#,##0");
    attendanceSheet.getRange(2, 7, rows.length, 1).setNumberFormat("[h]:mm");

    Logger.log("✅ 完全版勤怠システム：更新OK");

  } catch (err) {
    Logger.log("💥 updateAttendanceSheet ERROR: " + err);
  }
}


// === 時刻を「分」に変換 ===
function toMinutes(str) {
  try {
    const [h, m] = str.split(":").map(Number);
    return h * 60 + m;
  } catch (e) {
    return 0;
  }
}

// === 分 → "H:MM" 表示へ ===
function minutesToHHMM(min) {
  const h = Math.floor(min / 60);
  const m = min % 60;
  return `${h}:${m.toString().padStart(2, "0")}`;
}



