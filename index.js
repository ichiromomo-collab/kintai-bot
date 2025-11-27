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

// ===== 勤怠記録へ転記（出勤丸め・休憩優先・割増計算・分単位計算）=====
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

    // --- スタッフマスタ（名前 → {時給, 出勤丸め分}）---
    const staffData = staffSheet.getDataRange().getValues();
    staffData.shift();
    const staffMap = new Map();

    staffData.forEach(([name, wage, start]) => {
      if (!name || !wage) return;
      const startStr = start
        ? Utilities.formatDate(start, "Asia/Tokyo", "HH:mm")
        : "";
      staffMap.set(String(name).trim(), {
        wage: Number(wage),
        startMinutes: startStr ? toMinutes(startStr) : null,
      });
    });

    // --- 受信ログ（同じ日付＋名前で集約）---
    const logs = logSheet.getDataRange().getValues();
    logs.shift();

    const map = new Map(); // key: "日付_名前" → {date,name,in,out,rest}

    logs.forEach(row => {
      const [ts, name, action, date, time, rest] = row;
      if (!name || !date || !time) return;

      const key = `${date}_${name}`;
      const obj = map.get(key) || { date, name, in: "", out: "", rest: "" };

      if (action === "punch_in") obj.in = time;
      if (action === "punch_out") obj.out = time;
      if (rest) obj.rest = rest; // 受信ログの休憩があれば優先

      map.set(key, obj);
    });

    // --- 勤怠記録初期化 ---
    attendanceSheet.clearContents();
    attendanceSheet.appendRow(["日付","名前","出勤","退勤","労働時間","勤務金額","休憩時間"]);

    const rows = [];

    map.forEach(rec => {
      const staff = staffMap.get(String(rec.name).trim());
      if (!staff) return;

      // ===== 出勤・退勤・休憩を「分」に変換 =====
      let startMinutes = null;
      let endMinutes   = null;

      // 出勤（丸めロジック）
      if (rec.in) {
        const pressedMin = toMinutes(rec.in);              // 実際押した時間
        const scheduled  = staff.startMinutes;             // マスタ出勤時間（分）

        if (scheduled != null && pressedMin < scheduled) {
          // 予定より前 → 丸めて scheduled
          startMinutes = scheduled;
        } else {
          // 予定以降 → 押した時間そのまま
          startMinutes = pressedMin;
        }
      }

      // 退勤（そのまま）
      if (rec.out) {
        endMinutes = toMinutes(rec.out);
      }

      // 休憩
      const restStr = rec.rest ? rec.rest : "1:00"; // 受信ログ優先、なければ1:00
      const restMinutes = toMinutes(restStr);

      // ===== 労働時間（分単位） =====
      let workMinutes = 0;
      if (startMinutes != null && endMinutes != null) {
        workMinutes = Math.max(0, endMinutes - startMinutes - restMinutes);
      }

      // ===== 割増計算（8h超は1.25倍） =====
      const normalMinutes   = Math.min(workMinutes, 480);
      const overtimeMinutes = Math.max(0, workMinutes - 480);

      const money =
        (normalMinutes / 60 * staff.wage) +
        (overtimeMinutes / 60 * staff.wage * 1.25);

      // ===== 出力用の表示文字列 =====
      const startStr = startMinutes != null ? minutesToHHMM(startMinutes) : "";
      const endStr   = endMinutes   != null ? minutesToHHMM(endMinutes)   : "";

      rows.push([
        rec.date,
        rec.name,
        startStr,
        endStr,
        minutesToHHMM(workMinutes),
        money,
        restStr
      ]);
    });

    if (rows.length) {
      attendanceSheet.getRange(2, 1, rows.length, 7).setValues(rows);
      attendanceSheet.getRange(2, 5, rows.length, 1).setNumberFormat("[h]:mm"); // 労働時間
      attendanceSheet.getRange(2, 6, rows.length, 1).setNumberFormat("¥#,##0"); // 金額
      attendanceSheet.getRange(2, 7, rows.length, 1).setNumberFormat("[h]:mm"); // 休憩
    }

    Logger.log("✅ 勤怠記録 更新OK（丸めロジック修正版）");

  } catch (err) {
    Logger.log("💥 updateAttendanceSheet ERROR: " + err);
  }
}



// === 時刻を「分」に変換 ===
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

// === 分 → "H:MM" 表示へ ===
function minutesToHHMM(min) {
  const h = Math.floor(min / 60);
  const m = min % 60;
  return `${h}:${m.toString().padStart(2, "0")}`;
}

// ===== 月末処理：スタッフごとに個人シートを生成 =====
function exportMonthlySheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const attendance = ss.getSheetByName("勤怠記録");

  const data = attendance.getDataRange().getValues();
  data.shift(); // ヘッダー除去

  if (data.length === 0) return;

  // 今月を抽出
  const today = new Date();
  const year  = today.getFullYear();
  const month = today.getMonth() + 1;
  const ymStr = `${year}/${month.toString().padStart(2,"0")}`;

  // スタッフごとにデータまとめる
  const map = new Map();

  data.forEach(row => {
  const [date, name, start, end, work, money, rest] = row;
  if (!name || !date) return;

  // --- 今月データの抽出（Date型対応）---
  const y = date.getFullYear();
  const m = date.getMonth() + 1;

  if (y !== year || m !== month) return;

  if (!map.has(name)) map.set(name, []);
  map.get(name).push(row);
  });


  // 各スタッフのシート作成
  map.forEach((rows, name) => {

    const sheetName = `${name}_${year}${String(month).padStart(2,"0")}`;

    // 既存なら削除して作り直す
    const old = ss.getSheetByName(sheetName);
    if (old) ss.deleteSheet(old);

    const newSheet = ss.insertSheet(sheetName);

    // ヘッダー
    newSheet.appendRow(["日付","名前","出勤","退勤","労働時間","勤務金額","休憩時間"]);
    
    // 本文
    newSheet.getRange(2,1,rows.length,7).setValues(rows);

    // 合計行
    const totalRow = rows.length + 3;
    newSheet.getRange(totalRow, 4).setValue("【合計】");
    newSheet.getRange(totalRow, 5).setFormula(`=SUM(E2:E${rows.length+1})`);
    newSheet.getRange(totalRow, 6).setFormula(`=SUM(F2:F${rows.length+1})`);

    // フォーマット
    newSheet.getRange(2,5,rows.length,1).setNumberFormat("[h]:mm");
    newSheet.getRange(2,6,rows.length,1).setNumberFormat("¥#,##0");
    newSheet.getRange(totalRow,5).setNumberFormat("[h]:mm");
    newSheet.getRange(totalRow,6).setNumberFormat("¥#,##0");
  });

  Logger.log("✅ 個人別月次シートの出力完了");
}


