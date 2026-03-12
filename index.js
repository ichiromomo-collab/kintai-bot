// ====== 設定 ======
const SLACK_BOT_TOKEN = PropertiesService.getScriptProperties().getProperty("SLACK_BOT_TOKEN");
const CHANNEL_ID      = PropertiesService.getScriptProperties().getProperty("CHANNEL_ID");
const LOG_SHEET       = "受信ログ";
const SPREADSHEET_ID  = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");

const ATTENDANCE_LOG_SHEET = "勤怠確認ログ";
const OVERTIME_CHANNEL  = "C09946WKPDE";
const OVERTIME_SHEET    = "残業申請ログ";
const SCHEDULE_SHEET_OT = "スケジュール";
const MANAGER_ID_1      = PropertiesService.getScriptProperties().getProperty("MANAGER_SLACK_ID_1");
const MANAGER_ID_2      = PropertiesService.getScriptProperties().getProperty("MANAGER_SLACK_ID_2");
const MANAGER_ID_3      = PropertiesService.getScriptProperties().getProperty("MANAGER_SLACK_ID_3");


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
          { name: "punch_in",  text: "出勤",      type: "button", style: "primary" },
          { name: "punch_out", text: "退勤",      type: "button", style: "danger"  },
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


// ===== Slack API呼び出し共通関数 =====
function callSlackApi(method, params) {
  const url = "https://slack.com/api/" + method;
  const res = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json; charset=utf-8",
    headers: { Authorization: "Bearer " + SLACK_BOT_TOKEN },
    payload: JSON.stringify(params),
    muteHttpExceptions: true
  });
  const json = JSON.parse(res.getContentText());
  if (!json.ok) {
    Logger.log("⚠ Slack API エラー [" + method + "]: " + json.error);
    Logger.log(JSON.stringify(json));
  }
  return json;
}

// ===== 時刻変換ユーティリティ =====
function toMinutes(val) {
  if (!val) return null;
  if (val instanceof Date) {
    return val.getHours() * 60 + val.getMinutes();
  }
  const str = String(val).trim();
  const m = str.match(/^(\d{1,2}):(\d{2})/);
  if (!m) return null;
  return parseInt(m[1]) * 60 + parseInt(m[2]);
}

function timeToMinutes(str) {
  if (!str) return 0;
  const m = String(str).match(/^(\d{1,2}):(\d{2})/);
  if (!m) return 0;
  return parseInt(m[1]) * 60 + parseInt(m[2]);
}

function minutesToHHMM(min) {
  const h = Math.floor(min / 60);
  const m = min % 60;
  return h + ":" + String(m).padStart(2, "0");
}

function formatDate(date) {
  return Utilities.formatDate(date, "Asia/Tokyo", "yyyy/MM/dd");
}

function convertScheduleDateToYMD(dateStr, year) {
  // "3/11(火)" → "2026/03/11"
  const m = dateStr.match(/(\d{1,2})\/(\d{1,2})/);
  if (!m) return dateStr;
  return year + "/" + String(m[1]).padStart(2, "0") + "/" + String(m[2]).padStart(2, "0");
}


// ===== SlackからのPOSTを受け取る =====
function doPost(e) {
  Logger.log("doPost受信: " + (e.postData ? e.postData.contents.substring(0, 200) : "なし"));
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

    // --- モーダル送信（view_submission）---
    if (payload.type === "view_submission" && payload.view?.callback_id === "overtime_submit") {
      handleOvertimeSubmit(payload);
      return ContentService.createTextOutput(JSON.stringify({ response_action: "clear" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const action   = payload.actions?.[0]?.action_id || payload.actions?.[0]?.name || "";
    const userName = payload.user?.username || payload.user?.name || payload.user?.id || "unknown";

    // --- 残業申請モーダルを開く ---
    if (action === "open_overtime_modal") {
      handleOvertimeModalOpen(payload);
      return ContentService.createTextOutput(JSON.stringify({ text: "" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // --- スケジュール通りで完了 ---
    if (action === "schedule_as_is") {
      handleScheduleAsIs(payload);
      return ContentService.createTextOutput(JSON.stringify({ text: "" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // --- 残業承認 ---
    if (action === "approve_overtime") {
      handleOvertimeApprove(payload);
      return ContentService.createTextOutput(JSON.stringify({ text: "" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // --- 勤怠確認：できてます ---
    // --- 勤怠修正完了ボタン ---
    if (action === "attendance_fixed") {
      handleAttendanceFixed(payload);
      return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT);
    }

    // --- 今日のおむすび：緊急携帯・ステータス ---
   if (action.startsWith("oncall_") || action.startsWith("status_")) {
  handleTodayStatus(payload);
  return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT);
}

    // --- 勤怠確認：できてます（残業許可申請済み）---
    if (action === "attendance_overtime_ok") {
      const { staffName, dateStr } = JSON.parse(payload.actions?.[0]?.value || "{}");
      callSlackApi("chat.update", {
        channel: payload.channel?.id,
        ts: payload.message?.ts,
        text: `🔖 ${staffName} さん（${dateStr}）残業許可申請済み`,
        blocks: [{ type: "section", text: { type: "mrkdwn",
          text: `🔖 *${staffName} さん*
${dateStr} の勤怠打刻確認済み（残業許可申請済み）` }}]
      });
      logAttendanceCheck(staffName, dateStr, "できてます（残業許可申請済み）");
      const managerIds = [MANAGER_ID_1, MANAGER_ID_2, MANAGER_ID_3].filter(Boolean);
      managerIds.forEach(mid => {
        callSlackApi("chat.postMessage", {
          channel: mid,
          text: `🔖 残業許可申請済みの報告`,
          blocks: [{ type: "section", text: { type: "mrkdwn",
            text: `🔖 *残業許可申請済みの報告*

*スタッフ：* ${staffName}
*対象日：* ${dateStr}

本人から「残業許可申請済み」と回答がありました。` }}]
        });
      });
      return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT);
    }

    if (action === "attendance_ok") {
      const { staffName, dateStr } = JSON.parse(payload.actions?.[0]?.value || "{}");
      // メッセージを「確認済み」に更新してボタンを消す
      callSlackApi("chat.update", {
        channel: payload.channel?.id,
        ts: payload.message?.ts,
        text: `✅ ${staffName} さん（${dateStr}）打刻確認済み`,
        blocks: [{
          type: "section",
          text: { type: "mrkdwn", text: `✅ *${staffName} さん*
${dateStr} の勤怠打刻を確認しました。` }
        }]
      });
      logAttendanceCheck(staffName, dateStr, "できてます");
      return ContentService.createTextOutput("")
        .setMimeType(ContentService.MimeType.TEXT);
    }

    // --- 勤怠確認：できてません ---
    if (action === "attendance_ng") {
      const { staffName, dateStr } = JSON.parse(payload.actions?.[0]?.value || "{}");
      // メッセージを更新してボタンを消す
      callSlackApi("chat.update", {
        channel: payload.channel?.id,
        ts: payload.message?.ts,
        text: `❌ ${staffName} さん（${dateStr}）打刻未完了`,
        blocks: [{
          type: "section",
          text: { type: "mrkdwn", text: `❌ *${staffName} さん*
${dateStr} の勤怠打刻が未完了です。管理者に連絡済み。` }
        }]
      });
      // ログ記録
      logAttendanceCheck(staffName, dateStr, "できてません");
      handleAttendanceNG(payload);
      return ContentService.createTextOutput(JSON.stringify({ text: "" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // --- 出勤・退勤・オンコール ---
    const labelMap = {
      punch_in:  "出勤",
      punch_out: "退勤",
      oncall:    "オンコール"
    };
    const label = labelMap[action] || action;

    const resp = {
      response_type: "in_channel",
      replace_original: false,
      text: `✅ ${userName} さんが「${label}」を押しました！`,
    };

    const output = ContentService.createTextOutput(JSON.stringify(resp))
      .setMimeType(ContentService.MimeType.JSON);

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

    const lock = LockService.getScriptLock();
    lock.tryLock(3000);
    logSheet.appendRow([now, userName, action, dateStr, timeStr]);
    Logger.log("📝 受信ログに追記: " + userName + " / " + action);
    lock.releaseLock();

    // 退勤打刻のときにリアルタイムで残業チェック
    if (action === "punch_out") {
      checkOvertimeOnPunchOut(ss, userName, dateStr, timeStr);
    }

  } catch (err) {
    Logger.log("💥 saveLogOnly ERROR: " + err);
  }
}

// ===== 退勤打刻時リアルタイム残業チェック =====
function checkOvertimeOnPunchOut(ss, userName, dateStr, punchOutTime) {
  try {
    const scheduleSheet = ss.getSheetByName(SCHEDULE_SHEET_OT);
    const overtimeSheet = ss.getSheetByName(OVERTIME_SHEET);
    const staffSheet    = ss.getSheetByName("スタッフマスタ");

    if (!scheduleSheet || !overtimeSheet || !staffSheet) return;

    // スタッフマスタから日本語名を取得
    const staffData = staffSheet.getDataRange().getValues();
    staffData.shift();
    let staffName = userName;
    staffData.forEach(([id, , , , name]) => {
      if (String(id).trim() === userName) staffName = name || userName;
    });

    // スケジュールシートからその日の終了時間を検索
    const scheduleData = scheduleSheet.getDataRange().getValues();
    scheduleData.shift();
    let scheduleEndTime = null, scheduleStartTime = null;
    const today = new Date();

    scheduleData.forEach(row => {
      if (String(row[0]).trim() !== staffName) return;
      const convertedDate = convertScheduleDateToYMD(String(row[1]).trim(), today.getFullYear());
      if (convertedDate !== dateStr) return;
      const endRaw = row[3], startRaw = row[2];
      const end   = endRaw   instanceof Date ? Utilities.formatDate(endRaw,   "Asia/Tokyo", "HH:mm") : String(endRaw).trim();
      const start = startRaw instanceof Date ? Utilities.formatDate(startRaw, "Asia/Tokyo", "HH:mm") : String(startRaw).trim();
      if (!scheduleEndTime || end > scheduleEndTime) {
        scheduleEndTime = end; scheduleStartTime = start;
      }
    });

    if (!scheduleEndTime) return;

    // HH:mm:ss → HH:mm に変換して比較
    const punchOutHHMM = punchOutTime.substring(0, 5);
    const overtimeMin  = timeToMinutes(punchOutHHMM) - timeToMinutes(scheduleEndTime);
    if (overtimeMin <= 0) return;

    // 残業申請ログに既に記録済みかチェック
    const overtimeData = overtimeSheet.getDataRange().getValues();
    overtimeData.shift();
    const alreadyLogged = overtimeData.some(row =>
      String(row[0]) === dateStr && String(row[1]) === userName
    );
    if (alreadyLogged) return;

    // 残業申請ログに追記
    const newRowIndex = overtimeSheet.getLastRow() + 1;
    overtimeSheet.appendRow([
      dateStr, userName, staffName, punchOutHHMM, scheduleEndTime,
      overtimeMin, "未申請", "", "", "", "", 0, scheduleStartTime
    ]);

    // その場でSlackにボタン送信
    const h = Math.floor(overtimeMin / 60);
    const m = overtimeMin % 60;
    const overtimeStr = `${h > 0 ? h+"時間" : ""}${m}分`;

    callSlackApi("chat.postMessage", {
      channel: OVERTIME_CHANNEL,
      text: "⏰ 残業申請のお知らせ",
      blocks: [
        {
          type: "section",
          text: { type: "mrkdwn",
            text: `⏰ *${staffName} さんへ*\n\n本日（${dateStr}）スケジュール終了時間を ${overtimeStr} 超えて退勤しました。\n残業申請が必要な場合はマネーフォワードから申請してください。`
          }
        },
        {
          type: "actions",
          elements: [
            {
              type: "button",
              text: { type: "plain_text", text: "📝 残業申請する", emoji: true },
              style: "primary",
              action_id: "open_overtime_modal",
              value: JSON.stringify({
                staffId: userName, staffName,
                items: [{ dateStr, overtimeMin, rowIndex: newRowIndex }]
              })
            }
          ]
        }
      ]
    });

    Logger.log(`⏰ リアルタイム残業通知送信: ${staffName} / ${overtimeStr}`);

  } catch (err) {
    Logger.log("💥 checkOvertimeOnPunchOut ERROR: " + err);
  }
}



// ===== 勤怠記録へ転記 =====
function updateAttendanceSheet() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const logSheet = ss.getSheetByName("受信ログ");
    const attendanceSheet = ss.getSheetByName("勤怠記録");
    const staffSheet = ss.getSheetByName("スタッフマスタ");

    if (!logSheet || !attendanceSheet || !staffSheet) {
      Logger.log("⚠ シートが見つからない");
      return;
    }

    const staffData = staffSheet.getDataRange().getValues();
    staffData.shift();

    const staffMap = new Map();
    staffData.forEach(([id, wage, startTime, endTime, fullName]) => {
      if (!id) return;
      staffMap.set(String(id).trim(), {
        id: String(id).trim(),
        name: fullName || id,
        wage: Number(wage) || 0,
        startMinutes: toMinutes(startTime),
        endMinutes: toMinutes(endTime)
      });
    });

    const logs = logSheet.getDataRange().getValues();
    logs.shift();

    const map = new Map();
    logs.forEach(row => {
      const [ts, id, action, dateStr, timeStr, restStr, allowOver, early] = row;
      if (!id || !dateStr || !timeStr) return;

      const key = `${dateStr}_${String(id).trim()}`;
      const obj = map.get(key) || {
        date: dateStr, id: String(id).trim(),
        in: "", out: "", rest: "", allowOver: "", early: "", oncall: ""
      };

      if (action === "punch_in")  obj.in  = timeStr;
      if (action === "punch_out") obj.out = timeStr;
      if (action === "oncall")    obj.oncall = "OK";
      if (restStr)    obj.rest     = restStr;
      if (allowOver)  obj.allowOver = String(allowOver).trim();
      if (early)      obj.early     = String(early).trim();

      map.set(key, obj);
    });

    attendanceSheet.clearContents();
    attendanceSheet.appendRow(["日付","ID","名前","出勤","退勤","労働時間","勤務金額","休憩","残業許可","早出","オンコール"]);

    const rows = [];
    map.forEach(rec => {
      const staff = staffMap.get(String(rec.id).trim());
      if (!staff) return;

      const pressedStart = rec.in  ? toMinutes(rec.in)  : null;
      const pressedEnd   = rec.out ? toMinutes(rec.out) : null;

      let startMinutes = pressedStart;
      if (rec.early === "OK") {
        startMinutes = pressedStart;
      } else if (pressedStart != null && staff.startMinutes != null && pressedStart < staff.startMinutes) {
        startMinutes = staff.startMinutes;
      }

      let endMinutes = pressedEnd;
      const allowOverToday = (rec.allowOver === "OK");
      if (!allowOverToday && staff.endMinutes != null && endMinutes != null) {
        if (endMinutes > staff.endMinutes) endMinutes = staff.endMinutes;
      }

      let restStr, restMinutes;
      if (rec.rest) {
        restStr = rec.rest;
        restMinutes = toMinutes(restStr);
      } else {
        if (startMinutes != null && endMinutes != null && (endMinutes - startMinutes) < 360) {
          restStr = "0:00"; restMinutes = 0;
        } else {
          restStr = "1:00"; restMinutes = 60;
        }
      }

      let workMinutes = 0;
      if (startMinutes != null && endMinutes != null) {
        workMinutes = Math.max(0, endMinutes - startMinutes - restMinutes);
      }

      const ONCALL_FEE = 5000;
      const oncallFee = (rec.oncall === "OK") ? ONCALL_FEE : 0;

      const normal = Math.min(workMinutes, 480);
      const over   = Math.max(0, workMinutes - 480);
      const money  = (normal / 60 * staff.wage) + (over / 60 * staff.wage * 1.25) + oncallFee;

      rows.push([
        rec.date, staff.id, staff.name,
        startMinutes != null ? minutesToHHMM(startMinutes) : "",
        endMinutes   != null ? minutesToHHMM(endMinutes)   : "",
        minutesToHHMM(workMinutes), money, restStr,
        rec.allowOver || "", rec.early || "", rec.oncall || ""
      ]);
    });

    if (rows.length) {
      attendanceSheet.getRange(2, 1, rows.length, 11).setValues(rows);
      attendanceSheet.getRange(2, 6, rows.length, 1).setNumberFormat("[h]:mm");
      attendanceSheet.getRange(2, 7, rows.length, 1).setNumberFormat("¥#,##0");
      attendanceSheet.getRange(2, 8, rows.length, 1).setNumberFormat("[h]:mm");
    }

    const lastRow = attendanceSheet.getLastRow();
    const rangeI  = attendanceSheet.getRange(2, 9, Math.max(0, lastRow - 1), 1);
    const rule    = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("OK").setBackground("#ffd6d6").setRanges([rangeI]).build();
    attendanceSheet.setConditionalFormatRules([rule]);
    applyAttendanceFormatting(attendanceSheet);

    Logger.log("✅ 勤怠記録 更新OK");

  } catch (err) {
    Logger.log("💥 updateAttendanceSheet ERROR: " + (err.stack || err));
  }
}

function applyAttendanceFormatting(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  sheet.setConditionalFormatRules([]);
  const rules = [];
  const dataRows = Math.max(1, lastRow - 1);

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR(AND($D2="", $E2<>""), AND($D2<>"", $E2=""))')
    .setBackground("#F46A6A").setRanges([sheet.getRange(`D2:E${lastRow}`)]).build());

  ["D","E"].forEach(col => {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(${col}2<>"",${col}2<>0)`)
      .setBackground("#e6f4ea").setRanges([sheet.getRange(`${col}2:${col}${lastRow}`)]).build());
  });

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F2>=TIME(8,0,0)')
    .setBackground("#FFE566").setRanges([sheet.getRange(2, 6, dataRows, 1)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($F2>=TIME(5,30,0),$F2<TIME(8,0,0))')
    .setBackground("#FFF1AB").setRanges([sheet.getRange(2, 6, dataRows, 1)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$H2>=TIME(1,0,0)')
    .setBackground("#F48383").setRanges([sheet.getRange(2, 8, dataRows, 1)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($H2>0,$H2<TIME(1,0,0))')
    .setBackground("#F4B4B4").setRanges([sheet.getRange(2, 8, dataRows, 1)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("OK").setBackground("#66C4FF")
    .setRanges([sheet.getRange(2, 9, dataRows, 1)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($J2="OK",$D2<>"")')
    .setBackground("#F6ADC6").setRanges([sheet.getRange(`J2:J${lastRow}`)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("OK").setBackground("#d9e1f2")
    .setRanges([sheet.getRange(2, 11, dataRows, 1)]).build());

  sheet.setConditionalFormatRules(rules);
}


// ===== 残業申請ログシート初期設定（初回のみ実行） =====
function setupOvertimeSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(OVERTIME_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(OVERTIME_SHEET);
    sheet.appendRow([
      "対象日","スタッフID","スタッフ名","退勤打刻","スケジュール終了","残業時間(分)",
      "申請状況","残業理由","申請日時","承認者","承認日時","警告送信回数","スケジュール開始"
    ]);
    sheet.setFrozenRows(1);
    Logger.log("✅ 残業申請ログシート作成完了");
  }
}


// ===== 勤怠確認ログシート初期設定（初回のみ実行） =====
function setupAttendanceLogSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(ATTENDANCE_LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(ATTENDANCE_LOG_SHEET);
    sheet.appendRow(["対象日", "スタッフ名", "回答", "回答日時", "修正者"]);
    sheet.setFrozenRows(1);
    Logger.log("✅ 勤怠確認ログシート作成完了");
  }
}

// ===== 毎朝8時トリガー：残業チェック＆未申請警告 =====
function dailyOvertimeCheck() {
  const ss            = SpreadsheetApp.openById(SPREADSHEET_ID);
  const logSheet      = ss.getSheetByName(LOG_SHEET);
  const scheduleSheet = ss.getSheetByName(SCHEDULE_SHEET_OT);
  const overtimeSheet = ss.getSheetByName(OVERTIME_SHEET);
  const staffSheet    = ss.getSheetByName("スタッフマスタ");

  if (!logSheet || !scheduleSheet || !overtimeSheet || !staffSheet) {
    Logger.log("⚠ 必要なシートが見つかりません"); return;
  }

  // スタッフマスタ読み込み [ID, 時給, 出勤時刻, 終了時刻, 日本語名, SlackUserID]
  const staffData = staffSheet.getDataRange().getValues();
  staffData.shift();
  const staffMap = new Map();
  staffData.forEach(([id, , , , name, slackUserId]) => {
    if (id) staffMap.set(String(id).trim(), { name: name || id, slackUserId: slackUserId || "" });
  });

  // スケジュールシート読み込み [職員名, 日付, 最初の開始, 予定終了時間, 訪問件数]
  const scheduleData = scheduleSheet.getDataRange().getValues();
  scheduleData.shift();
  const scheduleMap = new Map();
  scheduleData.forEach(row => {
    const staffName = String(row[0]).trim();
    const dateStr   = String(row[1]).trim();
    const startRaw2 = row[2], endRaw2 = row[3];
    const startTime = startRaw2 instanceof Date ? Utilities.formatDate(startRaw2, "Asia/Tokyo", "HH:mm") : String(startRaw2).trim();
    const endTime   = endRaw2   instanceof Date ? Utilities.formatDate(endRaw2,   "Asia/Tokyo", "HH:mm") : String(endRaw2).trim();
    staffMap.forEach((info, id) => {
      if (info.name === staffName) {
        const key = `${dateStr}_${id}`;
        const existing = scheduleMap.get(key);
        scheduleMap.set(key, {
          start: (!existing || startTime < existing.start) ? startTime : existing.start,
          end:   (!existing || endTime   > existing.end)   ? endTime   : existing.end
        });
      }
    });
  });

  // 受信ログから退勤時間集計
  const logs = logSheet.getDataRange().getValues();
  logs.shift();
  const punchMap = new Map();
  logs.forEach(row => {
    const [, id, action, dateStr, timeStr] = row;
    if (!id || !dateStr || action !== "punch_out") return;
    const key = `${dateStr}_${String(id).trim()}`;
    const existing = punchMap.get(key);
    if (!existing || timeStr > existing.out) {
      punchMap.set(key, { out: timeStr, dateStr, id: String(id).trim() });
    }
  });

  // 既存残業申請ログ読み込み
  const overtimeData = overtimeSheet.getDataRange().getValues();
  overtimeData.shift();
  const overtimeMap = new Map();
  overtimeData.forEach((row, i) => {
    overtimeMap.set(`${row[0]}_${row[1]}`, i);
  });

  const today = new Date();
  punchMap.forEach((punch) => {
    let scheduleEndTime = null, scheduleStartTime = null;
    scheduleMap.forEach((schedule, scheduleKey) => {
      if (scheduleKey.endsWith(`_${punch.id}`)) {
        const datePart = scheduleKey.split("_")[0];
        if (convertScheduleDateToYMD(datePart, today.getFullYear()) === punch.dateStr) {
          scheduleEndTime   = schedule.end;
          scheduleStartTime = schedule.start;
        }
      }
    });

    if (!scheduleEndTime) return;
    const overtimeMin = timeToMinutes(punch.out) - timeToMinutes(scheduleEndTime);
    if (overtimeMin <= 0) return;

    const overtimeKey = `${punch.dateStr}_${punch.id}`;
    const staff = staffMap.get(punch.id) || { name: punch.id, slackUserId: "" };

    if (!overtimeMap.has(overtimeKey)) {
      overtimeSheet.appendRow([
        punch.dateStr, punch.id, staff.name, punch.out, scheduleEndTime,
        overtimeMin, "未申請", "", "", "", "", 0, scheduleStartTime
      ]);
    }
  });

  // 未申請をスタッフごとにまとめてSlack通知
  const updatedData = overtimeSheet.getDataRange().getValues();
  updatedData.shift();
  const pendingByStaff = new Map();

  updatedData.forEach((row, i) => {
    const [dateStr, staffId, staffName, , , overtimeMin, status, , , , , warnCount] = row;
    if (status !== "未申請" || dateStr === formatDate(today)) return;
    if (!pendingByStaff.has(staffId)) {
      pendingByStaff.set(staffId, {
        staffId, staffName,
        slackUserId: staffMap.get(staffId)?.slackUserId || "",
        items: []
      });
    }
    pendingByStaff.get(staffId).items.push({ dateStr, overtimeMin, rowIndex: i + 2, warnCount });
  });

  pendingByStaff.forEach((data) => sendOvertimeRequest(data, overtimeSheet));
}


// ===== Slackに残業申請ボタンを送信 =====
function sendOvertimeRequest(data, overtimeSheet) {
  const { staffName, items } = data;

  const itemLines = items.map(item => {
    const h = Math.floor(item.overtimeMin / 60);
    const m = item.overtimeMin % 60;
    return `　• ${item.dateStr}（残業 ${h > 0 ? h+"時間" : ""}${m}分）`;
  }).join("\n");

  const btnValue = JSON.stringify({
    staffId: data.staffId, staffName,
    items: items.map(i => ({ dateStr: i.dateStr, overtimeMin: i.overtimeMin, rowIndex: i.rowIndex }))
  });

  callSlackApi("chat.postMessage", {
    channel: OVERTIME_CHANNEL,
    text: "⏰ 残業申請のお知らせ",
    blocks: [
      {
        type: "section",
        text: { type: "mrkdwn",
          text: `⏰ *${staffName} さんへ*\n\n以下の日程でスケジュール終了時間を超えた記録があります。\n残業申請が必要な場合はマネーフォワードから申請してください。\n\n${itemLines}`
        }
      },
      {
        type: "actions",
        elements: [
          {
            type: "button",
            text: { type: "plain_text", text: "📝 残業申請する", emoji: true },
            style: "primary",
            action_id: "open_overtime_modal",
            value: btnValue
          }
        ]
      }
    ]
  });

  // 警告送信回数を更新
  items.forEach(item => {
    if (overtimeSheet && item.rowIndex) {
      const warnCell = overtimeSheet.getRange(item.rowIndex, 12);
      warnCell.setValue((warnCell.getValue() || 0) + 1);
    }
  });

  Logger.log(`⏰ 残業申請ボタン送信: ${staffName}`);
}
function handleAttendanceNG(payload) {
  const { staffName, staffId, dateStr } = JSON.parse(payload.actions?.[0]?.value || "{}");
  const presserName = payload.user?.name || payload.user?.username || "unknown";
  const presserUserId = payload.user?.id || "";

  // 押した本人にDMで修正申請の案内を送る（ボタン付き）
  callSlackApi("chat.postMessage", {
    channel: presserUserId,
    text: "勤怠修正申請のご案内",
    blocks: [
      {
        type: "section",
        text: { type: "mrkdwn",
          text: `⚠️ *勤怠修正申請が必要です*\n\n${dateStr} の打刻が未完了です。\n\n👉 マネーフォワードで修正申請をしてください。`
        }
      },
      {
        type: "actions",
        elements: [{
          type: "button",
          text: { type: "plain_text", text: "✅ 修正しました", emoji: true },
          style: "primary",
          action_id: "attendance_fixed",
          value: JSON.stringify({ staffName, staffId, dateStr })
        }]
      }
    ]
  });

  // 管理者にDMで通知
  const managerIds = [MANAGER_ID_1, MANAGER_ID_2, MANAGER_ID_3].filter(Boolean);
  managerIds.forEach(managerId => {
    callSlackApi("chat.postMessage", {
      channel: managerId,
      text: `⚠️ 勤怠未打刻の報告`,
      blocks: [{
        type: "section",
        text: { type: "mrkdwn",
          text: `⚠️ *勤怠未打刻の報告*\n\n*スタッフ：* ${staffName}\n*対象日：* ${dateStr}\n\n本人から「打刻できていない」と回答がありました。\n勤怠修正の対応をお願いします。`
        }
      }]
    });
  });

  Logger.log(`⚠️ 勤怠未打刻報告: ${staffName} / ${dateStr}`);
}


// ===== 勤怠確認ログ記録 =====
function logAttendanceCheck(staffName, dateStr, answer) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(ATTENDANCE_LOG_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(ATTENDANCE_LOG_SHEET);
      sheet.appendRow(["対象日", "スタッフ名", "回答", "回答日時", "修正者"]);
      sheet.setFrozenRows(1);
    }
    const nowStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");
    sheet.appendRow([dateStr, staffName, answer, nowStr]);
    Logger.log(`✅ 勤怠確認ログ記録: ${staffName} / ${answer}`);
  } catch(err) {
    Logger.log("💥 logAttendanceCheck ERROR: " + err);
  }
}


// ===== 「修正完了」ボタン処理 =====
function handleAttendanceFixed(payload) {
  try {
    const { staffName, staffId, dateStr } = JSON.parse(payload.actions?.[0]?.value || "{}");
    const approverName = payload.user?.name || payload.user?.username || "管理者";
    const nowStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");

    // ログに修正完了を記録
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const logSheet = ss.getSheetByName(ATTENDANCE_LOG_SHEET);
    if (logSheet) {
      // 既存の「できてません」行を探して修正完了に更新
      const data = logSheet.getDataRange().getValues();
      let updated = false;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === dateStr && data[i][1] === staffName && data[i][2] === "できてません") {
          logSheet.getRange(i + 1, 3).setValue("修正完了");
          logSheet.getRange(i + 1, 5).setValue(`${approverName}（${nowStr}）`);
          updated = true;
          break;
        }
      }
      if (!updated) {
        logSheet.appendRow([dateStr, staffName, "修正完了", nowStr, approverName]);
      }
    }

    // DMのメッセージを更新してボタンを消す
    callSlackApi("chat.update", {
      channel: payload.channel?.id,
      ts: payload.message?.ts,
      text: `✅ ${staffName} さん（${dateStr}）修正完了`,
      blocks: [{
        type: "section",
        text: { type: "mrkdwn",
          text: `✅ *${staffName} さん（${dateStr}）の勤怠修正完了*\n対応者：${approverName}（${nowStr}）`
        }
      }]
    });

    Logger.log(`✅ 勤怠修正完了: ${staffName} / ${dateStr} / 対応: ${approverName}`);
  } catch(err) {
    Logger.log("💥 handleAttendanceFixed ERROR: " + err);
  }
}


// ===== 毎朝8時：勤怠確認ボタン送信 =====
function dailyAttendanceCheck() {
  const ss            = SpreadsheetApp.openById(SPREADSHEET_ID);
  const scheduleSheet = ss.getSheetByName(SCHEDULE_DETAIL_SHEET);
  const staffSheet    = ss.getSheetByName("スタッフマスタ");

  if (!scheduleSheet || !staffSheet) {
    Logger.log("⚠ シートが見つかりません"); return;
  }

  // 土日はスキップ
  const today = new Date();
  const todayDow = today.getDay();
  if (todayDow === 0 || todayDow === 6) {
    Logger.log("土日のためスキップ"); return;
  }

  // 対象日：月曜日なら金曜日（3日前）、それ以外は昨日
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - (todayDow === 1 ? 3 : 1));
  const yesterdayStr  = Utilities.formatDate(yesterday, "Asia/Tokyo", "yyyy/MM/dd");
  const yesterdayDisp = Utilities.formatDate(yesterday, "Asia/Tokyo", "M/d");

  // スタッフマスタ読み込み
  const staffData = staffSheet.getDataRange().getValues();
  staffData.shift();
  const staffMap = new Map();
  staffData.forEach(([id, , , , name, slackUserId]) => {
    if (id) staffMap.set(String(name).trim(), { id: String(id).trim(), slackUserId: slackUserId || "" });
  });

  // スケジュール詳細シートから昨日勤務があったスタッフを取得
  const scheduleData = scheduleSheet.getDataRange().getValues();
  scheduleData.shift();
  const workedStaff = new Set();
  const excludeStaff = ["川畑 麻衣子", "岩崎 里沙"];

  scheduleData.forEach(row => {
    const staffName = String(row[0]).trim();
    let dateStr;
    if (row[1] instanceof Date) {
      dateStr = Utilities.formatDate(row[1], "Asia/Tokyo", "yyyy/MM/dd");
    } else {
      dateStr = String(row[1]).trim();
    }
    if (dateStr === yesterdayStr && !excludeStaff.includes(staffName)) workedStaff.add(staffName);
  });

  if (workedStaff.size === 0) {
    Logger.log("昨日の勤務者なし"); return;
  }

  // 対象スタッフごとにボタン送信
  let sentCount = 0;
  workedStaff.forEach(staffName => {
    const staff = staffMap.get(staffName);
    if (!staff) return;

    const mentionText = staff.slackUserId ? `<@${staff.slackUserId}>` : staffName + " さん";

    callSlackApi("chat.postMessage", {
      channel: CHANNEL_ID,
      text: `📋 ${mentionText} 昨日の勤怠確認`,
      blocks: [
        {
          type: "section",
          text: { type: "mrkdwn",
            text: `📋 ${mentionText}\n\n昨日（${yesterdayDisp}）の勤怠打刻はできていますか？`
          }
        },
        {
          type: "actions",
          elements: [
            {
              type: "button",
              text: { type: "plain_text", text: "✅ できてます", emoji: true },
              style: "primary",
              action_id: "attendance_ok",
              value: JSON.stringify({ staffName, staffId: staff.id, dateStr: yesterdayStr })
            },
            {
              type: "button",
              text: { type: "plain_text", text: "🔖 できてます（残業許可申請済み）", emoji: true },
              action_id: "attendance_overtime_ok",
              value: JSON.stringify({ staffName, staffId: staff.id, dateStr: yesterdayStr })
            },
            {
              type: "button",
              text: { type: "plain_text", text: "❌ できてません", emoji: true },
              style: "danger",
              action_id: "attendance_ng",
              value: JSON.stringify({ staffName, staffId: staff.id, dateStr: yesterdayStr })
            }
          ]
        }
      ]
    });
    logAttendanceCheck(staffName, yesterdayStr, "送信済み");
    sentCount++;
  });

  Logger.log(`✅ 勤怠確認ボタン送信完了: ${sentCount}名`);

  // 未解決の「できてません」を再催促
  const logSheet2 = ss.getSheetByName(ATTENDANCE_LOG_SHEET);
  if (!logSheet2) return;
  const logData = logSheet2.getDataRange().getValues();
  logData.shift();
  logData.forEach(row => {
    if (row[2] !== "できてません") return;
    const isFixed = logData.some(r => r[0] === row[0] && r[1] === row[1] && r[2] === "修正完了");
    if (isFixed) return;
    const managerIds = [MANAGER_ID_1, MANAGER_ID_2, MANAGER_ID_3].filter(Boolean);
    managerIds.forEach(mid => {
      callSlackApi("chat.postMessage", {
        channel: mid,
        text: `🔁 勤怠未修正の再催促`,
        blocks: [
          {
            type: "section",
            text: { type: "mrkdwn",
              text: `🔁 *勤怠未修正の再催促*\n\n*スタッフ：* ${row[1]}\n*対象日：* ${row[0]}\n\nまだ修正が完了していません。対応をお願いします。マネーフォワードで申請後、修正完了ボタンを押してください。`
            }
          },
          {
            type: "actions",
            elements: [{
              type: "button",
              text: { type: "plain_text", text: "✅ 修正完了", emoji: true },
              style: "primary",
              action_id: "attendance_fixed",
              value: JSON.stringify({ staffName: row[1], dateStr: row[0] })
            }]
          }
        ]
      });
    });
  });
}


// ===== 昼12時：未回答者に再送 =====
function noonAttendanceReminder() {
  const ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
  const logSheet  = ss.getSheetByName(ATTENDANCE_LOG_SHEET);
  const staffSheet = ss.getSheetByName("スタッフマスタ");
  if (!logSheet || !staffSheet) return;

  const today = new Date();
  const todayDow2 = today.getDay();
  // 土日はスキップ
  if (todayDow2 === 0 || todayDow2 === 6) {
    Logger.log("土日のためスキップ"); return;
  }

  const yesterday2 = new Date();
  yesterday2.setDate(yesterday2.getDate() - (todayDow2 === 1 ? 3 : 1));
  const yesterdayStr = Utilities.formatDate(yesterday2, "Asia/Tokyo", "yyyy/MM/dd");
  const yesterdayDisp = Utilities.formatDate(yesterday2, "Asia/Tokyo", "M/d");

  // ログから今日送信済みの人と回答済みの人を取得
  const logData = logSheet.getDataRange().getValues();
  logData.shift();

  const sentStaff    = new Set();
  const answeredStaff = new Set();

  logData.forEach(row => {
    const dateStr = row[0] instanceof Date
      ? Utilities.formatDate(row[0], "Asia/Tokyo", "yyyy/MM/dd")
      : String(row[0]).trim();
    const name    = String(row[1]);
    const answer  = String(row[2]);
    if (dateStr !== yesterdayStr) return;
    if (answer === "送信済み") sentStaff.add(name);
    if (["できてます", "できてません", "できてます（残業許可申請済み）"].includes(answer)) {
      answeredStaff.add(name);
    }
  });

  // 未回答者を抽出して再送
  sentStaff.forEach(staffName => {
    if (answeredStaff.has(staffName)) return;

    callSlackApi("chat.postMessage", {
      channel: CHANNEL_ID,
      text: `🔔 ${staffName} さん、勤怠確認の回答をお願いします`,
      blocks: [
        {
          type: "section",
          text: { type: "mrkdwn",
            text: `🔔 *${staffName} さんへ*\n\n昨日（${yesterdayDisp}）の勤怠確認がまだ回答されていません。\n以下のボタンを押してください。`
          }
        },
        {
          type: "actions",
          elements: [
            {
              type: "button",
              text: { type: "plain_text", text: "✅ できてます", emoji: true },
              style: "primary",
              action_id: "attendance_ok",
              value: JSON.stringify({ staffName, dateStr: yesterdayStr })
            },
            {
              type: "button",
              text: { type: "plain_text", text: "🔖 できてます（残業許可申請済み）", emoji: true },
              action_id: "attendance_overtime_ok",
              value: JSON.stringify({ staffName, dateStr: yesterdayStr })
            },
            {
              type: "button",
              text: { type: "plain_text", text: "❌ できてません", emoji: true },
              style: "danger",
              action_id: "attendance_ng",
              value: JSON.stringify({ staffName, dateStr: yesterdayStr })
            }
          ]
        }
      ]
    });

    Logger.log(`🔔 未回答再送: ${staffName}`);
  });
}


// ===== 設定：今日のおむすびチャンネル =====
const TODAY_CHANNEL = "C0AKRKRLCQ5"; // #今日のおむすび
const OMUSUBI_LOG_SHEET = "今日のおむすびログ";

// スタッフ定義
const STAFF_CONFIG = [
  { name: "川畑 麻衣子", type: "nurse" },
  { name: "岩崎 里沙",   type: "nurse" },
  { name: "仲村渠 長代", type: "nurse" },
  { name: "今村 俊貴",   type: "nurse" },
  { name: "知念 美穂",   type: "office" },
  { name: "米須 珠美",   type: "sales"  },
];

// ===== 毎朝：今日のおむすび投稿 =====
function postTodayOmusubi() {
  const today = new Date();
  const todayDisp = Utilities.formatDate(today, "Asia/Tokyo", "M/d(E)");
  const todayStr  = Utilities.formatDate(today, "Asia/Tokyo", "yyyy/MM/dd");

  const blocks = [
    {
      type: "header",
      text: { type: "plain_text", text: `📋 今日のおむすび（${todayDisp}）`, emoji: true }
    },
    { type: "divider" },
    {
      type: "section",
      text: { type: "mrkdwn", text: "📱 *緊急携帯当番を選んでください*" }
    },
    {
      type: "actions",
      elements: [
        {
          type: "button",
          text: { type: "plain_text", text: "川畑さん", emoji: true },
          action_id: "oncall_kawabata",
          value: "川畑 麻衣子"
        },
        {
          type: "button",
          text: { type: "plain_text", text: "岩崎さん", emoji: true },
          action_id: "oncall_iwasaki",
          value: "岩崎 里沙"
        }
      ]
    },
    { type: "divider" }
  ];

  // スタッフごとのステータスボタン
  STAFF_CONFIG.forEach(staff => {
    blocks.push({
      type: "section",
      text: { type: "mrkdwn", text: `👤 *${staff.name}* さんのステータス` }
    });

    let buttons = [];
    if (staff.type === "nurse") {
      buttons = [
        { text: "🏥 訪問開始", action_id: "status_visit_start", value: staff.name },
        { text: "🚗 移動中",   action_id: "status_moving",      value: staff.name },
        { text: "✅ 空き",     action_id: "status_free",        value: staff.name },
      ];
    } else if (staff.type === "office") {
      buttons = [
        { text: "🏢 事務所",   action_id: "status_office",  value: staff.name },
        { text: "🚗 外出中",   action_id: "status_out",     value: staff.name },
        { text: "✅ 空き",     action_id: "status_free",    value: staff.name },
      ];
    } else if (staff.type === "sales") {
      buttons = [
        { text: "📊 営業中",   action_id: "status_sales",   value: staff.name },
        { text: "🏢 事務所",   action_id: "status_office",  value: staff.name },
        { text: "🚗 外出中",   action_id: "status_out",     value: staff.name },
        { text: "✅ 空き",     action_id: "status_free",    value: staff.name },
      ];
    }

    blocks.push({
      type: "actions",
      elements: buttons.map(b => ({
        type: "button",
        text: { type: "plain_text", text: b.text, emoji: true },
        action_id: b.action_id + "_" + staff.name.replace(/\s/g, "_"),
        value: JSON.stringify({ staffName: staff.name, status: b.text })
      }))
    });

    blocks.push({ type: "divider" });
  });

  // tsを保存してステータスシート初期化
  const result = callSlackApi("chat.postMessage", {
    channel: TODAY_CHANNEL,
    text: `📋 今日のおむすび（${todayDisp}）`,
    blocks
  });

  if (result?.ok) {
    const ts = result.message?.ts || result.ts || "";
    initOmusubiLog(todayStr, ts);
  }

  Logger.log("✅ 今日のおむすび投稿完了");
}


// ===== ステータス・緊急携帯ボタン処理 =====
function handleTodayStatus(payload) {
  const action  = payload.actions?.[0]?.action_id || "";
  const nowStr  = Utilities.formatDate(new Date(), "Asia/Tokyo", "HH:mm");
  const todayStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");
  Logger.log("🔍 action_id=" + action + " value=" + JSON.stringify(payload.actions?.[0]?.value));

  // 緊急携帯
  if (action.startsWith("oncall_")) {
    const personName = payload.actions?.[0]?.value || "";
    updateOmusubiLog(todayStr, "oncall", personName);
    updateOmusubiMessage(todayStr);
    // ログとして流す
    callSlackApi("chat.postMessage", {
      channel: TODAY_CHANNEL,
      text: `📱 緊急携帯当番：${personName}`,
      blocks: [{ type: "section",
        text: { type: "mrkdwn", text: `📱 *緊急携帯当番*
${nowStr} 現在：*${personName}* が担当します` }
      }]
    });
    return;
  }

  // ステータス変更
  if (action.startsWith("status_")) {
    const actionValue = payload.actions?.[0]?.value || "";
    let staffName = "", status = "";
    try {
      // +をスペースに戻してからJSONパース
      const decoded = actionValue.replace(/\+/g, " ");
      const parsed = JSON.parse(decoded);
      staffName = parsed.staffName || parsed.name || "";
      status = parsed.status || "";
    } catch(e) {
      staffName = actionValue.replace(/\+/g, " ");
      status = action.replace("status_", "");
    }
    updateOmusubiLog(todayStr, staffName, status);
    updateOmusubiMessage(todayStr);

    // 訪問開始の場合はスケジュール一覧も投稿
    if (action.startsWith("status_visit_start")) {
      const scheduleLines = getTodaySchedule(staffName, todayStr);
      const scheduleText = scheduleLines.length > 0
        ? scheduleLines.map(r => `　• ${r.start}〜${r.end}　${r.patient}${r.kind ? "（"+r.kind+"）" : ""}`).join("\n")
        : "　（スケジュールなし）";
      const currentPatient = scheduleLines.length > 0 ? scheduleLines[0].patient : "";

      callSlackApi("chat.postMessage", {
        channel: TODAY_CHANNEL,
        text: `🏥 訪問開始：${staffName}`,
        blocks: [{ type: "section",
          text: { type: "mrkdwn",
            text: `🏥 *${staffName}* 訪問開始${currentPatient ? "：" + currentPatient + "さん" : ""}\n（${nowStr} 更新）\n\n*本日のスケジュール：*\n${scheduleText}`
          }
        }]
      });
      
    } else {
  // チャンネルに通知
  callSlackApi("chat.postMessage", {
    channel: TODAY_CHANNEL,
    text: `${status}　${staffName}`,
    blocks: [{ type: "section",
      text: { type: "mrkdwn", text: `${status}　*${staffName}*\n（${nowStr} 更新）` }
    }]
  });
  // DMに次のボタンを送る
  sendNextStatusButton(staffName, status);
}
  }
}

// ===== 今日のスケジュール取得（高速版） =====
function getTodaySchedule(staffName, todayStr) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SCHEDULE_DETAIL_SHEET);
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    // A列・B列だけ先に読んで対象行を特定
    const nameCol = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const dateCol = sheet.getRange(2, 2, lastRow - 1, 1).getValues();

    const targetRows = [];
    for (let i = 0; i < nameCol.length; i++) {
      const name = String(nameCol[i][0]).trim();
      if (name !== staffName) continue;
      const d = dateCol[i][0];
      const dateStr = d instanceof Date
        ? Utilities.formatDate(d, "Asia/Tokyo", "yyyy/MM/dd")
        : String(d).trim();
      if (dateStr === todayStr) targetRows.push(i + 2); // 実際の行番号
    }

    Logger.log("getTodaySchedule: staffName=" + staffName + " todayStr=" + todayStr + " targetRows=" + targetRows.length);
    if (nameCol.length > 0) {
      const sampleDate = dateCol[0][0];
      const sampleFormatted = sampleDate instanceof Date
        ? Utilities.formatDate(sampleDate, "Asia/Tokyo", "yyyy/MM/dd")
        : String(sampleDate);
      Logger.log("サンプル日付: " + JSON.stringify(sampleDate) + " → " + sampleFormatted);
    }
    if (targetRows.length === 0) return [];

    const rows = [];
    targetRows.forEach(rowNum => {
      const r = sheet.getRange(rowNum, 1, 1, 6).getValues()[0];
      const start   = r[2] instanceof Date ? Utilities.formatDate(r[2], "Asia/Tokyo", "HH:mm") : String(r[2]).trim();
      const end     = r[3] instanceof Date ? Utilities.formatDate(r[3], "Asia/Tokyo", "HH:mm") : String(r[3]).trim();
      const kind    = String(r[4]).trim();
      const patient = String(r[5]).trim();
      rows.push({ start, end, kind, patient });
    });

    rows.sort((a, b) => a.start.localeCompare(b.start));
    return rows;
  } catch(err) {
    Logger.log("💥 getTodaySchedule ERROR: " + err);
    return [];
  }
}

// ===== 非同期おむすびステータス処理 =====
function processOmusubiAsync() {
  try {
    const props = PropertiesService.getScriptProperties();
    const cacheKey = props.getProperty("PENDING_OMUSUBI_KEY");
    if (!cacheKey) return;

    const cache = CacheService.getScriptCache();
    const payloadStr = cache.get(cacheKey);
    if (!payloadStr) return;

    const payload = JSON.parse(payloadStr);
    handleTodayStatus(payload);

    // クリーンアップ
    cache.remove(cacheKey);
    props.deleteProperty("PENDING_OMUSUBI_KEY");

    // このトリガー自身を削除
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => {
      if (t.getHandlerFunction() === "processOmusubiAsync") {
        ScriptApp.deleteTrigger(t);
      }
    });
  } catch(err) {
    Logger.log("💥 processOmusubiAsync ERROR: " + err);
  }
}


// ===== おむすびログ初期化 =====
function initOmusubiLog(todayStr, ts) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(OMUSUBI_LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(OMUSUBI_LOG_SHEET);
    sheet.appendRow(["日付", "メッセージts", "oncall", ...STAFF_CONFIG.map(s => s.name)]);
    sheet.setFrozenRows(1);
  }
  // 今日の行を追加
  const row = [todayStr, ts, "未定", ...STAFF_CONFIG.map(() => "")];
  sheet.appendRow(row);
  Logger.log("✅ おむすびログ初期化: " + todayStr);
}

// ===== おむすびログ更新 =====
function updateOmusubiLog(todayStr, key, value) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(OMUSUBI_LOG_SHEET);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colIndex = headers.indexOf(key);
  if (colIndex < 0) return;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === todayStr) {
      sheet.getRange(i + 1, colIndex + 1).setValue(value);
      return;
    }
  }
}

// ===== おむすびメッセージ更新 =====
function updateOmusubiMessage(todayStr) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(OMUSUBI_LOG_SHEET);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  let ts = "", oncall = "未定", statusMap = {};

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === todayStr) {
      ts = String(data[i][1]);
      oncall = String(data[i][2]) || "未定";
      STAFF_CONFIG.forEach(staff => {
        const col = headers.indexOf(staff.name);
        if (col >= 0) statusMap[staff.name] = String(data[i][col]) || "";
      });
      break;
    }
  }

  if (!ts) return;

  const todayDisp = Utilities.formatDate(new Date(), "Asia/Tokyo", "M/d(E)");
  const nowStr    = Utilities.formatDate(new Date(), "Asia/Tokyo", "HH:mm");

  // スタッフ一覧テキスト生成（訪問開始の場合は患者名も表示）
  const staffLines = STAFF_CONFIG.map(staff => {
    const s = statusMap[staff.name] || "－";
    if (s && s.includes("訪問")) {
      const todaySchedule = getTodaySchedule(staff.name, todayStr);
      const nowMin = new Date().getHours() * 60 + new Date().getMinutes();
      const current = todaySchedule.find(r => {
        const [sh, sm] = r.start.split(":").map(Number);
        const [eh, em] = r.end.split(":").map(Number);
        return (sh * 60 + sm) <= nowMin && nowMin <= (eh * 60 + em);
      }) || todaySchedule[0];
      const patientInfo = current ? `（${current.patient}さん）` : "";
      return `🏥 訪問中${patientInfo}　${staff.name}`;
    }
    return `${s || "－"}　${staff.name}`;
  }).join("\n");;

  const blocks = [
    { type: "header", text: { type: "plain_text", text: `📋 今日のおむすび（${todayDisp}）`, emoji: true } },
    { type: "divider" },
    { type: "section", text: { type: "mrkdwn",
      text: `📱 *緊急携帯：${oncall}*

👥 *スタッフ状況*（${nowStr} 更新）
${staffLines}`
    }},
    { type: "divider" },
    { type: "section", text: { type: "mrkdwn", text: "📱 *緊急携帯当番を選んでください*" }},
    { type: "actions", elements: [
      { type: "button", text: { type: "plain_text", text: "川畑さん", emoji: true }, action_id: "oncall_kawabata", value: "川畑 麻衣子" },
      { type: "button", text: { type: "plain_text", text: "岩崎さん",  emoji: true }, action_id: "oncall_iwasaki",  value: "岩崎 里沙" }
    ]},
    { type: "divider" }
  ];

  // ステータスボタン
  STAFF_CONFIG.forEach(staff => {
    blocks.push({ type: "section", text: { type: "mrkdwn", text: `👤 *${staff.name}* さんのステータス` }});
    let buttons = [];
    if (staff.type === "nurse") {
      buttons = [
        { text: "🏥 訪問開始", action_id: "status_visit_start" },
        { text: "🚗 移動中",   action_id: "status_moving" },
        { text: "✅ 空き",     action_id: "status_free" },
      ];
    } else if (staff.type === "office") {
      buttons = [
        { text: "🏢 事務所",   action_id: "status_office" },
        { text: "🚗 外出中",   action_id: "status_out" },
        { text: "✅ 空き",     action_id: "status_free" },
      ];
    } else {
      buttons = [
        { text: "📊 営業中",   action_id: "status_sales" },
        { text: "🏢 事務所",   action_id: "status_office" },
        { text: "🚗 外出中",   action_id: "status_out" },
        { text: "✅ 空き",     action_id: "status_free" },
      ];
    }
    blocks.push({ type: "actions", elements: buttons.map(b => ({
      type: "button",
      text: { type: "plain_text", text: b.text, emoji: true },
      action_id: b.action_id + "_" + staff.name.replace(/\s/g, "_"),
      value: JSON.stringify({ staffName: staff.name, status: b.text })
    }))});
    blocks.push({ type: "divider" });
  });

  callSlackApi("chat.update", {
    channel: TODAY_CHANNEL,
    ts,
    text: `📋 今日のおむすび（${todayDisp}）`,
    blocks
  });
}


// ===== カイポケPDF→スケジュール詳細シート変換 =====
const KAIPOKE_FOLDER_ID = "1HCI8dn5IxizqTQ4lMSUnIV5rUSQrGDlO";
const SCHEDULE_DETAIL_SHEET = "スケジュール詳細";

const STAFF_LIST = ["米須 珠美", "岩崎 里沙", "川畑 麻衣子", "今村 俊貴", "知念 美穂", "仲村渠 長代", "新垣 早紀", "入谷 京子"];

function importKaipokePDF() {
  const folder = DriveApp.getFolderById(KAIPOKE_FOLDER_ID);
  const files = folder.getFilesByType(MimeType.PDF);

  if (!files.hasNext()) {
    Logger.log("❌ PDFが見つかりません");
    return;
  }

  // 最新のPDFを取得
  let latestFile = null;
  let latestDate = new Date(0);
  while (files.hasNext()) {
    const f = files.next();
    if (f.getDateCreated() > latestDate) {
      latestDate = f.getDateCreated();
      latestFile = f;
    }
  }

  Logger.log("📄 変換対象: " + latestFile.getName());

  // PDFをGoogleドキュメントに変換（DriveApp経由）
  const blob = latestFile.getBlob().setContentType(MimeType.PDF);
  const docFile = DriveApp.getFolderById(KAIPOKE_FOLDER_ID).createFile(blob);
  const converted = Drive.Files.copy(
    { title: latestFile.getName() + "_converted", mimeType: MimeType.GOOGLE_DOCS },
    docFile.getId(),
    { convert: true }
  );
  const docId = converted.id;
  docFile.setTrashed(true); // コピー元を削除

  Logger.log("📄 変換ドキュメントID: " + docId);

  // Docs APIでテーブルを取得
  parseAndWriteSchedule(docId);
  
  // 変換ドキュメントは確認のため残す
  // DriveApp.getFileById(docId).setTrashed(true);
}

function parseAndWriteSchedule(docId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SCHEDULE_DETAIL_SHEET);
  if (!sheet) sheet = ss.insertSheet(SCHEDULE_DETAIL_SHEET);
  sheet.clearContents();
  sheet.appendRow(["職員名", "日付", "開始", "終了", "区分", "患者名"]);

  // Docs REST APIでテーブル取得
  const token = ScriptApp.getOAuthToken();
  const url = "https://docs.googleapis.com/v1/documents/" + docId;
  const res = UrlFetchApp.fetch(url, { headers: { Authorization: "Bearer " + token } });
  const docData = JSON.parse(res.getContentText());

  // セルテキスト取得ヘルパー
  function getCellText(cell) {
    if (!cell || !cell.content) return "";
    return cell.content.map(p => {
      if (!p.paragraph || !p.paragraph.elements) return "";
      return p.paragraph.elements.map(e => (e.textRun && e.textRun.content) ? e.textRun.content : "").join("");
    }).join("").replace(/\n/g, "\n").trim();
  }

  const bodyContent = docData.body && docData.body.content ? docData.body.content : [];
  const rows = [];

  for (const elem of bodyContent) {
    if (!elem.table) continue;
    const tableRows = elem.table.tableRows || [];
    if (tableRows.length < 2) continue;

    // 1行目：ヘッダー（職員名 | 日付列...）
    const headerCells = tableRows[0].tableCells || [];
    const headerText0 = getCellText(headerCells[0]);
    if (!headerText0.includes("職員氏名")) continue; // スケジュールテーブルのみ

    // 2行目以降：職員行
    for (let r = 1; r < tableRows.length; r++) {
      const cells = tableRows[r].tableCells || [];
      if (cells.length < 2) continue;
      const staffName = getCellText(cells[0]).trim();
      if (!staffName) continue;

      // 日付ヘッダー取得（2行目のヘッダー行から）
      // 1列目=職員名、2列目以降=日付+訪問情報
      for (let c = 1; c < cells.length; c++) {
        const dateHeader = getCellText((tableRows[0].tableCells || [])[c] || {});
        const cellText = getCellText(cells[c]);
        if (!cellText) continue;

        // 日付パース（例：3/11(火)）
        const dateMatch = dateHeader.match(/(\d+)\/(\d+)/);
        if (!dateMatch) continue;
        const month = dateMatch[1].padStart(2, "0");
        const day = dateMatch[2].padStart(2, "0");
        const year = new Date().getFullYear();
        const dateStr = year + "/" + month + "/" + day;

        // セル内の訪問を改行で分割
        // 形式：「09:00～10:00 医 患者名」または「09:00～10:00 患者名」
        const lines = cellText.split(/\n/).map(l => l.trim()).filter(l => l);
        for (const line of lines) {
          // 形式：「09:30～10:00（予定・医）　患者名」
          const timeMatch = line.match(/(\d{1,2}:\d{2})[～~](\d{1,2}:\d{2})/);
          if (!timeMatch) continue;
          const start = timeMatch[1];
          const end   = timeMatch[2];
          const kindMatch = line.match(/[（(]予定[・･]([^）)]+)[）)]/);
          const kind = kindMatch ? kindMatch[1] : "";
          const patient = line.replace(/.*[）)]\s*/, "").trim();
          if (patient) rows.push([staffName, dateStr, start, end, kind, patient]);
        }
      }
    }
  }

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
    Logger.log("✅ " + rows.length + "件書き込み完了");
  } else {
    Logger.log("⚠ テーブルデータが取得できませんでした");
  }
}
function testSimplePost() {
  const result = callSlackApi("chat.postMessage", {
    channel: TODAY_CHANNEL,
    text: "テスト投稿"
  });
  Logger.log(JSON.stringify(result));
}
// ===== DMに次のステータスボタンを送る =====
function sendNextStatusButton(staffName, currentStatus) {
  // スタッフマスタからSlackユーザーIDを取得
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const staffSheet = ss.getSheetByName("スタッフマスタ");
  if (!staffSheet) return;

  const staffData = staffSheet.getDataRange().getValues();
  staffData.shift();
  let slackUserId = "";
  staffData.forEach(([id, , , , name, userId]) => {
    if (String(name).trim() === staffName) slackUserId = userId || "";
  });
  if (!slackUserId) return;

  // スタッフの種別を取得
  const staffConf = STAFF_CONFIG.find(s => s.name === staffName);
  if (!staffConf) return;

  // 次のボタンを種別ごとに定義
  let buttons = [];
  if (staffConf.type === "nurse") {
    buttons = [
      { text: "🏥 訪問開始", action_id: "status_visit_start" },
      { text: "🚗 移動中",   action_id: "status_moving" },
      { text: "✅ 空き",     action_id: "status_free" },
    ];
  } else if (staffConf.type === "office") {
    buttons = [
      { text: "🏢 事務所",   action_id: "status_office" },
      { text: "🚗 外出中",   action_id: "status_out" },
      { text: "✅ 空き",     action_id: "status_free" },
    ];
  } else {
    buttons = [
      { text: "📊 営業中",   action_id: "status_sales" },
      { text: "🏢 事務所",   action_id: "status_office" },
      { text: "🚗 外出中",   action_id: "status_out" },
      { text: "✅ 空き",     action_id: "status_free" },
    ];
  }

  callSlackApi("chat.postMessage", {
    channel: slackUserId,
    text: `現在：${currentStatus}　次のステータスを選んでください`,
    blocks: [
      {
        type: "section",
        text: { type: "mrkdwn",
          text: `現在：*${currentStatus}*\n次のステータスを選んでください` }
      },
      {
        type: "actions",
        elements: buttons.map(b => ({
          type: "button",
          text: { type: "plain_text", text: b.text, emoji: true },
          action_id: b.action_id + "_" + staffName.replace(/\s/g, "_"),
          value: JSON.stringify({ staffName, status: b.text })
        }))
      }
    ]
  });
}