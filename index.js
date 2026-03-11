// ====== 設定 ======
const SLACK_BOT_TOKEN = PropertiesService.getScriptProperties().getProperty("SLACK_BOT_TOKEN");
const CHANNEL_ID      = PropertiesService.getScriptProperties().getProperty("CHANNEL_ID");
const LOG_SHEET       = "受信ログ";
const SPREADSHEET_ID  = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");

const OVERTIME_CHANNEL  = "C09946WKPDE";
const OVERTIME_SHEET    = "残業申請ログ";
const SCHEDULE_SHEET_OT = "スケジュール";
const MANAGER_ID_1      = PropertiesService.getScriptProperties().getProperty("MANAGER_SLACK_ID_1");
const MANAGER_ID_2      = PropertiesService.getScriptProperties().getProperty("MANAGER_SLACK_ID_2");


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
  } catch (err) {
    Logger.log("💥 saveLogOnly ERROR: " + err);
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
    const startTime = String(row[2]).trim();
    const endTime   = String(row[3]).trim();
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
          text: `⏰ *${staffName} さんへ*\n\n以下の日程でスケジュール終了時間を超えた記録があります。\n残業申請をするか、スケジュール通りで完了するかをご選択ください。\n\n${itemLines}`
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
          },
          {
            type: "button",
            text: { type: "plain_text", text: "✅ スケジュール通りで完了", emoji: true },
            style: "danger",
            action_id: "schedule_as_is",
            value: JSON.stringify({
              staffId: data.staffId, staffName,
              items: items.map(i => ({ dateStr: i.dateStr, rowIndex: i.rowIndex }))
            })
          }
        ]
      }
    ]
  });

  items.forEach(item => {
    overtimeSheet.getRange(item.rowIndex, 12).setValue((item.warnCount || 0) + 1);
  });
}


// ===== 「スケジュール通りで完了」ボタン処理 =====
function handleScheduleAsIs(payload) {
  const ss            = SpreadsheetApp.openById(SPREADSHEET_ID);
  const overtimeSheet = ss.getSheetByName(OVERTIME_SHEET);
  const { staffName, items } = JSON.parse(payload.actions?.[0]?.value || "{}");
  const nowStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");
  const updatedDates = [];

  items.forEach(item => {
    overtimeSheet.getRange(item.rowIndex, 7).setValue("スケジュール通り");
    overtimeSheet.getRange(item.rowIndex, 9).setValue(nowStr);
    updatedDates.push(overtimeSheet.getRange(item.rowIndex, 1).getValue());
  });

  callSlackApi("chat.postMessage", {
    channel: OVERTIME_CHANNEL,
    text: "📅 スケジュール通りで完了",
    blocks: [{
      type: "section",
      text: { type: "mrkdwn",
        text: `📅 *${staffName} さん*\n${updatedDates.join("、")} の勤怠をスケジュール通りで完了しました。\n退勤時間はスケジュール終了時間で記録されます。`
      }
    }]
  });
}


// ===== 残業申請モーダルを開く =====
function handleOvertimeModalOpen(payload) {
  const { staffName, items } = JSON.parse(payload.actions?.[0]?.value || "{}");

  const dateOptions = items.map(item => {
    const h = Math.floor(item.overtimeMin / 60);
    const m = item.overtimeMin % 60;
    return {
      text: { type: "plain_text", text: `${item.dateStr}（残業 ${h > 0 ? h+"時間" : ""}${m}分）` },
      value: String(item.rowIndex)
    };
  });

  callSlackApi("views.open", {
    trigger_id: payload.trigger_id,
    view: {
      type: "modal",
      callback_id: "overtime_submit",
      private_metadata: payload.actions?.[0]?.value || "{}",
      title:  { type: "plain_text", text: "残業申請" },
      submit: { type: "plain_text", text: "申請する" },
      close:  { type: "plain_text", text: "キャンセル" },
      blocks: [
        {
          type: "section",
          text: { type: "mrkdwn", text: `*${staffName} さんの残業申請*\n申請する日程と理由を入力してください。` }
        },
        {
          type: "input",
          block_id: "target_dates",
          label: { type: "plain_text", text: "申請する日程（複数選択可）" },
          element: { type: "checkboxes", action_id: "dates_selected", options: dateOptions }
        },
        {
          type: "input",
          block_id: "overtime_reason",
          label: { type: "plain_text", text: "残業理由" },
          element: {
            type: "plain_text_input",
            action_id: "reason_input",
            multiline: true,
            placeholder: { type: "plain_text", text: "例：訪問記録の作成に時間がかかったため、〇〇様のバイタル対応のため など" }
          }
        }
      ]
    }
  });
}


// ===== モーダル送信（申請完了） =====
function handleOvertimeSubmit(payload) {
  const ss            = SpreadsheetApp.openById(SPREADSHEET_ID);
  const overtimeSheet = ss.getSheetByName(OVERTIME_SHEET);
  const values        = payload.view.state.values;
  const metadata      = JSON.parse(payload.view.private_metadata || "{}");
  const applicantName = metadata.staffName || "不明";
  const selectedOptions = values?.target_dates?.dates_selected?.selected_options || [];
  const reason          = values?.overtime_reason?.reason_input?.value || "";
  const nowStr          = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");

  if (selectedOptions.length === 0) return;

  const appliedDates = [];
  selectedOptions.forEach(opt => {
    const rowIndex = Number(opt.value);
    overtimeSheet.getRange(rowIndex, 7).setValue("申請済");
    overtimeSheet.getRange(rowIndex, 8).setValue(reason);
    overtimeSheet.getRange(rowIndex, 9).setValue(nowStr);
    appliedDates.push(overtimeSheet.getRange(rowIndex, 1).getValue());
  });

  callSlackApi("chat.postMessage", {
    channel: OVERTIME_CHANNEL,
    text: "✅ 残業申請が提出されました",
    blocks: [
      {
        type: "section",
        text: { type: "mrkdwn",
          text: `✅ *${applicantName} さんが残業申請を提出しました*\n\n📅 対象日：${appliedDates.join("、")}\n📝 理由：${reason}\n\n<@${MANAGER_ID_1}> <@${MANAGER_ID_2}> ご確認・承認をお願いします 🙏`
        }
      },
      {
        type: "actions",
        elements: [{
          type: "button",
          text: { type: "plain_text", text: "✅ 承認する", emoji: true },
          style: "primary",
          action_id: "approve_overtime",
          value: JSON.stringify({
            staffName: applicantName,
            rowIndexes: selectedOptions.map(o => Number(o.value)),
            dates: appliedDates
          })
        }]
      }
    ]
  });
}


// ===== 承認ボタン処理 =====
function handleOvertimeApprove(payload) {
  const ss            = SpreadsheetApp.openById(SPREADSHEET_ID);
  const overtimeSheet = ss.getSheetByName(OVERTIME_SHEET);
  const approverName  = payload.user?.name || "管理者";
  const { staffName, rowIndexes, dates } = JSON.parse(payload.actions?.[0]?.value || "{}");
  const nowStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");

  rowIndexes.forEach(rowIndex => {
    overtimeSheet.getRange(rowIndex, 7).setValue("承認済");
    overtimeSheet.getRange(rowIndex, 10).setValue(approverName);
    overtimeSheet.getRange(rowIndex, 11).setValue(nowStr);
  });

  callSlackApi("chat.postMessage", {
    channel: OVERTIME_CHANNEL,
    text: "🎉 残業申請が承認されました",
    blocks: [{
      type: "section",
      text: { type: "mrkdwn",
        text: `🎉 *${staffName} さんの残業申請が承認されました*\n\n📅 対象日：${dates.join("、")}\n👤 承認者：${approverName}（${nowStr}）`
      }
    }]
  });
}


// ====== 分変換ユーティリティ ======
function toMinutes(v) {
  try {
    if (v instanceof Date) return v.getHours() * 60 + v.getMinutes();
    if (typeof v === "string") {
      const [h, m] = v.split(":").map(Number);
      return h * 60 + m;
    }
    return 0;
  } catch(e) { return 0; }
}

function minutesToHHMM(min) {
  const h = Math.floor(min / 60);
  const m = min % 60;
  return `${h}:${m.toString().padStart(2, "0")}`;
}

function timeToMinutes(timeStr) {
  if (!timeStr) return 0;
  const [h, m] = String(timeStr).split(":").map(Number);
  return (h || 0) * 60 + (m || 0);
}

function formatDate(date) {
  return Utilities.formatDate(date, "Asia/Tokyo", "yyyy/MM/dd");
}

function convertScheduleDateToYMD(scheduleDateStr, year) {
  try {
    const match = scheduleDateStr.match(/(\d+)\/(\d+)/);
    if (!match) return "";
    return `${year}/${String(match[1]).padStart(2,"0")}/${String(match[2]).padStart(2,"0")}`;
  } catch(e) { return ""; }
}

function callSlackApi(method, body) {
  const response = UrlFetchApp.fetch(`https://slack.com/api/${method}`, {
    method: "post",
    contentType: "application/json",
    headers: { "Authorization": "Bearer " + SLACK_BOT_TOKEN },
    payload: JSON.stringify(body)
  });
  const result = JSON.parse(response.getContentText());
  if (!result.ok) Logger.log(`⚠ Slack API エラー [${method}]: ${result.error}`);
  return result;
}

function testAuth() {
  const id = "19V-S--MPEqAGgothYOfCRKNaq9-fuRLc-PYOJqpj6e8";
  const ss = SpreadsheetApp.openById(id);
  const sheet = ss.getSheets()[0];
  Logger.log("✅ 認証成功: " + sheet.getName());
}


// ===== ポップアップで年月を入力して出力 =====
function exportMonthlySheetsPrompt() {
  const text = Browser.inputBox("月次シート出力", "出力したい年月を 2025/11 の形式で入力してください。", Browser.Buttons.OK_CANCEL);
  if (text === "cancel") return;
  const match = text.match(/^(\d{4})\/(\d{1,2})$/);
  if (!match) { Browser.msgBox("⚠ 入力形式が正しくありません。\n例: 2025/11"); return; }
  exportMonthlySheets(Number(match[1]), Number(match[2]));
  Browser.msgBox(`📄 ${match[1]}年${match[2]}月 の個人シートを作成しました！`);
}


// ===== 月末個人シート出力 =====
function exportMonthlySheets(targetYear, targetMonth) {
  const ss         = SpreadsheetApp.openById(SPREADSHEET_ID);
  const attendance = ss.getSheetByName("勤怠記録");
  const data       = attendance.getDataRange().getValues();
  data.shift();

  const today = new Date();
  const year  = targetYear  || today.getFullYear();
  const month = targetMonth || (today.getMonth() + 1);

  const map = new Map();
  data.forEach(r => {
    const date = r[0];
    if (!(date instanceof Date)) return;
    if (date.getFullYear() !== year || date.getMonth() + 1 !== month) return;
    const id = r[1];
    if (!map.has(id)) map.set(id, { name: r[2], rows: [] });
    map.get(id).rows.push(r);
  });

  map.forEach((obj, id) => {
    const { name, rows } = obj;
    let oncallCount = 0;
    rows.forEach(r => { if (r[10] === "OK") oncallCount++; });

    const sheetName = `${name}_${year}${String(month).padStart(2, "0")}`;
    const old = ss.getSheetByName(sheetName);
    if (old) ss.deleteSheet(old);
    const sh = ss.insertSheet(sheetName);

    sh.appendRow(["日付","ID","名前","出勤","退勤","労働時間","勤務金額","休憩","残業許可","早出","オンコール"]);
    sh.getRange(2, 1, rows.length, 11).setValues(rows);
    sh.getRange(2, 4, rows.length, 1).setNumberFormat("h:mm");
    sh.getRange(2, 5, rows.length, 1).setNumberFormat("h:mm");
    sh.getRange(2, 6, rows.length, 1).setNumberFormat("[h]:mm");
    sh.getRange(2, 7, rows.length, 1).setNumberFormat("¥#,##0");
    sh.getRange(2, 8, rows.length, 1).setNumberFormat("[h]:mm");

    const totalRow   = rows.length + 3;
    const overtimeRow = totalRow + 1;
    const moneyRow   = totalRow + 2;
    const oncallRow  = moneyRow + 1;

    sh.getRange(totalRow, 3).setValue("【合計】");
    sh.getRange(totalRow, 6).setFormula(`=SUM(F2:F${rows.length + 1})`).setNumberFormat("[h]:mm");
    sh.getRange(overtimeRow, 3).setValue("残業時間");
    sh.getRange(overtimeRow, 6)
      .setFormula(`=SUM(FILTER(F2:F${rows.length+1},F2:F${rows.length+1}>TIME(8,0,0)))-TIME(8,0,0)*COUNT(FILTER(F2:F${rows.length+1},F2:F${rows.length+1}>TIME(8,0,0)))`)
      .setNumberFormat("[h]:mm");
    sh.getRange(moneyRow, 3).setValue("勤務金額 合計");
    sh.getRange(moneyRow, 7).setFormula(`=SUM(G2:G${rows.length + 1})`).setNumberFormat("¥#,##0");
    sh.getRange(oncallRow, 3).setValue("オンコール回数");
    sh.getRange(oncallRow, 6).setValue(oncallCount + " 回");
    sh.getRange(oncallRow + 1, 3).setValue("オンコール手当");
    sh.getRange(oncallRow + 1, 7).setValue(oncallCount * 5000).setNumberFormat("¥#,##0");

    applyStripeFormatting(sh);
    Logger.log(`📄 作成: ${sheetName}`);
  });
}


function applyStripeFormatting(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const lastCol = sheet.getLastColumn();
  const rules = [];

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=ROW()=1').setBackground('#F6DEE6').setBold(true)
    .setRanges([sheet.getRange(1, 1, 1, lastCol)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISODD(ROW()), $C2<>"【合計】")')
    .setBackground('#CDE6C7').setRanges([sheet.getRange(2, 1, lastRow - 1, lastCol)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$C2="【合計】"').setBackground('#FFF2CC').setBold(true)
    .setRanges([sheet.getRange(2, 1, lastRow, lastCol)]).build());

  sheet.setConditionalFormatRules(rules);
  Logger.log("🎉 個人シート完成！");
}