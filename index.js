// ============================================================
// 残業申請・警告システム（おむすび訪問看護ステーション）
// 既存の GAS コードに追記して使用してください
// ============================================================

// ====== 追加スクリプトプロパティ（設定が必要） ======
// MANAGER_SLACK_ID_1 = D098H53CYCX（例: U012XXXXXX）
// MANAGER_SLACK_ID_2 = D09958M3CUQ（例: U034XXXXXX）
// ※ SLACK_BOT_TOKEN / SPREADSHEET_ID / CHANNEL_ID は既存のまま流用

const OVERTIME_CHANNEL  = "C09946WKPDE";           // 勤怠希望チャンネル
const OVERTIME_SHEET    = "残業申請ログ";            // 残業申請記録シート名
const SCHEDULE_SHEET_OT = "スケジュール";            // カイポケPDF変換データシート名
const MANAGER_ID_1      = PropertiesService.getScriptProperties().getProperty("MANAGER_SLACK_ID_1");
const MANAGER_ID_2      = PropertiesService.getScriptProperties().getProperty("MANAGER_SLACK_ID_2");

// ============================================================
// 【初回のみ実行】残業申請ログシートを作成する
// ============================================================
function setupOvertimeSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(OVERTIME_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(OVERTIME_SHEET);
    sheet.appendRow([
      "対象日", "スタッフID", "スタッフ名",
      "退勤打刻", "スケジュール終了", "残業時間(分)",
      "申請状況",        // "未申請" / "申請済" / "承認済" / "スケジュール通り"
      "残業理由", "申請日時", "承認者", "承認日時",
      "警告送信回数", "スケジュール開始"  // 始業丸め用
    ]);
    sheet.setFrozenRows(1);
    Logger.log("✅ 残業申請ログシート作成完了");
  }
}

// ============================================================
// 【毎朝8時トリガーで実行】残業チェック＆未申請警告
// トリガー設定：時間ベース → 毎日 → 午前8時〜9時
// ============================================================
function dailyOvertimeCheck() {
  const ss            = SpreadsheetApp.openById(SPREADSHEET_ID);
  const logSheet      = ss.getSheetByName(LOG_SHEET);
  const scheduleSheet = ss.getSheetByName(SCHEDULE_SHEET_OT);
  const overtimeSheet = ss.getSheetByName(OVERTIME_SHEET);
  const staffSheet    = ss.getSheetByName("スタッフマスタ");

  if (!logSheet || !scheduleSheet || !overtimeSheet || !staffSheet) {
    Logger.log("⚠ 必要なシートが見つかりません");
    return;
  }

  // ===== スタッフマスタ → Map(ID → {name, slackUserId}) =====
  // スタッフマスタ列構成: [ID, 時給, 出勤時刻, 終了時刻, 日本語名, SlackUserID]
  const staffData = staffSheet.getDataRange().getValues();
  staffData.shift();
  const staffMap = new Map();
  staffData.forEach(([id, , , , name, slackUserId]) => {
    if (id) staffMap.set(String(id).trim(), { name: name || id, slackUserId: slackUserId || "" });
  });

  // ===== スケジュールシート → Map("日付_ID" → スケジュール終了時間) =====
  // スケジュールシート列構成: [職員名, 日付, 開始時間, 終了時間, 区分, 患者名]
  // ※ サマリーCSVをシートに貼った場合: [職員名, 日付, 最初の開始, 予定終了時間, 訪問件数]
  const scheduleData = scheduleSheet.getDataRange().getValues();
  scheduleData.shift();
  // key: "yyyy/MM/dd_staffId", value: { start: "HH:mm", end: "HH:mm" }
  const scheduleMap = new Map();

  scheduleData.forEach(row => {
    const staffName = String(row[0]).trim();
    const dateStr   = String(row[1]).trim();   // 例: "3/11(水)"
    const startTime = String(row[2]).trim();   // 最初の開始時間 例: "09:00"
    const endTime   = String(row[3]).trim();   // 予定終了時間   例: "17:30"

    // スタッフ名からIDを逆引き
    staffMap.forEach((info, id) => {
      if (info.name === staffName) {
        const key = `${dateStr}_${id}`;
        const existing = scheduleMap.get(key);
        scheduleMap.set(key, {
          // 最も早い開始時間を採用
          start: (!existing || startTime < existing.start) ? startTime : existing.start,
          // 最も遅い終了時間を採用
          end:   (!existing || endTime > existing.end)     ? endTime   : existing.end
        });
      }
    });
  });

  // ===== 受信ログから退勤時間を集計 =====
  const logs = logSheet.getDataRange().getValues();
  logs.shift();

  // 日付・IDごとに退勤時間を集める
  const punchMap = new Map(); // key: "yyyy/MM/dd_id", value: {out: "HH:mm", dateStr}
  logs.forEach(row => {
    const [, id, action, dateStr, timeStr] = row;
    if (!id || !dateStr || action !== "punch_out") return;
    const key = `${dateStr}_${String(id).trim()}`;
    // 最も遅い退勤を採用
    const existing = punchMap.get(key);
    if (!existing || timeStr > existing.out) {
      punchMap.set(key, { out: timeStr, dateStr, id: String(id).trim() });
    }
  });

  // ===== 既存の残業申請ログを読み込む =====
  const overtimeData = overtimeSheet.getDataRange().getValues();
  overtimeData.shift();
  // key: "日付_id", value: 行インデックス(0始まり)
  const overtimeMap = new Map();
  overtimeData.forEach((row, i) => {
    const key = `${row[0]}_${row[1]}`;
    overtimeMap.set(key, i);
  });

  // ===== 残業チェック：スケジュール終了 vs 退勤打刻 =====
  const today = new Date();

  punchMap.forEach((punch, punchKey) => {
    // スケジュールキーの形式に合わせる（受信ログは yyyy/MM/dd、スケジュールは 3/11(水) 形式）
    // スケジュールマップのキーをスキャンして日付マッチを探す
    let scheduleEndTime = null;
    let matchedScheduleKey = null;

    let scheduleStartTime = null;
    scheduleMap.forEach((schedule, scheduleKey) => {
      if (scheduleKey.endsWith(`_${punch.id}`)) {
        const scheduleDatePart = scheduleKey.split("_")[0];
        const convertedDate = convertScheduleDateToYMD(scheduleDatePart, today.getFullYear());
        if (convertedDate === punch.dateStr) {
          scheduleEndTime   = schedule.end;
          scheduleStartTime = schedule.start;
          matchedScheduleKey = scheduleKey;
        }
      }
    });

    if (!scheduleEndTime) return; // スケジュールなし → スキップ

    const punchOutMin  = timeToMinutes(punch.out);
    const scheduleMin  = timeToMinutes(scheduleEndTime);
    const overtimeMin  = punchOutMin - scheduleMin;

    if (overtimeMin <= 0) return; // 残業なし

    // ===== 残業申請ログに記録（未記録の場合のみ追加） =====
    const overtimeKey = `${punch.dateStr}_${punch.id}`;
    const staff = staffMap.get(punch.id) || { name: punch.id, slackUserId: "" };

    if (!overtimeMap.has(overtimeKey)) {
      // 新規残業記録
      overtimeSheet.appendRow([
        punch.dateStr,
        punch.id,
        staff.name,
        punch.out,
        scheduleEndTime,
        overtimeMin,
        "未申請",
        "",       // 残業理由
        "",       // 申請日時
        "",       // 承認者
        "",       // 承認日時
        0,        // 警告送信回数
        scheduleStartTime  // スケジュール開始（始業丸め用）
      ]);
      Logger.log(`📝 残業記録追加: ${staff.name} / ${punch.dateStr} / ${overtimeMin}分`);
    }
  });

  // ===== 未申請の残業に対してSlack通知 =====
  const updatedOvertimeData = overtimeSheet.getDataRange().getValues();
  updatedOvertimeData.shift();

  // スタッフごとに未申請をまとめる
  const pendingByStaff = new Map();

  updatedOvertimeData.forEach((row, i) => {
    const [dateStr, staffId, staffName, punchOut, schedEnd, overtimeMin, status, , , , , warnCount] = row;
    if (status !== "未申請") return;

    // 今日のデータは翌朝に送るので、昨日以前のみ対象
    if (dateStr === formatDate(today)) return;

    if (!pendingByStaff.has(staffId)) {
      pendingByStaff.set(staffId, { staffName, slackUserId: staffMap.get(staffId)?.slackUserId || "", items: [] });
    }
    pendingByStaff.get(staffId).items.push({ dateStr, overtimeMin, rowIndex: i + 2, warnCount });
  });

  // ===== Slack通知送信 =====
  pendingByStaff.forEach((data, staffId) => {
    sendOvertimeRequest(data, overtimeSheet);
  });
}

// ============================================================
// Slackに残業申請ボタンを送信する
// ============================================================
function sendOvertimeRequest(data, overtimeSheet) {
  const { staffName, items } = data;

  // 未申請リストのテキストを作成
  const itemLines = items.map(item => {
    const h = Math.floor(item.overtimeMin / 60);
    const m = item.overtimeMin % 60;
    const timeStr = h > 0 ? `${h}時間${m}分` : `${m}分`;
    return `　• ${item.dateStr}（残業 ${timeStr}）`;
  }).join("\n");

  const message = {
    channel: OVERTIME_CHANNEL,
    text: `⏰ 残業申請のお知らせ`,
    blocks: [
      {
        type: "section",
        text: {
          type: "mrkdwn",
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
            value: JSON.stringify({
              staffId: data.staffId || staffName,
              staffName: staffName,
              items: items.map(i => ({ dateStr: i.dateStr, overtimeMin: i.overtimeMin, rowIndex: i.rowIndex }))
            })
          },
          {
            type: "button",
            text: { type: "plain_text", text: "✅ スケジュール通りで完了", emoji: true },
            style: "danger",
            action_id: "schedule_as_is",
            value: JSON.stringify({
              staffId: data.staffId || staffName,
              staffName: staffName,
              items: items.map(i => ({ dateStr: i.dateStr, rowIndex: i.rowIndex }))
            })
          }
        ]
      }
    ]
  };

  callSlackApi("chat.postMessage", message);

  // 警告送信回数を更新
  items.forEach(item => {
    const cell = overtimeSheet.getRange(item.rowIndex, 12);
    cell.setValue((item.warnCount || 0) + 1);
  });

  Logger.log(`📨 残業申請通知送信: ${staffName} / ${items.length}件`);
}

// ============================================================
// 「✅ スケジュール通りで完了」ボタン処理
// → 退勤時間をスケジュール終了時間に自動カット＆申請状況を更新
// ============================================================
// ※ 既存 doPost に追記：
// if (action === "schedule_as_is") {
//   handleScheduleAsIs(payload);
//   return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.JSON);
// }

function handleScheduleAsIs(payload) {
  const ss            = SpreadsheetApp.openById(SPREADSHEET_ID);
  const overtimeSheet = ss.getSheetByName(OVERTIME_SHEET);
  const logSheet      = ss.getSheetByName(LOG_SHEET);

  const buttonValue   = JSON.parse(payload.actions?.[0]?.value || "{}");
  const { staffName, staffId, items } = buttonValue;
  const now    = new Date();
  const nowStr = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm");

  const updatedDates = [];

  items.forEach(item => {
    const rowIndex      = item.rowIndex;
    const scheduleEnd   = overtimeSheet.getRange(rowIndex, 5).getValue(); // スケジュール終了時間
    const scheduleStart = overtimeSheet.getRange(rowIndex, 13).getValue(); // スケジュール開始時間
    const dateStr       = overtimeSheet.getRange(rowIndex, 1).getValue();

    // 残業申請ログを「スケジュール通り」で更新
    overtimeSheet.getRange(rowIndex, 7).setValue("スケジュール通り");
    overtimeSheet.getRange(rowIndex, 9).setValue(nowStr);

    // 受信ログの退勤時間をスケジュール終了時間に上書き
    // ※ 受信ログを直接編集するのではなく、勤怠記録の更新時に反映される
    // updateAttendanceSheet() の allowOver 判定で自動カットされる想定

    updatedDates.push(dateStr);
  });

  const datesText = updatedDates.join("、");

  // チャンネルに通知
  const msg = {
    channel: OVERTIME_CHANNEL,
    text: `📅 スケジュール通りで完了`,
    blocks: [
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: `📅 *${staffName} さん*\n${datesText} の勤怠をスケジュール通りで完了しました。\n退勤時間はスケジュール終了時間で記録されます。`
        }
      }
    ]
  };

  callSlackApi("chat.postMessage", msg);
  Logger.log(`📅 スケジュール通り処理完了: ${staffName} / ${datesText}`);
}

// ============================================================
// モーダルを開く
// ============================================================
function handleOvertimeModalOpen(payload) {
  const triggerId = payload.trigger_id;
  const buttonValue = JSON.parse(payload.actions?.[0]?.value || "{}");
  const { staffName, items } = buttonValue;

  // 申請対象の日付リストをオプションとして作成
  const dateOptions = items.map(item => {
    const h = Math.floor(item.overtimeMin / 60);
    const m = item.overtimeMin % 60;
    const timeStr = h > 0 ? `${h}時間${m}分` : `${m}分`;
    return {
      text: { type: "plain_text", text: `${item.dateStr}（残業 ${timeStr}）` },
      value: String(item.rowIndex)
    };
  });

  const modal = {
    trigger_id: triggerId,
    view: {
      type: "modal",
      callback_id: "overtime_submit",
      private_metadata: JSON.stringify(buttonValue),
      title: { type: "plain_text", text: "残業申請" },
      submit: { type: "plain_text", text: "申請する" },
      close: { type: "plain_text", text: "キャンセル" },
      blocks: [
        {
          type: "section",
          text: { type: "mrkdwn", text: `*${staffName} さんの残業申請*\n申請する日程と理由を入力してください。` }
        },
        {
          type: "input",
          block_id: "target_dates",
          label: { type: "plain_text", text: "申請する日程（複数選択可）" },
          element: {
            type: "checkboxes",
            action_id: "dates_selected",
            options: dateOptions
          }
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
  };

  callSlackApi("views.open", modal);
}

// ============================================================
// モーダル送信（申請完了）の処理
// ============================================================
function handleOvertimeSubmit(payload) {
  const ss           = SpreadsheetApp.openById(SPREADSHEET_ID);
  const overtimeSheet = ss.getSheetByName(OVERTIME_SHEET);

  const values        = payload.view.state.values;
  const metadata      = JSON.parse(payload.view.private_metadata || "{}");
  const applicantName = metadata.staffName || "不明";

  // 選択された日程を取得
  const selectedOptions = values?.target_dates?.dates_selected?.selected_options || [];
  const reason          = values?.overtime_reason?.reason_input?.value || "";
  const now             = new Date();
  const nowStr          = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm");

  if (selectedOptions.length === 0) return;

  // 残業申請ログを更新
  const appliedDates = [];
  selectedOptions.forEach(opt => {
    const rowIndex = Number(opt.value);
    overtimeSheet.getRange(rowIndex, 7).setValue("申請済");   // 申請状況
    overtimeSheet.getRange(rowIndex, 8).setValue(reason);      // 残業理由
    overtimeSheet.getRange(rowIndex, 9).setValue(nowStr);      // 申請日時

    const dateStr = overtimeSheet.getRange(rowIndex, 1).getValue();
    appliedDates.push(dateStr);
  });

  // チャンネルに申請完了通知
  const datesText = appliedDates.join("、");
  const completeMsg = {
    channel: OVERTIME_CHANNEL,
    text: `✅ 残業申請が提出されました`,
    blocks: [
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: `✅ *${applicantName} さんが残業申請を提出しました*\n\n📅 対象日：${datesText}\n📝 理由：${reason}\n\n<@${MANAGER_ID_1}> <@${MANAGER_ID_2}> ご確認・承認をお願いします 🙏`
        }
      },
      {
        type: "actions",
        elements: [
          {
            type: "button",
            text: { type: "plain_text", text: "✅ 承認する", emoji: true },
            style: "primary",
            action_id: "approve_overtime",
            value: JSON.stringify({
              staffName: applicantName,
              rowIndexes: selectedOptions.map(o => Number(o.value)),
              dates: appliedDates
            })
          }
        ]
      }
    ]
  };

  callSlackApi("chat.postMessage", completeMsg);
  Logger.log(`✅ 残業申請完了: ${applicantName} / ${datesText}`);
}

// ============================================================
// 承認ボタン押下の処理
// ============================================================
// ※ 既存の doPost 内に以下を追記してください：
// if (action === "approve_overtime") {
//   handleOvertimeApprove(payload);
//   return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.JSON);
// }

function handleOvertimeApprove(payload) {
  const ss            = SpreadsheetApp.openById(SPREADSHEET_ID);
  const overtimeSheet = ss.getSheetByName(OVERTIME_SHEET);

  const approverName  = payload.user?.name || "管理者";
  const buttonValue   = JSON.parse(payload.actions?.[0]?.value || "{}");
  const { staffName, rowIndexes, dates } = buttonValue;
  const now    = new Date();
  const nowStr = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm");

  rowIndexes.forEach(rowIndex => {
    overtimeSheet.getRange(rowIndex, 7).setValue("承認済");   // 申請状況
    overtimeSheet.getRange(rowIndex, 10).setValue(approverName); // 承認者
    overtimeSheet.getRange(rowIndex, 11).setValue(nowStr);       // 承認日時
  });

  const datesText = dates.join("、");
  const approveMsg = {
    channel: OVERTIME_CHANNEL,
    text: `🎉 残業申請が承認されました`,
    blocks: [
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: `🎉 *${staffName} さんの残業申請が承認されました*\n\n📅 対象日：${datesText}\n👤 承認者：${approverName}（${nowStr}）`
        }
      }
    ]
  };

  callSlackApi("chat.postMessage", approveMsg);
  Logger.log(`🎉 残業承認完了: ${staffName} / ${datesText} / 承認者: ${approverName}`);
}

// ============================================================
// 既存 doPost に追記する処理（差し込み用）
// ============================================================
// 既存の doPost 関数の action 判定部分（labelMap の下あたり）に
// 以下のコードをまとめて追加してください：
//
// // --- 残業申請モーダルを開く ---
// if (action === "open_overtime_modal") {
//   handleOvertimeModalOpen(payload);
//   return ContentService.createTextOutput(JSON.stringify({ text: "" }))
//     .setMimeType(ContentService.MimeType.JSON);
// }
//
// // --- スケジュール通りで完了 ---
// if (action === "schedule_as_is") {
//   handleScheduleAsIs(payload);
//   return ContentService.createTextOutput(JSON.stringify({ text: "" }))
//     .setMimeType(ContentService.MimeType.JSON);
// }
//
// // --- 残業承認 ---
// if (action === "approve_overtime") {
//   handleOvertimeApprove(payload);
//   return ContentService.createTextOutput(JSON.stringify({ text: "" }))
//     .setMimeType(ContentService.MimeType.JSON);
// }
//
// // --- モーダル送信（view_submission） ---
// if (payload.type === "view_submission" && payload.view?.callback_id === "overtime_submit") {
//   handleOvertimeSubmit(payload);
//   return ContentService.createTextOutput(JSON.stringify({ response_action: "clear" }))
//     .setMimeType(ContentService.MimeType.JSON);
// }

// ============================================================
// ユーティリティ関数
// ============================================================

// Slack API を呼び出す汎用関数
function callSlackApi(method, body) {
  const response = UrlFetchApp.fetch(`https://slack.com/api/${method}`, {
    method: "post",
    contentType: "application/json",
    headers: { "Authorization": "Bearer " + SLACK_BOT_TOKEN },
    payload: JSON.stringify(body)
  });
  const result = JSON.parse(response.getContentText());
  if (!result.ok) {
    Logger.log(`⚠ Slack API エラー [${method}]: ${result.error}`);
  }
  return result;
}

// "HH:mm" → 分に変換
function timeToMinutes(timeStr) {
  if (!timeStr) return 0;
  const [h, m] = String(timeStr).split(":").map(Number);
  return (h || 0) * 60 + (m || 0);
}

// Date → "yyyy/MM/dd"
function formatDate(date) {
  return Utilities.formatDate(date, "Asia/Tokyo", "yyyy/MM/dd");
}

// "3/11(水)" → "2026/03/11" に変換
function convertScheduleDateToYMD(scheduleDateStr, year) {
  try {
    const match = scheduleDateStr.match(/(\d+)\/(\d+)/);
    if (!match) return "";
    const month = String(match[1]).padStart(2, "0");
    const day   = String(match[2]).padStart(2, "0");
    return `${year}/${month}/${day}`;
  } catch (e) {
    return "";
  }
}

// ============================================================
// セットアップ手順（初回のみ）
// ============================================================
// 1. setupOvertimeSheet() を手動実行 → 「残業申請ログ」シートを作成
// 2. スタッフマスタのF列に各スタッフの Slack User ID を追加
// 3. スクリプトプロパティに追加:
//    - MANAGER_SLACK_ID_1 = 川畑さんの Slack User ID
//    - MANAGER_SLACK_ID_2 = 岩崎さんの Slack User ID
// 4. 「スケジュール」シートを作成し、カイポケPDFから変換したサマリーCSVを毎週貼り付け
//    列構成: [職員名, 日付, 最初の開始, 予定終了時間, 訪問件数]
// 5. dailyOvertimeCheck() にトリガーを設定:
//    時間ベース → 毎日 → 午前8時〜9時
// 6. 既存の doPost 関数に、上記コメントの差し込みコードを追加
//
// ============================================================
// 始業時間の丸めについて
// ============================================================
// 打刻の出勤時間 < スケジュール開始時間 の場合
// → updateAttendanceSheet() の startMinutes 決定ロジックで
//    スケジュール開始時間（staffMap の startMinutes）に丸められます。
// ※ 早出申請（early=OK）がある場合のみ実打刻を採用します。
//
// ============================================================
// 退勤時間の自動カットについて
// ============================================================
// 「✅ スケジュール通りで完了」押下 or 残業申請なし の場合
// → 申請状況が "スケジュール通り" or "未申請" のまま
// → updateAttendanceSheet() の allowOver 判定（allowOverToday=false）で
//    スケジュール終了時間（staff.endMinutes）に自動カットされます。
// ============================================================
