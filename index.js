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
    const labelMap = {
   punch_in: "出勤",
   punch_out: "退勤",
   oncall: "オンコール"
   };
   const label = labelMap[action] || action;

    // Slackに即レス（ボタン押し確認）
   const resp = {
   response_type: "in_channel",
   replace_original: false,
     text: `✅ ${userName} さんが「${label}」を押しました！`,
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

// ===== 勤怠記録へ転記（出勤丸め・休憩は時間のみ・残業OKは受信ログで管理）=====
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

    // ===== スタッフマスタ → Map(ID → {漢字名, 時給, 丸め時刻, 定時}) =====
    const staffData = staffSheet.getDataRange().getValues();
    staffData.shift();

    const staffMap = new Map();
    staffData.forEach(([id, wage, startTime, endTime, fullName]) => {
      if (!id) return;

      staffMap.set(String(id).trim(), {
        id: String(id).trim(),
        name: fullName || id,
        wage: Number(wage) || 0,
        startMinutes: toMinutes(startTime), // 丸め開始
        endMinutes: toMinutes(endTime)      // 定時
      });
    });

    // ===== 受信ログを集約（同一日付＋同一ID）=====
    // 受信ログの列想定：
    // [ts, id, action, dateStr, timeStr, restStr, allowOver, early]
    const logs = logSheet.getDataRange().getValues();
    logs.shift();

    const map = new Map();

    logs.forEach(row => {
      const [ts, id, action, dateStr, timeStr, restStr, allowOver, early] = row;
      if (!id || !dateStr || !timeStr) return;

      const key = `${dateStr}_${String(id).trim()}`;
      const obj = map.get(key) || {
        date: dateStr,
        id: String(id).trim(),
        in: "",
        out: "",
        rest: "",
        allowOver: "" , //残業
        early: "" ,//早出
        oncall:""//オンコール
      };

      if (action === "punch_in")  obj.in  = timeStr;
      if (action === "punch_out") obj.out = timeStr;
      if (action === "oncall") obj.oncall = "OK";


      if (restStr) obj.rest = restStr;
      if (allowOver) obj.allowOver = String(allowOver).trim(); // "OK" 想定
      if (early) obj.early = String(early).trim();

      map.set(key, obj);
    });

    // ===== 勤怠記録 初期化（ここは消してOK。入力は受信ログだから問題なし）=====
    attendanceSheet.clearContents();
    attendanceSheet.appendRow(["日付","ID","名前","出勤","退勤","労働時間","勤務金額","休憩","残業許可"," 早出","オンコール"]);

    const rows = [];

    map.forEach(rec => {
      const staff = staffMap.get(String(rec.id).trim());
      if (!staff) return;

      const pressedStart = rec.in ? toMinutes(rec.in) : null;
      const pressedEnd   = rec.out ? toMinutes(rec.out) : null;

      // ==== 出勤時間決定 ====
     let startMinutes = pressedStart;

     // 早出OKなら実打刻を採用
     if (rec.early === "OK") {
       startMinutes = pressedStart;
     }
     // 早出でなければ丸め
     else if (
     pressedStart != null &&
     staff.startMinutes != null &&
      pressedStart < staff.startMinutes
     ) {
       startMinutes = staff.startMinutes;
      }


      // ==== 退勤 ====
      let endMinutes = pressedEnd;

      // ==== 残業許可（受信ログのOKを見る）====
      const allowOverToday = (rec.allowOver === "OK");

      // 残業NGの日は「定時」でカット（定時が設定されている場合）
      if (!allowOverToday && staff.endMinutes != null && endMinutes != null) {
        if (endMinutes > staff.endMinutes) endMinutes = staff.endMinutes;
      }

      // ==== 休憩（時間のみ管理）====
      // 休憩が受信ログに入ってればそれを優先。
      // なければ、労働が6時間未満→0分 / 6時間以上→60分
      let restStr;
      let restMinutes;

      if (rec.rest) {
        restStr = rec.rest;
        restMinutes = toMinutes(restStr);
      } else {
        if (startMinutes != null && endMinutes != null && (endMinutes - startMinutes) < 360) {
          restStr = "0:00";
          restMinutes = 0;
        } else {
          restStr = "1:00";
          restMinutes = 60;
        }
      }

      // ==== 労働時間 ====
      let workMinutes = 0;
      if (startMinutes != null && endMinutes != null) {
        workMinutes = Math.max(0, endMinutes - startMinutes - restMinutes);
      }

     // ==== オンコール手当 ====
        const ONCALL_FEE = 5000;
        let oncallFee = 0;

       if (rec.oncall === "OK") {
       oncallFee = ONCALL_FEE;
       }

      // ==== 金額（8時間超は1.25）====
      const normal = Math.min(workMinutes, 480);
      const over   = Math.max(0, workMinutes - 480);

      const money =
        (normal / 60 * staff.wage) +
        (over / 60 * staff.wage * 1.25) +
       oncallFee;
       

       rows.push([
        rec.date,
        staff.id,
        staff.name,
        startMinutes != null ? minutesToHHMM(startMinutes) : "",
        endMinutes   != null ? minutesToHHMM(endMinutes)   : "",
        minutesToHHMM(workMinutes),
        money,
        restStr,
        rec.allowOver || "",
         rec.early || "" ,
         rec.oncall|| ""
      ]);
    });

    // ===== 出力 =====
    if (rows.length) {
      attendanceSheet.getRange(2, 1, rows.length, 11).setValues(rows);
      attendanceSheet.getRange(2, 6, rows.length, 1).setNumberFormat("[h]:mm"); // 労働時間
      attendanceSheet.getRange(2, 7, rows.length, 1).setNumberFormat("¥#,##0"); // 金額
      attendanceSheet.getRange(2, 8, rows.length, 1).setNumberFormat("[h]:mm"); // 休憩
    }

     // ← ここで色付け復活
     // ===== 勤怠記録の色付け（段階グラデーション風） =====
     function applyAttendanceFormatting(sheet) {
     const lastRow = sheet.getLastRow();
     if (lastRow < 2) return;

     // 既存ルール全削除（重複防止）
     sheet.setConditionalFormatRules([]);
     const rules = [];
     const dataRows = Math.max(1, lastRow - 1);

     // ========= ① 出勤 or 退勤が片方欠けていたら警告（赤） =========
     rules.push(
     SpreadsheetApp.newConditionalFormatRule()
     .whenFormulaSatisfied(
      '=OR(AND($D2="", $E2<>""), AND($D2<>"", $E2=""))'
     )
     .setBackground("#F46A6A") // 警告赤
     .setRanges([sheet.getRange(`D2:E${lastRow}`)])
     .build()
     );

      // ========= ① 時間が入っているセル → 薄緑 =========
     const timeGreen = "#e6f4ea";

      ["D","E"].forEach(col => {
      rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND(${col}2<>"",${col}2<>0)`)
        .setBackground(timeGreen)
        .setRanges([sheet.getRange(`${col}2:${col}${lastRow}`)])
        .build()
       );
      });

      // ========= ② 労働時間（F列）黄色グラデーション =========
      rules.push(
     SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$F2>=TIME(8,0,0)')
      .setBackground("#FFE566") // 濃い黄
      .setRanges([sheet.getRange(2, 6, dataRows, 1)])
      .build()
      );

      rules.push(
     SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($F2>=TIME(5,30,0),$F2<TIME(8,0,0))')
      .setBackground("#FFF1AB") // 中黄
      .setRanges([sheet.getRange(2, 6, dataRows, 1)])
      .build()
      );

       // ========= ③ 休憩時間（H列）赤グラデーション =========
      rules.push(
     SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$H2>=TIME(1,0,0)')
      .setBackground("#F48383") // 濃赤
      .setRanges([sheet.getRange(2, 8, dataRows, 1)])
      .build()
      );

      rules.push(
      SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($H2>0,$H2<TIME(1,0,0))')
      .setBackground("#F4B4B4") // 薄赤
      .setRanges([sheet.getRange(2, 8, dataRows, 1)])
      .build()
       );

      // ========= ④ 残業許可 OK（I列） =========
      rules.push(
     SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("OK")
      .setBackground("#66C4FF")
      .setRanges([sheet.getRange(2, 9, dataRows, 1)])
      .build()
      );

      // ========= ⑤ 早出の日 → 出勤セルを色付け =========
     rules.push(
     SpreadsheetApp.newConditionalFormatRule()
     .whenFormulaSatisfied('=AND($J2="OK",$D2<>"")')
     .setBackground("#F6ADC6") // 
     .setRanges([sheet.getRange(`J2:J${lastRow}`)])
     .build()
     );

     // ========= オンコール（J列） =========
     rules.push(
     SpreadsheetApp.newConditionalFormatRule()
     .whenTextEqualTo("OK")
     .setBackground("#d9e1f2") // 薄い青
     .setRanges([sheet.getRange(2, 11, dataRows, 1)])
     .build()
     );


     sheet.setConditionalFormatRules(rules);
     }

      // 「残業許可=OK」だけ色付け（※毎回ルールを増やさないように、いったん置き換え）
      const lastRow = attendanceSheet.getLastRow();
      const rangeI = attendanceSheet.getRange(2, 9, Math.max(0, lastRow - 1), 1);

      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("OK")
        .setBackground("#ffd6d6")
        .setRanges([rangeI])
        .build();

      attendanceSheet.setConditionalFormatRules([rule]);
      applyAttendanceFormatting(attendanceSheet);

    

     Logger.log("✅ 勤怠記録 更新OK（残業OKは受信ログ管理）");

     } catch (err) {
     Logger.log("💥 updateAttendanceSheet ERROR: " + (err.stack || err));
     }
     }


// ====== 分変換 utilities ======
function toMinutes(v) {
  try {
    if (v instanceof Date) {
      return v.getHours() * 60 + v.getMinutes();
    }
    if (typeof v === "string") {
      const [h,m] = v.split(":").map(Number);
      return h*60 + m;
    }
    return 0;
  } catch(e) { return 0; }
}

function minutesToHHMM(min) {
  const h = Math.floor(min/60);
  const m = min % 60;
  return `${h}:${m.toString().padStart(2,"0")}`;
}


  // ===== ポップアップで年月を入力して出力 =====
  function exportMonthlySheetsPrompt() {

  // 入力を促すダイアログ
  const text = Browser.inputBox(
    "月次シート出力",
    "出力したい年月を 2025/11 の形式で入力してください。",
    Browser.Buttons.OK_CANCEL
  );

  if (text === "cancel") return;

  // 入力チェック
  const match = text.match(/^(\d{4})\/(\d{1,2})$/);
  if (!match) {
    Browser.msgBox("⚠ 入力形式が正しくありません。\n例: 2025/11");
    return;
  }

  const year = Number(match[1]);
  const month = Number(match[2]);

  // 実行
  exportMonthlySheets(year, month);

  Browser.msgBox(`📄 ${year}年${month}月 の個人シートを作成しました！`);
 }

  // ===== 月末個人シート（漢字名＋残業＋勤務金額） =====
  // exportMonthlySheets();          → 今月を出力
  // exportMonthlySheets(2025, 11);  → 2025年11月を出力
  function exportMonthlySheets(targetYear, targetMonth) {

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const attendance = ss.getSheetByName("勤怠記録");

  const data = attendance.getDataRange().getValues();
  data.shift(); // header除去

  // --- 引数が無ければ「今日の年月」を使う ---
  const today = new Date();
  const year  = targetYear  || today.getFullYear();
  const month = targetMonth || (today.getMonth() + 1);

  Logger.log(`📅 出力対象: ${year}年${month}月`);

  // --- スタッフごとにまとめる（IDごとに集計） ---
  const map = new Map();

  data.forEach(r => {
    // [日付, ID, 名前, 出勤, 退勤, 労働時間, 勤務金額, 休憩]
    const date = r[0];
    const id   = r[1];
    const fullName = r[2];

    if (!(date instanceof Date)) return;

    const y = date.getFullYear();
    const m = date.getMonth() + 1;

    if (y !== year || m !== month) return;

    if (!map.has(id)) {
      map.set(id, { name: fullName, rows: [] });
    }
    map.get(id).rows.push(r);
  });

  // ================= ==== シート出力 =====================
  map.forEach((obj, id) => {

    const name = obj.name;
    const rows = obj.rows;
    let oncallCount = 0;
    
    rows.forEach(r => {
    const oncall = r[10]; // K列（オンコール）
    if (oncall === "OK") oncallCount++;
   });

    const sheetName = `${name}_${year}${String(month).padStart(2, "0")}`;

    // 既存は削除
    const old = ss.getSheetByName(sheetName);
    if (old) ss.deleteSheet(old);

    const sh = ss.insertSheet(sheetName);

    // ヘッダー
    sh.appendRow(["日付", "ID", "名前", "出勤", "退勤", "労働時間", "勤務金額", "休憩","残業許可","早出","オンコール"]);

    // 本文
    sh.getRange(2, 1, rows.length, 11).setValues(rows);

    // ===== 自動フォーマット =====
    sh.getRange(2, 4, rows.length, 1).setNumberFormat("h:mm");     // 出勤
    sh.getRange(2, 5, rows.length, 1).setNumberFormat("h:mm");     // 退勤
    sh.getRange(2, 6, rows.length, 1).setNumberFormat("[h]:mm");   // 労働時間
    sh.getRange(2, 7, rows.length, 1).setNumberFormat("¥#,##0");   // 勤務金額
    sh.getRange(2, 8, rows.length, 1).setNumberFormat("[h]:mm");   // 休憩

    // ===== 合計行 =====
    const totalRow = rows.length + 3;

    // ラベル
    sh.getRange(totalRow, 3).setValue("【合計】");

    // 労働時間 合計
    sh.getRange(totalRow, 6)
      .setFormula(`=SUM(F2:F${rows.length + 1})`)
      .setNumberFormat("[h]:mm");

    // ===== 残業時間（8h超） =====
    const overtimeRow = totalRow + 1;
    sh.getRange(overtimeRow, 3).setValue("残業時間");

    sh.getRange(overtimeRow, 6)
      .setFormula(
        `=SUM(FILTER(F2:F${rows.length + 1}, F2:F${rows.length + 1} > TIME(8,0,0)))` +
        ` - TIME(8,0,0) * COUNT(FILTER(F2:F${rows.length + 1}, F2:F${rows.length + 1} > TIME(8,0,0)))`
      )
      .setNumberFormat("[h]:mm");

    // ===== 勤務金額 合計 =====
    const moneyRow = totalRow + 2;
    sh.getRange(moneyRow, 3).setValue("勤務金額 合計");

    sh.getRange(moneyRow, 7)
      .setFormula(`=SUM(G2:G${rows.length + 1})`)
      .setNumberFormat("¥#,##0");


      //オンコール回数
      const oncallRow = moneyRow + 1;

     sh.getRange(oncallRow, 3).setValue("オンコール回数");
     sh.getRange(oncallRow, 6).setValue(oncallCount + " 回");

     sh.getRange(oncallRow + 1, 3).setValue("オンコール手当");
     sh.getRange(oncallRow + 1, 7)
     .setValue(oncallCount * 5000)
     .setNumberFormat("¥#,##0");

      function applyStripeFormatting(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const rules = [];

  // 偶数行ストライプ
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=ISEVEN(ROW())')
      .setBackground('#f5f5f5')
      .setRanges([sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn())])
      .build()
  );

  // 合計行を強調
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$C2="【合計】"')
      .setBackground('#fff2cc')
      .setBold(true)
      .setRanges([sheet.getRange(1, 1, lastRow, sheet.getLastColumn())])
      .build()
  );

  sheet.setConditionalFormatRules(rules);
}

 


     Logger.log(`📄 作成: ${sheetName}`);
     });

      Logger.log("🎉 個人シート（年月指定対応） 完成！");
    }