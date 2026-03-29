function updateAndSyncSchedule() {
  const currentSS = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. スクリプトプロパティから全設定を取得 ---
  const props = PropertiesService.getScriptProperties().getProperties();
  const geminiApiKey = props['Gemini_key'];

  // カラー設定（小文字に統一して比較準備）
  const colorAttend = (props['Color_Attend'] || "#ff0000").toLowerCase();
  const colorMain   = (props['Color_Main']   || "#00ffff").toLowerCase();

  // --- IDまたはURLからIDを抽出するヘルパー ---
  // Google Drive URL の /d/XXXXX 形式を優先し、次に ?id=XXXXX、最後に25文字以上の英数字列にフォールバック
  const extractId = (input) => {
    if (!input || typeof input !== 'string') return input;
    const slashD = input.match(/\/d\/([-\w]{25,})/);
    if (slashD) return slashD[1];
    const queryId = input.match(/[?&]id=([-\w]{25,})/);
    if (queryId) return queryId[1];
    const plain = input.match(/[-\w]{25,}/);
    return plain ? plain[0] : input;
  };

  // --- 2. スクリプトプロパティから各種設定を取得 ---
  const sourceFileId = extractId(props['Inport_URL'] || "");
  const targetSSId   = extractId(props['Export_URL']  || "");
  const syncLineNo   = parseInt(props['SyncLine_No']  || "1");
  const workFolderId = extractId(props['Work_Dir']    || "");

  const syncTag = "-- Sync_Scheduleによって自動登録 --";
  const now = new Date();
  const syncTimestamp = Utilities.formatDate(now, "JST", "yyyy/MM/dd HH:mm");

  // キャッシュと集計
  let addressCache = {};
  let geminiCount = 0;

  // --- 3. 稼働カレンダーの取得 ---
  const workCalendarId = (props['Work_CalendarID'] || "").trim();
  let calendar = null;
  if (workCalendarId) {
    calendar = CalendarApp.getCalendarById(workCalendarId);
    if (!calendar) {
      console.error("【エラー】Work_CalendarID のカレンダーが見つかりません（ID: " + workCalendarId + "）。処理を中断します。");
      return;
    }
    console.log("カレンダー接続完了 \n接続カレンダー名：【" + calendar.getName() + "】 \nカレンダーID:" + workCalendarId);
  } else {
    console.warn("【警告】Work_CalendarID が未設定のため、プライマリカレンダーを使用します。");
    calendar = CalendarApp.getCalendarById('primary');
  }

  // --- 4. 希望休カレンダーの取得 ---
  const reqOffCalendarId = (props['Req_Off_CalendarID'] || "").trim();
  let reqOffCalendar = null;
  if (reqOffCalendarId) {
    try {
      reqOffCalendar = CalendarApp.getCalendarById(reqOffCalendarId);
      if (!reqOffCalendar) {
        console.warn("【警告】Req_Off_CalendarID のカレンダーが見つかりません（ID: " + reqOffCalendarId + "）。希望休登録をスキップします。");
      } else {
        console.log("カレンダー接続完了 \n接続カレンダー名：【" + reqOffCalendar.getName() + "】 \nカレンダーID:" + reqOffCalendarId);
      }
    } catch (e) {
      console.warn("【警告】Req_Off_CalendarID の取得に失敗しました: " + e.message);
    }
  } else {
    console.log("【情報】Req_Off_CalendarID が未設定のため、希望休登録をスキップします。");
  }

  // --- 5. Gemini接続確認 ---
  console.log("【Gemini接続確認開始】モデル: gemini-2.5-flash");
  let isGeminiAvailable = geminiApiKey ? testGeminiConnection(geminiApiKey) : false;

  // --- 6. ソースファイル取得 ---
  // 共有ドライブ上のファイルにも対応するため Drive.Files.get() で取得する
  let sourceFileBlob;
  try {
    sourceFileBlob = DriveApp.getFileById(sourceFileId).getBlob();
  } catch (e1) {
    // DriveApp で失敗した場合は Drive API（共有ドライブ対応）でリトライ
    try {
      console.warn("DriveApp での取得失敗。Drive API で再試行します: " + e1.message);
      const fileMeta = Drive.Files.get(sourceFileId, { supportsAllDrives: true, fields: "id,name,mimeType" });
      const fileRes = UrlFetchApp.fetch(
        "https://www.googleapis.com/drive/v3/files/" + sourceFileId + "?alt=media&supportsAllDrives=true",
        { headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }
      );
      if (fileRes.getResponseCode() !== 200) {
        console.error("ソースファイルの取得に失敗しました（HTTP " + fileRes.getResponseCode() + "）");
        return;
      }
      sourceFileBlob = fileRes.getBlob().setName(fileMeta.name);
      console.log("【Drive API で取得成功】" + fileMeta.name);
    } catch (e2) {
      console.error("ソースファイルの取得に失敗しました: " + e2.message);
      return;
    }
  }

  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  // ProtectUntil: スクリプトプロパティの日数（未設定時は0日＝今日から即変更対象）
  const protectDays = parseInt(props['Protect_Until'] || "0");
  const protectUntil = new Date(today.getTime() + (protectDays * 24 * 60 * 60 * 1000));
  // Sync_Until: スクリプトプロパティの週数（未設定時は12週）
  const syncWeeks = parseInt(props['Sync_Until'] || "12");
  const syncUntil = new Date(today.getTime() + (syncWeeks * 7 * 24 * 60 * 60 * 1000));
  console.log("【保護期間】" + protectDays + "日間（" + Utilities.formatDate(protectUntil, "JST", "yyyy/MM/dd") + "以降を変更対象）");
  console.log("【同期範囲】" + syncWeeks + "週間（" + Utilities.formatDate(syncUntil, "JST", "yyyy/MM/dd") + "まで）");

  for (let m = 0; m <= 4; m++) {
    const targetDate = new Date(now.getFullYear(), now.getMonth() + m, 1);
    const yearMonth = Utilities.formatDate(targetDate, "JST", "yyyyMM");
    const logLabel = Utilities.formatDate(targetDate, "JST", "yyyy年MM月");
    
    if (targetDate > syncUntil) break;

    let tempFileId = null;

    // --- 5. Excelから最新データを読み込み、currentSSのシートを常に更新する ---
    let localSheet = currentSS.getSheetByName(yearMonth);
    try {
      // ExcelをDrive APIで一時的にGoogle Sheetsに変換
      const metadata = {
        name: "temp_convert_" + yearMonth,
        mimeType: "application/vnd.google-apps.spreadsheet",
        parents: [workFolderId]
      };
      const tempFile = Drive.Files.create(metadata, sourceFileBlob);
      tempFileId = tempFile.id;
      const tempSS = SpreadsheetApp.openById(tempFileId);
      const tempSheet = tempSS.getSheetByName(yearMonth);

      if (tempSheet) {
        const srcRange = tempSheet.getDataRange();
        if (!localSheet) {
          // シートが存在しない場合：新規作成
          localSheet = currentSS.insertSheet(yearMonth);
          console.log(logLabel + ": currentSSに新規シートを作成しました。");
        } else {
          // シートが既存の場合：既存データをクリアして最新データで上書き
          localSheet.clearContents();
          localSheet.clearFormats();
          console.log(logLabel + ": currentSSの既存シートをExcelの最新データで更新しました。");
        }
        // 値・書式をcurrentSSのシートに反映
        const destRange = localSheet.getRange(1, 1, srcRange.getNumRows(), srcRange.getNumColumns());
        destRange.setValues(srcRange.getValues());
        destRange.setFontColors(srcRange.getFontColors());
        destRange.setBackgrounds(srcRange.getBackgrounds());
      } else {
        console.log(logLabel + ": Excelに該当シートなし。スキップします。");
        Drive.Files.remove(tempFileId);
        tempFileId = null;
        continue;
      }
    } catch (e) {
      console.error(logLabel + "のExcel変換・コピー失敗: " + e.message);
      if (tempFileId) { Drive.Files.remove(tempFileId); tempFileId = null; }
      continue;
    }

    const endOfMonth = new Date(targetDate.getFullYear(), targetDate.getMonth() + 1, 0, 23, 59, 59);
    const deleteStart = (protectUntil > targetDate) ? protectUntil : targetDate;
    const deleteEnd = (syncUntil < endOfMonth) ? syncUntil : endOfMonth;
    
    if (deleteStart < deleteEnd) {
      // メインカレンダーの既存同期イベントを削除
      calendar.getEvents(deleteStart, deleteEnd).forEach(event => {
        if (event.getDescription().indexOf(syncTag) !== -1) event.deleteEvent();
      });

      // 希望休カレンダーの既存同期イベントを削除
      if (reqOffCalendar) {
        reqOffCalendar.getEvents(deleteStart, deleteEnd).forEach(event => {
          if (event.getDescription().indexOf(syncTag) !== -1) event.deleteEvent();
        });
      }

      // --- すべてのデータ・色情報をcurrentSSのシートから取得 ---
      const lastCol = localSheet.getLastColumn();
      const allData = localSheet.getDataRange().getValues();
      const datesRow = allData[0];
      const myValues = allData[syncLineNo - 1];
      const fontColors = localSheet.getRange(syncLineNo, 1, 1, lastCol).getFontColors()[0];
      const bgColors = localSheet.getRange(syncLineNo, 1, 1, lastCol).getBackgrounds()[0];
      const colorReqOff = "#ffff00"; // 黄色

      let createCount = 0;
      let reqOffCount = 0;

      for (let i = 0; i < datesRow.length; i++) {
        const day = parseInt(datesRow[i]);
        if (isNaN(day)) continue;

        const eventDate = new Date(targetDate.getFullYear(), targetDate.getMonth(), day);
        if (eventDate < protectUntil || eventDate > syncUntil) continue;

        const content = myValues[i] ? String(myValues[i]).trim() : "";
        const bgColor = bgColors[i] ? bgColors[i].toLowerCase() : "";

        // --- 希望休の登録（背景色が黄色のとき）---
        if (bgColor === colorReqOff) {
          if (reqOffCalendar) {
            reqOffCalendar.createAllDayEvent("希望休", eventDate, {
              description: syncTag + "\n同期日時：" + syncTimestamp
            });
            reqOffCount++;
            console.log("【希望休登録】" + Utilities.formatDate(eventDate, "JST", "yyyy/MM/dd"));
          } else {
            console.warn("【希望休スキップ】カレンダー未取得のため登録できません: " + Utilities.formatDate(eventDate, "JST", "yyyy/MM/dd"));
          }
        }

        // --- 通常の稼働予定登録（セルに内容があるとき）---
        if (!content) continue;

        let color = fontColors[i] ? fontColors[i].toLowerCase() : "";
        
        // --- 接頭辞（プレフィックス）の判定ロジック ---
        let prefix = "";
        if (color === colorMain) {
          prefix = "【メイン】設置@";
        } else if (color === colorAttend) {
          prefix = "稼働@";
        } else {
          prefix = "設置@";
        }
        
        let result;
        if (addressCache[content]) {
          result = addressCache[content];
          // キャッシュ内にすでに【メイン】判定がある場合は、今回の色が違ってもプレフィックスを上書き
          if (result.isMain) prefix = "【メイン】設置@";
        } else {
          result = getTripleSearchAddress(content, geminiApiKey, isGeminiAvailable);

          // 【メイン】属性をキャッシュに保持
          result.isMain = (color === colorMain);
          
          addressCache[content] = result;
          if (result.isGeminiUsed) {
            geminiCount++;
            Utilities.sleep(2000); 
          }
        }
        
        calendar.createAllDayEvent(prefix + content, eventDate, {
          description: syncTag + "\n" + result.log + "\n同期日時：" + syncTimestamp,
          location: result.address
        });
        createCount++;
      }

      if (createCount > 0) console.log(logLabel + ": " + createCount + " 件登録。");
      if (reqOffCount > 0) console.log(logLabel + ": 希望休 " + reqOffCount + " 件登録。");
    }
    if (tempFileId) Drive.Files.remove(tempFileId);
  }
  console.log("【全工程終了】総Gemini使用回数: " + geminiCount + "回");
}

/**
 * カレンダーURL/IDからカレンダーIDを抽出する
 * Google カレンダーの embed/ical URL に含まれる src= パラメータに対応
 * 例: https://calendar.google.com/calendar/embed?src=xxx%40group.calendar.google.com&ctz=...
 */
function extractCalendarId(input) {
  if (!input || typeof input !== 'string') return input;
  // src= パラメータからIDを取得（%40 → @ にデコード）
  const srcMatch = input.match(/[?&]src=([^&]+)/);
  if (srcMatch) return decodeURIComponent(srcMatch[1]);
  // すでにメールアドレス形式のIDならそのまま返す
  if (input.indexOf('@') !== -1) return input;
  // それ以外は25文字以上の連続文字列をIDとみなす
  const idMatch = input.match(/[-\w]{25,}/);
  return idMatch ? idMatch[0] : input;
}

/**
 * 接続テスト（グローバルエンドポイント & gemini-2.5-flash）
 */
function testGeminiConnection(apiKey) {
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent';
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'POST', contentType: 'application/json', headers: { 'x-goog-api-key': apiKey },
      payload: JSON.stringify({ contents: [{ parts: [{ text: 'connection test' }] }] }), muteHttpExceptions: true
    });
    if (response.getResponseCode() === 200) {
      console.log("【接続成功】Gemini 2.5 Flash 通信確認完了。");
      return true;
    }
    console.error(`【接続失敗】応答コード: ${response.getResponseCode()}`);
  } catch (e) {
    console.error("【接続エラー】" + e.message);
  }
  return false;
}

/**
 * 住所検索
 * 検索順序:
 *   1. Googleマップ単体検索 → ROOFTOP/RANGE_INTERPOLATED で一意確定なら即返却
 *   2. 施設名＋住所で精密検索（findPrecisePlace）→ 一意確定なら返却
 *   3. 上記で一意にならない場合のみ Gemini を使用
 */
function getTripleSearchAddress(placeName, apiKey, isGeminiAvailable) {
  const query = placeName.replace(/[0-9０-９]+台.*$/, "").replace(/[（(].*[）)]/g, "").trim();

  // --- Step1: Googleマップ単体検索 ---
  let mapAddress = "";
  let mapLocationType = "";
  try {
    const results = Maps.newGeocoder().setLanguage('ja').setRegion('jp').geocode(query);
    if (results.status === "OK") {
      const bestMatch = results.results[0];
      mapAddress = bestMatch.formatted_address.replace(/^日本、/, "");
      mapLocationType = bestMatch.geometry.location_type;
    }
  } catch (e) {}

  const isDetailed = (t) => t === "ROOFTOP" || t === "RANGE_INTERPOLATED";

  if (isDetailed(mapLocationType)) {
    return { address: mapAddress, log: "Googleマップで住所設定（精度: " + mapLocationType + "）", isGeminiUsed: false };
  }

  // --- Step2: 施設名＋住所で精密検索 ---
  if (mapAddress) {
    const cleanName = placeName.replace(/[0-9０-９]+台.*$/, "").replace(/[（(].*[）)]/g, "").trim();
    try {
      const results2 = Maps.newGeocoder().setLanguage('ja').setRegion('jp').geocode(cleanName + " " + mapAddress);
      if (results2.status === "OK") {
        const best2 = results2.results[0];
        const type2 = best2.geometry.location_type;
        if (isDetailed(type2)) {
          const preciseAddress = best2.formatted_address.replace(/^日本、/, "");
          return { address: preciseAddress, log: "精密場所検索で住所設定（精度: " + type2 + "）", isGeminiUsed: false };
        }
      }
    } catch (e) {
      console.warn("精密場所検索失敗: " + e.message);
    }
  }

  // --- Step3: Googleマップで一意に特定できなかった場合のみ Gemini を使用 ---
  console.log("[Gemini使用] \"" + placeName + "\" はマップで一意特定できず（精度: " + (mapLocationType || "取得不可") + "）");
  if (isGeminiAvailable) {
    const geminiRes = callGeminiWithRetry(placeName, apiKey);
    if (geminiRes.success) return { address: geminiRes.address, log: "Geminiで住所設定", isGeminiUsed: true };
    if (mapAddress) return { address: mapAddress, log: "Gemini不明のためGoogleマップ情報を引き継ぎ", isGeminiUsed: true };
  }

  return { address: mapAddress, log: mapAddress ? "Googleマップ情報を引き継ぎ" : "住所特定不可", isGeminiUsed: false };
}

/**
 * Gemini API リトライロジック（30秒・60秒・90秒の固定3回）
 */
function callGeminiWithRetry(placeName, apiKey) {
  const retryWaits = [30000, 60000, 90000]; // 30秒・60秒・90秒

  // 1回目の試行
  let result = getGeminiAddress(placeName, apiKey);
  if (result.success || result.address === "不明") return result;

  // リトライ（最大3回）
  for (let i = 0; i < retryWaits.length; i++) {
    if (!result.isRetryable) break;
    const waitSec = retryWaits[i] / 1000;
    console.warn(`[Retry ${i + 1}] 429エラー。${waitSec}秒待機後に再試行します...`);
    Utilities.sleep(retryWaits[i]);
    result = getGeminiAddress(placeName, apiKey);
    if (result.success || result.address === "不明") return result;
  }

  return { address: "", success: false };
}

/**
 * Gemini API 実体（gemini-2.5-flash）
 */
function getGeminiAddress(placeName, apiKey) {
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent';
  const prompt = `${placeName} の住所のみ出力し、不明なら「不明」と出力して`;
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'POST', contentType: 'application/json', headers: { 'x-goog-api-key': apiKey },
      payload: JSON.stringify({ 
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.1 }
      }), 
      muteHttpExceptions: true
    });
    
    const code = response.getResponseCode();
    if (code === 200) {
      const resultText = JSON.parse(response.getContentText()).candidates[0].content.parts[0].text.trim();
      console.log(`[Gemini Response] ${placeName} -> ${resultText}`);
      return { address: resultText === "不明" ? "" : resultText, success: resultText !== "不明", isRetryable: false };
    } else if (code === 429) {
      return { address: "", success: false, isRetryable: true };
    }
  } catch (e) {
    console.error("Gemini例外: " + e.message);
  }
  return { address: "", success: false, isRetryable: false };
}

/**
 * 診断用：希望休が登録されない原因を特定する
 * GASエディタからこの関数を単独実行し、ログを確認してください
 */
function debugReqOff() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties().getProperties();

  // ---- [診断1] カレンダーID確認 ----
  const calId = (props['Req_Off_CalendarID'] || "").trim();
  console.log("━━━ [診断1] スクリプトプロパティ ━━━");
  console.log("Req_Off_CalendarID = \"" + calId + "\"");
  if (!calId) {
    console.error("→ ❌ Req_Off_CalendarID が未設定です。スクリプトプロパティに登録してください。");
    return;
  }

  // ---- [診断2] カレンダー接続確認 ----
  console.log("━━━ [診断2] カレンダー接続 ━━━");
  let cal = null;
  try {
    cal = CalendarApp.getCalendarById(calId);
    if (cal) { console.log("【" + cal.getName() + " 接続完了】"); console.log("【" + calId + "】"); }
    else     console.error("→ ❌ getCalendarById が null を返しました。IDが正しいか、このアカウントからアクセス可能か確認してください。");
  } catch (e) {
    console.error("→ ❌ 例外発生: " + e.message);
  }

  // ---- [診断3] syncLineNo 確認 ----
  console.log("━━━ [診断3] スクリプトプロパティ（SyncLine_No） ━━━");
  const syncLineNo = parseInt(props['SyncLine_No'] || "1");
  console.log("SyncLine_No = " + syncLineNo);
  if (isNaN(syncLineNo)) { console.error("→ ❌ SyncLine_No が数値ではありません"); return; }

  // ---- [診断4] 背景色スキャン（本番と同じロジックで確認）----
  // 本番では「currentSSに同名シートがあればそれを使い、なければExcelを変換してコピー」する。
  // 診断でも同じ順序で確認する。
  console.log("━━━ [診断4] 背景色スキャン ━━━");
  const extractId = (input) => {
    if (!input || typeof input !== 'string') return input;
    const m = input.match(/[-\w]{25,}/);
    return m ? m[0] : input;
  };
  const sourceFileId = extractId(props['Inport_URL'] || "");
  const workFolderId = extractId(props['Work_Dir']   || "");

  const now = new Date();
  // 当月と翌月をチェック
  for (let m = 0; m <= 1; m++) {
    const d = new Date(now.getFullYear(), now.getMonth() + m, 1);
    const yearMonth = Utilities.formatDate(d, "JST", "yyyyMM");

    // まずcurrentSSの既存シートを確認
    let sheet = ss.getSheetByName(yearMonth);
    if (sheet) {
      console.log("シート " + yearMonth + " → currentSSに既存シートあり");
    } else {
      // なければExcelをDrive APIで変換して確認
      console.log("シート " + yearMonth + " → currentSSになし。Excelから変換して確認します...");
      let tempFileId = null;
      try {
        // 共有ドライブ対応: DriveApp 失敗時は Drive API でフォールバック
        let debugBlob;
        try {
          debugBlob = DriveApp.getFileById(sourceFileId).getBlob();
        } catch (e1) {
          const fileRes = UrlFetchApp.fetch(
            "https://www.googleapis.com/drive/v3/files/" + sourceFileId + "?alt=media&supportsAllDrives=true",
            { headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }
          );
          if (fileRes.getResponseCode() !== 200) throw new Error("ファイル取得失敗 HTTP " + fileRes.getResponseCode());
          const fileMeta = Drive.Files.get(sourceFileId, { supportsAllDrives: true, fields: "name" });
          debugBlob = fileRes.getBlob().setName(fileMeta.name);
        }
        const metadata = {
          name: "debug_temp_" + yearMonth,
          mimeType: "application/vnd.google-apps.spreadsheet",
          parents: [workFolderId]
        };
        const tempFile = Drive.Files.create(metadata, debugBlob);
        tempFileId = tempFile.id;
        const tempSS = SpreadsheetApp.openById(tempFileId);
        sheet = tempSS.getSheetByName(yearMonth);
        if (!sheet) {
          console.log("  → Excelにも " + yearMonth + " シートなし。スキップします。");
          Drive.Files.remove(tempFileId);
          continue;
        }
        console.log("  → ✅ Excelから変換成功");
      } catch (e) {
        console.error("  → ❌ Excel変換失敗: " + e.message);
        if (tempFileId) try { Drive.Files.remove(tempFileId); } catch(_) {}
        continue;
      }
    }

    const lastCol = sheet.getLastColumn();
    const bgs = sheet.getRange(syncLineNo, 1, 1, lastCol).getBackgrounds()[0];
    const yellowCells = bgs.map((c, i) => c.toLowerCase() === "#ffff00" ? (i+1)+"列目("+c+")" : null).filter(Boolean);
    const nonDefault = bgs.filter(c => c.toLowerCase() !== "#ffffff" && c !== "").slice(0, 10);

    console.log("  黄色セル数: " + yellowCells.length);
    if (yellowCells.length > 0) {
      console.log("  → ✅ 黄色セル: " + yellowCells.join(", "));
    } else {
      console.log("  → ❌ 黄色セルなし。色付きセル（先頭10件）: " + (nonDefault.length ? nonDefault.join(", ") : "なし（全セルが白または無色）"));
    }
  }

  // ---- [診断5] 保護期間の確認 ----
  console.log("━━━ [診断5] 保護期間 ━━━");
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const protectDays = parseInt(props['Protect_Until'] || "0");
  const protectUntil = new Date(today.getTime() + (protectDays * 24 * 60 * 60 * 1000));
  console.log("今日: " + Utilities.formatDate(today, "JST", "yyyy/MM/dd"));
  console.log("Protect_Until = " + protectDays + "日 → 変更対象開始日: " + Utilities.formatDate(protectUntil, "JST", "yyyy/MM/dd"));
  console.log("→ この日付以降のセルが登録・更新されます");
  // ---- [診断6] Work_CalendarID 確認 ----
  console.log("━━━ [診断6] Work_CalendarID ━━━");
  const workCalId = (props['Work_CalendarID'] || "").trim();
  console.log("Work_CalendarID = \"" + workCalId + "\"");
  if (!workCalId) {
    console.warn("→ ⚠️ 未設定のため、プライマリカレンダーを使用します。");
  } else {
    const wCal = CalendarApp.getCalendarById(workCalId);
    if (wCal) { console.log("【" + wCal.getName() + " 接続完了】"); console.log("【" + workCalId + "】"); }
    else      console.error("→ ❌ カレンダーが見つかりません。IDを確認してください。");
  }
}