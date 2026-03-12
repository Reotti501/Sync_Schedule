function updateAndSyncSchedule() {
  const currentSS = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = currentSS.getSheetByName("Config");
  
  if (!configSheet) {
    console.error("エラー: 'Config' シートが見つかりません。");
    return;
  }

  // --- 1. スクリプトプロパティから各種設定を取得 ---
  const props = PropertiesService.getScriptProperties().getProperties();
  const geminiApiKey = props['GEMINI_API_KEY'];
  const workDirId = props['Work_Dir'];
  
  // カラー設定を取得（小文字に統一して比較準備）
  const colorAttend = (props['Color_Attend'] || "#ff0000").toLowerCase();
  const colorMain = (props['Color_Main'] || "#00ffff").toLowerCase();

  // --- IDまたはURLからIDを抽出するヘルパー ---
  const extractId = (input) => {
    if (!input || typeof input !== 'string') return input;
    const match = input.match(/[-\w]{25,}/);
    return match ? match[0] : input;
  };

  // --- 希望休カレンダーの取得 ---
  const reqOffCalendarId = (props['Req_Off_CalendarID'] || "").trim();
  console.log("【Req_Off_CalendarID】" + reqOffCalendarId);
  let reqOffCalendar = null;
  if (reqOffCalendarId) {
    try {
      reqOffCalendar = CalendarApp.getCalendarById(reqOffCalendarId);
      if (!reqOffCalendar) {
        console.warn("【警告】Req_Off_URL のカレンダーが見つかりません（ID: " + reqOffCalendarId + "）。希望休登録をスキップします。");
      } else {
        console.log("【希望休カレンダー接続成功】" + reqOffCalendar.getName());
      }
    } catch (e) {
      console.warn("【警告】Req_Off_URL の取得に失敗しました: " + e.message);
    }
  } else {
    console.log("【情報】Req_Off_URL が未設定のため、希望休登録をスキップします。");
  }

  // --- 2. Configシートから設定を読み込む ---
  const configValues = configSheet.getDataRange().getValues();
  let conf = {};
  configValues.forEach(row => { conf[row[0]] = row[1]; });

  const sourceFileId = extractId(conf["Inport_URL"]);
  const targetSSId = extractId(conf["Export_URL"]);
  const syncLineNo = parseInt(conf["SyncLine_No."]);
  const workFolderId = extractId(workDirId || conf["WorkFolder_URL"]);

  const syncTag = "-- Sync_Scheduleによって自動登録されました --";
  const now = new Date();
  const syncTimestamp = Utilities.formatDate(now, "JST", "yyyy/MM/dd HH:mm");
  
  // キャッシュと集計
  let addressCache = {};
  let geminiCount = 0;

  // --- 3. Gemini接続確認テスト ---
  console.log("【Gemini接続確認開始】モデル: gemini-2.5-flash");
  let isGeminiAvailable = geminiApiKey ? testGeminiConnection(geminiApiKey) : false;

  // --- 4. 実行準備 ---
  let sourceFile;
  try {
    sourceFile = DriveApp.getFileById(sourceFileId);
  } catch (e) {
    console.error("ソースファイルの取得に失敗しました: " + e.message);
    return;
  }

  const calendar = CalendarApp.getCalendarById('primary');
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  const protectUntil = new Date(today.getTime() + (7 * 24 * 60 * 60 * 1000));
  const syncUntil = new Date(today.getTime() + (12 * 7 * 24 * 60 * 60 * 1000));

  for (let m = 0; m <= 4; m++) {
    const targetDate = new Date(now.getFullYear(), now.getMonth() + m, 1);
    const yearMonth = Utilities.formatDate(targetDate, "JST", "yyyyMM");
    const logLabel = Utilities.formatDate(targetDate, "JST", "yyyy年MM月");
    
    if (targetDate > syncUntil) break;

    let tempFileId = null;

    // --- 5. currentSSに同名シートがなければExcelから値をコピーして作成 ---
    let localSheet = currentSS.getSheetByName(yearMonth);
    if (!localSheet) {
      try {
        // ExcelをDrive APIで一時的にGoogle Sheetsに変換
        const metadata = {
          name: "temp_convert_" + yearMonth,
          mimeType: "application/vnd.google-apps.spreadsheet",
          parents: [workFolderId]
        };
        const tempFile = Drive.Files.create(metadata, sourceFile.getBlob());
        tempFileId = tempFile.id;
        const tempSS = SpreadsheetApp.openById(tempFileId);
        const tempSheet = tempSS.getSheetByName(yearMonth);

        if (tempSheet) {
          // 値・書式をcurrentSSに新規シートとしてコピー
          const newSheet = currentSS.insertSheet(yearMonth);
          const srcRange = tempSheet.getDataRange();
          const destRange = newSheet.getRange(1, 1, srcRange.getNumRows(), srcRange.getNumColumns());
          destRange.setValues(srcRange.getValues());
          destRange.setFontColors(srcRange.getFontColors());
          destRange.setBackgrounds(srcRange.getBackgrounds());
          localSheet = newSheet;
          console.log(logLabel + ": currentSSに新規シートを作成しました。");
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
    } else {
      console.log(logLabel + ": currentSSに既存シートあり。そのまま使用します。");
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
          
          // 精密場所検索
          if (result.address) {
            const placeDetail = findPrecisePlace(content, result.address);
            if (placeDetail) {
              result.address = placeDetail;
              result.log += " -> 精密場所検索を適用";
            }
          }
          
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
 * 企業名と住所から正確な場所情報を取得する
 */
function findPrecisePlace(name, address) {
  const cleanName = name.replace(/[0-9０-９]+台.*$/, "").replace(/[（(].*[）)]/g, "").trim();
  const searchQuery = `${cleanName} ${address}`;
  
  try {
    const results = Maps.newGeocoder().setLanguage('ja').setRegion('jp').geocode(searchQuery);
    if (results.status === "OK") {
      const bestMatch = results.results[0];
      return bestMatch.formatted_address.replace(/^日本、/, "");
    }
  } catch (e) {
    console.warn(`プレイス検索失敗: ${searchQuery} - ${e.message}`);
  }
  return null;
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
 * 住所検索（Maps精度判定付）
 */
function getTripleSearchAddress(placeName, apiKey, isGeminiAvailable) {
  let query = placeName.replace(/[0-9０-９]+台.*$/, "").replace(/[（(].*[）)]/g, "").trim();
  let mapAddress = "";
  let isMapDetailed = false;

  try {
    const results = Maps.newGeocoder().setLanguage('ja').setRegion('jp').geocode(query);
    if (results.status === "OK") {
      let bestMatch = results.results[0];
      mapAddress = bestMatch.formatted_address.replace(/^日本、/, "");
      let type = bestMatch.geometry.location_type;
      if (type === "ROOFTOP" || type === "RANGE_INTERPOLATED") isMapDetailed = true;
    }
  } catch (e) {}

  if (isMapDetailed) return { address: mapAddress, log: "Googleマップで住所設定", isGeminiUsed: false };

  if (isGeminiAvailable) {
    let geminiRes = callGeminiWithRetry(placeName, apiKey);
    if (geminiRes.success) return { address: geminiRes.address, log: "Geminiで住所設定", isGeminiUsed: true };
    if (mapAddress) return { address: mapAddress, log: "Gemini不明のためGoogleマップ情報を引き継ぎ", isGeminiUsed: true };
  }

  return { address: mapAddress, log: mapAddress ? "Googleマップ情報を引き継ぎ" : "住所特定不可", isGeminiUsed: isGeminiAvailable };
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
    if (cal) console.log("→ ✅ 接続成功: " + cal.getName());
    else     console.error("→ ❌ getCalendarById が null を返しました。IDが正しいか、このアカウントからアクセス可能か確認してください。");
  } catch (e) {
    console.error("→ ❌ 例外発生: " + e.message);
  }

  // ---- [診断3] Config・syncLineNo 確認 ----
  console.log("━━━ [診断3] Config シート ━━━");
  const configSheet = ss.getSheetByName("Config");
  if (!configSheet) { console.error("→ ❌ Config シートが見つかりません"); return; }
  let conf = {};
  configSheet.getDataRange().getValues().forEach(row => { conf[row[0]] = row[1]; });
  const syncLineNo = parseInt(conf["SyncLine_No."]);
  console.log("SyncLine_No. = " + syncLineNo);
  if (isNaN(syncLineNo)) { console.error("→ ❌ SyncLine_No. が数値ではありません"); return; }

  // ---- [診断4] 背景色スキャン（本番と同じロジックで確認）----
  // 本番では「currentSSに同名シートがあればそれを使い、なければExcelを変換してコピー」する。
  // 診断でも同じ順序で確認する。
  console.log("━━━ [診断4] 背景色スキャン ━━━");
  const extractId = (input) => {
    if (!input || typeof input !== 'string') return input;
    const m = input.match(/[-\w]{25,}/);
    return m ? m[0] : input;
  };
  const sourceFileId = extractId(conf["Inport_URL"]);
  const workFolderId = extractId(props['Work_Dir'] || conf["WorkFolder_URL"]);

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
        const sourceFile = DriveApp.getFileById(sourceFileId);
        const metadata = {
          name: "debug_temp_" + yearMonth,
          mimeType: "application/vnd.google-apps.spreadsheet",
          parents: [workFolderId]
        };
        const tempFile = Drive.Files.create(metadata, sourceFile.getBlob());
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
  const protectUntil = new Date(today.getTime() + (7 * 24 * 60 * 60 * 1000));
  console.log("今日: " + Utilities.formatDate(today, "JST", "yyyy/MM/dd"));
  console.log("登録開始日（7日後）: " + Utilities.formatDate(protectUntil, "JST", "yyyy/MM/dd"));
  console.log("→ この日付より前のセルは保護期間のためスキップされます");
}