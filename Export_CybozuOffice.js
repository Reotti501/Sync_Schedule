/**
 * Export_CybozuOffice.gs
 * Sync_Scheduleで同期した1週間分の予定をサイボウズOffice形式のCSVとして
 * Google DriveのWork_Dirフォルダに出力する。
 * 毎週1回トリガーで実行することを想定。
 *
 * CSVフィールド仕様:
 *   イベントID : 空欄
 *   開始日付   : カレンダー登録日付
 *   開始時刻   : 空欄
 *   終了日付   : カレンダー登録日付
 *   終了時刻   : 空欄
 *   予定       : 接頭辞部分（「設置@」「稼働@」「【メイン】設置@」）
 *   予定詳細   : 案件名部分（接頭辞を除いたタイトル）
 *   メモ       : 空欄
 *   参加者     : currentSSのSyncLine_No行目のA列の値
 *   施設       : 空欄
 */
function exportCybozuOfficeCsv() {

  // --- 1. スクリプトプロパティから設定を取得 ---
  const props          = PropertiesService.getScriptProperties().getProperties();
  const workCalendarId = (props['Work_CalendarID'] || "").trim();

  // Work_Dir はURL・IDどちらでも対応（Sync_Scheduleと同じextractIdロジック）
  const extractId = function(input) {
    if (!input || typeof input !== 'string') return input;
    const slashD = input.match(/\/d\/([-\w]{25,})/);
    if (slashD) return slashD[1];
    const queryId = input.match(/[?&]id=([-\w]{25,})/);
    if (queryId) return queryId[1];
    const plain = input.match(/[-\w]{25,}/);
    return plain ? plain[0] : input;
  };
  const workFolderId = extractId(props['Work_Dir'] || "");
  const syncLineNo     = parseInt(props['SyncLine_No'] || "1");
  const protectDays    = parseInt(props['Protect_Until'] || "7");
  const syncTag        = "-- Sync_Scheduleによって自動登録 --";

  // --- 2. 稼働カレンダーの取得 ---
  let calendar;
  if (workCalendarId) {
    calendar = CalendarApp.getCalendarById(workCalendarId);
    if (!calendar) {
      console.error("【エラー】Work_CalendarID のカレンダーが見つかりません（ID: " + workCalendarId + ")。処理を中断します。");
      return;
    }
    console.log("【" + calendar.getName() + " 接続完了】");
    console.log("【" + workCalendarId + "】");
  } else {
    console.warn("【警告】Work_CalendarID が未設定のため、プライマリカレンダーを使用します。");
    calendar = CalendarApp.getCalendarById('primary');
  }

  // --- 3. 参加者：currentSSのSyncLine_No行目のA列から取得 ---
  const currentSS  = SpreadsheetApp.getActiveSpreadsheet();
  const now        = new Date();
  const yearMonth  = Utilities.formatDate(now, "JST", "yyyyMM");
  const localSheet = currentSS.getSheetByName(yearMonth);
  let participant  = "";
  if (localSheet) {
    const val = localSheet.getRange(syncLineNo, 1).getValue();
    participant = val ? String(val).trim() : "";
  } else {
    console.warn("【警告】シート " + yearMonth + " が見つかりません。参加者は空欄になります。");
  }
  console.log("【参加者】" + (participant || "（空欄）"));

  // --- 4. 出力対象期間（今日からProtect_Until日数分）---
  const today      = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const exportUntil = new Date(today.getTime() + (protectDays * 24 * 60 * 60 * 1000));

  console.log("【出力期間】" +
    Utilities.formatDate(today,       "JST", "yyyy/MM/dd") + " 〜 " +
    Utilities.formatDate(exportUntil, "JST", "yyyy/MM/dd") +
    "（" + protectDays + "日間）");

  // --- 5. イベント取得（Sync_Scheduleで登録したもののみ）---
  const events = calendar.getEvents(today, exportUntil).filter(function(event) {
    return event.getDescription().indexOf(syncTag) !== -1;
  });

  console.log("対象イベント数: " + events.length + " 件");

  if (events.length === 0) {
    console.log("出力対象のイベントがありませんでした。");
    return;
  }

  // --- 6. タイトルを「接頭辞」と「案件名」に分割するヘルパー ---
  // 例: "【メイン】設置@ABC商事" -> prefix="設置", caseName="ABC商事"
  //     "稼働@ABC商事"           -> prefix="稼働", caseName="ABC商事"
  //     "設置@ABC商事"           -> prefix="設置", caseName="ABC商事"
  // @と【メイン】は除いて出力する
  const splitTitle = function(title) {
    const match = title.match(/^(?:【メイン】)?(設置|稼働)@([\s\S]*)$/);
    if (match) return { prefix: match[1], caseName: match[2] };
    return { prefix: "", caseName: title };
  };

  // --- 7. CSV行の組み立て ---
  const csvEscape = function(value) {
    const str = value == null ? "" : String(value);
    return '"' + str.replace(/"/g, '""') + '"';
  };

  const header = [
    "イベントID", "開始日付", "開始時刻", "終了日付", "終了時刻",
    "予定", "予定詳細", "メモ", "参加者", "施設"
  ].map(csvEscape).join(",");

  const rows = events.map(function(event) {
    const startDate    = event.getStartTime();
    const startDateStr = Utilities.formatDate(startDate, "JST", "yyyy/MM/dd");
    const parts        = splitTitle(event.getTitle());

    return [
      "",              // イベントID（空欄）
      startDateStr,    // 開始日付
      "",              // 開始時刻（空欄）
      startDateStr,    // 終了日付（終日イベントのため開始日付と同じ）
      "",              // 終了時刻（空欄）
      parts.prefix,    // 予定（接頭辞部分）
      parts.caseName,  // 予定詳細（案件名部分）
      "",              // メモ（空欄）
      participant,     // 参加者
      ""               // 施設（空欄）
    ].map(csvEscape).join(",");
  });

  const csvContent = "\uFEFF" + header + "\n" + rows.join("\n");

  // --- 8. Google DriveのWork_Dirフォルダに保存 ---
  const fileName = "cybozu_export_" +
    Utilities.formatDate(today,       "JST", "yyyyMMdd") + "_" +
    Utilities.formatDate(exportUntil, "JST", "yyyyMMdd") + ".csv";

  try {
    const folder = DriveApp.getFolderById(workFolderId);

    // 同名ファイルが既にあれば上書き（削除して再作成）
    const existing = folder.getFilesByName(fileName);
    while (existing.hasNext()) existing.next().setTrashed(true);

    const file = folder.createFile(fileName, csvContent, MimeType.PLAIN_TEXT);
    console.log("【CSV出力完了】" + fileName);
    console.log("ファイルURL: " + file.getUrl());
  } catch (e) {
    console.error("【エラー】CSVの保存に失敗しました: " + e.message);
  }
}