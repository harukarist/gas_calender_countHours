// シートを取得
const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
const sheetSchedules = spreadSheet.getSheetByName('予定');
const sheetSettings = spreadSheet.getSheetByName('設定');
const sheetPJCodes = spreadSheet.getSheetByName('PJコード一覧');

// 設定値を取得
const calendarIds = sheetSettings.getRange('B1').getValue().split(','); // カレンダーID配列
const reExclusion = new RegExp(sheetSettings.getRange('B5').getValue()); // 除外ワード
const allDayExclusion = sheetSettings.getRange('B6').getValue(); // 終日の予定除外有無
const startDate = new Date(sheetSettings.getRange('B2').getValue()); // 取得開始日
const endDate = new Date(startDate);
endDate.setMonth(endDate.getMonth() + 1); // 取得終了日
const breakTimeThreshold = sheetSettings.getRange('B7').getValue(); // 休憩時間付与のしきい値
const pjNamesRange = sheetPJCodes.getRange(2,1,sheetPJCodes.getLastRow(),1); //PJ名一覧


function getCalendarEvents() {
  // 最終行・最終列を取得
  let lastRow = sheetSchedules.getLastRow();
  // let lastColumn = sheetSchedules.getLastColumn();
  const lastColumn = 8;
  const pjNameColumn = 8;

  // シートの予定データの値をクリア
  if(lastRow){
    sheetSchedules.getRange(2, 1, lastRow, lastColumn).clearContent();
  }

  // PJ名リスト（プルダウンメニュー）の入力規則を設定
  const listRule = SpreadsheetApp.newDataValidation()
  .requireValueInRange(pjNamesRange)
  .build(); //PJ名リストの入力規則
  sheetSchedules.getRange(2, pjNameColumn, lastRow, 1).setDataValidation(listRule);

  // 条件付き書式を設定
  let range = sheetSchedules.getRange(2, pjNameColumn, lastRow, 2);
  const emptyRule = SpreadsheetApp.newConditionalFormatRule()
  .whenCellEmpty() //セルが空白の場合
  .setBackground("#B6E1CD")
  .setRanges([range])
  .build();
  const rules = sheetSchedules.getConditionalFormatRules();
  rules.push(emptyRule);
  sheetSchedules.setConditionalFormatRules(rules);


  // 配列初期化
  let table = new Array();

  // カレンダー数分ループ処理
  for (let i = 0; i < calendarIds.length; i++) {
    // 取得結果の配列を追記
    table = table.concat(
      // fetchSchedulesを呼び出して予定一覧を取得
      fetchSchedules(calendarIds[i],reExclusion,allDayExclusion,startDate,endDate)
    );
  }

  if (table.length) {
    // シートに出力
    sheetSchedules.getRange(2, 1, table.length, table[0].length).setValues(table);
    Logger.log(`${table.length}件の予定を取得しました。`);
    spreadSheet.toast(`${table.length}件の予定を取得しました。`, 'Googleカレンダー取得完了', 5); // 完了メッセージ表示
  } else {
    Logger.log(`${table.length}件の予定を取得しました。`);
    spreadSheet.toast('取得結果が0件です。', 'Googleカレンダー取得完了', 5); // エラーメッセージ表示
  }

}

/**
 * シートを開いた際にGAS実行用メニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi(); // UIクラス取得
  const menu = ui.createMenu('GASメニュー'); // スプレッドシートにメニューを追加
  menu.addItem('Googleカレンダー取得', 'getCalendarEvents'); // 関数セット
  menu.addToUi(); // スプレッドシートに反映
}

/**
 * GoogleカレンダーからgetEvents
 * @param {Array} calendarIds 取得対象カレンダーIDの配列
 * @return {Array} 取得結果の二次元配列
 */
function fetchSchedules(calendarId) {
  const schedules = new Array(); // 配列初期化
  const calendar = CalendarApp.getCalendarById(calendarId); // カレンダー
  const calendarName = calendar.getName(); // カレンダー名
  const events = calendar.getEvents(startDate, endDate); // 範囲内の予定を取得

  // 各予定のデータを配列に追加
  for (let i = 0; i < events.length; i++) {
    // isExclusionで除外判定し、除外対象の場合はスキップ
    if (isExclusion(events[i], reExclusion, allDayExclusion)) continue;

    // 予定の開始時刻・終了時刻を取得
    let start = events[i].getStartTime();
    let end = events[i].getEndTime();

    // 予定データの配列を作成
    let event = new Array();
    event.push(calendarName); // カレンダー名
    event.push(events[i].getTitle()); // 件名
    event.push(start); // 開始日時
    event.push(end); // 終了日時
    event.push(start.getDate()); // 月
    event.push(getOperatingTime(start, end)); // 時間数
    // event.push(events[i].getDescription()); // 詳細

    // 予定一覧の配列に追加
    schedules.push(event);
  }
  // 予定一覧を返す
  return schedules;
}

/**
 * 取得対象の切り分け
 * @param {CalendarEvent} schedule 個別のCalendarEventクラス
 * @return {boolean} 真偽値
 */
function isExclusion(event) {
  // 終日イベントはスキップ
  if (allDayExclusion && event.isAllDayEvent()) return true;

  // 除外ワードを含む場合はスキップ
  if (reExclusion.test(event.getTitle())) return true;

  return false;
}

/**
 * 経過時間数の計算
 * @param {Date} start 予定の開始日時
 * @param {Date} end 予定の終了日時
 * @return {number} 経過時間数
 */
function getOperatingTime(start, end) {
  // 時間数算出
  const time = (end - start) / 1000 / 60 / 60;

  // 休憩時間の減算
  const operatingTime = time >= breakTimeThreshold ? time - 1 : time;

  return operatingTime;
}


