// ==================================================================
// fiscalYear.gs ― 年度判定・年度別スプレッドシート解決・製造間接費シート自動作成
// ==================================================================
// 【概要】
//   売上予定日の年度に応じて、書き込み先を「年度ごとの別スプレッドシート（ファイル）」
//   に振り分ける。年度ファイルは手動作成し、IDをスクリプトプロパティに登録しておく。
//     例）令和7年度 → スクリプトプロパティ「DEST_SS_ID_R7」にファイルIDを登録
//         令和8年度 → 「DEST_SS_ID_R8」
//   各ファイル内の書き込み先シートは共通の固定名（CONFIG_PROJECT.DEST_SHEET_SALES）。
//
//   ※ 売上進捗表シートの自動作成は廃止（ファイルは手動作成のため）。
//      製造間接費入力シートは、新年度の売上予定日が入った時に自動作成する（従来通り）。
//
// 【年度定義】日本の和暦年度(7月始まり)。2026/7/1〜2027/6/30 → 令和8年度 → R8
// 【製造間接費シート命名規則】
//   {西暦}製造間接費入力シート（事務員専用）
//   (西暦 = 年度終了の年。例: R7年度=2025/7〜2026/6 → 「2026」)
//
// ★今回の修正（2026/06）
//   1. advanceMfgSheetMonths_：売上月度(A列)が「日付セル」でも年を繰り上げるよう対応。
//      （表示書式「yyyy年M月」のせいで文字列に見えるが実体は日付、というケースで
//        年が変わらない不具合への対策）
//   2. clearMfgSheetData_：複製シートに H1「転記済み」見出しも自動セット
//      （複製シートが自己完結し、「転記済み列を準備」を毎回押さなくてよくなる）
// ==================================================================

const FY_CONFIG = {
  // 年度の開始月 (7 = 7月始まり)
  FISCAL_START_MONTH: 7,

  // 和暦変換の基準 (令和元年 = 2019年 → REIWA_BASE_YEAR + N が西暦)
  REIWA_BASE_YEAR: 2018,

  // スクリプトプロパティの年度別IDキーの接頭辞（例: DEST_SS_ID_R7）
  DEST_SS_ID_KEY_PREFIX: "DEST_SS_ID_R",

  // 製造間接費シート
  MFG_SHEET_NAME_TEMPLATE: "{YEAR}製造間接費入力シート（事務員専用）",
  MFG_SHEET_MATCH_PREFIX:  /^(\d{4})製造間接費入力シート/,

  // 製造間接費シートのクリア定義
  MFG_CLEAR_COL_START: 2,  // B列
  MFG_CLEAR_COL_COUNT: 5   // B〜F列(物件No,顧客名,完成売上高,製造原価計,売上総利益)
};

// ==================================================================
// [PUBLIC] 売上予定日 → 書き込み先スプレッドシートID を解決
// ==================================================================
/**
 * 売上予定日(Date)から、対応する年度ファイルのID等を返す。
 * 日付が無効な場合は「今日」の年度で解決する。
 * @param {Date} salesDate
 * @return {{reiwaYear:number, ssId:string, sheetName:string}}
 *         ssId が "" の場合は、その年度のファイルが未登録。
 */
function resolveDestSsId(salesDate) {
  const d = (salesDate instanceof Date && !isNaN(salesDate)) ? salesDate : new Date();
  const reiwaYear = calcFiscalYearReiwa_(d);
  return {
    reiwaYear: reiwaYear,
    ssId: getDestSsIdForFiscalYear_(reiwaYear),
    sheetName: CONFIG_PROJECT.DEST_SHEET_SALES
  };
}

/**
 * 製造間接費シート名から、対応する年度ファイルのID等を返す（overhead.gs 用）。
 * 例: "2027製造間接費入力シート（事務員専用）" → 令和8年度のファイル
 * @param {string} mfgSheetName
 * @return {{reiwaYear:number, ssId:string, sheetName:string}|null}
 */
function resolveDestContextFromMfgSheet(mfgSheetName) {
  const match = String(mfgSheetName || "").match(/^(\d{4})製造間接費/);
  if (!match) return null;

  const calendarYear = parseInt(match[1], 10);
  const reiwaYear = calendarYear - FY_CONFIG.REIWA_BASE_YEAR - 1;  // 西暦は「年度終了の年」
  if (reiwaYear <= 0) return null;

  return {
    reiwaYear: reiwaYear,
    ssId: getDestSsIdForFiscalYear_(reiwaYear),
    sheetName: CONFIG_PROJECT.DEST_SHEET_SALES
  };
}

/** 年度(令和N)に対応する書き込み先ファイルIDをスクリプトプロパティから取得（未登録は ""）。 */
function getDestSsIdForFiscalYear_(reiwaYear) {
  const key = FY_CONFIG.DEST_SS_ID_KEY_PREFIX + reiwaYear;  // 例: DEST_SS_ID_R7
  const id = PropertiesService.getScriptProperties().getProperty(key);
  return id ? String(id).trim() : "";
}

// ==================================================================
// [PRIVATE] 年度算出
// ==================================================================
/** Dateから和暦年度(7月始まり)を算出する。例: 2025/8 → 7（令和7年度） */
function calcFiscalYearReiwa_(date) {
  const year = date.getFullYear();
  const month = date.getMonth() + 1;
  const fiscalCalendarYear = (month >= FY_CONFIG.FISCAL_START_MONTH) ? year : year - 1;
  return fiscalCalendarYear - FY_CONFIG.REIWA_BASE_YEAR;
}

/** 和暦年度 → 製造間接費シートの西暦(年度終了側)。例: R7 → 2026 */
function reiwaToMfgCalendarYear_(reiwaYear) {
  return FY_CONFIG.REIWA_BASE_YEAR + reiwaYear + 1;
}

// ==================================================================
// [PUBLIC] 製造間接費シートの自動作成（新年度の売上予定日が入った時）
// ==================================================================
/**
 * 指定年度の製造間接費入力シートが無ければ、最新の同シートを複製して作成する。
 * 売上進捗表（別ファイル）には一切触れない。
 * @param {number} reiwaYear
 */
function ensureMfgSheetForFiscalYear_(reiwaYear) {
  const calendarYear = reiwaToMfgCalendarYear_(reiwaYear);
  const newMfgSheetName = FY_CONFIG.MFG_SHEET_NAME_TEMPLATE.replace("{YEAR}", calendarYear);
  const currentSs = SpreadsheetApp.getActiveSpreadsheet();

  if (currentSs.getSheetByName(newMfgSheetName)) return;  // 既に存在

  const templateMfgSheet = findLatestMfgSheet_(currentSs);
  if (!templateMfgSheet) {
    Logger.log(`❌ 複製元の製造間接費シートが見つかりません`);
    return;
  }

  Logger.log(`📋 「${templateMfgSheet.getName()}」を複製して「${newMfgSheetName}」を作成`);
  const newMfgSheet = templateMfgSheet.copyTo(currentSs);
  newMfgSheet.setName(newMfgSheetName);
  moveSheetWithoutSwitching_(currentSs, newMfgSheet, templateMfgSheet.getIndex());

  clearMfgSheetData_(newMfgSheet);

  // 売上月度(A列)の年を、複製元との差分ぶん繰り上げる（例: 2025年7月 → 2026年7月）
  const tplMatch = String(templateMfgSheet.getName()).match(/^(\d{4})/);
  const tplCalYear = tplMatch ? parseInt(tplMatch[1], 10) : (calendarYear - 1);
  const yearDelta = calendarYear - tplCalYear;
  advanceMfgSheetMonths_(newMfgSheet, yearDelta);

  Logger.log(`✅ 製造間接費シート「${newMfgSheetName}」の作成完了`);
}

/**
 * 製造間接費シートの売上月度(A列)に含まれる月度の年を、yearDelta ぶん繰り上げる。
 *   ・日付セル（表示書式で "yyyy年M月" に見えるもの）→ 年を +delta（書式は維持）。
 *   ・文字列セル "YYYY年M月" → 年を +delta。
 *   ・数式セル／「合計」など月度でないセル → 触らない。
 * ★日付セルにも対応したのが今回の修正点。
 */
function advanceMfgSheetMonths_(sheet, yearDelta) {
  if (!yearDelta) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const col = MFG_CONFIG.SRC_COL.MONTH;  // A列(1)
  const range = sheet.getRange(2, col, lastRow - 1, 1);
  const values = range.getValues();
  const formulas = range.getFormulas();

  let updated = 0;
  for (let i = 0; i < values.length; i++) {
    if (formulas[i][0]) continue;  // 数式は触らない
    const v = values[i][0];
    const cell = sheet.getRange(i + 2, col);

    if (v instanceof Date && !isNaN(v)) {
      // 日付セル：年を delta 進める（セルの表示書式「yyyy年M月」はそのまま維持される）
      const nd = new Date(v.getTime());
      nd.setFullYear(nd.getFullYear() + yearDelta);
      cell.setValue(nd);
      updated++;
    } else {
      // 文字列セル："YYYY年M月" の年を進める
      const newV = bumpYearInTextBy_(v, yearDelta);
      if (newV !== null) {
        cell.setValue(newV);
        updated++;
      }
    }
  }
  Logger.log(`📅 製造間接費シートの月度を ${yearDelta >= 0 ? "+" : ""}${yearDelta}年 に更新: ${updated}件`);
}

/** 文字列中の「YYYY年M月」を見つけ、年を delta ぶん足して返す。含まなければ null。 */
function bumpYearInTextBy_(value, delta) {
  if (value === null || value === undefined || value === "") return null;
  const s = String(value);
  const m = s.match(/(\d{4})年(\d{1,2})月/);
  if (!m) return null;
  const newYear = parseInt(m[1], 10) + delta;
  return s.replace(/(\d{4})年(\d{1,2})月/, newYear + "年" + m[2] + "月");
}

/** シートを指定インデックスの直前に並べ替える(処理後に元のアクティブシートへ戻す)。 */
function moveSheetWithoutSwitching_(ss, sheet, targetIndex) {
  const originalActive = ss.getActiveSheet();
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(targetIndex);
  if (originalActive) {
    ss.setActiveSheet(originalActive);
  }
}

/** 現在のSS内から、最も新しい製造間接費入力シートを返す。 */
function findLatestMfgSheet_(ss) {
  const pattern = FY_CONFIG.MFG_SHEET_MATCH_PREFIX;
  let latestSheet = null;
  let latestYear = -1;

  ss.getSheets().forEach(function (sheet) {
    const m = sheet.getName().match(pattern);
    if (m) {
      const y = parseInt(m[1], 10);
      if (y > latestYear) {
        latestYear = y;
        latestSheet = sheet;
      }
    }
  });

  return latestSheet;
}

// ==================================================================
// [PRIVATE] データクリア(製造間接費)
// ==================================================================
function clearMfgSheetData_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const numRows = lastRow - 1;
  const startCol = FY_CONFIG.MFG_CLEAR_COL_START;   // B(2)
  const colCount = FY_CONFIG.MFG_CLEAR_COL_COUNT;   // 5列(B〜F)

  const bToFRange = sheet.getRange(2, startCol, numRows, colCount);
  const bToFValues = bToFRange.getValues();
  const bToFFormulas = bToFRange.getFormulas();

  for (let i = 0; i < bToFValues.length; i++) {
    for (let c = 0; c < colCount; c++) {
      if (bToFFormulas[i][c]) continue;  // 数式セルは温存
      const v = bToFValues[i][c];
      if (v !== "" && v !== null && v !== undefined) {
        sheet.getRange(i + 2, startCol + c).clearContent();
      }
    }
  }

  // H列(チェックボックス)を再設定し、コピー時に残るチェックを全てオフにする
  const checkCol = MFG_CONFIG.SRC_COL.CHECK;        // H列(8)
  // ★複製シートが自己完結するよう、H1見出し「転記済み」も自動セット
  sheet.getRange(1, checkCol).setValue("転記済み");
  const checkRange = sheet.getRange(2, checkCol, numRows, 1);
  checkRange.insertCheckboxes();  // チェックボックス化（複製直後でも確実に機能させる）
  checkRange.uncheck();           // 全てオフ（false）にする

  Logger.log(`🧹 製造間接費シートのデータクリア完了: ${numRows}行(数式温存・チェックボックス全オフ・H1見出しセット)`);
}

// ==================================================================
// [TEST] 動作確認用
// ==================================================================
function testFiscalYearLogic() {
  const cases = [
    new Date("2025/06/30"), new Date("2025/07/01"),
    new Date("2026/06/30"), new Date("2026/07/01"),
    new Date("2027/06/30"), new Date("2027/07/01")
  ];

  Logger.log("=== 年度判定・ファイルID解決テスト ===");
  cases.forEach(function (d) {
    const info = resolveDestSsId(d);
    const mfgYear = reiwaToMfgCalendarYear_(info.reiwaYear);
    const mfgName = FY_CONFIG.MFG_SHEET_NAME_TEMPLATE.replace("{YEAR}", mfgYear);
    Logger.log(
      `${Utilities.formatDate(d, "Asia/Tokyo", "yyyy/MM/dd")} → R${info.reiwaYear}年度 ` +
      `→ ファイルID=${info.ssId || "（未登録）"} / シート=${info.sheetName} / 製造間接費=${mfgName}`
    );
  });
}