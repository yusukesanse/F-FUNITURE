// ==================================================================
// ★★★ fiscalYear.gs - 年度自動切替・新年度シート自動作成 ★★★
// ==================================================================
// 【概要】
// 売上予定日が新年度(7/1以降)になった瞬間、転記先スプレッドシートに
// 新年度の「R{N}年度全社売上進捗表（集計）」シートを自動作成する。
// 同時に、現スプレッドシートに新年度の製造間接費入力シートも作成する。
//
// 【新年度シート作成の流れ】
//   1. 既存の最新年度シートを複製
//   2. データ行(A列に連番があり、合計/累計でない行)の B〜K 列を
//      clearContent でクリア(L列=粗利高は数式のため温存)
//   3. データ行D列・合計行B列の月度を「+1年」に書き換え
//      (例: 2025年7月 → 2026年7月。暦年をすべて+1する)
//   ※ 数式・項目見出し行・合計行・累計行・月度ヘッダーは温存
//
// 【年度定義】
//   日本の和暦年度(7月始まり)
//   2026/7/1 〜 2027/6/30 → 令和8年度 → R8
//
// 【シート命名規則】
//   売上進捗表  : R{N}年度全社売上進捗表（集計）
//   製造間接費  : {西暦}製造間接費入力シート（事務員専用）
//                 (西暦 = 年度終了の年。例: R7年度=2025/7〜2026/6 → 「2026」)
//
// 【設定カスタマイズ】
//   FY_CONFIG オブジェクトを編集することで命名規則・基準月を変更可能
// ==================================================================

const FY_CONFIG = {
  // 年度の開始月 (7 = 7月始まり)
  FISCAL_START_MONTH: 7,

  // 売上進捗表シート名のテンプレート ({N}が和暦年度に置換される / 全角カッコ)
  SALES_SHEET_NAME_TEMPLATE: "R{N}年度全社売上進捗表（集計）",

  // 製造間接費シート名のテンプレート ({YEAR}が西暦に置換される / 全角カッコ)
  MFG_SHEET_NAME_TEMPLATE: "{YEAR}製造間接費入力シート（事務員専用）",

  // テンプレート(複製元)を探すための前方一致パターン
  //   カッコの全角/半角に依存しないよう「（集計）」より前で判定する
  SALES_SHEET_MATCH_PREFIX: /^R(\d+)年度全社売上進捗表/,
  MFG_SHEET_MATCH_PREFIX:   /^(\d{4})製造間接費入力シート/,

  // 和暦変換の基準 (令和元年 = 2019年 → REIWA_BASE_YEAR + N が西暦)
  REIWA_BASE_YEAR: 2018,

  // ---- 売上進捗表シートの列定義(1始まり) ----
  // 月度の所在: データ行はD列、合計行はB列
  COL_A_SEQ: 1,          // A列: 連番(データ行の目印)
  COL_B_STORE: 2,        // B列: 店舗 / 合計行では月度
  COL_D_SALES_MONTH: 4,  // D列: 売上年月(データ行の月度)
  CLEAR_COL_START: 2,    // クリア開始列(B列)
  CLEAR_COL_END: 11,     // クリア終了列(K列。L列=粗利高は数式のため除外)

  // 合計行/累計行の判定ラベル
  TOTAL_KEYWORD: "合計",
  RUIKEI_KEYWORD: "累計",

  // ---- 製造間接費シートのクリア定義 ----
  MFG_CLEAR_COL_START: 2,  // B列
  MFG_CLEAR_COL_COUNT: 5   // B〜F列(物件No,顧客名,完成売上高,製造原価計,売上総利益)
};

// ==================================================================
// [PUBLIC] 売上予定日から年度シート名を取得 (存在しなければ作成)
// ==================================================================
/**
 * 売上予定日(Dateオブジェクト)から、対応する年度の売上進捗表シート名を返す。
 * 該当シートが転記先SSに存在しなければ自動作成し、同時に製造間接費シートも
 * 現在のスプレッドシートに自動作成する。
 *
 * @param {Date} salesDate - 売上予定日
 * @return {string} 対応する年度シート名(例: "R8年度全社売上進捗表（集計）")
 */
function resolveSalesSheetName(salesDate) {
  if (!(salesDate instanceof Date) || isNaN(salesDate)) {
    return CONFIG_PROJECT.DEST_SHEET_SALES;
  }

  const fy = calcFiscalYearReiwa_(salesDate);
  const sheetName = FY_CONFIG.SALES_SHEET_NAME_TEMPLATE.replace("{N}", fy);

  const destSs = SpreadsheetApp.openById(CONFIG_PROJECT.DESTINATION_SS_ID);
  if (destSs.getSheetByName(sheetName)) {
    return sheetName;
  }

  Logger.log(`📅 新年度シート「${sheetName}」が存在しないため自動作成します`);
  createNewFiscalYearSheets_(fy, destSs);

  return sheetName;
}

/**
 * 製造間接費シート名から対応する売上進捗表シート名を返す。
 * 例: "2027製造間接費入力シート（事務員専用）" → "R8年度全社売上進捗表（集計）"
 *
 * @param {string} mfgSheetName - 製造間接費シート名
 * @return {string|null} 対応する売上進捗表シート名 (判定不能ならnull)
 */
function resolveSalesSheetNameFromMfgSheet(mfgSheetName) {
  const match = String(mfgSheetName || "").match(/^(\d{4})製造間接費/);
  if (!match) return null;

  const calendarYear = parseInt(match[1], 10);
  // 製造間接費シートの西暦は「年度終了の年」(例: 2026シート → R7年度)
  const fy = calendarYear - FY_CONFIG.REIWA_BASE_YEAR - 1;
  if (fy <= 0) return null;

  return FY_CONFIG.SALES_SHEET_NAME_TEMPLATE.replace("{N}", fy);
}

// ==================================================================
// [PRIVATE] 年度算出
// ==================================================================
/**
 * Dateから和暦年度(7月始まり)を算出する。
 * 例: 2026/6/30 → R7年度(7)、2026/7/1 → R8年度(8)
 */
function calcFiscalYearReiwa_(date) {
  const year = date.getFullYear();
  const month = date.getMonth() + 1;
  const fiscalCalendarYear = (month >= FY_CONFIG.FISCAL_START_MONTH) ? year : year - 1;
  return fiscalCalendarYear - FY_CONFIG.REIWA_BASE_YEAR;
}

/**
 * 和暦年度から、対応する製造間接費シートの西暦を返す。
 * 例: R8年度 → 2027 (年度終了側の年)
 */
function reiwaToMfgCalendarYear_(reiwaYear) {
  return FY_CONFIG.REIWA_BASE_YEAR + reiwaYear + 1;
}

// ==================================================================
// [PRIVATE] 新年度シート作成
// ==================================================================
/**
 * 指定された和暦年度の新年度シートを作成する。
 * 1. 転記先SSに「R{N}年度全社売上進捗表（集計）」を作成(複製→クリア→月度+1年)
 * 2. 現在のSSに「{西暦}製造間接費入力シート（事務員専用）」を作成(複製→クリア)
 */
function createNewFiscalYearSheets_(reiwaYear, destSs) {
  // --- 1. 売上進捗表シートの作成 ---
  const newSalesSheetName = FY_CONFIG.SALES_SHEET_NAME_TEMPLATE.replace("{N}", reiwaYear);

  if (!destSs.getSheetByName(newSalesSheetName)) {
    const templateSalesSheet = findLatestSalesSheet_(destSs);
    if (!templateSalesSheet) {
      Logger.log(`❌ 複製元の売上進捗表シートが見つかりません`);
      return;
    }

    Logger.log(`📋 「${templateSalesSheet.getName()}」を複製して「${newSalesSheetName}」を作成`);
    const newSheet = templateSalesSheet.copyTo(destSs);
    newSheet.setName(newSalesSheetName);
    // シートを複製元の直前に並べ替える(アクティブシートは元に戻して画面遷移を防ぐ)
    moveSheetWithoutSwitching_(destSs, newSheet, templateSalesSheet.getIndex());

    clearSheetData_(newSheet);        // データ行 B〜K をクリア(数式L列は温存)
    advanceSheetMonths_(newSheet);    // 月度を +1年 に書き換え
    Logger.log(`✅ 売上進捗表「${newSalesSheetName}」の作成完了`);
  }

  // --- 2. 製造間接費シートの作成 ---
  const calendarYear = reiwaToMfgCalendarYear_(reiwaYear);
  const newMfgSheetName = FY_CONFIG.MFG_SHEET_NAME_TEMPLATE.replace("{YEAR}", calendarYear);
  const currentSs = SpreadsheetApp.getActiveSpreadsheet();

  if (!currentSs.getSheetByName(newMfgSheetName)) {
    const templateMfgSheet = findLatestMfgSheet_(currentSs);
    if (!templateMfgSheet) {
      Logger.log(`❌ 複製元の製造間接費シートが見つかりません`);
      return;
    }

    Logger.log(`📋 「${templateMfgSheet.getName()}」を複製して「${newMfgSheetName}」を作成`);
    const newMfgSheet = templateMfgSheet.copyTo(currentSs);
    newMfgSheet.setName(newMfgSheetName);
    // シートを複製元の直前に並べ替える(アクティブシートは元に戻して画面遷移を防ぐ)
    moveSheetWithoutSwitching_(currentSs, newMfgSheet, templateMfgSheet.getIndex());

    clearMfgSheetData_(newMfgSheet);
    Logger.log(`✅ 製造間接費シート「${newMfgSheetName}」の作成完了`);
  }
}

/**
 * シートを指定インデックスの直前に並べ替える。
 * 並べ替えには対象シートを一時的にアクティブ化する必要があるが、
 * 処理後に元のアクティブシートへ戻すことで画面遷移を防ぐ。
 *
 * @param {Spreadsheet} ss        - 対象スプレッドシート
 * @param {Sheet} sheet           - 移動するシート
 * @param {number} targetIndex    - 移動先インデックス(複製元シートのgetIndex())
 */
function moveSheetWithoutSwitching_(ss, sheet, targetIndex) {
  // 現在表示中のシートを記憶
  const originalActive = ss.getActiveSheet();

  // 並べ替え(moveActiveSheetは1始まりのpositionを取る)
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(targetIndex);

  // 元のアクティブシートに戻す → 画面が動かない
  if (originalActive) {
    ss.setActiveSheet(originalActive);
  }
}

/**
 * 転記先SS内から、最も新しい年度の売上進捗表シートを返す(複製元として使う)。
 * カッコの全角/半角に依存しないよう「R<数字>年度全社売上進捗表」の前方一致で判定。
 * ※「受注進捗表」とは別物なので誤マッチしない。
 */
function findLatestSalesSheet_(destSs) {
  const pattern = FY_CONFIG.SALES_SHEET_MATCH_PREFIX;
  let latestSheet = null;
  let latestFy = -1;

  destSs.getSheets().forEach(function(sheet) {
    const m = sheet.getName().match(pattern);
    if (m) {
      const fy = parseInt(m[1], 10);
      if (fy > latestFy) {
        latestFy = fy;
        latestSheet = sheet;
      }
    }
  });

  return latestSheet;
}

/**
 * 転記先SS内の年度進捗表シート名を「新しい年度順(降順)」で返す。
 * 段階的検索(まず最新年度→なければ過年度)に使う。
 *
 * @param {Spreadsheet} destSs
 * @param {string} [preferredFirst] - 最優先で先頭に置きたいシート名(任意)
 * @return {Array<string>} シート名の配列(新しい年度が先頭)
 */
function getFiscalSalesSheetNames_(destSs, preferredFirst) {
  const pattern = FY_CONFIG.SALES_SHEET_MATCH_PREFIX;
  const list = [];

  destSs.getSheets().forEach(function(sheet) {
    const name = sheet.getName();
    const m = name.match(pattern);
    if (m) {
      list.push({ name: name, fy: parseInt(m[1], 10) });
    }
  });

  // 年度の降順(新しい年度が先頭)
  list.sort(function(a, b) { return b.fy - a.fy; });

  let names = list.map(function(item) { return item.name; });

  // preferredFirst が指定され、リストに含まれていれば先頭へ移動
  if (preferredFirst) {
    names = names.filter(function(n) { return n !== preferredFirst; });
    names.unshift(preferredFirst);
  }

  return names;
}

/**
 * 現在のSS内から、最も新しい製造間接費入力シートを返す。
 */
function findLatestMfgSheet_(ss) {
  const pattern = FY_CONFIG.MFG_SHEET_MATCH_PREFIX;
  let latestSheet = null;
  let latestYear = -1;

  ss.getSheets().forEach(function(sheet) {
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
// [PRIVATE] データクリア(売上進捗表)
// ==================================================================
/**
 * 売上進捗表シートのデータ行をクリアする。
 * - 対象行 : A列に連番(数字)があり、合計行・累計行でない行(=案件データ行)
 * - クリア列: B〜K列 (L列=粗利高は数式のため触らない → 自動で0になる)
 * - 温存   : 項目見出し行 / 合計行 / 累計行 / 月度ヘッダー / A列連番 / 数式セル / L列
 *
 * ※ setValues での一括書き戻しは数式を計算結果値に化けさせるため使わない。
 *   クリア対象セルだけを clearContent() で個別に消す。
 */
function clearSheetData_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const startCol = FY_CONFIG.CLEAR_COL_START;  // B(2)
  const endCol = FY_CONFIG.CLEAR_COL_END;      // K(11)
  const numCols = endCol - startCol + 1;

  // A列〜K列の値・数式をまとめて取得
  const range = sheet.getRange(1, 1, lastRow, endCol);
  const values = range.getValues();
  const formulas = range.getFormulas();

  let clearedRows = 0;

  for (let i = 0; i < values.length; i++) {
    const rowNum = i + 1;
    const row = values[i];

    // データ行判定: A列が連番(数字) かつ 合計/累計行でない
    if (!isDataRow_(row)) continue;

    // B〜K列のうち、数式でないセルだけをクリア
    let hasChange = false;
    for (let c = startCol; c <= endCol; c++) {
      const colIdx = c - 1;
      if (formulas[i][colIdx]) continue;  // 数式セルは温存
      const v = row[colIdx];
      if (v !== "" && v !== null && v !== undefined) {
        sheet.getRange(rowNum, c).clearContent();
        hasChange = true;
      }
    }
    if (hasChange) clearedRows++;
  }

  Logger.log(`🧹 売上進捗表のデータクリア完了: ${clearedRows}行(B〜K列。L列数式・ヘッダー類は温存)`);
}

/**
 * 行がデータ行(案件行)かどうかを判定する。
 * 条件: A列が数字(連番) かつ 合計行・累計行でない。
 */
function isDataRow_(rowValues) {
  const aVal = rowValues[FY_CONFIG.COL_A_SEQ - 1];
  const bVal = String(rowValues[FY_CONFIG.COL_B_STORE - 1] || "").trim();

  // A列が連番(数字)か
  const aIsNumber = (typeof aVal === "number")
    || /^\d+$/.test(String(aVal || "").trim());
  if (!aIsNumber) return false;

  // 合計行(B列に「合計」) / 累計行(B列に「累計」)は除外
  if (bVal.indexOf(FY_CONFIG.TOTAL_KEYWORD) !== -1) return false;
  if (bVal.indexOf(FY_CONFIG.RUIKEI_KEYWORD) !== -1) return false;

  return true;
}

// ==================================================================
// [PRIVATE] 月度の +1年 書き換え(売上進捗表)
// ==================================================================
/**
 * 複製した新年度シートの月度表記を「+1年」に書き換える。
 * - データ行のD列(売上年月): "2025年7月" → "2026年7月"
 * - 合計行のB列            : "2025年7月" → "2026年7月"(「合計」等の後続文字は保持)
 * 暦年をすべて +1 する(7〜12月→翌年, 1〜6月→翌年 いずれも +1)。
 * ※ 数式セルは書き換えない(値セルのみ対象)。
 */
function advanceSheetMonths_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const dCol = FY_CONFIG.COL_D_SALES_MONTH; // D(4)
  const bCol = FY_CONFIG.COL_B_STORE;       // B(2)
  const maxCol = Math.max(dCol, bCol);

  const range = sheet.getRange(1, 1, lastRow, maxCol);
  const values = range.getValues();
  const formulas = range.getFormulas();

  let updated = 0;

  for (let i = 0; i < values.length; i++) {
    const rowNum = i + 1;

    // --- D列(データ行の売上年月) ---
    if (!formulas[i][dCol - 1]) {
      const newD = bumpYearInText_(values[i][dCol - 1]);
      if (newD !== null) {
        sheet.getRange(rowNum, dCol).setValue(newD);
        updated++;
      }
    }

    // --- B列(合計行の月度。「2025年7月 合計」等) ---
    if (!formulas[i][bCol - 1]) {
      const newB = bumpYearInText_(values[i][bCol - 1]);
      if (newB !== null) {
        sheet.getRange(rowNum, bCol).setValue(newB);
        updated++;
      }
    }
  }

  Logger.log(`📅 月度を+1年に更新: ${updated}セル`);
}

/**
 * 文字列中の「YYYY年M月」を見つけ、年を+1して返す。
 * 月度表記が含まれなければ null を返す(=書き換え不要)。
 * 例: "2025年7月"       → "2026年7月"
 *     "2025年7月 合計"  → "2026年7月 合計"(後続文字は保持)
 */
function bumpYearInText_(value) {
  if (value === null || value === undefined || value === "") return null;
  const s = String(value);
  const m = s.match(/(\d{4})年(\d{1,2})月/);
  if (!m) return null;

  const newYear = parseInt(m[1], 10) + 1;
  const month = m[2];
  // 元の「YYYY年M月」部分だけを置換し、前後の文字(「 合計」等)は保持
  return s.replace(/(\d{4})年(\d{1,2})月/, newYear + "年" + month + "月");
}

// ==================================================================
// [PRIVATE] データクリア(製造間接費)
// ==================================================================
/**
 * 製造間接費入力シートのデータ行をクリアする。
 * - クリア対象: B〜F列(物件No,顧客名,完成売上高,製造原価計,売上総利益) + I列(転記先行)
 * - 温存対象  : A列(売上月度) / ヘッダー行 / 数式
 * - チェックボックス: 複製直後の確実性のため H列に再設定(全て false に初期化)
 *
 * ※ クリア対象セルだけを clearContent() で個別に消し、数式は温存する。
 * ※ 列番号は MFG_CONFIG を参照(設定の一元管理を維持)。
 */
function clearMfgSheetData_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const numRows = lastRow - 1;
  const startCol = FY_CONFIG.MFG_CLEAR_COL_START;   // B(2)
  const colCount = FY_CONFIG.MFG_CLEAR_COL_COUNT;   // 5列(B〜F)

  // B〜F列の値・数式を取得
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

  // H列(チェックボックス)を再設定 → 複製直後でも確実に機能させる
  const checkCol = MFG_CONFIG.SRC_COL.CHECK;        // H列(8)
  const destRowCol = MFG_CONFIG.SRC_COL.DEST_ROW;   // I列(9)

  sheet.getRange(2, checkCol, numRows, 1).insertCheckboxes();  // 全て false に初期化

  // I列(転記先行)をクリア
  sheet.getRange(2, destRowCol, numRows, 1).clearContent();

  Logger.log(`🧹 製造間接費シートのデータクリア完了: ${numRows}行(数式温存・チェックボックス再設定済み)`);
}

// ==================================================================
// [TEST] 動作確認用
// ==================================================================
function testFiscalYearLogic() {
  const cases = [
    new Date("2025/06/30"),
    new Date("2025/07/01"),
    new Date("2026/06/30"),
    new Date("2026/07/01"),
    new Date("2027/06/30"),
    new Date("2027/07/01")
  ];

  Logger.log("=== 年度判定テスト ===");
  cases.forEach(function(d) {
    const fy = calcFiscalYearReiwa_(d);
    const sheetName = FY_CONFIG.SALES_SHEET_NAME_TEMPLATE.replace("{N}", fy);
    const mfgYear = reiwaToMfgCalendarYear_(fy);
    const mfgName = FY_CONFIG.MFG_SHEET_NAME_TEMPLATE.replace("{YEAR}", mfgYear);
    Logger.log(`${Utilities.formatDate(d, "Asia/Tokyo", "yyyy/MM/dd")} → R${fy}年度 → ${sheetName} / ${mfgName}`);
  });

  Logger.log("\n=== 製造間接費シート名 → 売上進捗表名 変換テスト ===");
  ["2026製造間接費入力シート（事務員専用）", "2027製造間接費入力シート（事務員専用）"].forEach(function(name) {
    Logger.log(`${name} → ${resolveSalesSheetNameFromMfgSheet(name)}`);
  });

  Logger.log("\n=== 月度+1年 変換テスト ===");
  ["2025年7月", "2025年7月 合計", "2026年1月", "累計", "店舗", ""].forEach(function(v) {
    Logger.log(`[${v}] → [${bumpYearInText_(v)}]`);
  });
}

/**
 * テンプレート検索の動作確認用。
 */
function debugCreateR8() {
  const destSs = SpreadsheetApp.openById(CONFIG_PROJECT.DESTINATION_SS_ID);

  Logger.log("=== 現在のシート一覧 ===");
  destSs.getSheets().forEach(function(s) {
    Logger.log("[" + s.getName() + "]");
  });

  const name = FY_CONFIG.SALES_SHEET_NAME_TEMPLATE.replace("{N}", 8);
  Logger.log("生成したシート名: [" + name + "]");

  const found = destSs.getSheetByName(name);
  Logger.log("getSheetByName結果: " + (found ? "見つかった" : "見つからない"));

  const tmpl = findLatestSalesSheet_(destSs);
  Logger.log("複製元テンプレート: " + (tmpl ? "[" + tmpl.getName() + "]" : "なし"));
}