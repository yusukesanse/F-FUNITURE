// ==================================================================
// overhead.gs ― 製造間接費 → 売上進捗表 への「差分転記」（メニューバー方式）
// ==================================================================
// 【動作】
//   ・メニューバー「製造間接費 ＞ 売上進捗表へ転記」を押すと実行（自動転記なし）。
//   ・"今開いている"製造間接費入力シートのデータを、その年度に対応する
//     売上進捗表シートへ書き込む（年度はシート名から自動判定）。
//   ・各行を、売上月度（A列）の月ブロックに【新規行として追加】する。
//       売上（J列）   ← 売上総利益
//       粗利率（K列） ← 100%（=1）
//       粗利高（L列） ← 売上総利益
//   ・「転記済み」チェック（H列）が付いた行はスルー。未チェックの行だけ追加し、
//     完了したらチェックを付ける（＝何度押しても未転記だけ追記。重複防止）。
//
// 【依存】同一プロジェクト内の config.gs（CONFIG_PROJECT）／salesTransfer.gs
//   （findMonthBlock・findOrCreateTargetRow・formatMonthString）／fiscalYear.gs
//   （resolveDestContextFromMfgSheet）を利用します。
//
// 【設定】列番号（1始まり）は実シートに合わせて調整してください。
//   SRC_COL.CHECK と SRC_COL.DEST_ROW のキーは fiscalYear.gs も参照するため残しています。
//
// ★今回の変更（2026/06）
//   メニューから「（初回設定）転記済み列を準備」を削除。
//   新年度シートは fiscalYear.gs の複製時に
//   「チェックボックス全オフ・H1見出し『転記済み』」まで自動整備されるため、
//   通常運用で初回設定ボタンは不要。
//   ※ setupOverheadSheet 関数は残しているので、最初の元シートを一度だけ整える等の
//     特別なケースでは GASエディタから手動実行できます。
// ==================================================================

const MFG_CONFIG = {

  // --- 転記元：製造間接費入力シートの列（画像2の構成） ---
  SRC_COL: {
    MONTH:        1,  // A: 売上月度（例 "2025年7月"）
    BUKKEN_NO:    2,  // B: 物件No（空のことあり）
    CUSTOMER:     3,  // C: 顧客名
    SALES:        4,  // D: 完成売上高（転記しない）
    COST:         5,  // E: 製造原価計（転記しない）
    GROSS_PROFIT: 6,  // F: 売上総利益 ← 転記する値
    GROSS_RATE:   7,  // G: 売上総利益率（転記しない・参考）
    CHECK:        8,  // H: 「転記済み」チェックボックス（新設）
    DEST_ROW:     9   // I: 未使用（overhead.gsでは書き込まない。fiscalYear.gsのシート初期化が参照するためキーのみ残置）
  },
  SRC_DATA_START_ROW: 2,                          // データ開始行（1行目＝見出し）
  SRC_SHEET_PATTERN: /^\d{4}製造間接費入力シート/, // 製造間接費入力シートの判定

  GROSS_MARGIN_RATE_VALUE: 1  // 粗利率（K列）に入れる値。1 = 100%
  // ※ 転記先は年度別ファイル（fiscalYear.gs の resolveDestContextFromMfgSheet で解決）
};

// ==================================================================
// メニュー追加（スプレッドシートを開いた時に自動実行）
// ==================================================================
// ※ プロジェクト内に onOpen は1つだけにしてください。
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("製造間接費")
    .addItem("売上進捗表へ転記", "transferOverhead")
    .addToUi();
}

// ==================================================================
// ボタン本体：差分転記（未転記の行だけ、月ブロックへ新規追加）
// ==================================================================
function transferOverhead() {
  const ui = SpreadsheetApp.getUi();
  try {
    const src = MFG_CONFIG.SRC_COL;
    const cols = CONFIG_PROJECT.DESTINATION_COLUMNS;
    const srcSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const srcName = srcSheet.getName();

    // 1) 製造間接費入力シート以外では実行しない
    if (!MFG_CONFIG.SRC_SHEET_PATTERN.test(srcName)) {
      ui.alert("製造間接費入力シートを開いてから実行してください。");
      return;
    }

    // 2) 対応する年度ファイル（スプレッドシート）を製造間接費シート名から特定
    const destInfo = resolveDestContextFromMfgSheet(srcName);
    if (!destInfo) {
      ui.alert(`「${srcName}」に対応する年度を判定できませんでした。`);
      return;
    }
    if (!destInfo.ssId) {
      ui.alert(
        `令和${destInfo.reiwaYear}年度の売上進捗表スプレッドシートが未登録です。\n` +
        `スクリプトプロパティ「DEST_SS_ID_R${destInfo.reiwaYear}」にファイルIDを設定してください。`
      );
      return;
    }
    let destSs;
    try {
      destSs = SpreadsheetApp.openById(destInfo.ssId);
    } catch (err) {
      ui.alert(`令和${destInfo.reiwaYear}年度の転記先ファイル（ID: ${destInfo.ssId}）が開けません。`);
      return;
    }
    const destSheet = destSs.getSheetByName(destInfo.sheetName);
    if (!destSheet) {
      ui.alert(`転記先シート「${destInfo.sheetName}」が令和${destInfo.reiwaYear}年度ファイルに見つかりません。`);
      return;
    }
    const destSheetName = destInfo.sheetName;

    // 3) 転記元データを一括取得
    const lastRow = srcSheet.getLastRow();
    if (lastRow < MFG_CONFIG.SRC_DATA_START_ROW) {
      ui.alert("転記するデータがありません。");
      return;
    }
    const numRows = lastRow - MFG_CONFIG.SRC_DATA_START_ROW + 1;
    const maxCol = Math.max(src.MONTH, src.BUKKEN_NO, src.CUSTOMER,
                            src.GROSS_PROFIT, src.CHECK, src.DEST_ROW);
    const srcValues = srcSheet
      .getRange(MFG_CONFIG.SRC_DATA_START_ROW, 1, numRows, maxCol)
      .getValues();

    // 4) 差分転記
    let written = 0, skippedDone = 0, skippedEmpty = 0, skippedTotal = 0, noBlock = 0;
    const START = cols.STORE;                // 2
    const END = cols.GROSS_MARGIN_AMOUNT;    // 12
    const width = END - START + 1;

    for (let i = 0; i < srcValues.length; i++) {
      const row = srcValues[i];
      const srcRowNum = MFG_CONFIG.SRC_DATA_START_ROW + i;

      const isDone = row[src.CHECK - 1] === true;
      if (isDone) { skippedDone++; continue; }   // 転記済みはスルー

      const monthRaw = row[src.MONTH - 1];
      // A列に「合計」が入っている行（集計行）はスルー
      if (String(monthRaw).indexOf("合計") !== -1) { skippedTotal++; continue; }

      const grossProfit = row[src.GROSS_PROFIT - 1];
      const hasMonth = monthRaw !== "" && monthRaw !== null;
      const hasProfit = grossProfit !== "" && grossProfit !== null;
      if (!hasMonth || !hasProfit) { skippedEmpty++; continue; }  // 月度/売上総利益が空

      const bukken = String(row[src.BUKKEN_NO - 1]).trim();
      const customer = String(row[src.CUSTOMER - 1]).trim();
      const monthStr = mfgMonthToBlockString_(monthRaw);

      // 月ブロックを特定（毎回読み直して行挿入のズレを防ぐ）
      const destData = destSheet.getDataRange().getValues();
      const monthBlock = findMonthBlock(destData, monthStr, cols);
      if (!monthBlock) { noBlock++; continue; }  // その月の枠が無い

      const targetRow = findOrCreateTargetRow(destSheet, destData, monthBlock, cols);

      // 行を書き込み（STORE〜粗利高 の範囲を一括）
      const vals = new Array(width).fill("");
      vals[cols.SALES_MONTH_COL - START]      = monthStr;     // D: 売上年月
      vals[cols.OBJECT_NO - START]            = bukken;       // E: 物件No（空でも可）
      vals[cols.PROJECT_NAME - START]         = customer;     // G: 顧客名
      vals[cols.SALES_COL - START]            = grossProfit;  // J: 売上 ← 売上総利益
      vals[cols.GROSS_MARGIN_RATE_AB - START] = MFG_CONFIG.GROSS_MARGIN_RATE_VALUE; // K: 粗利率=100%
      vals[cols.GROSS_MARGIN_AMOUNT - START]  = grossProfit;  // L: 粗利高 ← 売上総利益
      destSheet.getRange(targetRow, START, 1, width).setValues([vals]);

      // 転記済みチェックを付ける
      srcSheet.getRange(srcRowNum, src.CHECK).setValue(true);
      SpreadsheetApp.flush();
      written++;
    }

    // 5) 結果を表示
    ui.alert(
      "転記完了",
      `転記先：${destSheetName}\n\n`
      + `・書き込み：${written}件\n`
      + `・スキップ（転記済み）：${skippedDone}件\n`
      + `・スキップ（合計行）：${skippedTotal}件\n`
      + `・スキップ（月度/売上総利益が空）：${skippedEmpty}件\n`
      + `・月ブロックなし（その月の枠が売上進捗表に無い）：${noBlock}件`,
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert("エラー", `転記中にエラーが発生しました。\n\n${err.message}`, ui.ButtonSet.OK);
  }
}

// ==================================================================
// 「転記済み」チェックボックス列（H列）を用意する
//   ※ メニューからは外しました。新年度シートは fiscalYear.gs の複製時に自動整備されます。
//     最初の元シートを一度だけ整える等の場合は、GASエディタでこの関数を直接実行してください。
// ==================================================================
function setupOverheadSheet() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!MFG_CONFIG.SRC_SHEET_PATTERN.test(sheet.getName())) {
    ui.alert("製造間接費入力シートを開いてから実行してください。");
    return;
  }
  const startRow = MFG_CONFIG.SRC_DATA_START_ROW;
  const lastRow = Math.max(sheet.getLastRow(), startRow);
  const numRows = lastRow - startRow + 1;

  sheet.getRange(1, MFG_CONFIG.SRC_COL.CHECK).setValue("転記済み");      // H1 見出し
  sheet.getRange(startRow, MFG_CONFIG.SRC_COL.CHECK, numRows, 1).insertCheckboxes(); // H列にチェックボックス
  SpreadsheetApp.flush();
  ui.alert(`H列に「転記済み」チェックボックスを設定しました（${numRows}行）。`);
}

// ==================================================================
// 売上月度の値を、売上進捗表の月ブロック見出しと同じ表記に正規化する
//   "2025年7月" / "2025年07月" / "2025/7" / Date → "2025年7月"
// ==================================================================
function mfgMonthToBlockString_(v) {
  if (v instanceof Date && !isNaN(v)) return formatMonthString(v);
  const m = String(v).match(/(\d{4})\D+(\d{1,2})/);
  if (m) return m[1] + "年" + parseInt(m[2], 10) + "月";
  return String(v).trim();
}