/**
 * ============================================================
 * 製造間接費入力シート → R7年度全社売上進捗表（集計）転記処理
 * ============================================================
 *
 * 【概要】
 * 「2026製造間接費入力シート（事務員専用）」のB〜F列にデータが入力されたら、
 * 別スプレッドシートの「R7年度全社売上進捗表（集計）」の該当月ブロックに
 * 以下のマッピングで転記する。
 *
 *   転記元B（物件No）    → 転記先E（物件No）
 *   転記元C（顧客名）    → 転記先G（顧客名）
 *   転記元F（売上総利益）→ 転記先L（粗利高）
 *
 * 【転記先の構造】
 *   - A列：番号
 *   - B列：店舗 ／ 合計行では月度（例：「2025年7月」）
 *   - C列：担当者 ／ 合計行では「合計」
 *   - D列：売上年月 / E列：物件No / F列：工務店 / G列：お客様名
 *   - L列：粗利高
 *   各月ブロックの最終行は B列=月度 + C列=「合計」で判定。
 *   データ行の月度はD列に入る（ブロック識別自体は合計行で行う）。
 *
 * 【連携列（転記元に自動追加）】
 *   H列：転記済みチェックボックス
 *   I列：転記先の行番号
 *
 * 【トリガー設定】
 * installable onEdit トリガーとして `handleManufacturingOverheadEdit`
 * を直接設定する（別スプレッドシートへの書き込みがあるため、シンプルな
 * onEdit では権限不足）。
 * ============================================================
 */

// ============================================================
// 設定（CONFIG）
// ============================================================
const MFG_DEFAULT_SCRIPT_PROPERTIES = {
  MFG_SOURCE_SHEET_NAME: "2026製造間接費入力シート（事務員専用）",
  MFG_DEST_SHEET_NAME: "R7年度全社売上進捗表（集計）",
  MFG_DEST_TOTAL_MARKER: "合計",
  MFG_DEST_RUIKEI_MARKER: "累計",
};

function getRequiredScriptPropertyMfg_(key) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (value) {
    return value;
  }
  if (Object.prototype.hasOwnProperty.call(MFG_DEFAULT_SCRIPT_PROPERTIES, key)) {
    Logger.log('⚠ Script Property "' + key + '" が未設定のため初期値を使用します。');
    return MFG_DEFAULT_SCRIPT_PROPERTIES[key];
  }
  throw new Error(
    'Script Properties に "' + key + '" が未設定です。'
    + ' Apps Script エディタ > プロジェクトの設定 > スクリプト プロパティ で設定してください。'
  );
}

const MFG_CONFIG = {
  // 転記元シート（現在のスプレッドシート内のタブ）
  SOURCE_SHEET_NAME: getRequiredScriptPropertyMfg_('MFG_SOURCE_SHEET_NAME'),

  // 転記先スプレッドシート（別ファイル）
  DEST_SPREADSHEET_ID: getRequiredScriptPropertyMfg_('MFG_DEST_SPREADSHEET_ID'),
  DEST_SHEET_NAME: getRequiredScriptPropertyMfg_('MFG_DEST_SHEET_NAME'),

  // 転記元の列番号（1始まり）
  SRC_COL: {
    MONTH: 1,         // A: 売上月度
    BUKKEN_NO: 2,     // B: 物件No
    CUSTOMER: 3,      // C: 顧客名
    SALES: 4,         // D: 完成売上高（転記しない）
    COST: 5,          // E: 製造原価計（転記しない）
    GROSS_PROFIT: 6,  // F: 売上総利益
    CHECK: 8,         // H: 転記済みチェック
    DEST_ROW: 9,      // I: 転記先行番号
  },

  // 転記先の列番号（1始まり）
  DEST_COL: {
    STORE: 2,         // B: 店舗 ／ 合計行では月度
    PERSON: 3,        // C: 担当者 ／ 合計行では「合計」
    SALES_MONTH: 4,   // D: 売上年月
    BUKKEN_NO: 5,     // E: 物件No
    CUSTOMER: 7,      // G: お客様名
    GROSS_PROFIT: 12, // L: 粗利高
  },

  // 転記先の合計行判定マーカー
  DEST_TOTAL_MARKER: getRequiredScriptPropertyMfg_('MFG_DEST_TOTAL_MARKER'),

  // 転記先の累計行判定マーカー
  DEST_RUIKEI_MARKER: getRequiredScriptPropertyMfg_('MFG_DEST_RUIKEI_MARKER'),
};

// ============================================================
// メインエントリーポイント（onEdit トリガーに直接設定する関数）
// ============================================================
/**
 * インストーラブル onEdit トリガーに直接設定する関数。
 * @param {Object} e - onEdit イベントオブジェクト
 */
function handleManufacturingOverheadEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();

  // 「{西暦}製造間接費入力シート(事務員専用)」形式のシート全てを対象に
  if (!/^\d{4}製造間接費入力シート[((]事務員専用[))]$/.test(sheet.getName())) return;

  const col = e.range.getColumn();
  const row = e.range.getRow();

  // ヘッダー行は無視
  if (row <= 1) return;

  // B〜F列 または H列の編集のみ対象
  const isDataEdit = (col >= MFG_CONFIG.SRC_COL.BUKKEN_NO && col <= MFG_CONFIG.SRC_COL.GROSS_PROFIT);
  const isCheckEdit = (col === MFG_CONFIG.SRC_COL.CHECK);

  if (!isDataEdit && !isCheckEdit) return;

  // 無限ループ防止：スクリプト自身による編集をスキップ
  if (isScriptEdit_()) {
    Logger.log('⏭ スクリプトによる編集のためスキップ');
    return;
  }

  // ロック（多重実行防止）
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) {
    Logger.log('⚠ ロック取得失敗、処理をスキップ');
    return;
  }

  const startTime = new Date().getTime();
  try {
    setScriptEditFlag_(true);

    if (isCheckEdit) {
      handleCheckboxToggle_(sheet, row, e);
    } else {
      handleDataEdit_(sheet, row);
    }

    const elapsed = new Date().getTime() - startTime;
    Logger.log(`✅ 処理完了（${elapsed}ms）`);
  } catch (err) {
    Logger.log('❌ エラー: ' + err.message + '\n' + err.stack);
  } finally {
    setScriptEditFlag_(false);
    lock.releaseLock();
  }
}

// ============================================================
// データ編集時の処理（B〜F列の編集）
// ============================================================
function handleDataEdit_(sheet, row) {
  const rowData = getSourceRowData_(sheet, row);
  const srcSheetName = sheet.getName();  // 追加

  Logger.log(`📝 データ編集: 行=${row}, 月=${rowData.month}, 物件No=${rowData.bukkenNo}, 顧客=${rowData.customer}, 粗利=${rowData.grossProfit}`);

  // スキップ判定①：B列が「◯◯◯◯年◯月」形式（集計行等）
  if (isMonthFormat_(rowData.bukkenNo)) {
    Logger.log('⏭ B列が月度形式のためスキップ');
    return;
  }

  // B・C・F列すべて空白なら「転記削除」扱い
  if (isEmptyRow_(rowData)) {
    // 既に転記済みなら転記先を削除し、H/Iをクリア
    if (rowData.destRow) {
      deleteDestRow_(rowData, srcSheetName);
      Logger.log(`🗑 データ削除を検知: 転記先行 ${rowData.destRow} を削除`);
    }

    // H列（チェック）とI列（転記先行）を初期化
    sheet.getRange(row, MFG_CONFIG.SRC_COL.CHECK).setValue(false);
    sheet.getRange(row, MFG_CONFIG.SRC_COL.DEST_ROW).clearContent();
    Logger.log('🧹 転記元のH列/I列をクリア');
    return;
  }

  // H列がチェック済み かつ I列に転記先行番号あり → 更新処理
  if (rowData.isChecked && rowData.destRow) {
    updateDestRow_(rowData, srcSheetName);
  } else {
    // 新規追加処理
    insertNewRow_(sheet, row, rowData);
  }
}

// ============================================================
// チェックボックス操作時の処理（H列の編集）
// ============================================================
function handleCheckboxToggle_(sheet, row, e) {
  const rowData = getSourceRowData_(sheet, row);
  const srcSheetName = sheet.getName();  // ★この行を追加
  const newValue = e.value === 'TRUE' || e.value === true;

  Logger.log(`☑ チェック操作: 行=${row}, 新状態=${newValue}, 転記先行=${rowData.destRow}`);

  if (newValue === true) {
    // チェックをONにした

    if (rowData.destRow) {
      // 既に転記先行がある → 何もしない（既に転記済み）
      Logger.log('⏭ 既に転記済みのため処理なし');
      return;
    }

    // データがあれば新規追加
    if (isMonthFormat_(rowData.bukkenNo)) {
      Logger.log('⏭ B列が月度形式のためスキップ');
      return;
    }

    if (isEmptyRow_(rowData)) {
      Logger.log('⏭ データ空のため新規追加なし');
      return;
    }

    insertNewRow_(sheet, row, rowData);
  } else {
    // チェックをOFFにした → 転記先行を削除

    if (!rowData.destRow) {
      Logger.log('⏭ 転記先行の記録がないためスキップ');
      return;
    }

    deleteDestRow_(rowData, srcSheetName);

    // I列の行番号をクリア
    sheet.getRange(row, MFG_CONFIG.SRC_COL.DEST_ROW).clearContent();
    Logger.log(`🗑 転記先行 ${rowData.destRow} を削除し、I列をクリア`);
  }
}

// ============================================================
// 新規追加処理
// ============================================================
function insertNewRow_(sheet, srcRow, rowData) {
  const destSheet = getDestSheet_(sheet.getName());
  if (!destSheet) return;

  // 該当月ブロックの空行を探す
  const targetRow = findEmptyRowInMonthBlock_(destSheet, rowData.month);

  if (!targetRow) {
    Logger.log(`⏭ 転記先に「${rowData.month}」ブロックが見つからない、または空きなしのため処理なし`);
    return;
  }

  // 転記先に書き込み
  writeToDestRow_(destSheet, targetRow, rowData);

  // 転記元の H列にチェック、I列に行番号を記録
  sheet.getRange(srcRow, MFG_CONFIG.SRC_COL.CHECK).setValue(true);
  sheet.getRange(srcRow, MFG_CONFIG.SRC_COL.DEST_ROW).setValue(targetRow);

  Logger.log(`✨ 新規追加完了: 転記元行=${srcRow} → 転記先行=${targetRow}`);
}

// ============================================================
// 更新処理
// ============================================================
function updateDestRow_(rowData, srcSheetName) {
  const destSheet = getDestSheet_(srcSheetName);
  if (!destSheet) return;
  writeToDestRow_(destSheet, rowData.destRow, rowData);
  Logger.log(`🔄 更新完了: 転記先行=${rowData.destRow}`);
}

// ============================================================
// 削除処理（転記先の行全体をクリア）
// ============================================================
function deleteDestRow_(rowData, srcSheetName) {
  const destSheet = getDestSheet_(srcSheetName);
  if (!destSheet) return;
  const lastCol = destSheet.getLastColumn();
  destSheet.getRange(rowData.destRow, 1, 1, lastCol).clearContent();
}

// ============================================================
// 転記先への書き込み（E・G・L列）
// ============================================================
function writeToDestRow_(destSheet, destRow, rowData) {
  destSheet.getRange(destRow, MFG_CONFIG.DEST_COL.BUKKEN_NO).setValue(rowData.bukkenNo);
  destSheet.getRange(destRow, MFG_CONFIG.DEST_COL.CUSTOMER).setValue(rowData.customer);
  destSheet.getRange(destRow, MFG_CONFIG.DEST_COL.GROSS_PROFIT).setValue(rowData.grossProfit);
}

// ============================================================
// 転記先シート取得
// ============================================================
function getDestSheet_(srcSheetName) {
  try {
    const destSs = SpreadsheetApp.openById(MFG_CONFIG.DEST_SPREADSHEET_ID);

    // srcSheetName から対応する年度シート名を解決(なければ自動作成)
    let destSheetName = MFG_CONFIG.DEST_SHEET_NAME;  // フォールバック
    if (srcSheetName) {
      const resolved = resolveSalesSheetNameFromMfgSheet(srcSheetName);
      if (resolved) {
        // 対応する売上進捗表がなければ作成する必要があるが、
        // 通常は売上予定日編集時に作成済みのはず。念のため存在チェック。
        if (destSs.getSheetByName(resolved)) {
          destSheetName = resolved;
        } else {
          Logger.log(`⚠ 転記先シート「${resolved}」が未作成。営業進捗管理側で売上予定日を編集してください`);
          return null;
        }
      }
    }

    const destSheet = destSs.getSheetByName(destSheetName);
    if (!destSheet) {
      Logger.log(`❌ 転記先シート「${destSheetName}」が見つかりません`);
      return null;
    }
    return destSheet;
  } catch (err) {
    Logger.log(`❌ 転記先スプレッドシートを開けません: ${err.message}`);
    return null;
  }
}

// ============================================================
// 該当月ブロックの空行を探す（合計行ベースで判定）
// ============================================================
/**
 * 転記先シートから、指定月度ブロックの「E列（物件No）が空の最初の行」を返す。
 * 該当月ブロックが存在しない、または空きがない場合は null を返す。
 *
 * 月ブロックの判定方法：
 *   - B列が月度（例：「2025年7月」） + C列が「合計」 の行がブロックの末尾
 *   - その1つ前の合計行の直後（or シート先頭）から、次の合計行の1つ前までが当該月ブロック
 *
 * 例：
 *   行4～行40：2025年7月のデータ行
 *   行41：B列「2025年7月」、C列「合計」 ← 2025年7月ブロックの末尾
 *   行42～行80：2025年8月のデータ行
 *   行81：B列「2025年8月」、C列「合計」 ← 2025年8月ブロックの末尾
 *
 * @param {Sheet} destSheet
 * @param {string} targetMonth 例：'2026年1月'
 * @return {number|null} 空行の行番号（1始まり）
 */
function findEmptyRowInMonthBlock_(destSheet, targetMonth) {
  if (!targetMonth) return null;

  const lastRow = destSheet.getLastRow();
  if (lastRow < 1) return null;

  // E列/G列の空白判定に加えて、合計行・累計行判定にも使うため B〜L を取得
  const maxCol = MFG_CONFIG.DEST_COL.GROSS_PROFIT; // L列まで
  const values = destSheet.getRange(1, 1, lastRow, maxCol).getValues();

  // Step 1: 該当月の合計行（B=月度、C=合計）を特定
  let targetTotalRowIndex = -1; // 0始まり
  for (let i = 0; i < values.length; i++) {
    const bCell = normalizeMonth_(values[i][MFG_CONFIG.DEST_COL.STORE - 1]);
    const cCell = String(values[i][MFG_CONFIG.DEST_COL.PERSON - 1] || '').trim();

    const isTotalRow = isMonthFormat_(values[i][MFG_CONFIG.DEST_COL.STORE - 1])
      && cCell === MFG_CONFIG.DEST_TOTAL_MARKER;

    if (!isTotalRow) continue;

    if (bCell === targetMonth) {
      targetTotalRowIndex = i;
      break;
    }
  }

  if (targetTotalRowIndex === -1) {
    Logger.log(`⚠ 「${targetMonth}」の合計行が見つかりません`);
    return null;
  }

  // Step 2: 該当月の合計行の上方向を走査し、直近の「累計」行を特定
  let ruikeiRowIndex = -1; // 0始まり
  for (let i = targetTotalRowIndex - 1; i >= 0; i--) {
    const bCellRaw = String(values[i][MFG_CONFIG.DEST_COL.STORE - 1] || '').trim();
    if (bCellRaw === MFG_CONFIG.DEST_RUIKEI_MARKER) {
      ruikeiRowIndex = i;
      break;
    }
  }

  if (ruikeiRowIndex === -1) {
    Logger.log(`⚠ 「${targetMonth}」の直前に「${MFG_CONFIG.DEST_RUIKEI_MARKER}」行が見つかりません`);
    return null;
  }

  // セクション定義：
  //   開始 = 累計行の次行
  //   終了 = 合計行の1つ前
  const blockStartRow = ruikeiRowIndex + 2; // 1始まり
  const blockEndRow = targetTotalRowIndex;  // 1始まり

  if (blockStartRow > blockEndRow) {
    Logger.log(`⚠ 「${targetMonth}」ブロックの範囲が不正: start=${blockStartRow}, end=${blockEndRow}`);
    return null;
  }

  // Step 3: ブロック範囲内で E列（物件No）と G列（顧客名）が空の最初の行を探す
  const eColIdx = MFG_CONFIG.DEST_COL.BUKKEN_NO - 1;
  const gColIdx = MFG_CONFIG.DEST_COL.CUSTOMER - 1;
  for (let r = blockStartRow; r <= blockEndRow; r++) {
    const row = values[r - 1];
    const isWritable = isWritableRow_(row, eColIdx, gColIdx);
    if (isWritable) {
      return r;
    }
  }

  Logger.log(`⚠ 「${targetMonth}」セクション（行${blockStartRow}～${blockEndRow}）に記入可能行がありません`);
  return null;
}

function isWritableRow_(rowValues, eColIdx, gColIdx) {
  const eVal = rowValues[eColIdx];
  const gVal = rowValues[gColIdx];
  const isEEmpty = (eVal === '' || eVal === null || eVal === undefined);
  const isGEmpty = (gVal === '' || gVal === null || gVal === undefined);
  return isEEmpty && isGEmpty;
}

// ============================================================
// 最初のデータ行を動的に探す
// ============================================================
/**
 * 画像によると、3行目までがヘッダー、4行目以降がデータ。
 * 安全のため、先頭から順に「A列が数字 or B列にデータがあり、合計行でない」
 * 最初の行を探す。見つからなければデフォルトで 4 を返す。
 */
function findFirstDataRow_(values) {
  for (let i = 0; i < values.length; i++) {
    const aCell = values[i][0];
    const bCell = values[i][MFG_CONFIG.DEST_COL.STORE - 1];
    const cCell = String(values[i][MFG_CONFIG.DEST_COL.PERSON - 1] || '').trim();

    // A列が数字
    const aIsNumber = (typeof aCell === 'number') || /^\d+$/.test(String(aCell || '').trim());
    // 合計行ではない
    const isTotalRow = isMonthFormat_(bCell) && cCell === MFG_CONFIG.DEST_TOTAL_MARKER;

    if (aIsNumber && !isTotalRow) {
      return i + 1;  // 1始まり
    }
  }
  return 4;  // デフォルト（画像の構造より）
}

// ============================================================
// 転記元の行データを取得
// ============================================================
function getSourceRowData_(sheet, row) {
  const values = sheet.getRange(row, 1, 1, MFG_CONFIG.SRC_COL.DEST_ROW).getValues()[0];

  return {
    row: row,
    month: normalizeMonth_(values[MFG_CONFIG.SRC_COL.MONTH - 1]),
    bukkenNo: values[MFG_CONFIG.SRC_COL.BUKKEN_NO - 1],
    customer: values[MFG_CONFIG.SRC_COL.CUSTOMER - 1],
    grossProfit: values[MFG_CONFIG.SRC_COL.GROSS_PROFIT - 1],
    isChecked: values[MFG_CONFIG.SRC_COL.CHECK - 1] === true,
    destRow: Number(values[MFG_CONFIG.SRC_COL.DEST_ROW - 1]) || null,
  };
}

// ============================================================
// 空行判定（B・C・F列がすべて空白なら空行とみなす）
// ============================================================
function isEmptyRow_(rowData) {
  const b = String(rowData.bukkenNo || '').trim();
  const c = String(rowData.customer || '').trim();
  const f = String(rowData.grossProfit || '').trim();
  return b === '' && c === '' && f === '';
}

// ============================================================
// 月度形式判定（「2025年10月」のような形式か）
// ============================================================
function isMonthFormat_(value) {
  if (value === null || value === undefined || value === '') return false;
  if (value instanceof Date) return true;  // Date型も月度として扱う

  const s = String(value).trim();
  return /^\d{4}年\d{1,2}月$/.test(s);
}

// ============================================================
// 月度の正規化（「2026年1月」形式に統一）
// ============================================================
function normalizeMonth_(value) {
  if (value === null || value === undefined || value === '') return '';

  // Date 型の場合
  if (value instanceof Date) {
    const year = value.getFullYear();
    const month = value.getMonth() + 1;
    return year + '年' + month + '月';
  }

  let s = String(value).trim();

  // 全角数字 → 半角数字（安全な個別変換）
  const fullToHalf = {
    '0': '0', '1': '1', '2': '2', '3': '3', '4': '4',
    '5': '5', '6': '6', '7': '7', '8': '8', '9': '9'
  };
  s = s.replace(/[0-9]/g, function(ch) {
    return fullToHalf[ch] || ch;
  });

  // 「2026年01月」→「2026年1月」（ゼロ埋めを削除して統一）
  const match = s.match(/^(\d{4})年0?(\d{1,2})月$/);
  if (match) {
    return match[1] + '年' + parseInt(match[2], 10) + '月';
  }

  return s;
}

// ============================================================
// 無限ループ防止：スクリプトによる編集フラグ管理
// ============================================================
const MFG_SCRIPT_EDIT_FLAG = 'MFG_SCRIPT_EDIT_FLAG';

function setScriptEditFlag_(value) {
  const props = PropertiesService.getScriptProperties();
  if (value) {
    props.setProperty(MFG_SCRIPT_EDIT_FLAG, String(new Date().getTime()));
  } else {
    props.deleteProperty(MFG_SCRIPT_EDIT_FLAG);
  }
}

function isScriptEdit_() {
  const props = PropertiesService.getScriptProperties();
  const flag = props.getProperty(MFG_SCRIPT_EDIT_FLAG);
  if (!flag) return false;

  const elapsed = new Date().getTime() - Number(flag);
  if (elapsed > 10000) {
    props.deleteProperty(MFG_SCRIPT_EDIT_FLAG);
    return false;
  }
  return true;
}

// ============================================================
// 初期設定（H列にチェックボックスを設定）
// ============================================================
function setupManufacturingOverheadSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MFG_CONFIG.SOURCE_SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(`シート「${MFG_CONFIG.SOURCE_SHEET_NAME}」が見つかりません`);
    return;
  }

  const lastRow = Math.max(sheet.getLastRow(), 2);
  const numRows = lastRow - 1;

  const checkRange = sheet.getRange(2, MFG_CONFIG.SRC_COL.CHECK, numRows, 1);
  checkRange.insertCheckboxes();

  sheet.getRange(1, MFG_CONFIG.SRC_COL.CHECK).setValue('転記済');
  sheet.getRange(1, MFG_CONFIG.SRC_COL.DEST_ROW).setValue('転記先行');

  SpreadsheetApp.getUi().alert('H列にチェックボックスを設定しました。\nI列のヘッダーも追加しました。');
  Logger.log(`✅ 初期設定完了: H列${numRows}行にチェックボックス設定`);
}

// ============================================================
// 手動実行用：現在のアクティブセルの行を再転記
// ============================================================
function manualTransferCurrentRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getName() !== MFG_CONFIG.SOURCE_SHEET_NAME) {
    SpreadsheetApp.getUi().alert(`転記元シート「${MFG_CONFIG.SOURCE_SHEET_NAME}」で実行してください`);
    return;
  }

  const row = sheet.getActiveCell().getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert('データ行を選択してください');
    return;
  }

  try {
    setScriptEditFlag_(true);
    handleDataEdit_(sheet, row);
    SpreadsheetApp.getUi().alert(`行${row} の転記処理を実行しました。ログを確認してください。`);
  } finally {
    setScriptEditFlag_(false);
  }
}

// ============================================================
// デバッグ用：転記先シートの月ブロック構造を確認
// ============================================================
/**
 * 転記先シートの合計行を検出してログ出力する。
 * 月ブロックが正しく認識されているか確認用。
 */
function debugListMonthBlocks(srcSheetName) {
  // 引数がなければ最新の製造間接費シートから推定
  const name = srcSheetName || MFG_CONFIG.SOURCE_SHEET_NAME;
  const destSheet = getDestSheet_(name);
  if (!destSheet) return;

  const lastRow = destSheet.getLastRow();
  const maxCol = MFG_CONFIG.DEST_COL.BUKKEN_NO;
  const values = destSheet.getRange(1, 1, lastRow, maxCol).getValues();

  Logger.log('=== 月ブロック合計行の検出結果 ===');
  let count = 0;
  for (let i = 0; i < values.length; i++) {
    const bCell = values[i][MFG_CONFIG.DEST_COL.STORE - 1];
    const cCell = String(values[i][MFG_CONFIG.DEST_COL.PERSON - 1] || '').trim();
    const isTotal = isMonthFormat_(bCell) && cCell === MFG_CONFIG.DEST_TOTAL_MARKER;

    if (isTotal) {
      count++;
      Logger.log(`行${i + 1}: B=「${bCell}」(正規化:${normalizeMonth_(bCell)}) / C=「${cCell}」`);
    }
  }
  Logger.log(`=== 合計 ${count} 個のブロックを検出 ===`);
}