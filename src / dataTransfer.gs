
// ==================================================================
// ★★★ main.gs - メイン処理 ★★★
// ==================================================================
// [INDEX]
// [ENTRY]    handleProjectUpdate
// [HANDLER]  handleDateTrigger / handleStatusTrigger / handleAmountTrigger / handleTextTrigger
// [PATTERN]  handleSameGroupTransition / handleOtherDealSync / handleOtherDealToFinalStage / handleFinalStageToOtherDeal
// [TEST]     testWriteToSheet / testConfig
// [UTILITY]  旧 lib.gs 統合関数群（createTimer 以降）

// ------------------------------------------------------------------
// [ENTRY]
// ------------------------------------------------------------------
function handleProjectUpdate(e) {
  const timer = createTimer();
  logStart("handleProjectUpdate");
  
  if (!e || !e.range) {
    Logger.log("❌ エラー: イベントオブジェクトが不正です。");
    return;
  }
  
  const range = e.range;
  const sheet = range.getSheet();
  const C = CONFIG_PROJECT;
  const sheetNameTrimmed = sheet.getName().trim();
  
  Logger.log(`📋 編集シート: ${sheetNameTrimmed}, 行: ${range.getRow()}, 列: ${range.getColumn()}`);
  
  if (!C.SOURCE_SHEET_NAMES.includes(sheetNameTrimmed) || range.getRow() <= 1) {
    Logger.log("⏭️ 対象外のシートまたはヘッダー行のため終了");
    return;
  }

  timer.lap("初期チェック完了");

  const lastCol = sheet.getLastColumn();
  const rowNum = range.getRow();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rowValues = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
  timer.lap("転記元データ取得");

  const editedColHeader = headers[range.getColumn() - 1];
  Logger.log(`✏️ 編集された列: ${editedColHeader}`);

  const triggerConfig = getTriggerConfig(editedColHeader);
  if (!triggerConfig) {
    Logger.log("⏭️ トリガー対象外の列のため終了");
    return;
  }
  Logger.log(`🎯 トリガー検出: ${triggerConfig.key} (${triggerConfig.type})`);

  const sourceData = extractSourceData(headers, rowValues, e, editedColHeader);
  sourceData.sheetNameTrimmed = sheetNameTrimmed;
  Logger.log(`📦 取得データ: 物件No=${sourceData.objectNumber}, 案件名=${sourceData.customerName}, ステージ=${sourceData.currentStatus}`);
  
  if (!sourceData.objectNumber) {
    Logger.log("⚠️ 処理中断: 物件ナンバーが空です。");
    if (triggerConfig.type === "date" && range.getValue() !== "" && range.getValue() !== null) {
      SpreadsheetApp.getUi().alert(
        "⚠️ 物件ナンバーが入力されていません", 
        "売上進捗表への転記には物件ナンバーが必要です。\nZohoのデータを確認してください", 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return;
  }

  timer.lap("データ解析完了");
  
  Logger.log(`📂 転記先SS (${C.DESTINATION_SS_ID}) を開いています...`);
  let destSs;
  try {
    destSs = SpreadsheetApp.openById(C.DESTINATION_SS_ID);
  } catch (err) {
    Logger.log(`❌ エラー: 転記先SSが開けません - ${err}`);
    SpreadsheetApp.getUi().alert("転記先のスプレッドシートIDが不正またはアクセスできません。");
    return;
  }
  timer.lap("転記先SS接続");

  const context = { C, destSs, sourceData, range, timer, triggerConfig };

  try {
    switch (triggerConfig.type) {
      case "date":   handleDateTrigger(context);   break;
      case "status": handleStatusTrigger(context); break;
      case "amount": handleAmountTrigger(context); break;
      case "text":   handleTextTrigger(context);   break;
      default:
        Logger.log(`⚠️ 未定義のトリガータイプ: ${triggerConfig.type}`);
    }
  } catch (error) {
    Logger.log(`❌ エラー: ${error.message}`);
    Logger.log(`📜 スタック: ${error.stack}`);
    SpreadsheetApp.getUi().alert(`エラーが発生しました: ${error.message}`);
  }
  
  logComplete(timer);
}

// ------------------------------------------------------------------
// [HANDLER] トリガータイプ別
// ------------------------------------------------------------------

function handleDateTrigger(context) {
  const { C, destSs, sourceData, range, timer, triggerConfig } = context;
  
  const editedCellValue = range.getValue();
  if (triggerConfig.clearOnEmpty && (editedCellValue === "" || editedCellValue === null)) {
    Logger.log(`🗑️ 日付列がクリアされました。転記先のデータを削除します。`);
    clearTargetRow(sourceData.objectNumber, destSs, C.DEST_SHEET_SALES);
    SpreadsheetApp.flush();
    return;
  }
  
  const targetDate = sourceData.salesDate;
  if (!isDate(targetDate)) {
    Logger.log(`⚠️ 処理中断: 日付が無効です。`);
    return;
  }

  const stageInfo = analyzeStage(sourceData.currentStatus);
  const searchMonthString = formatMonthString(targetDate);
  Logger.log(`📅 売上予定日: ${targetDate}, 月度: ${searchMonthString}`);
  Logger.log(`📊 ステージ判定: isFinalStage(A/B)=${stageInfo.isFinalStage}, isOtherDeal(C/D)=${stageInfo.isOtherDeal}`);

  const destSheet = destSs.getSheetByName(C.DEST_SHEET_SALES);
  if (!destSheet) {
    Logger.log(`❌ シート「${C.DEST_SHEET_SALES}」が見つかりません。`);
    SpreadsheetApp.getUi().alert(`転記先に「${C.DEST_SHEET_SALES}」が見つかりません。`);
    return;
  }
  timer.lap("転記先シート取得");

  Logger.log(`🗑️ 既存データをクリア中...`);
  clearTargetRow(sourceData.objectNumber, destSs, C.DEST_SHEET_SALES);
  timer.lap("既存データクリア");

  const destData = destSheet.getDataRange().getValues();
  const cols = C.DESTINATION_COLUMNS;
  timer.lap("転記先データ読み込み");

  let targetRowNumber;
  
  if (stageInfo.isOtherDeal) {
    Logger.log("🔍 行決定: C/Dランク → その他商談案件合計ブロック");
    targetRowNumber = findExistingRowByObjectNo(destData, sourceData.objectNumber, cols);
    if (!targetRowNumber) {
      targetRowNumber = findOtherDealWriteRow(destData, cols);
    }
  } else if (stageInfo.isFinalStage) {
    Logger.log("🔍 行決定: A/Bランク → 月ブロック");
    const monthBlock = findMonthBlock(destData, searchMonthString, cols);
    if (!monthBlock) {
      Logger.log(`❌ 処理中断: 「${searchMonthString}」のブロックが見つかりません。`);
      return;
    }
    Logger.log(`📍 月ブロック範囲: ${monthBlock.start} - ${monthBlock.end}`);
    targetRowNumber = findExistingRowByObjectNo(destData, sourceData.objectNumber, cols);
    if (!targetRowNumber) {
      targetRowNumber = findOrCreateTargetRow(destSheet, destData, monthBlock, cols);
    }
  } else {
    Logger.log(`⚠️ 処理中断: 未定義ステージ (${sourceData.currentStatus})`);
    return;
  }

  if (!targetRowNumber) {
    Logger.log(`❌ 処理中断: 追記行を決定できませんでした。`);
    return;
  }

  Logger.log(`📍 ターゲット行: ${targetRowNumber}`);
  timer.lap("行番号決定");

  const values = buildWriteData(sourceData, { salesMonthString: searchMonthString });
  writeToRow(destSheet, targetRowNumber, values);
  timer.lap("データ書き込み");
  
  SpreadsheetApp.flush();
  timer.lap("flush完了");
}

function handleStatusTrigger(context) {
  const { C, destSs, sourceData, timer } = context;
  const cols = C.DESTINATION_COLUMNS;
  
  const currentStatus = sourceData.currentStatus;
  const oldStatus = sourceData.oldStatus;
  const objectNumber = sourceData.objectNumber;

  Logger.log(`🔄 ステージ変更処理: ${oldStatus} → ${currentStatus}`);

  const currentStage = analyzeStage(currentStatus);
  const oldStage = analyzeStage(oldStatus);

  if (currentStage.isClearTarget) {
    Logger.log(`🗑️ 「${currentStatus}」へ変更 → データクリア`);
    clearTargetRow(objectNumber, destSs, C.DEST_SHEET_SALES);
    SpreadsheetApp.flush();
    return;
  }

  const destSheet = destSs.getSheetByName(C.DEST_SHEET_SALES);
  if (!destSheet) {
    Logger.log(`❌ シートが見つかりません`);
    return;
  }

  // パターン0: A ↔ B
  if (currentStage.isFinalStage && oldStage.isFinalStage && 
      (oldStage.isStageA !== currentStage.isStageA)) {
    Logger.log(`📝 パターン0: A↔B同期 (${oldStatus} → ${currentStatus})`);
    handleSameGroupTransition(destSheet, sourceData, cols, timer);
    return;
  }

  // パターン1: C ↔ D
  if (currentStage.isOtherDeal && oldStage.isOtherDeal) {
    Logger.log(`📝 パターン1: C↔D同期`);
    handleOtherDealSync(destSs, destSheet, sourceData, cols, timer);
    return;
  }

  // パターン2: C/D → A/B
  if (currentStage.isFinalStage && oldStage.isOtherDeal) {
    Logger.log(`📝 パターン2: C/D→A/B移動`);
    handleOtherDealToFinalStage(destSs, destSheet, sourceData, cols, timer);
    return;
  }

  // パターン3: A/B → C/D
  if (currentStage.isOtherDeal && !oldStage.isOtherDeal) {
    Logger.log(`📝 パターン3: A/B→C/D移動`);
    handleFinalStageToOtherDeal(destSs, destSheet, sourceData, cols, oldStage, timer);
    return;
  }
}

function handleAmountTrigger(context) {
  const { C, destSs, sourceData, timer } = context;
  const cols = C.DESTINATION_COLUMNS;
  
  const objectNumber = sourceData.objectNumber;
  const newAmount = sourceData.estimatedAmount || 0;
  Logger.log(`💰 見積金額変更処理: 物件No=${objectNumber}, 新金額=${newAmount}`);

  const stageInfo = analyzeStage(sourceData.currentStatus);

  if (!stageInfo.isFinalStage && !stageInfo.isOtherDeal) {
    Logger.log(`⏭️ 転記対象外のステージ (${sourceData.currentStatus}) のためスキップ`);
    return;
  }

  if (!isDate(sourceData.salesDate)) {
    Logger.log(`⏭️ 売上予定日がないためスキップ`);
    return;
  }

  const destSheet = destSs.getSheetByName(C.DEST_SHEET_SALES);
  if (!destSheet) {
    Logger.log(`❌ シートが見つかりません`);
    return;
  }

  const destData = destSheet.getDataRange().getValues();
  timer.lap("転記先データ読み込み");

  let targetRowNumber;
  if (stageInfo.isOtherDeal) {
    targetRowNumber = findExistingRowInOtherDeal(destData, objectNumber, cols);
  } else {
    targetRowNumber = findExistingRowByObjectNo(destData, objectNumber, cols);
  }

  if (!targetRowNumber) {
    Logger.log(`ℹ️ 転記先に該当行なし（物件No: ${objectNumber}）→ 処理スキップ`);
    return;
  }

  Logger.log(`📍 既存行発見: ${targetRowNumber}`);
  updateExistingRowAmount(destSheet, targetRowNumber, sourceData);
  SpreadsheetApp.flush();
  timer.lap("見積金額更新完了");
  Logger.log(`✅ 見積金額更新完了: ${targetRowNumber}行目`);
}

/**
 * テキスト列編集時の処理
 * 粗利率の場合: STAGE_GROSS_MARGIN_MAPPING / STAGE_GROSS_AMOUNT_MAPPING でステージ別に列を切り替え
 *   A/B → 粗利率:J列、粗利額:L列
 *   C   → 粗利率:J列、粗利額:L列
 *   D   → 粗利率:J列、粗利額:K列
 */
function handleTextTrigger(context) {
  const { C, destSs, sourceData, timer, triggerConfig } = context;

  // 数式の再計算を待つ
  SpreadsheetApp.flush();
  Utilities.sleep(1000); // 数式再計算待ち

  const cols = C.DESTINATION_COLUMNS;
  const objectNumber = sourceData.objectNumber;
  Logger.log(`📝 テキスト変更処理: ${triggerConfig.header}, 物件No=${objectNumber}`);

  const stageInfo = analyzeStage(sourceData.currentStatus);

  if (!stageInfo.isFinalStage && !stageInfo.isOtherDeal) {
    Logger.log(`⏭️ 転記対象外のステージのためスキップ`);
    return;
  }

  if (!isDate(sourceData.salesDate)) {
    Logger.log(`⏭️ 売上予定日がないためスキップ`);
    return;
  }

  const destSheet = destSs.getSheetByName(C.DEST_SHEET_SALES);
  if (!destSheet) {
    Logger.log(`❌ シートが見つかりません`);
    return;
  }

  const destData = destSheet.getDataRange().getValues();
  timer.lap("転記先データ読み込み");

  let targetRowNumber;
  if (stageInfo.isOtherDeal) {
    targetRowNumber = findExistingRowInOtherDeal(destData, objectNumber, cols);
  } else {
    targetRowNumber = findExistingRowByObjectNo(destData, objectNumber, cols);
  }

  // 修正後
  if (!targetRowNumber) {
    Logger.log(`ℹ️ 転記先に該当行なし → 新規作成`);
    const destData2 = destSheet.getDataRange().getValues();
    
    if (stageInfo.isOtherDeal) {
      // C/D: その他商談ブロックに新規追記
      targetRowNumber = findOtherDealWriteRow(destData2, cols);
    } else if (stageInfo.isFinalStage) {
      // A/B: 月ブロックに新規追記
      const searchMonthString = formatMonthString(sourceData.salesDate);
      const monthBlock = findMonthBlock(destData2, searchMonthString, cols);
      if (!monthBlock) {
        Logger.log(`❌ 月ブロック「${searchMonthString}」が見つかりません → スキップ`);
        return;
      }
      targetRowNumber = findOrCreateTargetRow(destSheet, destData2, monthBlock, cols);
    }

    if (!targetRowNumber) {
      Logger.log(`❌ 追記行を決定できません → スキップ`);
      return;
    }

    // 基本データを新規書き込み
    const salesMonthString = formatMonthString(sourceData.salesDate);
    const values = buildWriteData(sourceData, { salesMonthString: salesMonthString });
    writeToRow(destSheet, targetRowNumber, values);
    Logger.log(`📍 新規行作成: ${targetRowNumber}`);
  }

  Logger.log(`📍 既存行発見: ${targetRowNumber}`);

  const triggerKey = triggerConfig.key;
  const mapping = C.TEXT_TRIGGER_MAPPING[triggerKey];

  if (mapping && mapping.sourceKey === "grossMarginRate") {
    // 粗利率の書き込み
    const grossRateColName = C.STAGE_GROSS_MARGIN_MAPPING[sourceData.currentStatus];
    if (grossRateColName && cols[grossRateColName]) {
      const targetCol = cols[grossRateColName];
      const rateValue = sourceData.grossMarginRate || "";
      destSheet.getRange(targetRowNumber, targetCol).setValue(rateValue);
      Logger.log(`📝 粗利率 → ${grossRateColName}(${targetCol}列): ${rateValue}`);
    } else {
      Logger.log(`⚠️ ステージ「${sourceData.currentStatus}」の粗利率列が未定義`);
    }

    // A/Bの場合は売上列(J列)に見積金額を書き込む
    if (stageInfo.isFinalStage && isDate(sourceData.salesDate)) {
      const salesCol = cols.SALES_COL;
      const salesAmount = sourceData.estimatedAmount || 0;
      destSheet.getRange(targetRowNumber, salesCol).setValue(salesAmount);
      Logger.log(`📝 売上(J列): ${salesAmount}`);
    }

    // 粗利額の書き込み
    const grossAmountColName = C.STAGE_GROSS_AMOUNT_MAPPING[sourceData.currentStatus];
    if (grossAmountColName && cols[grossAmountColName]) {
      const extraCol = cols[grossAmountColName];
      const amountValue = sourceData.grossMarginAmount || 0;
      destSheet.getRange(targetRowNumber, extraCol).setValue(amountValue);
      Logger.log(`📝 粗利額 → ${grossAmountColName}(${extraCol}列): ${amountValue}`);
    } else {
      Logger.log(`⚠️ ステージ「${sourceData.currentStatus}」の粗利額列が未定義`);
    }

  } else if (mapping && cols[mapping.destCol]) {
    // ===== 工務店名など通常のテキストトリガー =====
    const targetCol = cols[mapping.destCol];
    const newValue = sourceData[mapping.sourceKey] || "";
    destSheet.getRange(targetRowNumber, targetCol).setValue(newValue);
    Logger.log(`📝 ${mapping.destCol} (${targetCol}列) を更新: ${newValue}`);
  }

  SpreadsheetApp.flush();
  timer.lap("テキスト更新完了");
  Logger.log(`✅ 更新完了: ${targetRowNumber}行目`);
}

// ------------------------------------------------------------------
// [PATTERN] ステージ変更パターン別
// ------------------------------------------------------------------

function handleSameGroupTransition(destSheet, sourceData, cols, timer) {
  const destData = destSheet.getDataRange().getValues();
  timer.lap("転記先データ読み込み");
  const targetRowNumber = findExistingRowByObjectNo(destData, sourceData.objectNumber, cols);
  if (targetRowNumber) {
    Logger.log(`📍 既存行発見: ${targetRowNumber}`);
    updateExistingRowAmount(destSheet, targetRowNumber, sourceData);
    SpreadsheetApp.flush();
    timer.lap("A↔B同期書き込み完了");
    Logger.log(`✅ A↔B同期完了`);
  } else {
    Logger.log(`ℹ️ 既存行なし → 売上予定日入力時に新規作成されます`);
  }
}

function handleOtherDealSync(destSs, destSheet, sourceData, cols, timer) {
  const C = CONFIG_PROJECT;
  Logger.log(`💰 使用金額（ソースシート）: ${sourceData.estimatedAmount}`);
  
  if (!isDate(sourceData.salesDate)) {
    Logger.log(`⚠️ 売上予定日がないためスキップ`);
    return;
  }
  
  const destData = destSheet.getDataRange().getValues();
  timer.lap("転記先データ読み込み");
  
  let targetRowNumber = findExistingRowInOtherDeal(destData, sourceData.objectNumber, cols);
  
  if (targetRowNumber) {
    Logger.log(`📍 既存行更新: ${targetRowNumber}`);
    updateExistingRowAmount(destSheet, targetRowNumber, sourceData);
  } else {
    targetRowNumber = findOtherDealWriteRow(destData, cols);
    if (!targetRowNumber) {
      Logger.log(`❌ 追記行を決定できません`);
      return;
    }
    Logger.log(`📍 新規追記: ${targetRowNumber}`);
    const searchMonthString = formatMonthString(sourceData.salesDate);
    const values = buildWriteData(sourceData, { salesMonthString: searchMonthString });
    writeToRow(destSheet, targetRowNumber, values);
  }
  
  SpreadsheetApp.flush();
  timer.lap("書き込み完了");
}

function handleOtherDealToFinalStage(destSs, destSheet, sourceData, cols, timer) {
  const C = CONFIG_PROJECT;
  const START_COL = cols.STORE;
  const END_COL = cols.SALES_COL;
  const numCols = END_COL - START_COL + 1;
  
  Logger.log(`💰 使用金額（ソースシート）: ${sourceData.estimatedAmount}`);
  
  const destData = destSheet.getDataRange().getValues();
  const existingRowOtherDeal = findExistingRowInOtherDeal(destData, sourceData.objectNumber, cols);
  
  if (existingRowOtherDeal) {
    Logger.log(`🗑️ 合計ブロックの行 ${existingRowOtherDeal} をクリア`);
    destSheet.getRange(existingRowOtherDeal, START_COL, 1, numCols).clearContent();
  }

  if (isDate(sourceData.salesDate)) {
    Logger.log(`📝 月ブロックへ再追記`);
    clearTargetRow(sourceData.objectNumber, destSs, C.DEST_SHEET_SALES);
    writeToMonthBlock(destSs, C.DEST_SHEET_SALES, sourceData.salesDate, sourceData);
  }
  
  SpreadsheetApp.flush();
  timer.lap("書き込み完了");
}

function handleFinalStageToOtherDeal(destSs, destSheet, sourceData, cols, oldStage, timer) {
  const C = CONFIG_PROJECT;
  const START_COL = cols.STORE;
  const END_COL = cols.SALES_COL;
  const numCols = END_COL - START_COL + 1;
  
  Logger.log(`💰 使用金額（ソースシート）: ${sourceData.estimatedAmount}`);
  
  if (oldStage.isFinalStage) {
    Logger.log(`🗑️ 月ブロックをクリア`);
    clearTargetRow(sourceData.objectNumber, destSs, C.DEST_SHEET_SALES);
  }

  if (!isDate(sourceData.salesDate)) {
    Logger.log(`⚠️ 売上予定日がないためスキップ`);
    return;
  }
  
  const destData = destSheet.getDataRange().getValues();
  let targetRowNumber = findExistingRowInOtherDeal(destData, sourceData.objectNumber, cols);
  if (!targetRowNumber) {
    targetRowNumber = findOtherDealWriteRow(destData, cols);
  }
  if (!targetRowNumber) {
    Logger.log(`❌ 追記行を決定できません`);
    return;
  }

  Logger.log(`📍 追記先: ${targetRowNumber}`);
  const searchMonthString = formatMonthString(sourceData.salesDate);
  const values = buildWriteData(sourceData, { salesMonthString: searchMonthString });
  writeToRow(destSheet, targetRowNumber, values);
  
  SpreadsheetApp.flush();
  timer.lap("書き込み完了");
}

// ------------------------------------------------------------------
// [TEST]
// ------------------------------------------------------------------

function testWriteToSheet() {
  const timer = createTimer();
  const C = CONFIG_PROJECT;
  Logger.log("テスト開始");
  const destSs = SpreadsheetApp.openById(C.DESTINATION_SS_ID);
  timer.lap("SS接続");
  const destSheet = destSs.getSheetByName(C.DEST_SHEET_SALES);
  timer.lap("シート取得");
  if (!destSheet) {
    Logger.log("テスト失敗: シートが見つかりません");
    return;
  }
  destSheet.getRange(2, 2).setValue("テスト書き込み");
  timer.lap("書き込み");
  SpreadsheetApp.flush();
  timer.lap("flush");
  const writtenValue = destSheet.getRange(2, 2).getValue();
  Logger.log(`書き込み後の値: ${writtenValue}`);
  destSheet.getRange(2, 2).clearContent();
  SpreadsheetApp.flush();
  Logger.log(`テスト完了 (総時間: ${timer.total()}ms)`);
}

function testConfig() {
  const C = CONFIG_PROJECT;
  Logger.log("=== CONFIG_PROJECT 設定確認 ===");
  Logger.log(`転記先SS ID: ${C.DESTINATION_SS_ID}`);
  Logger.log(`転記先シート: ${C.DEST_SHEET_SALES}`);
  Logger.log(`転記元シート: ${C.SOURCE_SHEET_NAMES.join(", ")}`);
  
  Logger.log("\n--- トリガー列 ---");
  for (const [key, config] of Object.entries(C.TRIGGER_COLUMNS)) {
    Logger.log(`  ${key}: ${config.header} (${config.type})`);
  }
  
  Logger.log("\n--- テキストトリガーマッピング ---");
  for (const [key, mapping] of Object.entries(C.TEXT_TRIGGER_MAPPING)) {
    Logger.log(`  ${key} → destCol: ${mapping.destCol}, sourceKey: ${mapping.sourceKey}`);
  }
  
  Logger.log("\n--- ステージ別見積金額列 ---");
  for (const [stage, col] of Object.entries(C.STAGE_AMOUNT_MAPPING)) {
    Logger.log(`  ${stage} → ${col}`);
  }
  
  Logger.log("\n--- ステージ別粗利率列 ---");
  for (const [stage, col] of Object.entries(C.STAGE_GROSS_MARGIN_MAPPING)) {
    Logger.log(`  ${stage} → ${col}`);
  }
  
  Logger.log("\n--- ステージ別粗利額列 ---");
  for (const [stage, col] of Object.entries(C.STAGE_GROSS_AMOUNT_MAPPING)) {
    Logger.log(`  ${stage} → ${col}`);
  }
  
  Logger.log("\n--- 店舗マップ ---");
  for (const [channel, store] of Object.entries(C.STORE_MAP)) {
    Logger.log(`  ${channel} → ${store}`);
  }
}

// ------------------------------------------------------------------
// [UTILITY] 旧 lib.gs 統合
// ------------------------------------------------------------------

// ==================================================================
// タイマー・ログ関連
// ==================================================================

function createTimer() {
  const startTime = new Date().getTime();
  let lastLap = startTime;
  return {
    lap: function(label) {
      const now = new Date().getTime();
      const lapTime = now - lastLap;
      const totalTime = now - startTime;
      Logger.log(`⏱️ [${label}] ${lapTime}ms (累計: ${totalTime}ms)`);
      lastLap = now;
      return lapTime;
    },
    total: function() {
      return new Date().getTime() - startTime;
    }
  };
}

function logStart(functionName) {
  Logger.log("========================================");
  Logger.log(`🚀 ${functionName}: 処理開始`);
  Logger.log("========================================");
}

function logComplete(timer) {
  Logger.log("========================================");
  Logger.log(`✅ 処理完了 (総時間: ${timer.total()}ms)`);
  Logger.log("========================================");
}

// ==================================================================
// データ変換関連
// ==================================================================

function coerceToDate(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === "number") {
    try {
      const d = new Date((v - 25569) * 86400000);
      return isNaN(d) ? null : d;
    } catch (e) { return null; }
  }
  if (typeof v === "string" && v.trim()) {
    try {
      const d = new Date(v.split(" ")[0].replace(/\//g, "-"));
      return isNaN(d) ? null : d;
    } catch (e) { return null; }
  }
  return null;
}

function coerceToNumber(v) {
  if (v === null || v === "" || v === undefined) return 0;
  const numVal = Number(v);
  return isNaN(numVal) ? 0 : numVal;
}

function isDate(d) {
  return d instanceof Date && !isNaN(d);
}

function formatMonthString(date) {
  if (!isDate(date)) return "";
  return Utilities.formatDate(date, CONFIG_PROJECT.TIMEZONE, CONFIG_PROJECT.MONTH_FORMAT);
}

// ==================================================================
// データ取得関連
// ==================================================================

function extractSourceData(headers, rowValues, e, editedHeader) {
  const C = CONFIG_PROJECT;
  const getIdx = (h) => headers.indexOf(h);
  const getData = (h) => {
    const i = getIdx(h);
    return i !== -1 ? rowValues[i] : null;
  };

  const result = {};

  for (const [key, colDef] of Object.entries(C.SOURCE_COLUMNS)) {
    let value = getData(colDef.header);
    if (colDef.type === "number") {
      value = coerceToNumber(value);
    } else if (colDef.type === "date") {
      value = coerceToDate(value);
    } else if (value !== null) {
      value = String(value || "").trim();
    }
    result[key] = value;
  }

  const statusHeader = C.SOURCE_COLUMNS.STATUS?.header;
  const isStatusEdit = (editedHeader === statusHeader);
  result.currentStatus = String(getData(statusHeader) || "").trim();
  result.oldStatus = isStatusEdit ? String(e?.oldValue || "").trim() : result.currentStatus;

  // エイリアス
  result.objectNumber = result.OBJECT_NUMBER;
  result.customerName = result.CUSTOMER_NAME;
  result.estimatedAmount = result.ESTIMATED_AMOUNT;
  result.salesDate = result.SALES_DATE;
  result.salesRep = result.SALES_REP;
  result.channel = result.CHANNEL;
  result.companyName = result.COMPANY_NAME;
  result.grossMarginRate = result.GROSS_MARGIN_RATE;
  result.grossMarginAmount = result.GROSS_MARGIN_AMOUNT;

  return result;
}

function getTriggerConfig(editedHeader) {
  const C = CONFIG_PROJECT;
  for (const [key, triggerDef] of Object.entries(C.TRIGGER_COLUMNS)) {
    if (triggerDef.header === editedHeader) {
      return { key, ...triggerDef };
    }
  }
  return null;
}

function getStoreName(channelValue) {
  const C = CONFIG_PROJECT;
  return C.STORE_MAP[String(channelValue || "").trim()] || "";
}

// ==================================================================
// ステージ判定関連
// ==================================================================

function analyzeStage(status) {
  const C = CONFIG_PROJECT;
  return {
    isFinalStage: C.STATUS_FINAL_STAGE.includes(status),
    isOtherDeal: C.STATUS_OTHER_DEAL.includes(status),
    isClearTarget: C.STATUS_CLEAR_TARGET.includes(status),
    isStageA: status === C.STAGES.A,
    isStageB: status === C.STAGES.B,
    isStageC: status === C.STAGES.C,
    isStageD: status === C.STAGES.D
  };
}

function getAmountColumnForStage(status) {
  const C = CONFIG_PROJECT;
  return C.STAGE_AMOUNT_MAPPING[status] || null;
}

// ==================================================================
// 行検索関連
// ==================================================================

function findMonthBlock(data, searchMonthString, cols) {
  const colIdx = cols.MONTH - 1;
  let monthHeaderRow = -1;
  for (let i = 0, len = data.length; i < len; i++) {
    if (String(data[i][colIdx]).includes(searchMonthString)) {
      monthHeaderRow = i;
      break;
    }
  }
  if (monthHeaderRow === -1) return null;

  let prevMonthHeaderRow = -1;
  for (let i = monthHeaderRow - 1; i >= 0; i--) {
    const val = String(data[i][colIdx]);
    if (val.includes("年") && val.includes("月")) {
      prevMonthHeaderRow = i;
      break;
    }
  }

  const blockEnd = monthHeaderRow - 1;
  const blockStart = (prevMonthHeaderRow === -1) ? 0 : prevMonthHeaderRow + 1;
  return { start: blockStart, end: blockEnd };
}

function findExistingRowByObjectNo(data, objectNumber, cols) {
  const searchObjectNo = String(objectNumber);
  const colIdx = cols.OBJECT_NO - 1;
  for (let i = 0, len = data.length; i < len; i++) {
    const v = data[i][colIdx];
    if (v !== "" && String(v) === searchObjectNo) {
      return i + 1;
    }
  }
  return null;
}

function findExistingRowInOtherDeal(data, objectNumber, cols) {
  const TOTAL_LABEL = CONFIG_PROJECT.OTHER_DEAL_LABEL;
  const STORE_COL = cols.STORE - 1;

  let totalRowIdx = -1;
  for (let i = 0, len = data.length; i < len; i++) {
    if (String(data[i][STORE_COL]).trim() === TOTAL_LABEL) {
      totalRowIdx = i;
      break;
    }
  }
  if (totalRowIdx === -1) return null;

  const searchObjectNo = String(objectNumber);
  const objColIdx = cols.OBJECT_NO - 1;
  const monthColIdx = cols.MONTH - 1;

  for (let i = totalRowIdx - 1; i >= 0; i--) {
    const v = data[i][objColIdx];
    if (v !== "" && String(v) === searchObjectNo) {
      return i + 1;
    }
    const monthLabel = String(data[i][monthColIdx]);
    if (monthLabel.includes("年") && monthLabel.includes("月")) break;
  }
  return null;
}

function findOtherDealWriteRow(data, cols) {
  const TOTAL_LABEL = CONFIG_PROJECT.OTHER_DEAL_LABEL;
  const STORE_COL = cols.STORE - 1;
  const PROJECT_NAME_COL = cols.PROJECT_NAME - 1;

  let totalRowIdx = -1;
  for (let i = 0, len = data.length; i < len; i++) {
    if (String(data[i][STORE_COL]).trim() === TOTAL_LABEL) {
      totalRowIdx = i;
      break;
    }
  }
  if (totalRowIdx === -1) return null;

  for (let i = totalRowIdx - 1; i >= 0; i--) {
    if (String(data[i][PROJECT_NAME_COL]).trim() !== "") {
      return i + 2;
    }
    const monthLabel = String(data[i][cols.MONTH - 1]);
    if (monthLabel.includes("年") && monthLabel.includes("月")) break;
  }
  return totalRowIdx + 1;
}

function findOrCreateTargetRow(destSheet, data, monthBlock, cols) {
  const RUIKEI_LABEL = CONFIG_PROJECT.RUIKEI_LABEL;
  const START_COL_IDX = cols.STORE - 1;
  const END_COL_IDX = cols.SALES_COL - 1;

  let ruikeiRowIdx = -1;
  for (let i = monthBlock.start; i <= monthBlock.end; i++) {
    if (String(data[i][START_COL_IDX]).includes(RUIKEI_LABEL)) {
      ruikeiRowIdx = i;
      break;
    }
  }

  if (ruikeiRowIdx !== -1) {
    const searchStartIdx = ruikeiRowIdx + 1;
    let nextMonthHeaderIdx = data.length;
    for (let i = searchStartIdx, len = data.length; i < len; i++) {
      const val = String(data[i][cols.MONTH - 1]);
      if (val.includes("年") && val.includes("月")) {
        nextMonthHeaderIdx = i;
        break;
      }
    }
    for (let i = searchStartIdx; i < nextMonthHeaderIdx; i++) {
      if (isRowEmpty(data[i], START_COL_IDX, END_COL_IDX)) {
        Logger.log(`📍 空白行発見: ${i + 1}`);
        return i + 1;
      }
    }
    Logger.log(`📈 空白行なし → 行挿入 (${nextMonthHeaderIdx + 1}の前)`);
    destSheet.insertRowBefore(nextMonthHeaderIdx + 1);
    return nextMonthHeaderIdx + 1;
  }

  if (monthBlock.start <= monthBlock.end) {
    for (let i = monthBlock.start; i <= monthBlock.end; i++) {
      if (isRowEmpty(data[i], START_COL_IDX, END_COL_IDX)) {
        return i + 1;
      }
    }
  }
  return monthBlock.end + 2;
}

function isRowEmpty(rowData, startColIdx, endColIdx) {
  for (let col = startColIdx; col <= endColIdx; col++) {
    const v = rowData[col];
    if (v !== "" && v !== null && v !== undefined) return false;
  }
  return true;
}

// ==================================================================
// シート操作関連
// ==================================================================

function ensureRowExists(sheet, rowNumber) {
  const lastRow = sheet.getMaxRows();
  if (rowNumber > lastRow) {
    Logger.log(`📈 行追加: ${lastRow} → ${rowNumber}`);
    sheet.insertRowsAfter(lastRow, rowNumber - lastRow);
  }
}

function clearTargetRow(objectNumber, destSs, sheetName) {
  const C = CONFIG_PROJECT;
  const cols = C.DESTINATION_COLUMNS;
  const sheet = destSs.getSheetByName(sheetName);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const targetRowNumber = findExistingRowByObjectNo(data, objectNumber, cols);

  if (targetRowNumber) {
    const START_COL = cols.STORE;
    const END_COL = cols.SALES_COL;
    Logger.log(`🗑️ クリア: ${targetRowNumber}行目`);
    sheet.getRange(targetRowNumber, START_COL, 1, END_COL - START_COL + 1).clearContent();
  } else {
    Logger.log(`ℹ️ クリア対象なし (物件No: ${objectNumber})`);
  }
}

// ==================================================================
// データ書き込み関連
// ==================================================================

function buildWriteData(sourceData, options = {}) {
  const C = CONFIG_PROJECT;
  const cols = C.DESTINATION_COLUMNS;
  const stageInfo = analyzeStage(sourceData.currentStatus);

  const START_COL = cols.STORE;
  const END_COL = cols.SALES_COL;
  const numCols = END_COL - START_COL + 1;

  const values = new Array(numCols).fill("");

  values[cols.STORE - START_COL] = getStoreName(sourceData.channel);
  values[cols.SALES_REP_COL - START_COL] = sourceData.salesRep || "";
  values[cols.SALES_MONTH_COL - START_COL] = options.salesMonthString || "";
  values[cols.OBJECT_NO - START_COL] = sourceData.objectNumber || "";
  values[cols.BTB_COMPANY - START_COL] = sourceData.companyName || "";
  values[cols.PROJECT_NAME - START_COL] = sourceData.customerName || "";

  if (options.includeAmount !== false) {
    const amount = sourceData.estimatedAmount || 0;
    const amountColName = getAmountColumnForStage(sourceData.currentStatus);
    if (amountColName && cols[amountColName]) {
      values[cols[amountColName] - START_COL] = amount;
      Logger.log(`📊 ${sourceData.currentStatus} → ${amountColName}に ${amount} を記入`);
    }

    if (isDate(sourceData.salesDate) && stageInfo.isFinalStage) {
      values[cols.SALES_COL - START_COL] = amount;
      Logger.log(`📊 売上列(J列)に ${amount} を記入`);
    }
  }

  return values;
}

function writeToRow(destSheet, rowNumber, values) {
  const C = CONFIG_PROJECT;
  const cols = C.DESTINATION_COLUMNS;
  const START_COL = cols.STORE;
  const numCols = values.length;
  ensureRowExists(destSheet, rowNumber);
  destSheet.getRange(rowNumber, START_COL, 1, numCols).setValues([values]);
  Logger.log(`📝 書き込み完了: ${rowNumber}行目`);
}

function updateExistingRowAmount(destSheet, rowNumber, sourceData) {
  const C = CONFIG_PROJECT;
  const cols = C.DESTINATION_COLUMNS;
  const stageInfo = analyzeStage(sourceData.currentStatus);

  const START_COL = cols.STORE;
  const END_COL = cols.SALES_COL;
  const numCols = END_COL - START_COL + 1;

  const values = destSheet.getRange(rowNumber, START_COL, 1, numCols).getValues()[0];

  values[cols.PLANNED_CONTRACT_COL - START_COL] = "";
  values[cols.CONTRACT_SALES_COL - START_COL] = "";

  if (stageInfo.isFinalStage) {
    values[cols.SALES_COL - START_COL] = "";
  }

  const amount = sourceData.estimatedAmount || 0;
  const amountColName = getAmountColumnForStage(sourceData.currentStatus);
  if (amountColName && cols[amountColName]) {
    values[cols[amountColName] - START_COL] = amount;
    Logger.log(`📊 金額更新: ${sourceData.currentStatus} → ${amountColName}に ${amount}`);
  }

  if (isDate(sourceData.salesDate) && stageInfo.isFinalStage) {
    values[cols.SALES_COL - START_COL] = amount;
  }

  values[cols.STORE - START_COL] = getStoreName(sourceData.channel);
  values[cols.SALES_REP_COL - START_COL] = sourceData.salesRep || "";
  values[cols.PROJECT_NAME - START_COL] = sourceData.customerName || "";

  destSheet.getRange(rowNumber, START_COL, 1, numCols).setValues([values]);
}

function writeToMonthBlock(destSs, sheetName, targetDate, sourceData) {
  const C = CONFIG_PROJECT;
  const cols = C.DESTINATION_COLUMNS;
  const destSheet = destSs.getSheetByName(sheetName);
  if (!destSheet) return;

  const destData = destSheet.getDataRange().getValues();
  const searchMonthString = formatMonthString(targetDate);

  const monthBlock = findMonthBlock(destData, searchMonthString, cols);
  if (!monthBlock) {
    Logger.log(`⚠️ 月ブロック「${searchMonthString}」が見つかりません`);
    return;
  }

  let targetRowNumber = findExistingRowByObjectNo(destData, sourceData.objectNumber, cols);
  if (!targetRowNumber) {
    targetRowNumber = findOrCreateTargetRow(destSheet, destData, monthBlock, cols);
  }

  const values = buildWriteData(sourceData, { salesMonthString: searchMonthString });
  writeToRow(destSheet, targetRowNumber, values);
}