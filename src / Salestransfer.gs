// ==================================================================
// salesTransfer.gs ― 営業進捗管理 → 売上進捗表 への転記（編集トリガー本体）
// ==================================================================
// [INDEX]
// [ENTRY]    handleProjectUpdate
// [HANDLER]  handleDateTrigger / handleStatusTrigger / handleAmountTrigger
//            / handleTextTrigger / handleCostTrigger（原価）
// [PATTERN]  handleSameGroupTransition / handleOtherDealSync
//            / handleOtherDealToFinalStage / handleFinalStageToOtherDeal
// [TEST]     testWriteToSheet / testConfig
// [UTILITY]  ユーティリティ関数群（createTimer 以降）
//
// ※ 設定は config.gs（CONFIG_PROJECT）を参照。年度別ファイルの解決は
//   fiscalYear.gs の resolveDestSsId に委譲（年度ごとに別スプレッドシートへ書き込む）。
//
// ★今回の修正点（2026/06）
//   1. 原価を削除したとき：粗利率・粗利高を「予定値」に戻す（実粗利の残骸を消す）。
//   2. 売上予定日クリア／ステージ→問い合わせ等の削除時：年度ファイル解決の前に
//      「全年度ファイル（DEST_SS_ID_R*）を横断」して物件Noの行を B〜L列まで消す。
//      （売上予定日が空だと年度を特定できず、別年度ファイルのデータが消えない問題への対策）
//   3. clearTargetRow の消去範囲を J列→L列（粗利率・粗利高を含む）に拡張。

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

  // ★★★ 削除系は「年度ファイルの解決」より前に処理する ★★★
  //   売上予定日が空になると年度を特定できず（＝今日の年度＝誤ったファイル）を
  //   見てしまうため、削除時は全年度ファイルを横断して物件Noの行を消す。
  //   （年度ファイルのIDが解決できなくても確実に削除できるようにするのが目的）
  if (triggerConfig.type === "date" && triggerConfig.clearOnEmpty) {
    const cellVal = range.getValue();
    if (cellVal === "" || cellVal === null) {
      Logger.log(`🗑️ 売上予定日クリア → 全年度ファイルを横断して削除`);
      clearTargetRowAllYears(sourceData.objectNumber);
      logComplete(timer);
      return;
    }
  }
  if (triggerConfig.type === "status") {
    const st = analyzeStage(sourceData.currentStatus);
    if (st.isClearTarget) {
      Logger.log(`🗑️ ステージ「${sourceData.currentStatus}」→ 全年度ファイルを横断して削除`);
      clearTargetRowAllYears(sourceData.objectNumber);
      logComplete(timer);
      return;
    }
  }

  // 売上予定日の年度から、書き込み先スプレッドシート（年度別ファイル）を解決
  const destInfo = resolveDestSsId(sourceData.salesDate);
  if (!destInfo.ssId) {
    const msg = `令和${destInfo.reiwaYear}年度の売上進捗表スプレッドシートが未登録です。\n`
      + `スクリプトプロパティ「DEST_SS_ID_R${destInfo.reiwaYear}」にファイルIDを設定してください。`;
    Logger.log(`❌ ${msg}`);
    SpreadsheetApp.getUi().alert(msg);
    return;
  }
  Logger.log(`📅 書き込み先: 令和${destInfo.reiwaYear}年度ファイル / シート=${destInfo.sheetName}`);

  let destSs;
  try {
    destSs = SpreadsheetApp.openById(destInfo.ssId);
  } catch (err) {
    Logger.log(`❌ エラー: 転記先SSが開けません - ${err}`);
    SpreadsheetApp.getUi().alert(
      `令和${destInfo.reiwaYear}年度の転記先ファイル（ID: ${destInfo.ssId}）が開けません。IDを確認してください。`
    );
    return;
  }
  timer.lap("転記先SS接続");

  // 売上予定日トリガー時は、新年度の製造間接費シートを自動作成（無ければ）
  if (triggerConfig.type === "date") {
    ensureMfgSheetForFiscalYear_(destInfo.reiwaYear);
  }

  // 書き込み先シート名は全ファイル共通の固定名（C.DEST_SHEET_SALES）
  const context = { C, destSs, sourceData, range, timer, triggerConfig };

  try {
    switch (triggerConfig.type) {
      case "date":   handleDateTrigger(context);   break;
      case "status": handleStatusTrigger(context); break;
      case "amount": handleAmountTrigger(context); break;
      case "text":   handleTextTrigger(context);   break;
      case "cost":   handleCostTrigger(context);   break;
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

  // ※ 売上予定日のクリア（削除）は handleProjectUpdate 側で全年度横断削除を実行済み。
  //    ここに到達するのは「有効な日付が入力された」ケースのみ。
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

  // 既存データクリア → 全年度横断（年度を跨ぐ日付変更で旧年度に行が残るのを防ぐ）
  Logger.log(`🗑️ 既存データをクリア中...（全年度横断）`);
  clearTargetRowAllYears(sourceData.objectNumber);
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

  // ★行作成時にも予定粗利率・粗利高を書き込む（入力順に依存しないようにするため）
  applyGrossMargin_(destSheet, targetRowNumber, sourceData, cols);
  timer.lap("粗利率・粗利高 書き込み");

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

  // 念のための保険（通常は handleProjectUpdate 側で処理済み）
  if (currentStage.isClearTarget) {
    Logger.log(`🗑️ 「${currentStatus}」へ変更 → 全年度ファイルから削除`);
    clearTargetRowAllYears(objectNumber);
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

  // 見積金額が変わると粗利額（自動計算）も変わるため、再計算を待って粗利率・粗利高も更新
  refreshGrossMarginAmount_(context, sourceData);
  applyGrossMargin_(destSheet, targetRowNumber, sourceData, cols);

  SpreadsheetApp.flush();
  timer.lap("見積金額更新完了");
  Logger.log(`✅ 見積金額更新完了: ${targetRowNumber}行目`);
}

/**
 * テキスト列編集時の処理
 * 粗利率の場合: STAGE_GROSS_MARGIN_MAPPING / STAGE_GROSS_AMOUNT_MAPPING でステージ別に列を切り替え
 *   A/B → 粗利率:K列、粗利額:L列
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
    // 粗利率（予定）の書き込み
    const grossRateColName = C.STAGE_GROSS_MARGIN_MAPPING[sourceData.currentStatus];
    if (grossRateColName && cols[grossRateColName]) {
      const targetCol = cols[grossRateColName];
      const rateValue = sourceData.grossMarginRate || "";
      destSheet.getRange(targetRowNumber, targetCol).setValue(rateValue);
      Logger.log(`📝 予定粗利率 → ${grossRateColName}(${targetCol}列): ${rateValue}`);
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

/**
 * 原価入力／削除時の処理。
 * ・原価あり → 実粗利率＝（見積−原価）/見積、実粗利額＝見積×実粗利率 を算出し、
 *   売上進捗表の粗利率列を実粗利率に、粗利額列を実粗利額に上書き（予定→実）。
 * ・原価が空（削除）→ 粗利率・粗利高を「予定値」に戻す（実粗利の残骸を消す）。★今回追加
 *
 * ※ Q列（実粗利率（自動計算）＝数式）の値を読まず、同じ式をコードで再現する。
 *   これにより数式再計算待ち（flush+sleep）が不要で、確実かつ高速に処理できる。
 */
function handleCostTrigger(context) {
  const { C, destSs, sourceData, timer } = context;
  const cols = C.DESTINATION_COLUMNS;
  const objectNumber = sourceData.objectNumber;

  const amount = coerceToNumber(sourceData.estimatedAmount);  // 見積金額
  const cost   = coerceToNumber(sourceData.cost);             // 原価
  Logger.log(`🧮 原価変更処理: 物件No=${objectNumber}, 見積=${amount}, 原価=${cost}`);

  const stageInfo = analyzeStage(sourceData.currentStatus);
  if (!stageInfo.isFinalStage && !stageInfo.isOtherDeal) {
    Logger.log(`⏭️ 転記対象外のステージ (${sourceData.currentStatus}) のためスキップ`);
    return;
  }
  if (!isDate(sourceData.salesDate)) {
    Logger.log(`⏭️ 売上予定日がないためスキップ（年度ファイルを特定できません）`);
    return;
  }

  const destSheet = destSs.getSheetByName(C.DEST_SHEET_SALES);
  if (!destSheet) {
    Logger.log(`❌ シートが見つかりません`);
    return;
  }

  const destData = destSheet.getDataRange().getValues();
  timer.lap("転記先データ読み込み");

  const targetRowNumber = stageInfo.isOtherDeal
    ? findExistingRowInOtherDeal(destData, objectNumber, cols)
    : findExistingRowByObjectNo(destData, objectNumber, cols);

  if (!targetRowNumber) {
    Logger.log(`ℹ️ 売上進捗表に該当行なし（物件No: ${objectNumber}）→ スキップ`);
    return;
  }
  Logger.log(`📍 既存行発見: ${targetRowNumber}`);

  // ★ 原価が空（削除）または見積0 → 粗利率・粗利高を予定値に戻す
  if (cost <= 0 || amount <= 0) {
    Logger.log(`↩️ 原価が空 → 粗利率・粗利高を予定値に戻します`);
    revertGrossMarginToPlanned_(destSheet, targetRowNumber, sourceData, cols);
    SpreadsheetApp.flush();
    timer.lap("原価クリア→予定値復帰");
    return;
  }

  // 原価あり → 実粗利率・実粗利額に上書き（予定→実）
  const actualRate = roundTo_((amount - cost) / amount, 8);
  const actualGrossAmount = Math.round(amount * actualRate);
  Logger.log(`📊 実粗利率=${actualRate}, 実粗利額=${actualGrossAmount}`);

  const rateColName = C.STAGE_GROSS_MARGIN_MAPPING[sourceData.currentStatus];
  if (rateColName && cols[rateColName]) {
    destSheet.getRange(targetRowNumber, cols[rateColName]).setValue(actualRate);
    Logger.log(`📝 実粗利率 → ${rateColName}(${cols[rateColName]}列): ${actualRate}`);
  } else {
    Logger.log(`⚠️ ステージ「${sourceData.currentStatus}」の粗利率列が未定義`);
  }

  const amtColName = C.STAGE_GROSS_AMOUNT_MAPPING[sourceData.currentStatus];
  if (amtColName && cols[amtColName]) {
    destSheet.getRange(targetRowNumber, cols[amtColName]).setValue(actualGrossAmount);
    Logger.log(`📝 実粗利額 → ${amtColName}(${cols[amtColName]}列): ${actualGrossAmount}`);
  } else {
    Logger.log(`⚠️ ステージ「${sourceData.currentStatus}」の粗利額列が未定義`);
  }

  SpreadsheetApp.flush();
  timer.lap("原価転記完了");
  Logger.log(`✅ 原価転記完了: ${targetRowNumber}行目`);
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
    applyGrossMargin_(destSheet, targetRowNumber, sourceData, cols);
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

  // 金額だけでなく粗利率・粗利高も書き込む（C/Dで空になる問題の対策）
  applyGrossMargin_(destSheet, targetRowNumber, sourceData, cols);

  SpreadsheetApp.flush();
  timer.lap("書き込み完了");
}

function handleOtherDealToFinalStage(destSs, destSheet, sourceData, cols, timer) {
  const C = CONFIG_PROJECT;
  const START_COL = cols.STORE;
  const END_COL = cols.GROSS_MARGIN_AMOUNT;  // ★ L列まで（粗利率・粗利高も消す）
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
  const END_COL = cols.GROSS_MARGIN_AMOUNT;  // ★ L列まで
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
  applyGrossMargin_(destSheet, targetRowNumber, sourceData, cols);

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

  Logger.log("\n--- ステージ別 見積金額列 ---");
  for (const [stage, col] of Object.entries(C.STAGE_AMOUNT_MAPPING)) {
    Logger.log(`  ${stage} → ${col}`);
  }

  Logger.log("\n--- ステージ別 粗利率列 ---");
  for (const [stage, col] of Object.entries(C.STAGE_GROSS_MARGIN_MAPPING)) {
    Logger.log(`  ${stage} → ${col}`);
  }

  Logger.log("\n--- ステージ別 粗利額列 ---");
  for (const [stage, col] of Object.entries(C.STAGE_GROSS_AMOUNT_MAPPING)) {
    Logger.log(`  ${stage} → ${col}`);
  }

  Logger.log("\n--- 店舗マップ ---");
  for (const [channel, store] of Object.entries(C.STORE_MAP)) {
    Logger.log(`  ${channel} → ${store}`);
  }
}

/** ★追加：登録済みの年度ファイル一覧を確認する診断用関数 */
function testRegisteredYearFiles() {
  const list = getAllRegisteredDestSs_();
  Logger.log("=== 登録済み年度ファイル（DEST_SS_ID_R*）===");
  if (list.length === 0) {
    Logger.log("  ❌ 1件も登録されていません。スクリプトプロパティを確認してください。");
    return;
  }
  list.forEach(function (yf) {
    Logger.log(`  R${yf.reiwaYear} → ${yf.ssId}`);
  });
}

// ------------------------------------------------------------------
// [UTILITY]
// ------------------------------------------------------------------

// ==================================================================
// タイマー・ログ関連
// ==================================================================

function createTimer() {
  const startTime = new Date().getTime();
  let lastLap = startTime;
  return {
    lap: function (label) {
      const now = new Date().getTime();
      const lapTime = now - lastLap;
      const totalTime = now - startTime;
      Logger.log(`⏱️ [${label}] ${lapTime}ms (累計: ${totalTime}ms)`);
      lastLap = now;
      return lapTime;
    },
    total: function () {
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

/** 小数を指定桁で四捨五入する（実粗利率の桁合わせに使用） */
function roundTo_(num, digits) {
  const p = Math.pow(10, digits);
  return Math.round(num * p) / p;
}

/**
 * 予定粗利率・粗利高を、ステージ別の列に書き込む。
 * 行作成時（handleDateTrigger）にも呼び、入力順に依存せず粗利率・粗利高が入るようにする。
 * ソース側の値が空のときは上書きしない（既存の実粗利率などを消さないため）。
 */
function applyGrossMargin_(destSheet, targetRow, sourceData, cols) {
  const C = CONFIG_PROJECT;
  const status = sourceData.currentStatus;
  const rate = sourceData.grossMarginRate;        // 予定粗利率
  const amount = sourceData.grossMarginAmount;    // 粗利額（自動計算）

  const isEmpty = function (v) { return v === "" || v === null || v === undefined; };

  // 粗利率（A/B→K列、C/D→J列）
  const rateColName = C.STAGE_GROSS_MARGIN_MAPPING[status];
  if (rateColName && cols[rateColName] && !isEmpty(rate)) {
    destSheet.getRange(targetRow, cols[rateColName]).setValue(rate);
    Logger.log(`📝 予定粗利率 → ${rateColName}(${cols[rateColName]}列): ${rate}`);
  }

  // 粗利高（A/B/C→L列、D→K列）
  const amtColName = C.STAGE_GROSS_AMOUNT_MAPPING[status];
  if (amtColName && cols[amtColName] && !isEmpty(amount)) {
    destSheet.getRange(targetRow, cols[amtColName]).setValue(amount);
    Logger.log(`📝 粗利高 → ${amtColName}(${cols[amtColName]}列): ${amount}`);
  }
}

/**
 * ★今回追加：原価が削除されたとき、粗利率・粗利高を「予定値」に戻す。
 *   粗利率列 ← 予定粗利率（手入力）。空ならクリア。
 *   粗利高列 ← 予定粗利額＝見積金額×予定粗利率（コードで算出。数式の再計算待ち不要）。
 *            予定粗利率か見積金額が空ならクリア。
 *   ※ 数式値（粗利額（自動計算））を読まずコードで計算するのは、原価削除直後の
 *     再計算タイミングに左右されず、確実な値を書くため。
 */
function revertGrossMarginToPlanned_(destSheet, targetRow, sourceData, cols) {
  const C = CONFIG_PROJECT;
  const status = sourceData.currentStatus;
  const isEmpty = function (v) { return v === "" || v === null || v === undefined; };

  const rate = sourceData.grossMarginRate;                  // 予定粗利率（手入力）
  const amount = coerceToNumber(sourceData.estimatedAmount); // 見積金額

  // 粗利率列：予定粗利率があればそれを、なければクリア
  const rateColName = C.STAGE_GROSS_MARGIN_MAPPING[status];
  if (rateColName && cols[rateColName]) {
    const cell = destSheet.getRange(targetRow, cols[rateColName]);
    if (isEmpty(rate)) { cell.clearContent(); } else { cell.setValue(rate); }
    Logger.log(`↩️ 粗利率を予定値へ: ${isEmpty(rate) ? "(クリア)" : rate}`);
  }

  // 粗利高列：予定粗利額＝見積×予定粗利率。どちらか空ならクリア
  const amtColName = C.STAGE_GROSS_AMOUNT_MAPPING[status];
  if (amtColName && cols[amtColName]) {
    const cell = destSheet.getRange(targetRow, cols[amtColName]);
    if (isEmpty(rate) || amount <= 0) {
      cell.clearContent();
      Logger.log(`↩️ 粗利高を予定値へ: (クリア)`);
    } else {
      const plannedAmount = Math.round(amount * coerceToNumber(rate));
      cell.setValue(plannedAmount);
      Logger.log(`↩️ 粗利高を予定値へ: ${plannedAmount}`);
    }
  }
}

/**
 * 見積金額の変更後、ソースの「粗利額（自動計算）」の再計算を待って読み直し、
 * sourceData.grossMarginAmount を最新値に更新する。
 * （予定粗利率は手入力で見積変更の影響を受けないため、粗利額のみ更新）
 */
function refreshGrossMarginAmount_(context, sourceData) {
  SpreadsheetApp.flush();
  Utilities.sleep(1000); // 数式（粗利額（自動計算））の再計算待ち
  const sheet = context.range.getSheet();
  const rowNum = context.range.getRow();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rowValues = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];

  const amtHeader = CONFIG_PROJECT.SOURCE_COLUMNS.GROSS_MARGIN_AMOUNT.header; // "粗利額（自動計算）"
  const idx = headers.indexOf(amtHeader);
  if (idx !== -1) {
    sourceData.grossMarginAmount = rowValues[idx];
    Logger.log(`🔄 粗利額（自動計算）再読込: ${sourceData.grossMarginAmount}`);
  }
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
  result.grossMarginRate = result.GROSS_MARGIN_RATE;       // 予定粗利率
  result.grossMarginAmount = result.GROSS_MARGIN_AMOUNT;
  result.cost = result.COST;                               // 原価
  result.actualGrossMarginRate = result.ACTUAL_GROSS_MARGIN_RATE; // 実粗利率（参考）

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
  // ★ 最初の月ブロック(前月ヘッダーなし)の場合は、0ではなくデータ開始行を起点にする
  const dataStartIdx = (CONFIG_PROJECT.DATA_START_ROW || 4) - 1;  // 0始まりインデックス
  const blockStart = (prevMonthHeaderRow === -1) ? dataStartIdx : prevMonthHeaderRow + 1;
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

/**
 * 指定した1ファイル内の固定名シートから、物件Noの行を消す（B〜L列）。
 * ※ 売上予定日が確定していて年度ファイルが正しく解決できる場面で使う。
 *   削除（売上予定日クリア等）では年度を取り違えるため、clearTargetRowAllYears を使う。
 */
function clearTargetRow(objectNumber, destSs, sheetName) {
  const C = CONFIG_PROJECT;
  const cols = C.DESTINATION_COLUMNS;

  const sheet = destSs.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`ℹ️ クリア: シート「${sheetName}」が見つかりません`);
    return;
  }

  const data = sheet.getDataRange().getValues();
  const targetRowNumber = findExistingRowByObjectNo(data, objectNumber, cols);

  if (targetRowNumber) {
    const START_COL = cols.STORE;
    const END_COL = cols.GROSS_MARGIN_AMOUNT;  // ★ L列まで（粗利率・粗利高も消す）
    Logger.log(`🗑️ クリア: 「${sheetName}」${targetRowNumber}行目 (物件No: ${objectNumber})`);
    sheet.getRange(targetRowNumber, START_COL, 1, END_COL - START_COL + 1).clearContent();
    return;
  }

  Logger.log(`ℹ️ クリア対象なし: 物件No ${objectNumber} は「${sheetName}」に見つかりませんでした`);
}

/**
 * ★今回追加：登録済みの全年度ファイル {reiwaYear, ssId} を年度降順で返す。
 *   スクリプトプロパティのキー "DEST_SS_ID_R{N}"（例 DEST_SS_ID_R7, DEST_SS_ID_R8）を走査する。
 */
function getAllRegisteredDestSs_() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const prefix = FY_CONFIG.DEST_SS_ID_KEY_PREFIX; // "DEST_SS_ID_R"
  const list = [];
  Object.keys(props).forEach(function (key) {
    if (key.indexOf(prefix) !== 0) return;
    const reiwaYear = parseInt(key.substring(prefix.length), 10);
    const ssId = String(props[key] || "").trim();
    if (!isNaN(reiwaYear) && ssId) list.push({ reiwaYear: reiwaYear, ssId: ssId });
  });
  list.sort(function (a, b) { return b.reiwaYear - a.reiwaYear; }); // 新しい年度順
  return list;
}

/**
 * ★今回追加：全年度ファイルを横断して、物件Noの行を消す（B〜L列）。
 *   売上予定日が空でも、データがどの年度ファイルにあっても確実に消すための関数。
 *   物件Noはユニークなので、見つけた時点で打ち切る（通常は1ファイル目で完結＝高速）。
 *   @return {boolean} 1件でも消したら true
 */
function clearTargetRowAllYears(objectNumber) {
  const C = CONFIG_PROJECT;
  const cols = C.DESTINATION_COLUMNS;
  const sheetName = C.DEST_SHEET_SALES;
  const yearFiles = getAllRegisteredDestSs_();

  if (yearFiles.length === 0) {
    Logger.log(`⚠️ 登録済みの年度ファイルがありません（スクリプトプロパティ DEST_SS_ID_R* 未設定）`);
    return false;
  }

  for (let k = 0; k < yearFiles.length; k++) {
    const yf = yearFiles[k];
    let ss;
    try {
      ss = SpreadsheetApp.openById(yf.ssId);
    } catch (err) {
      Logger.log(`⚠️ R${yf.reiwaYear}ファイルが開けません（スキップ）: ${err}`);
      continue;
    }
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;

    const data = sheet.getDataRange().getValues();
    const targetRowNumber = findExistingRowByObjectNo(data, objectNumber, cols);
    if (targetRowNumber) {
      const START_COL = cols.STORE;
      const END_COL = cols.GROSS_MARGIN_AMOUNT;  // L列まで
      sheet.getRange(targetRowNumber, START_COL, 1, END_COL - START_COL + 1).clearContent();
      SpreadsheetApp.flush();
      Logger.log(`🗑️ クリア: R${yf.reiwaYear}「${sheetName}」${targetRowNumber}行目 (物件No: ${objectNumber})`);
      return true; // 物件Noはユニーク → 打ち切り
    }
  }

  Logger.log(`ℹ️ クリア対象なし: 物件No ${objectNumber} はどの年度ファイルにも見つかりませんでした`);
  return false;
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
  applyGrossMargin_(destSheet, targetRowNumber, sourceData, cols);
}