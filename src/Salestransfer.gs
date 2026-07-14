// ==================================================================
// salesTransfer.gs ― 営業進捗管理 → 売上進捗表 への転記（編集トリガー本体）
// ==================================================================
// [INDEX]
// [ENTRY]    handleProjectUpdate
// [HANDLER]  handleDateTrigger / handleStatusTrigger / handleAmountTrigger
//            / handleTextTrigger / handleCostTrigger（原価）
// [HELPER]   resolveOrCreateTargetRow_ / executeTransfer_ / getDestSalesSheet_
//            / writeBaseRow_ / clearRowRange_
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
//
// ★今回の修正点（2026/07）― C/D と一部 A/B が黙って書き込まれない不具合の対策
//   ① handleStatusTrigger を「旧ステージ×新ステージ」の分岐から、現ステージだけで
//      転記先を決める方式に全面書き換え（e.oldValue に依存しない）。
//   ② どのハンドラも resolveOrCreateTargetRow_() で「該当行が無ければ作る」ようにし、
//      「該当行なし → 黙ってスキップ」を廃止（必ず理由をログに残す）。
//   ③ C/D は売上予定日が無くても書き込む（config.gs の
//      REQUIRE_SALES_DATE_FOR_OTHER_DEAL=true で従来動作に戻せる）。A/B は月ブロック
//      特定のため引き続き売上予定日が必須。
//   ④ 旧パターン関数（handleSameGroupTransition / handleOtherDealSync /
//      handleOtherDealToFinalStage / handleFinalStageToOtherDeal）を削除。
//
// ★今回の修正点（2026/07 仕上げ）― 潜在バグ修正 + 構造リファクタ
//   A. handleStatusTrigger の転記直前に clearTargetRowAllYears() を呼び、別年度ファイル
//      に残った旧行（C/D を「今日の年度」で作った後 A/B＋別年度へ変更した等）を掃除。
//   B. デッドコード削除：clearTargetRow / writeToMonthBlock（呼び出し元なし）、
//      A で不要になった removeFromOtherDeal_ / removeFromMonthBlocks_、および
//      それらのみが使っていた findExistingRowInOtherDeal（呼び出し元ゼロ）。
//   C. DRY 化：転記先シート取得を getDestSalesSheet_() に、標準転記
//      （行解決→金額更新→粗利適用→flush）を executeTransfer_() に共通化。
//      クリアは clearRowRange_() 経由に統一。

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
  const { C, destSs, sourceData, timer } = context;
  const cols = C.DESTINATION_COLUMNS;

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

  if (!stageInfo.isFinalStage && !stageInfo.isOtherDeal) {
    Logger.log(`⏭️ 転記対象外ステージ（${sourceData.currentStatus}）→ 書き込みなし`);
    return;
  }

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

  // 行を決定（無ければ作成し、基本データを書き込む）
  const targetRowNumber = resolveOrCreateTargetRow_(destSheet, sourceData, stageInfo, cols);
  if (!targetRowNumber) {
    // 理由（A/B で売上予定日なし・月ブロックなし 等）は resolveOrCreateTargetRow_ 内でログ済み
    Logger.log(`❌ 処理中断: 追記行を決定できませんでした。`);
    return;
  }
  Logger.log(`📍 ターゲット行: ${targetRowNumber}`);
  timer.lap("行番号決定");

  // 金額・売上年月・物件No 等を最新化（新規作成／既存行のどちらでも整合させる）
  updateExistingRowAmount(destSheet, targetRowNumber, sourceData);
  timer.lap("データ書き込み");

  // ★行作成時にも予定粗利率・粗利高を書き込む（入力順に依存しないようにするため）
  applyGrossMargin_(destSheet, targetRowNumber, sourceData, cols);
  timer.lap("粗利率・粗利高 書き込み");

  SpreadsheetApp.flush();
  timer.lap("flush完了");
}

/**
 * ステージ変更時の処理（全面書き換え）。
 *
 * 旧実装は「旧ステージ×新ステージ」の4パターンで分岐し、どこにも当てはまらない
 * 変更（例: E 追客 / 問い合わせ / 失注 / 空欄 → A・B）を黙って捨てていた。さらに
 * 複数セルの貼り付け・一括変更では e.oldValue が取れず oldStatus が空になり、A/B が
 * 全滅していた。
 *
 * 新実装は「旧ステージに依存せず、現ステージだけで転記先を決める」方式:
 *   ・現ステージが A/B → 月ブロックへ。以前 C/D として「その他商談」に居た残骸を掃除。
 *   ・現ステージが C/D → その他商談ブロックへ。以前 A/B として月ブロックに居た残骸を掃除。
 *   ・該当行が無ければ resolveOrCreateTargetRow_() が作る（黙ってスキップしない）。
 *   ・A/B/C/D 以外（削除対象を除く）は転記対象外である旨を必ずログに残す。
 */
function handleStatusTrigger(context) {
  const { C, destSs, sourceData, timer } = context;
  const cols = C.DESTINATION_COLUMNS;

  const currentStatus = sourceData.currentStatus;
  const oldStatus = sourceData.oldStatus;
  const objectNumber = sourceData.objectNumber;

  Logger.log(`🔄 ステージ変更処理: ${oldStatus || "(旧ステージ不明)"} → ${currentStatus}`);

  const currentStage = analyzeStage(currentStatus);

  // 念のための保険（通常は handleProjectUpdate 側で処理済み）
  if (currentStage.isClearTarget) {
    Logger.log(`🗑️ 「${currentStatus}」へ変更 → 全年度ファイルから削除`);
    clearTargetRowAllYears(objectNumber);
    SpreadsheetApp.flush();
    return;
  }

  const destSheet = getDestSalesSheet_(destSs, C.DEST_SHEET_SALES);
  if (!destSheet) return;

  // --- 現ステージ A/B: 月ブロックへ転記 ---
  if (currentStage.isFinalStage) {
    Logger.log(`📝 現ステージ A/B → 月ブロックへ転記（旧ステージは不問）`);
    // ★ 転記前に全年度ファイルから旧行を掃除（別年度ファイルに残った C/D 行なども消す）。
    //   現年度ファイル内の「その他商談ブロック」の残骸もこれで消える。
    clearTargetRowAllYears(objectNumber);

    const row = executeTransfer_(destSheet, sourceData, currentStage, cols, {});
    if (!row) return;  // A/B は売上予定日が必須。理由は resolveOrCreateTargetRow_ 内でログ済み
    timer.lap("A/B転記完了");
    Logger.log(`✅ A/B転記完了: ${row}行目`);
    return;
  }

  // --- 現ステージ C/D: その他商談ブロックへ転記 ---
  if (currentStage.isOtherDeal) {
    Logger.log(`📝 現ステージ C/D → その他商談ブロックへ転記（旧ステージは不問）`);
    // ★ 転記前に全年度ファイルから旧行を掃除（別年度ファイルに残った A/B 行なども消す）。
    //   現年度ファイル内の「月ブロック」の残骸もこれで消える。
    clearTargetRowAllYears(objectNumber);

    const row = executeTransfer_(destSheet, sourceData, currentStage, cols, {});
    if (!row) return;  // REQUIRE_SALES_DATE_FOR_OTHER_DEAL=true で売上予定日なし等。理由はログ済み
    timer.lap("C/D転記完了");
    Logger.log(`✅ C/D転記完了: ${row}行目`);
    return;
  }

  // --- それ以外（E 追客 など）: 転記対象外。黙って return せず理由を残す ---
  Logger.log(`⏭️ 転記対象外ステージ「${currentStatus}」→ 書き込みなし（A/B/C/D と削除対象「${C.STATUS_CLEAR_TARGET.join("・")}」以外は転記しません）`);
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

  const destSheet = getDestSalesSheet_(destSs, C.DEST_SHEET_SALES);
  if (!destSheet) return;
  timer.lap("転記先シート取得");

  // 該当行が無ければ作り、金額・粗利率・粗利高を更新（refreshGross=true で粗利額を再計算）
  const row = executeTransfer_(destSheet, sourceData, stageInfo, cols, { refreshGross: true, context: context });
  if (!row) return;  // 理由は resolveOrCreateTargetRow_ 内でログ済み

  timer.lap("見積金額更新完了");
  Logger.log(`✅ 見積金額更新完了: ${row}行目`);
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

  const destSheet = getDestSalesSheet_(destSs, C.DEST_SHEET_SALES);
  if (!destSheet) return;
  timer.lap("転記先シート取得");

  // 該当行が無ければ作る（C/D は売上予定日なしでも作成、A/B は売上予定日必須）
  const targetRowNumber = resolveOrCreateTargetRow_(destSheet, sourceData, stageInfo, cols);
  if (!targetRowNumber) {
    // 理由は resolveOrCreateTargetRow_ 内でログ済み
    return;
  }
  Logger.log(`📍 対象行: ${targetRowNumber}`);

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

  const destSheet = getDestSalesSheet_(destSs, C.DEST_SHEET_SALES);
  if (!destSheet) return;
  timer.lap("転記先シート取得");

  // 該当行が無ければ作る（C/D は売上予定日なしでも作成、A/B は売上予定日必須）
  const targetRowNumber = resolveOrCreateTargetRow_(destSheet, sourceData, stageInfo, cols);
  if (!targetRowNumber) {
    // 理由は resolveOrCreateTargetRow_ 内でログ済み
    return;
  }
  Logger.log(`📍 対象行: ${targetRowNumber}`);

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
// [HELPER] 行の解決・作成／掃除（ステージ変更・各トリガーで共通利用）
// ------------------------------------------------------------------

/**
 * 転記先の行を解決する。既存行があればその行番号を返す。無ければ:
 *   ・C/D → findOtherDealWriteRow() で「その他商談ブロック」に追記行を確保
 *   ・A/B → findMonthBlock() + findOrCreateTargetRow() で「月ブロック」に追記行を確保
 * を行い、writeBaseRow_() で基本データを書き込んでから行番号を返す。
 *
 * A/B は売上予定日が無いと月ブロックを特定できないため、無ければ null を返し理由をログする。
 * C/D は既定（REQUIRE_SALES_DATE_FOR_OTHER_DEAL=false）では売上予定日が空でも書き込む
 * （年度ファイルは呼び出し元で「今日の年度」に解決済み＝警告をログする）。true にすると従来動作。
 *
 * @return {number|null} 行番号（1始まり）。決定できなければ null（理由はログ済み）。
 */
function resolveOrCreateTargetRow_(destSheet, sourceData, stageInfo, cols) {
  const C = CONFIG_PROJECT;
  const objectNumber = sourceData.objectNumber;
  const destData = destSheet.getDataRange().getValues();

  // 既存行があればそれを使う（物件Noはユニーク）
  const existing = findExistingRowByObjectNo(destData, objectNumber, cols);
  if (existing) {
    Logger.log(`📍 既存行発見: ${existing}行目 (物件No: ${objectNumber})`);
    return existing;
  }

  // --- 新規作成 ---
  if (stageInfo.isOtherDeal) {
    // C/D: その他商談ブロック
    if (C.REQUIRE_SALES_DATE_FOR_OTHER_DEAL && !isDate(sourceData.salesDate)) {
      Logger.log(`❌ C/D 売上予定日が必須設定（REQUIRE_SALES_DATE_FOR_OTHER_DEAL=true）で空 → 書き込みスキップ`);
      return null;
    }
    if (!isDate(sourceData.salesDate)) {
      Logger.log(`⚠️ C/D 売上予定日が空 → 年度ファイルは「今日の年度」で解決されています（売上年月は空欄で登録）`);
    }
    const row = findOtherDealWriteRow(destData, cols);
    if (!row) {
      Logger.log(`❌ その他商談ブロックの追記行を決定できません → 書き込みスキップ`);
      return null;
    }
    writeBaseRow_(destSheet, row, sourceData);
    Logger.log(`📍 その他商談ブロックに新規行作成: ${row}行目`);
    return row;
  }

  if (stageInfo.isFinalStage) {
    // A/B: 月ブロック（売上予定日が必須）
    if (!isDate(sourceData.salesDate)) {
      Logger.log(`❌ A/B は売上予定日が無いと月ブロックを特定できません → 書き込みスキップ（売上予定日を入力してください）`);
      return null;
    }
    const searchMonthString = formatMonthString(sourceData.salesDate);
    const monthBlock = findMonthBlock(destData, searchMonthString, cols);
    if (!monthBlock) {
      Logger.log(`❌ 月ブロック「${searchMonthString}」が見つかりません → 書き込みスキップ`);
      return null;
    }
    const row = findOrCreateTargetRow(destSheet, destData, monthBlock, cols);
    if (!row) {
      Logger.log(`❌ 月ブロックの追記行を決定できません → 書き込みスキップ`);
      return null;
    }
    writeBaseRow_(destSheet, row, sourceData);
    Logger.log(`📍 月ブロックに新規行作成: ${row}行目`);
    return row;
  }

  Logger.log(`⚠️ 転記対象外ステージ（${sourceData.currentStatus}）→ 行を作成しません`);
  return null;
}

/**
 * 基本データ（店舗・担当・売上年月・物件No・工務店名・案件名・見積金額）を1行に書き込む。
 * buildWriteData + writeToRow のラッパー。売上予定日が無ければ売上年月は空欄。
 */
function writeBaseRow_(destSheet, row, sourceData) {
  const salesMonthString = isDate(sourceData.salesDate) ? formatMonthString(sourceData.salesDate) : "";
  const values = buildWriteData(sourceData, { salesMonthString: salesMonthString });
  writeToRow(destSheet, row, values);
}

/** 指定行の B〜L 列（STORE 〜 GROSS_MARGIN_AMOUNT）をクリアする。 */
function clearRowRange_(sheet, row, cols) {
  const START_COL = cols.STORE;                // B
  const END_COL = cols.GROSS_MARGIN_AMOUNT;    // L
  sheet.getRange(row, START_COL, 1, END_COL - START_COL + 1).clearContent();
}

/**
 * 転記先の固定名シートを取得する共通処理。無ければログして null を返す。
 * （各ハンドラで重複していたシート取得＋存在チェックを1箇所に集約）
 */
function getDestSalesSheet_(destSs, sheetName) {
  const sheet = destSs.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`❌ シート「${sheetName}」が見つかりません`);
  }
  return sheet;
}

/**
 * 標準転記の共通処理：
 *   resolveOrCreateTargetRow_（行解決／作成）→ updateExistingRowAmount（金額・売上年月等）
 *   →（任意）refreshGrossMarginAmount_（粗利額の再計算待ち）→ applyGrossMargin_（粗利率・粗利高）
 *   → SpreadsheetApp.flush()
 * を1関数にまとめる。handleStatusTrigger / handleAmountTrigger の重複を削減する。
 *
 * @param {Object} options
 *        options.refreshGross {boolean} 見積金額変更時など、粗利額（自動計算）を再読込してから適用する
 *        options.context      {Object}  refreshGross=true のとき必須（range を含むトリガーコンテキスト）
 * @return {number|null} 転記した行番号。行を決定できなければ null（理由は resolveOrCreateTargetRow_ がログ済み）。
 */
function executeTransfer_(destSheet, sourceData, stageInfo, cols, options) {
  const opts = options || {};
  const row = resolveOrCreateTargetRow_(destSheet, sourceData, stageInfo, cols);
  if (!row) {
    SpreadsheetApp.flush();  // クリア等の未反映書き込みを確定させてから抜ける
    return null;
  }
  updateExistingRowAmount(destSheet, row, sourceData);
  if (opts.refreshGross && opts.context) {
    // 見積金額が変わると粗利額（自動計算）も変わるため、再計算を待って読み直す
    refreshGrossMarginAmount_(opts.context, sourceData);
  }
  applyGrossMargin_(destSheet, row, sourceData, cols);
  SpreadsheetApp.flush();
  return row;
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
 * 登録済みの全年度ファイル {reiwaYear, ssId} を年度降順で返す。
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
      clearRowRange_(sheet, targetRowNumber, cols);  // B〜L列をクリア（クリア処理を1経路に統一）
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

  // ★ 物件No と 売上年月 も更新する。
  //   ステージ変更だけで（売上予定日入力より前に）行を作った場合に、これらが空のまま
  //   残るのを防ぐ。売上予定日が無いとき（C/D等）は売上年月を上書きしない。
  values[cols.OBJECT_NO - START_COL] = sourceData.objectNumber || "";
  if (isDate(sourceData.salesDate)) {
    values[cols.SALES_MONTH_COL - START_COL] = formatMonthString(sourceData.salesDate);
  }

  destSheet.getRange(rowNumber, START_COL, 1, numCols).setValues([values]);
}