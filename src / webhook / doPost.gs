/**
 * ==========================================
 * 全体設定（config.gs）
 * ==========================================
 * 設定変更はこのファイルだけで完結するように設計
 * シートの構成が変わった場合は、COL_NAMESの値（右側の文字列）を
 * シートのヘッダー行と完全一致するように修正してください。
 */
const CONFIG = {

  // ----------------------------------------
  // スクリプトプロパティ関連
  // ----------------------------------------
  // スプレッドシートIDを保存しているプロパティのキー名
  PROP_KEY_SS_ID: "SPREADSHEET_ID",


  // ----------------------------------------
  // シート名
  // ----------------------------------------
  // Webhookデータを書き込む対象のシート名
  SHEET_NAME: "営業進捗管理",


  // ----------------------------------------
  // 列名（スプレッドシートのヘッダー行と完全一致させる）
  // ----------------------------------------
  COL_NAMES: {
    // --- 空白判定対象の6列（新規追加行を決める基準） ---
    PROPERTY_NUMBER:        "物件ナンバー",
    CHANNEL:                "チャネル",
    VISIT_YEAR:             "来場年",
    VISIT_DATE:             "来場日",
    STAGE:                  "ステージ",
    PROJECT_NAME:           "案件名",

    // --- 工務店・担当者関連 ---
    BTB_COMPANY:            "工務店（B2B）",
    COMPANY_NAME:           "工務店名",
    PERSON_IN_CHARGE_1:     "担当者1",
    PERSON_IN_CHARGE_2:     "担当者2",

    // --- 金額・粗利関連 ---
    ESTIMATE_AMOUNT:        "見積金額",
    GROSS_MARGIN_RATE:      "粗利率",
    GROSS_MARGIN_AMOUNT:    "粗利額（自動計算）",
    SALES_SCHEDULED_DATE:   "売上予定日",

    // --- アプローチ・架電関連 ---
    APPROACH:               "アプローチ",
    CALL_SCHEDULED_DATE:    "架電予定日",
    CALL_RESULT:            "架電結果",
    TRIGGER:                "きっかけ",

    // --- 問い合わせ関連 ---
    HAS_PAST:               "過去の資料請求",
    INQUIRY_ROUTE:          "問い合わせ経路",
    ACCURACY:               "確度",

    // --- 顧客情報 ---
    CURRENT_ADDRESS:        "現住所",
    EMAIL:                  "メールアドレス",
    PHONE:                  "電話番号",

    // --- 施工関連 ---
    CONSTRUCTION_PLACE_1:   "施工場所1",
    CONSTRUCTION_PLACE_2:   "施工場所2",
    CONTENT:                "内容",
    INFO:                   "情報",

    // --- 受注・失注 ---
    ORDER_REASON:           "受注・失注理由",
    ORDER_REASON_REFERENCE: "受注・失注理由（参考）"
  }
};

/**
 * ==========================================
 * Webhookデータ処理関数群
 * ==========================================
 */

/**
 * Webhookの生データを統一されたキーに正規化する
 * @param {Object} e - Webhookイベントオブジェクト
 * @returns {Object} 正規化されたデータオブジェクト
 */
function getNormalizedData(e) {
  if (!e || !e.parameter) {
    throw new Error("Webhookパラメータが空です。");
  }

  const p = e.parameter;

  // 電話番号: mobile_phoneを優先、空ならphone
  const mobilePhone = (p.mobile_phone || "").trim();
  const phone = (p.phone || "").trim();
  const phoneNumber = mobilePhone !== "" ? mobilePhone : phone;

  // 来場日: 様々なフォーマットに対応して正規化
  const visitDateRaw = (p.inquiry_date || "").trim();
  const visitDateInfo = parseFlexibleDate(visitDateRaw);

  return {
    // --- 必須項目 ---
    propertyNumber:  (p.property_number || "").trim(),
    customerName:    (p.customer_name || "").trim(),
    channel:         (p.showroom || "").trim(),

    // --- 日付関連 ---
    visitDate:       visitDateInfo.formattedDate,  // "YYYY/MM/DD" or 空
    visitYear:       visitDateInfo.year,           // "YYYY年" or 空

    // --- ステージ・工務店 ---
    stage:           (p.stage || "").trim(),
    btbCompany:      (p.sorting || "").trim(),
    companyName:     (p.company_name || "").trim(),

    // --- 顧客情報 ---
    address:         (p.address || "").trim(),
    email:           (p.mail || "").trim(),
    phoneNumber:     phoneNumber,

    // --- 問い合わせ情報 ---
    inquiryClass:    (p.inquiry_class || "").trim(),
    foundOut:        (p.found_out || "").trim(),
    inquiryDetail:   (p.inquiry_detail || "").trim(),
    contactDetail:   (p.contact_detail || "").trim()
  };
}


/**
 * 様々なフォーマットの日付文字列をDateオブジェクト経由で正規化する
 * 対応フォーマット例:
 *   - 2025-11-15
 *   - 2025/11/15
 *   - 2025年11月15日
 *   - 2025.11.15
 *   - November 15, 2025
 *   - 等、Date()コンストラクタで解釈可能なもの
 * @param {string} dateStr - 日付文字列
 * @returns {Object} { formattedDate: "YYYY/MM/DD", year: "YYYY年" }（解釈不能なら両方空）
 */
function parseFlexibleDate(dateStr) {
  if (!dateStr) return { formattedDate: "", year: "" };

  // 全角数字を半角に変換
  let normalized = dateStr.replace(/[０-９]/g, (s) =>
    String.fromCharCode(s.charCodeAt(0) - 0xFEE0)
  );

  // 日本語フォーマット（"2025年11月15日"）を "2025/11/15" に変換
  normalized = normalized
    .replace(/年/g, "/")
    .replace(/月/g, "/")
    .replace(/日/g, "")
    .replace(/\./g, "/")    // "2025.11.15" → "2025/11/15"
    .replace(/-/g, "/")     // "2025-11-15" → "2025/11/15"
    .trim();

  const date = new Date(normalized);
  if (isNaN(date.getTime())) {
    Logger.log(`⚠️ 日付の解釈に失敗: "${dateStr}"`);
    return { formattedDate: "", year: "" };
  }

  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");

  return {
    formattedDate: `${y}/${m}/${d}`,
    year: `${y}年`
  };
}


/**
 * 正規化データをシートのヘッダー名キーのオブジェクトに変換する
 * シートに書き込む値だけを含む（マッピング対象外の列は含まない=上書き時に既存値を保持）
 * @param {Object} webhookData - 正規化済みデータ
 * @returns {Object} { "シートヘッダー名": 値, ... }
 */
function formatWriteData(webhookData) {
  const C = CONFIG.COL_NAMES;
  const result = {};

  // 値が空でないものだけセット（空値で既存データを上書きしないため）
  // ただし、新規行の場合はそのまま空欄になるので問題なし
  const setIfPresent = (key, value) => {
    if (value !== undefined && value !== null && value !== "") {
      result[key] = value;
    }
  };

  setIfPresent(C.PROPERTY_NUMBER, webhookData.propertyNumber);
  setIfPresent(C.CHANNEL,         webhookData.channel);
  setIfPresent(C.VISIT_YEAR,      webhookData.visitYear);
  setIfPresent(C.VISIT_DATE,      webhookData.visitDate);
  setIfPresent(C.STAGE,           webhookData.stage);
  setIfPresent(C.BTB_COMPANY,     webhookData.btbCompany);
  setIfPresent(C.COMPANY_NAME,    webhookData.companyName);
  setIfPresent(C.PROJECT_NAME,    webhookData.customerName);  // 案件名=顧客名
  setIfPresent(C.TRIGGER,         webhookData.foundOut);
  setIfPresent(C.INQUIRY_ROUTE,   webhookData.inquiryClass);
  setIfPresent(C.CURRENT_ADDRESS, webhookData.address);
  setIfPresent(C.EMAIL,           webhookData.email);
  setIfPresent(C.PHONE,           webhookData.phoneNumber);
  setIfPresent(C.CONTENT,         webhookData.inquiryDetail);
  setIfPresent(C.INFO,            webhookData.contactDetail);

  return result;
}


/**
 * 過去の資料請求の有無と上書き対象行を判定する
 * 判定ロジック:
 *   - 物件ナンバーが一致する行があれば → その行を上書き対象とする
 *   - 顧客名が一致する行が（上書き対象以外に）あれば → 過去の資料請求「有り」
 * @param {Array<Array>} sheetData - シート全体のデータ
 * @param {Object} headerMap - ヘッダー名 → 列インデックスのマップ
 * @param {Object} webhookData - 正規化済みデータ
 * @returns {Object} { updateRowIndex: number|null, hasPastInquiry: boolean }
 *   updateRowIndex: 上書き対象の行番号(1-based)、新規追加の場合はnull
 *   hasPastInquiry: 過去に同じ顧客名の問い合わせがあればtrue
 */
function checkPastInquiry(sheetData, headerMap, webhookData) {
  const C = CONFIG.COL_NAMES;
  const propertyNumberColIdx = headerMap[C.PROPERTY_NUMBER];
  const projectNameColIdx    = headerMap[C.PROJECT_NAME];

  if (propertyNumberColIdx === undefined) {
    Logger.log(`⚠️ 「${C.PROPERTY_NUMBER}」列が見つかりません。`);
  }
  if (projectNameColIdx === undefined) {
    Logger.log(`⚠️ 「${C.PROJECT_NAME}」列が見つかりません。`);
  }

  let updateRowIndex = null;
  let pastInquiryCount = 0;  // 同じ顧客名の行数をカウント

  const inputPropertyNumber = String(webhookData.propertyNumber || "").trim();
  const inputCustomerName   = String(webhookData.customerName || "").trim();

  // ヘッダー行(0行目)をスキップして1行目から走査
  for (let i = 1; i < sheetData.length; i++) {
    const currentRow = sheetData[i];  // ★ 変数名を row → currentRow に変更

    // 物件ナンバーの照合 → 上書き対象を決定
    if (propertyNumberColIdx !== undefined && inputPropertyNumber !== "") {
      const rowPropertyNumber = String(currentRow[propertyNumberColIdx] || "").trim();
      if (rowPropertyNumber !== "" && rowPropertyNumber === inputPropertyNumber) {
        updateRowIndex = i + 1;  // 1-basedに変換
        Logger.log(`🎯 物件ナンバー一致: ${i + 1}行目 (${rowPropertyNumber})`);
      }
    }

    // 顧客名の照合 → カウントだけしておく
    if (projectNameColIdx !== undefined && inputCustomerName !== "") {
      const rowCustomerName = String(currentRow[projectNameColIdx] || "").trim();
      if (rowCustomerName !== "" && rowCustomerName === inputCustomerName) {
        pastInquiryCount++;
      }
    }
  }

  // 過去の資料請求判定:
  // - 上書き対象がある場合: 上書き対象以外に同名行があれば「有り」
  //   → 上書き対象は自分自身なので、count >= 2 で「有り」
  // - 上書き対象がない場合: 同名行が1つでもあれば「有り」
  let hasPastInquiry;
  if (updateRowIndex !== null) {
    hasPastInquiry = pastInquiryCount >= 2;
  } else {
    hasPastInquiry = pastInquiryCount >= 1;
  }

  Logger.log(`📊 照合結果: 上書き行=${updateRowIndex || 'なし'}, 顧客名一致数=${pastInquiryCount}, 過去の資料請求=${hasPastInquiry ? '有り' : '無し'}`);
  return { updateRowIndex, hasPastInquiry };
}

/**
 * 処理時間計測用ヘルパー
 */
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

/**
 * 指定された複数の列がすべて空白である最初の行のインデックスを返す
 * @param {Array<Array>} sheetData - シート全体のデータ（2次元配列、1行目はヘッダー）
 * @param {Object} headerMap - ヘッダー名 → 列インデックスのマップ
 * @param {Array<string>} checkColumnNames - 空白判定対象の列名の配列
 * @returns {number} 0-basedの行インデックス（書き込み対象行）
 */
function findFirstEmptyRow(sheetData, headerMap, checkColumnNames) {
  // 判定対象列のインデックスを取得（ヘッダーに存在する列のみ）
  const checkColIndexes = checkColumnNames
    .map(name => {
      const idx = headerMap[name];
      if (idx === undefined) {
        Logger.log(`⚠️ 判定対象列「${name}」がヘッダーに見つかりません。スキップします。`);
      }
      return idx;
    })
    .filter(idx => idx !== undefined);

  if (checkColIndexes.length === 0) {
    throw new Error("空白判定対象の列が1つも見つかりませんでした。");
  }

  // ヘッダー(0行目)をスキップして1行目から走査
  for (let i = 1; i < sheetData.length; i++) {
    const row = sheetData[i];
    const allEmpty = checkColIndexes.every(colIdx => {
      const cell = row[colIdx];
      return cell === "" || cell === null || cell === undefined;
    });
    if (allEmpty) {
      Logger.log(`🔎 空白行を検出: ${i + 1}行目（指定列がすべて空白）`);
      return i;
    }
  }

  // すべての行が埋まっている場合は最終行の次
  Logger.log(`🔎 空白行なし。最終行の次（${sheetData.length + 1}行目）に追加します。`);
  return sheetData.length;
}

/**
 * Webhookリクエストを受信してスプレッドシートにデータを書き込むメイン関数
 * @param {Object} e - Webhookイベントオブジェクト
 * @returns {GoogleAppsScript.Content.TextOutput} JSON形式のレスポンス
 */
function doPost(e) {
  const timer = createTimer();
  
  Logger.log("========================================");
  Logger.log("🚀 doPost: Webhook受信開始");
  Logger.log("========================================");
  
  try {
    // リクエストデータのログ出力（デバッグ用）
    if (e && e.postData) {
      Logger.log(`📥 受信データタイプ: ${e.postData.type}`);
      Logger.log(`📥 受信データ長: ${e.postData.length} bytes`);
    }
    
    // スプレッドシートIDを取得
    Logger.log("📂 スプレッドシートIDを取得中...");
    const ssId = PropertiesService.getScriptProperties().getProperty(CONFIG.PROP_KEY_SS_ID);
    if (!ssId) throw new Error("スプレッドシートIDがプロパティに設定されていません。");
    Logger.log(`📂 スプレッドシートID: ${ssId.substring(0, 10)}...`);
    
    const ss = SpreadsheetApp.openById(ssId);
    timer.lap("スプレッドシート接続");

    // webhookデータを統一されたキーに変換
    Logger.log("🔄 Webhookデータを正規化中...");
    const webhookData = getNormalizedData(e);
    Logger.log(`📦 正規化データ: 顧客名=${webhookData.customerName}, 物件No=${webhookData.propertyNumber}`);
    timer.lap("データ正規化");

    // 書き込むシート（ターゲットシート）を取得
    Logger.log(`📋 シート「${CONFIG.SHEET_NAME}」を取得中...`);
    const targetSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!targetSheet) throw new Error("書き込む対象のシートが見つかりませんでした。");
    Logger.log(`📋 シート取得完了`);
    timer.lap("シート取得");

    // ターゲットシートのデータを2次元配列で取得
    Logger.log("📖 シートデータを読み込み中...");
    const targetSheetData = targetSheet.getDataRange().getValues();
    if (targetSheetData.length === 0) throw new Error("対象シートにヘッダー行がありません。");
    Logger.log(`📖 読み込み完了: ${targetSheetData.length}行 × ${targetSheetData[0].length}列`);
    timer.lap("シートデータ読み込み");

    // targetSheetDataからヘッダー行を取得
    const headers = targetSheetData[0];
    const headerMap = headers.reduce((map, header, i) => { map[header] = i; return map; }, {});
    Logger.log(`📋 ヘッダー列数: ${headers.length}`);

    // 書き込み用データ本体を作成
    Logger.log("📝 書き込みデータを作成中...");
    const writeDataMap = formatWriteData(webhookData);
    Logger.log(`📝 書き込みデータ項目数: ${Object.keys(writeDataMap).length}`);
    timer.lap("書き込みデータ作成");

    // 過去の資料請求の有無を判断
    Logger.log("🔍 過去の資料請求を確認中...");
    const { updateRowIndex, hasPastInquiry } = checkPastInquiry(targetSheetData, headerMap, webhookData);
    const hasPastStatus = hasPastInquiry ? '有り' : '無し';
    Logger.log(`🔍 チェック完了: 過去の資料請求=${hasPastStatus}, 更新対象行=${updateRowIndex || 'なし'}`);
    timer.lap("過去資料請求チェック");

    // updateRowIndex があれば UPDATE、なければ APPEND
    let rowValues;
    let targetRowIdx;
    let actionType;

    if (updateRowIndex !== null) {
      // ■■■ 上書き処理 ■■■
      actionType = "UPDATE";
      targetRowIdx = updateRowIndex - 1;  // 1-based → 0-based に変換
      rowValues = [...targetSheetData[targetRowIdx]];
      Logger.log(`✏️ [${webhookData.customerName}] 物件No. ${webhookData.propertyNumber} が一致。行 ${updateRowIndex} を上書き。`);
    } else {
      // ■■■ 新規追加処理 ■■■
      actionType = "APPEND";

      // 指定6列がすべて空白の最初の行を探す
      const checkColumns = [
        CONFIG.COL_NAMES.PROPERTY_NUMBER,  // 物件ナンバー
        CONFIG.COL_NAMES.CHANNEL,          // チャネル
        CONFIG.COL_NAMES.VISIT_YEAR,       // 来場年
        CONFIG.COL_NAMES.VISIT_DATE,       // 来場日
        CONFIG.COL_NAMES.STAGE,            // ステージ
        CONFIG.COL_NAMES.PROJECT_NAME      // 案件名
      ];
      targetRowIdx = findFirstEmptyRow(targetSheetData, headerMap, checkColumns);

      // 既存行（途中の空白行）の場合は既存の値を保持、最終行の次なら新規配列
      if (targetRowIdx < targetSheetData.length) {
        rowValues = [...targetSheetData[targetRowIdx]];
      } else {
        rowValues = new Array(headers.length).fill("");
      }
      
      // "過去の資料請求"の値を設定
      const hasPastStatusIdx = headerMap[CONFIG.COL_NAMES.HAS_PAST];
      if (hasPastStatusIdx !== undefined) {
        rowValues[hasPastStatusIdx] = hasPastStatus;
      }
      Logger.log(`➕ [${webhookData.customerName}] 一致なし。行 ${targetRowIdx + 1} に新規追加。過去の資料請求: ${hasPastStatus}`);
    }

    // 共通データの反映処理
    Logger.log("🔧 データマッピング中...");
    let mappedCount = 0;
    for (const [header, value] of Object.entries(writeDataMap)) {
      const colIdx = headerMap[header];
      if (colIdx !== undefined) {
        rowValues[colIdx] = value;
        mappedCount++;
      }
    }
    Logger.log(`📝 マッピング完了: ${mappedCount}項目`);
    timer.lap("データマッピング");

    // スプレッドシートに書き込み (rowIndexは0-basedなので +1 する)
    Logger.log(`💾 スプレッドシートに書き込み中... (${targetRowIdx + 1}行目)`);
    targetSheet.getRange(targetRowIdx + 1, 1, 1, rowValues.length).setValues([rowValues]);
    timer.lap("データ書き込み");
    
    // 変更を即座に反映
    SpreadsheetApp.flush();
    timer.lap("flush完了");

    // 正常終了を通知
    Logger.log("========================================");
    Logger.log(`✅ 処理成功！`);
    Logger.log(`   アクション: ${actionType}`);
    Logger.log(`   対象行: ${targetRowIdx + 1}行目`);
    Logger.log(`   来場年: ${writeDataMap[CONFIG.COL_NAMES.VISIT_YEAR]}`);
    Logger.log(`   総処理時間: ${timer.total()}ms`);
    Logger.log("========================================");
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "SUCCESS", 
      message: "データが正常に処理されました。",
      action: actionType,
      row: targetRowIdx + 1,
      processingTime: timer.total()
    })).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    Logger.log("========================================");
    Logger.log(`❌ エラー発生！`);
    Logger.log(`   メッセージ: ${err.message}`);
    Logger.log(`   スタック: ${err.stack}`);
    Logger.log(`   処理時間: ${timer.total()}ms`);
    Logger.log("========================================");
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "ERROR", 
      message: err.message,
      processingTime: timer.total()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}