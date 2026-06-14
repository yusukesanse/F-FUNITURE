// ==================================================================
// ★★★ config.gs - 設定の一元管理 ★★★
// ==================================================================
// 転記列やトリガー列を追加する場合は、このファイルのみ編集してください
//
// 【重要】スプレッドシートIDなどの機密情報は Script Properties で管理しています。
// 初回セットアップ時は GAS エディタの「プロジェクトの設定」→「スクリプト プロパティ」
// から必要な値を設定してください。
// ==================================================================
// [INDEX]
// [DEFAULTS] CONFIG_DEFAULT_SCRIPT_PROPERTIES（機密情報は含まない）
// [CONFIG]   CONFIG_PROJECT（main / Manufacturing 両方で参照）
// [SETUP]    setupScriptProperties / setupScriptPropertiesForce

/**
 * Script Properties のデフォルト値
 * ※ 機密情報（スプレッドシートID等）はここに記載しない。
 *   それらはスクリプトプロパティに直接設定する運用とする。
 *   ここに記載するのは、Gitに公開しても問題ない一般的な業務設定のみ。
 */
const CONFIG_DEFAULT_SCRIPT_PROPERTIES = {
  // --- シート名（業務固有だが機密ではない） ---
  MFG_SOURCE_SHEET_NAME: "2026製造間接費入力シート（事務員専用）",
  MFG_DEST_SHEET_NAME:   "R7年度全社売上進捗表（集計）",
  MFG_DEST_TOTAL_MARKER: "合計",
  DEST_SHEET_SALES:      "R7年度全社売上進捗表（集計）",
  SOURCE_SHEET_NAMES:    "営業進捗管理",

  // --- ラベル・フォーマット ---
  RUIKEI_LABEL:     "累計",
  OTHER_DEAL_LABEL: "その他商談案件合計",
  TIMEZONE:         "Asia/Tokyo",
  MONTH_FORMAT:     "yyyy年M月"

  // ※ 以下の機密情報は Script Properties に直接設定すること:
  //   - DESTINATION_SS_ID
  //   - MFG_DEST_SPREADSHEET_ID
  //   - SPREADSHEET_ID
  //   - SS_ID
};

// 追加導入した設定だけを一括登録したいときに使うマップ
const CONFIG_ADDITIONAL_SCRIPT_PROPERTIES = {
  MFG_DEST_RUIKEI_MARKER: "累計"
};

/**
 * 機密情報用: Script Properties から必須の値を取得する。
 * デフォルト値へのフォールバックはせず、未設定時は必ずエラーを投げる。
 * （スプレッドシートID等の機密情報専用）
 */
function getRequiredSecretConfig_(key) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (value) {
    return value;
  }
  throw new Error(
    'Script Properties に "' + key + '" が未設定です。'
    + ' これは機密情報のため、Apps Script エディタ > プロジェクトの設定 > '
    + 'スクリプト プロパティ で必ず設定してください。'
  );
}

/**
 * 業務設定用: Script Properties から値を取得し、なければデフォルト値を返す。
 * （シート名やラベル等、Gitに公開可能な設定用）
 */
function getRequiredScriptPropertyConfig_(key) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (value) {
    return value;
  }
  if (Object.prototype.hasOwnProperty.call(CONFIG_DEFAULT_SCRIPT_PROPERTIES, key)) {
    Logger.log('⚠ Script Property "' + key + '" が未設定のため初期値を使用します。');
    return CONFIG_DEFAULT_SCRIPT_PROPERTIES[key];
  }
  throw new Error(
    'Script Properties に "' + key + '" が未設定です。'
    + ' Apps Script エディタ > プロジェクトの設定 > スクリプト プロパティ で設定してください。'
  );
}

function getScriptPropertyWithDefaultConfig_(key, defaultValue) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  return value ? value : defaultValue;
}

var CONFIG_PROJECT = {

  // ==================================================================
  // 接続先設定（機密情報: Script Properties から必須取得）
  // ==================================================================

  /** 転記先スプレッドシートID（機密情報） */
  DESTINATION_SS_ID: getRequiredSecretConfig_("DESTINATION_SS_ID"),

  /** 転記先シート名 */
  DEST_SHEET_SALES: getRequiredScriptPropertyConfig_("DEST_SHEET_SALES"),

  /** 転記元シート名（複数可） */
  SOURCE_SHEET_NAMES: getRequiredScriptPropertyConfig_("SOURCE_SHEET_NAMES").split(",").map(function(name) {
    return String(name || "").trim();
  }).filter(function(name) {
    return name !== "";
  }),

  // ==================================================================
  // ★★★ トリガー列設定 ★★★
  // 編集時に処理を実行する列を定義
  // ==================================================================

  TRIGGER_COLUMNS: {
    SALES_DATE: {
      header: "売上予定日",
      type: "date",
      clearOnEmpty: true
    },
    STATUS: {
      header: "ステージ",
      type: "status",
      trackOldValue: true
    },
    ESTIMATED_AMOUNT: {
      header: "見積金額",
      type: "amount"
    },
    BTB_COMPANY: {
      header: "工務店名",
      type: "text"
    },
    GROSS_MARGIN_RATE: {
      header: "粗利率",
      type: "text"
    }
  },

  // ==================================================================
  // ★★★ テキストトリガーのマッピング ★★★
  // ==================================================================

  TEXT_TRIGGER_MAPPING: {
    BTB_COMPANY: {
      destCol: "BTB_COMPANY",
      sourceKey: "companyName"
    },
    GROSS_MARGIN_RATE: {
      destCol: "SALES_COL",       // 実質未使用（main.gsでSTAGE_GROSS_MARGIN_MAPPINGを参照）
      sourceKey: "grossMarginRate"
    }
  },

  // ==================================================================
  // ★★★ 転記元列マッピング ★★★
  // ==================================================================

  SOURCE_COLUMNS: {
    OBJECT_NUMBER:       { header: "物件ナンバー", required: true },
    CUSTOMER_NAME:       { header: "案件名",       required: false },
    ESTIMATED_AMOUNT:    { header: "見積金額",     required: false, type: "number" },
    SALES_DATE:          { header: "売上予定日",   required: false, type: "date" },
    STATUS:              { header: "ステージ",     required: false },
    SALES_REP:           { header: "担当者1",      required: false },
    CHANNEL:             { header: "チャネル",     required: false },
    COMPANY_NAME:        { header: "工務店名",     required: false },
    GROSS_MARGIN_RATE:   { header: "粗利率",       required: false, type: "number" },
    GROSS_MARGIN_AMOUNT: { header: "粗利額（自動計算）", required: false, type: "number" }
  },

  // ==================================================================
  // ★★★ 転記先列マッピング ★★★
  // 転記先シートの列番号を定義（1始まり）
  //
  // 列の役割まとめ:
  //   H列(8):  契約予定       ← B/Dステージの見積金額
  //   I列(9):  契約売上       ← A/Cステージの見積金額
  //   J列(10): 売上/粗利率    ← A/Bは売上金額、C/Dは粗利率
  //   K列(11): 粗利率/粗利額  ← A/Bは粗利率、Dは粗利額
  //   L列(12): 粗利額         ← A/B/Cは粗利額
  // ==================================================================

  DESTINATION_COLUMNS: {
    MONTH:                2,   // B列: 月度ヘッダー（検索用）
    STORE:                2,   // B列: 店舗
    SALES_REP_COL:        3,   // C列: 担当者
    SALES_MONTH_COL:      4,   // D列: 売上月
    OBJECT_NO:            5,   // E列: 物件No
    BTB_COMPANY:          6,   // F列: 工務店名
    PROJECT_NAME:         7,   // G列: 案件名
    PLANNED_CONTRACT_COL: 8,   // H列: 契約予定（B/Dステージ用）
    CONTRACT_SALES_COL:   9,   // I列: 契約売上（A/Cステージ用）
    SALES_COL:           10,   // J列: 売上（A/B用）/ 粗利率（C/D用）
    GROSS_MARGIN_RATE_AB:11,   // K列: 粗利率（A/B用）/ 粗利額（D用）
    GROSS_MARGIN_AMOUNT: 12    // L列: 粗利額（A/B/C用）
  },
  /** データ行の開始行(1-2行目=タイトル, 3行目=見出し, 4行目〜=データ) */
  DATA_START_ROW: 4,

  // ==================================================================
  // ★★★ 転記データマッピング ★★★
  // ==================================================================

  DATA_MAPPING: {
    storeName:    "STORE",
    salesRep:     "SALES_REP_COL",
    salesMonth:   "SALES_MONTH_COL",
    objectNumber: "OBJECT_NO",
    companyName:  "BTB_COMPANY",
    customerName: "PROJECT_NAME"
  },

  // ==================================================================
  // ★★★ ステージ定義 ★★★
  // ==================================================================

  STAGES: {
    A: "A 契約済み",
    B: "B プラン契約済み",
    C: "C 商談中",
    D: "D プラン・見積もり",
    E: "E 追客",
    F: "問い合わせ・見学のみ",
    N: "失注"
  },

  /** A/Bステージ（月ブロックに転記） */
  STATUS_FINAL_STAGE: ["A 契約済み", "B プラン契約済み"],

  /** C/Dステージ（その他商談ブロックに転記） */
  STATUS_OTHER_DEAL: ["C 商談中", "D プラン・見積もり"],

  /** クリア対象ステージ（転記先データを削除） */
  STATUS_CLEAR_TARGET: ["問い合わせ・見学のみ", "失注"],

  // ==================================================================
  // ★★★ ステージ別見積金額列マッピング ★★★
  // ==================================================================

  STAGE_AMOUNT_MAPPING: {
    "A 契約済み":         "CONTRACT_SALES_COL",   // I列
    "B プラン契約済み":   "PLANNED_CONTRACT_COL", // H列
    "C 商談中":           "CONTRACT_SALES_COL",   // I列
    "D プラン・見積もり": "PLANNED_CONTRACT_COL"  // H列
  },

  // ==================================================================
  // ★★★ ステージ別粗利率列マッピング ★★★
  // A/B → K列(11)、C/D → J列(10)
  // ==================================================================

  STAGE_GROSS_MARGIN_MAPPING: {
    "A 契約済み":         "GROSS_MARGIN_RATE_AB", // K列(11): 粗利率
    "B プラン契約済み":   "GROSS_MARGIN_RATE_AB", // K列(11): 粗利率
    "C 商談中":           "SALES_COL",            // J列(10): 粗利率（C/D用）
    "D プラン・見積もり": "SALES_COL"             // J列(10): 粗利率（C/D用）
  },

  // ==================================================================
  // ★★★ ステージ別粗利額列マッピング ★★★
  // A → L列、B → L列、C → L列、D → K列
  // ==================================================================

  STAGE_GROSS_AMOUNT_MAPPING: {
    "A 契約済み":         "GROSS_MARGIN_AMOUNT", // L列
    "B プラン契約済み":   "GROSS_MARGIN_AMOUNT", // L列
    "C 商談中":           "GROSS_MARGIN_AMOUNT", // L列
    "D プラン・見積もり": "GROSS_MARGIN_RATE_AB" // K列
  },

  // ==================================================================
  // ★★★ 店舗マッピング ★★★
  // ==================================================================

  STORE_MAP: {
    "岐阜":     "岐阜",
    "清澄白河": "清澄",
    "南青山":   "南青山",
    "湘南辻堂": "湘南"
  },

  // ==================================================================
  // その他設定
  // ==================================================================

  /** 累計行のラベル */
  RUIKEI_LABEL: getScriptPropertyWithDefaultConfig_("RUIKEI_LABEL", "累計"),

  /** その他商談ブロックのラベル */
  OTHER_DEAL_LABEL: getScriptPropertyWithDefaultConfig_("OTHER_DEAL_LABEL", "その他商談案件合計"),

  /** タイムゾーン */
  TIMEZONE: getScriptPropertyWithDefaultConfig_("TIMEZONE", "Asia/Tokyo"),

  /** 月度フォーマット */
  MONTH_FORMAT: getScriptPropertyWithDefaultConfig_("MONTH_FORMAT", "yyyy年M月")
};

/**
 * Script Properties に必要な初期設定を一括登録する（既存値は保持）。
 * 初回セットアップ時に1回だけ実行してください。
 * ※ 機密情報（スプレッドシートID等）は別途手動で設定すること。
 */
function setupScriptProperties() {
  setScriptProperties_(false);
}

/**
 * Script Properties に必要な初期設定を一括登録する（既存値も上書き）。
 * 値をリセットしたい場合のみ実行してください。
 * ※ 機密情報（スプレッドシートID等）は対象外。
 */
function setupScriptPropertiesForce() {
  setScriptProperties_(true);
}

/**
 * 追加導入した設定のみを一括登録する（既存は上書きしない）。
 * 例: 新しい設定キーを後から導入したときに使用。
 */
function setupAdditionalScriptProperties() {
  setScriptPropertiesFromMap_(CONFIG_ADDITIONAL_SCRIPT_PROPERTIES, false, "additional");
}

function setScriptProperties_(forceOverwrite) {
  setScriptPropertiesFromMap_(
    CONFIG_DEFAULT_SCRIPT_PROPERTIES,
    forceOverwrite,
    forceOverwrite ? "force" : "safe"
  );
}

function setScriptPropertiesFromMap_(sourceMap, forceOverwrite, modeLabel) {
  const props = PropertiesService.getScriptProperties();
  const current = props.getProperties();
  const next = {};
  let updatedCount = 0;
  let skippedCount = 0;

  Object.keys(sourceMap).forEach(function(key) {
    const exists = Object.prototype.hasOwnProperty.call(current, key) && current[key] !== "";
    if (exists && !forceOverwrite) {
      skippedCount++;
      return;
    }
    next[key] = sourceMap[key];
    updatedCount++;
  });

  if (Object.keys(next).length > 0) {
    props.setProperties(next, false);
  }

  Logger.log(
    "Script Properties 設定完了: 更新=" + updatedCount
    + ", スキップ=" + skippedCount
    + ", mode=" + modeLabel
  );
}

/**
 * 機密情報のスクリプトプロパティが設定されているか確認する診断関数。
 * 新しい環境にデプロイした際、最初に実行して設定漏れがないか確認すること。
 */
function checkSecretProperties() {
  const requiredKeys = [
    "DESTINATION_SS_ID",
    "MFG_DEST_SPREADSHEET_ID"
    // 他に機密情報があれば追加
  ];

  const props = PropertiesService.getScriptProperties();
  const missing = [];
  const present = [];

  requiredKeys.forEach(function(key) {
    const value = props.getProperty(key);
    if (value) {
      present.push(key);
    } else {
      missing.push(key);
    }
  });

  Logger.log("========================================");
  Logger.log("🔐 機密情報スクリプトプロパティ診断");
  Logger.log("========================================");
  Logger.log("✅ 設定済み: " + present.length + "件");
  present.forEach(function(k) { Logger.log("  - " + k); });

  if (missing.length > 0) {
    Logger.log("❌ 未設定: " + missing.length + "件");
    missing.forEach(function(k) { Logger.log("  - " + k); });
    Logger.log("");
    Logger.log("⚠️ GASエディタ > プロジェクトの設定 > スクリプトプロパティ から設定してください。");
  } else {
    Logger.log("🎉 すべて設定済みです。");
  }
  Logger.log("========================================");
}

  