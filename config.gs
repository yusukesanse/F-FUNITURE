// ==================================================================
// config.gs ― 複数ファイルが横断利用する共通設定だけを置くファイル
// ==================================================================
// 【方針】2つ以上のファイルが参照する設定のみここに集約する。
//   1ファイル専用の設定は各ファイル側に置く:
//     ・Webhook受信       → webhook.gs（WEBHOOK_CONFIG）
//     ・年度シート自動作成 → fiscalYear.gs（FY_CONFIG）
//     ・製造間接費の転記   → overhead.gs（MFG_CONFIG）
//
// 【機密情報】スプレッドシートID は Script Properties（スクリプトに直書きせず
//   GAS側に保存する設定領域）で管理する。
//   GASエディタ ＞ プロジェクトの設定 ＞ スクリプトプロパティ で設定:
//     DESTINATION_SS_ID … 転記先（売上進捗表）スプレッドシートID
//     SPREADSHEET_ID    … 営業進捗管理スプレッドシートID（webhook.gs が使用）
// ==================================================================

/** Script Properties から必須の機密値を取得（未設定ならエラー） */
function getSecret_(key) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  if (!v) {
    throw new Error('Script Property "' + key + '" が未設定です。'
      + ' GASエディタ ＞ プロジェクトの設定 ＞ スクリプトプロパティ で設定してください。');
  }
  return v;
}

const CONFIG_PROJECT = {

  // --- 接続先（機密は Script Properties から遅延取得） ---
  // 年度ごとに別スプレッドシート（手動作成）に書き込む。年度別IDは
  // スクリプトプロパティ「DEST_SS_ID_R{年度}」（例 DEST_SS_ID_R7, DEST_SS_ID_R8）に登録する。
  // 解決は fiscalYear.gs の resolveDestSsId / getDestSsIdForFiscalYear_ が行う。
  // ↓ DESTINATION_SS_ID はテスト関数・後方互換用（通常フローでは年度別IDを使用）。
  get DESTINATION_SS_ID() { return getSecret_("DESTINATION_SS_ID"); },

  /** 各年度ファイル共通の「書き込み先シート名」（全ファイルで同じタブ名にすること） */
  DEST_SHEET_SALES: "全社売上進捗表（集計）",

  /** 転記元シート（複数可） */
  SOURCE_SHEET_NAMES: ["営業進捗管理"],

  // --- トリガー列（この列を手入力で編集すると転記処理が走る） ---
  TRIGGER_COLUMNS: {
    SALES_DATE:        { header: "売上予定日", type: "date",   clearOnEmpty: true },
    STATUS:            { header: "ステージ",   type: "status", trackOldValue: true },
    ESTIMATED_AMOUNT:  { header: "見積金額",   type: "amount" },
    BTB_COMPANY:       { header: "工務店名",   type: "text" },
    GROSS_MARGIN_RATE: { header: "予定粗利率", type: "text" },   // ← 旧「粗利率」
    COST:              { header: "原価",       type: "cost" }    // ★今回追加
  },

  // テキストトリガーの転記先列マッピング
  TEXT_TRIGGER_MAPPING: {
    BTB_COMPANY:       { destCol: "BTB_COMPANY", sourceKey: "companyName" },
    GROSS_MARGIN_RATE: { destCol: "SALES_COL",   sourceKey: "grossMarginRate" }
  },

  // --- 転記元の列（ヘッダー名で検索） ---
  SOURCE_COLUMNS: {
    OBJECT_NUMBER:            { header: "物件ナンバー",         required: true },
    CUSTOMER_NAME:            { header: "案件名",               required: false },
    ESTIMATED_AMOUNT:         { header: "見積金額",             required: false, type: "number" },
    SALES_DATE:               { header: "売上予定日",           required: false, type: "date" },
    STATUS:                   { header: "ステージ",             required: false },
    SALES_REP:                { header: "担当者1",              required: false },
    CHANNEL:                  { header: "チャネル",             required: false },
    COMPANY_NAME:             { header: "工務店名",             required: false },
    GROSS_MARGIN_RATE:        { header: "予定粗利率",           required: false, type: "number" }, // ← 旧「粗利率」
    GROSS_MARGIN_AMOUNT:      { header: "粗利額（自動計算）",   required: false, type: "number" },
    COST:                     { header: "原価",                 required: false, type: "number" }, // ★追加
    ACTUAL_GROSS_MARGIN_RATE: { header: "実粗利率（自動計算）", required: false, type: "number" }   // ★追加
  },

  // --- 転記先の列番号（1始まり） ---
  //   H(8):契約予定 ／ I(9):契約売上 ／ J(10):売上 or 粗利率
  //   K(11):粗利率 or 粗利額 ／ L(12):粗利額
  DESTINATION_COLUMNS: {
    MONTH: 2, STORE: 2, SALES_REP_COL: 3, SALES_MONTH_COL: 4,
    OBJECT_NO: 5, BTB_COMPANY: 6, PROJECT_NAME: 7,
    PLANNED_CONTRACT_COL: 8, CONTRACT_SALES_COL: 9, SALES_COL: 10,
    GROSS_MARGIN_RATE_AB: 11, GROSS_MARGIN_AMOUNT: 12
  },
  /** データ行の開始行（1-2行目=タイトル, 3行目=見出し, 4行目〜=データ） */
  DATA_START_ROW: 4,

  DATA_MAPPING: {
    storeName: "STORE", salesRep: "SALES_REP_COL", salesMonth: "SALES_MONTH_COL",
    objectNumber: "OBJECT_NO", companyName: "BTB_COMPANY", customerName: "PROJECT_NAME"
  },

  // --- ステージ定義 ---
  STAGES: {
    A: "A 契約済み", B: "B プラン契約済み", C: "C 商談中",
    D: "D プラン・見積もり", E: "E 追客", F: "問い合わせ・見学のみ", N: "失注"
  },
  STATUS_FINAL_STAGE:  ["A 契約済み", "B プラン契約済み"],          // 月ブロックに転記
  STATUS_OTHER_DEAL:   ["C 商談中", "D プラン・見積もり"],          // その他商談ブロックに転記
  STATUS_CLEAR_TARGET: ["問い合わせ・見学のみ", "失注"],            // 転記先を削除

  // ステージ別 見積金額列
  STAGE_AMOUNT_MAPPING: {
    "A 契約済み": "CONTRACT_SALES_COL",   "B プラン契約済み": "PLANNED_CONTRACT_COL",
    "C 商談中":   "CONTRACT_SALES_COL",   "D プラン・見積もり": "PLANNED_CONTRACT_COL"
  },
  // ステージ別 粗利率列（A/B→K列, C/D→J列）
  STAGE_GROSS_MARGIN_MAPPING: {
    "A 契約済み": "GROSS_MARGIN_RATE_AB", "B プラン契約済み": "GROSS_MARGIN_RATE_AB",
    "C 商談中":   "SALES_COL",            "D プラン・見積もり": "SALES_COL"
  },
  // ステージ別 粗利額列（A/B/C→L列, D→K列）
  STAGE_GROSS_AMOUNT_MAPPING: {
    "A 契約済み": "GROSS_MARGIN_AMOUNT",  "B プラン契約済み": "GROSS_MARGIN_AMOUNT",
    "C 商談中":   "GROSS_MARGIN_AMOUNT",  "D プラン・見積もり": "GROSS_MARGIN_RATE_AB"
  },

  // --- 店舗マッピング（チャネル値 → 店舗名） ---
  STORE_MAP: { "岐阜": "岐阜", "清澄白河": "清澄", "南青山": "南青山", "湘南辻堂": "湘南" },

  // --- その他（横断利用するラベル等） ---
  RUIKEI_LABEL:     "累計",
  OTHER_DEAL_LABEL: "その他商談案件合計",
  TIMEZONE:         "Asia/Tokyo",
  MONTH_FORMAT:     "yyyy年M月"
};

/** 機密スクリプトプロパティの設定状況を確認する診断用関数（任意で実行） */
function checkSecretProperties() {
  const keys = ["DESTINATION_SS_ID", "SPREADSHEET_ID"];
  const props = PropertiesService.getScriptProperties();
  Logger.log("🔐 機密プロパティ診断");
  keys.forEach(function (k) {
    Logger.log((props.getProperty(k) ? "  ✅ 設定済み " : "  ❌ 未設定 ") + k);
  });
}