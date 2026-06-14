/**
 * ============================================================
 * doPost.gs — Zoho CRM Webhook 受信 → 「営業進捗管理」シート書き込み
 * ============================================================
 * 【設計方針】
 *  - 列はヘッダー文字列で検索して書き込む（列の追加・並び替えに強い）。
 *  - マッピング対象の列だけを setValue で書く。
 *    → 数式列（粗利額（自動計算）/ 実粗利率（自動計算））や、
 *      担当者が手入力する列（見積金額・原価 等）は一切触らない。
 *  - 1リクエスト＝1行なので setValue（個別書き込み）で十分。
 *
 * 【カスタマイズ】
 *  - 受け取る項目を増やすときは DIRECT_MAP に1行追加するだけ。
 *    （左＝シートのヘッダー名 / 右＝Zoho のパラメータ名）
 *  - 加工が必要な項目（電話番号・来場日/来場年）だけ buildWriteMap_ で個別処理。
 * ============================================================
 */

// ============================================================
// 設定
// ============================================================
const CONFIG = {
  PROP_KEY_SS_ID: "SPREADSHEET_ID",   // スプレッドシートIDを保存したスクリプトプロパティのキー名
  SHEET_NAME:     "営業進捗管理",      // 書き込み先シート

  COL_MATCH:    "物件ナンバー",        // この列が一致したら UPDATE（一致しなければ APPEND）
  COL_PROJECT:  "案件名",              // 過去の資料請求の照合に使う列
  COL_HAS_PAST: "過去の資料請求",      // 過去の資料請求の結果（有り/無し）を書く列

  // 新規行を探すときの空白判定列（すべて空ならその行に書き込む）
  CHECK_EMPTY_COLS: ["物件ナンバー", "チャネル", "来場年", "来場日", "ステージ", "案件名"],
};

// シートのヘッダー名 → Zoho のパラメータ名（値をそのまま入れる項目）
const DIRECT_MAP = {
  "物件ナンバー":   "property_number",
  "チャネル":       "showroom",
  "ステージ":       "stage",
  "工務店（B2B）":  "sorting",
  "工務店名":       "company_name",
  "案件名":         "customer_name",
  "きっかけ":       "found_out",
  "問い合わせ経路": "inquiry_class",
  "現住所":         "address",
  "メールアドレス": "mail",
  "内容":           "inquiry_detail",
  "情報":           "contact_detail",
};

// ============================================================
// メイン：Webhook 受信
// ============================================================
function doPost(e) {
  const t0 = Date.now();
  try {
    if (!e || !e.parameter) throw new Error("Webhook パラメータが空です。");

    const ssId = PropertiesService.getScriptProperties().getProperty(CONFIG.PROP_KEY_SS_ID);
    if (!ssId) throw new Error("スプレッドシートIDがスクリプトプロパティに未設定です。");

    const sheet = SpreadsheetApp.openById(ssId).getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) throw new Error(`シート「${CONFIG.SHEET_NAME}」が見つかりません。`);

    const data = sheet.getDataRange().getValues();
    if (data.length === 0) throw new Error("シートにヘッダー行がありません。");

    const headerMap = buildHeaderMap_(data[0]);   // ヘッダー名 → 列インデックス(0始まり)

    // 書き込む値を { ヘッダー名: 値 } の形で用意
    const writeMap = buildWriteMap_(e.parameter);

    // 全データ行を1回走査し、物件ナンバー一致行と案件名一致数を取得
    const { updateRow, nameMatchCount } = scanRows_(data, headerMap, writeMap);

    // 書き込み先の行（1始まり）を決定
    let targetRow, action;
    if (updateRow) {
      targetRow = updateRow;
      action = "UPDATE";
    } else {
      targetRow = findFirstEmptyRow_(data, headerMap, CONFIG.CHECK_EMPTY_COLS);
      action = "APPEND";
      // 過去の資料請求は新規追加時のみ記入
      writeMap[CONFIG.COL_HAS_PAST] = (nameMatchCount >= 1) ? "有り" : "無し";
    }

    // マッピング対象の列だけを書き込む（数式列・手入力列は触らない）
    writeMappedCells_(sheet, targetRow, headerMap, writeMap);
    SpreadsheetApp.flush();

    Logger.log(`✅ ${action}: ${targetRow}行目に書き込み（${Date.now() - t0}ms）`);
    return jsonOut_({ status: "SUCCESS", action: action, row: targetRow });

  } catch (err) {
    Logger.log(`❌ エラー: ${err.message}\n${err.stack || ""}`);
    return jsonOut_({ status: "ERROR", message: err.message });
  }
}

// ============================================================
// ヘルパー
// ============================================================

/** ヘッダー配列 → { ヘッダー名(トリム済): 列インデックス(0始まり) } */
function buildHeaderMap_(headers) {
  const map = {};
  headers.forEach(function(h, i) { map[String(h).trim()] = i; });
  return map;
}

/** Webhook パラメータ → { ヘッダー名: 値 } */
function buildWriteMap_(p) {
  const clean = function(v) { return String(v || "").trim(); };
  const map = {};

  // 直接マッピング項目
  Object.keys(DIRECT_MAP).forEach(function(header) {
    map[header] = clean(p[DIRECT_MAP[header]]);
  });

  // 電話番号：携帯(mobile_phone)優先、無ければ固定(phone)
  map["電話番号"] = clean(p.mobile_phone) || clean(p.phone);

  // 来場日・来場年：問い合わせ日(inquiry_date)を正規化して分解
  const d = parseFlexibleDate_(clean(p.inquiry_date));
  map["来場日"] = d.formattedDate;   // "YYYY/MM/DD"
  map["来場年"] = d.year;            // "YYYY年"

  return map;
}

/**
 * 全データ行を1回走査して以下を返す。
 *  - updateRow:      物件ナンバーが一致した行(1始まり)。無ければ null
 *  - nameMatchCount: 案件名が一致した行数
 */
function scanRows_(data, headerMap, writeMap) {
  const propCol = headerMap[CONFIG.COL_MATCH];
  const nameCol = headerMap[CONFIG.COL_PROJECT];
  const inProp = writeMap[CONFIG.COL_MATCH] || "";
  const inName = writeMap[CONFIG.COL_PROJECT] || "";

  let updateRow = null;
  let nameMatchCount = 0;

  for (let i = 1; i < data.length; i++) {   // 0行目＝ヘッダーは除外
    if (propCol != null && inProp !== "" &&
        String(data[i][propCol]).trim() === inProp) {
      updateRow = i + 1;
    }
    if (nameCol != null && inName !== "" &&
        String(data[i][nameCol]).trim() === inName) {
      nameMatchCount++;
    }
  }
  return { updateRow: updateRow, nameMatchCount: nameMatchCount };
}

/** 指定列がすべて空白の最初の行(1始まり)。無ければ最終行の次 */
function findFirstEmptyRow_(data, headerMap, checkCols) {
  const idxs = checkCols
    .map(function(h) { return headerMap[h]; })
    .filter(function(i) { return i != null; });

  for (let i = 1; i < data.length; i++) {
    const allEmpty = idxs.every(function(c) {
      const v = data[i][c];
      return v === "" || v === null || v === undefined;
    });
    if (allEmpty) return i + 1;
  }
  return data.length + 1;
}

/** writeMap の各項目を、ヘッダー名で見つけた列に書き込む（空値はスキップ＝既存値を温存） */
function writeMappedCells_(sheet, row, headerMap, writeMap) {
  Object.keys(writeMap).forEach(function(header) {
    const col = headerMap[header];
    const val = writeMap[header];
    if (col == null) {
      Logger.log(`⚠ ヘッダー「${header}」が見つからないためスキップ`);
      return;
    }
    if (val === "" || val === null || val === undefined) return;  // 既存値を上書きしない
    sheet.getRange(row, col + 1).setValue(val);  // col は0始まり → +1
  });
}

/**
 * 様々な書式の日付文字列を正規化する。
 * 対応例: 2025-11-15 / 2025/11/15 / 2025年11月15日 / 2025.11.15
 * @return {{formattedDate:string, year:string}} 解釈不能なら両方 ""
 */
function parseFlexibleDate_(dateStr) {
  if (!dateStr) return { formattedDate: "", year: "" };

  // 全角数字 → 半角
  let s = dateStr.replace(/[０-９]/g, function(ch) {
    return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0);
  });
  // 区切りを「/」に統一
  s = s.replace(/年/g, "/").replace(/月/g, "/").replace(/日/g, "")
       .replace(/\./g, "/").replace(/-/g, "/").trim();

  const date = new Date(s);
  if (isNaN(date.getTime())) {
    Logger.log(`⚠ 日付の解釈に失敗: "${dateStr}"`);
    return { formattedDate: "", year: "" };
  }
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return { formattedDate: `${y}/${m}/${d}`, year: `${y}年` };
}

/** JSON レスポンス生成 */
function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}