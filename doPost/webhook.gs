/**
 * ============================================================
 * webhook.gs ― Zoho CRM Webhook 受信 → シートへ書き込み
 * ============================================================
 * 【役割】このファイルだけで完結する。doPost が Zoho からのデータを
 *   シートへ書き込む。その後の転記は人の手入力編集で salesTransfer.gs の
 *   編集トリガーが行う（doPost の setValue では編集トリガーは発火しない
 *   ＝二重処理にならない）。
 *
 * 【書き込み先の振り分け】
 *   ご希望のショールーム（Zohoの showroom／シート上は「チャネル」列）が「資料請求」の
 *   場合は「資料請求管理」シートへ書き込む。該当しなければ既定の「営業進捗管理」へ。
 *   ルールは WEBHOOK_CONFIG.ROUTES で設定（上から評価し最初に一致したものを採用）。
 *   資料請求管理シートは営業進捗管理と同じヘッダー構成のため、フィールドマップ等は共通。
 *
 * 【過去の資料請求の判定】★今回修正
 *   「過去の資料請求」は “その顧客が以前に資料請求したか” を表す。資料請求の記録は
 *   常に「資料請求管理」シートに入るため、書き込み先が営業進捗管理／資料請求管理の
 *   どちらであっても、判定の参照先は WEBHOOK_CONFIG.REQUEST_SHEET_NAME（資料請求管理）
 *   に固定する。資料請求管理シート内に同じ案件名が既に1件以上あれば「有り」。
 *   （新規追記時は、今回の行は数に含めない＝既存行のみを数える）
 *
 * 【設計】列はヘッダー文字列で検索（列の追加・並び替えに強い）。
 *   マッピング対象の列だけ setValue する。数式列・手入力列は触らない。
 *
 * 【カスタマイズ】
 *   ・受け取る項目を増やす → WEBHOOK_FIELD_MAP に1行追加
 *   ・書き込み先を増やす   → WEBHOOK_CONFIG.ROUTES に1行追加
 * ============================================================
 */

// ============================================================
// 設定（このファイル専用）
// ============================================================
const WEBHOOK_CONFIG = {
  PROP_KEY_SS_ID: "SPREADSHEET_ID",   // 営業進捗管理スプレッドシートIDのプロパティ名

  DEFAULT_SHEET_NAME: "営業進捗管理",  // 既定の書き込み先（どのルートにも一致しない場合）

  // 資料請求の記録シート（「過去の資料請求」判定の参照先＝下のROUTESの資料請求の書き込み先と同じにする）
  REQUEST_SHEET_NAME: "資料請求管理",

  // --- 書き込み先の振り分けルール（上から評価し、最初に一致したものを採用）---
  //   param  … Zohoのパラメータ名
  //   equals … その値（前後の空白は無視して完全一致で判定）
  //   sheet  … 書き込み先シート名
  // ※ Zohoが送る値が「資料請求」と完全一致しない（前後に文字が付く等）場合は、
  //   equals の文字列を実際の値に合わせてください。部分一致にしたい場合はご相談を。
  ROUTES: [
    { param: "showroom", equals: "資料請求", sheet: "資料請求管理" }
  ],

  COL_MATCH:    "物件ナンバー",        // 一致したら UPDATE（しなければ APPEND）
  COL_PROJECT:  "案件名",              // 過去の資料請求の照合に使う列（資料請求管理内で照合）
  COL_HAS_PAST: "過去の資料請求",      // 過去の資料請求の結果（有り/無し）を書く列

  // 新規行を探すときの空白判定列（すべて空ならその行に書き込む）
  CHECK_EMPTY_COLS: ["物件ナンバー", "チャネル", "来場年", "来場日", "ステージ", "案件名"]
};

// シートのヘッダー名 → Zoho のパラメータ名（値をそのまま入れる項目）
const WEBHOOK_FIELD_MAP = {
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
  "情報":           "contact_detail"
};

// ============================================================
// メイン：Webhook 受信
// ============================================================
function doPost(e) {
  const t0 = Date.now();
  try {
    if (!e || !e.parameter) throw new Error("Webhook パラメータが空です。");

    const ssId = PropertiesService.getScriptProperties().getProperty(WEBHOOK_CONFIG.PROP_KEY_SS_ID);
    if (!ssId) throw new Error("スプレッドシートIDがスクリプトプロパティに未設定です。");

    const ss = SpreadsheetApp.openById(ssId);

    // 問い合わせ分類などで書き込み先シートを決める（該当なしは既定シート）
    const targetSheetName = resolveTargetSheetName_(e.parameter);

    const sheet = ss.getSheetByName(targetSheetName);
    if (!sheet) throw new Error(`シート「${targetSheetName}」が見つかりません。`);

    const data = sheet.getDataRange().getValues();
    if (data.length === 0) throw new Error(`シート「${targetSheetName}」にヘッダー行がありません。`);

    const headerMap = buildHeaderMap_(data[0]);   // ヘッダー名 → 列インデックス(0始まり)
    const writeMap  = buildWriteMap_(e.parameter); // { ヘッダー名: 値 }

    // 物件ナンバー一致行（更新対象）をターゲットシートから探す
    const updateRow = findUpdateRow_(data, headerMap, writeMap);

    let targetRow, action;
    if (updateRow) {
      targetRow = updateRow;
      action = "UPDATE";
    } else {
      targetRow = findFirstEmptyRow_(data, headerMap, WEBHOOK_CONFIG.CHECK_EMPTY_COLS);
      action = "APPEND";

      // ★過去の資料請求は「資料請求管理」シートで案件名一致を数えて判定（書き込み先に依存しない）
      const pastCount = countPastRequests_(
        ss, writeMap[WEBHOOK_CONFIG.COL_PROJECT], targetSheetName, data, headerMap
      );
      writeMap[WEBHOOK_CONFIG.COL_HAS_PAST] = (pastCount >= 1) ? "有り" : "無し";
      Logger.log(`📋 過去の資料請求: 資料請求管理に同名 ${pastCount}件 → ${writeMap[WEBHOOK_CONFIG.COL_HAS_PAST]}`);
    }

    writeMappedCells_(sheet, targetRow, headerMap, writeMap);
    SpreadsheetApp.flush();

    Logger.log(`✅ ${action}: 「${targetSheetName}」${targetRow}行目に書き込み（${Date.now() - t0}ms）`);
    return jsonOut_({ status: "SUCCESS", action: action, sheet: targetSheetName, row: targetRow });

  } catch (err) {
    Logger.log(`❌ エラー: ${err.message}\n${err.stack || ""}`);
    return jsonOut_({ status: "ERROR", message: err.message });
  }
}

// ============================================================
// ヘルパー
// ============================================================

/**
 * Webhookパラメータから書き込み先シート名を決める。
 *   WEBHOOK_CONFIG.ROUTES を上から評価し、最初に一致したルールのシートを採用。
 *   どれにも一致しなければ DEFAULT_SHEET_NAME。
 */
function resolveTargetSheetName_(p) {
  const routes = WEBHOOK_CONFIG.ROUTES || [];
  for (let i = 0; i < routes.length; i++) {
    const r = routes[i];
    const actual = String(p[r.param] || "").trim();
    if (actual !== "" && actual === String(r.equals).trim()) {
      Logger.log(`🔀 振り分け: ${r.param}="${actual}" → シート「${r.sheet}」`);
      return r.sheet;
    }
  }
  return WEBHOOK_CONFIG.DEFAULT_SHEET_NAME;
}

/**
 * ★今回修正：「過去の資料請求」を判定する。
 *   資料請求管理シート内に同じ案件名が既に何件あるかを数える（今回の新規行は含めない）。
 *   書き込み先が資料請求管理ならすでに読み込んだデータを再利用し、それ以外なら
 *   資料請求管理シートを開いて数える。
 *   @param {Spreadsheet} ss
 *   @param {string} projectName  今回の案件名
 *   @param {string} targetSheetName  今回の書き込み先シート名
 *   @param {Array}  targetData       書き込み先シートの全データ（再利用用）
 *   @param {Object} targetHeaderMap  書き込み先シートのヘッダーマップ（再利用用）
 *   @return {number} 一致件数
 */
function countPastRequests_(ss, projectName, targetSheetName, targetData, targetHeaderMap) {
  const name = String(projectName || "").trim();
  if (name === "") return 0;

  const reqSheetName = WEBHOOK_CONFIG.REQUEST_SHEET_NAME;

  // 書き込み先が資料請求管理なら、すでに読み込んだデータを再利用（余分な読み込みを避ける）
  if (targetSheetName === reqSheetName) {
    return countNameMatches_(targetData, targetHeaderMap[WEBHOOK_CONFIG.COL_PROJECT], name);
  }

  // それ以外は資料請求管理シートを開いて数える
  const reqSheet = ss.getSheetByName(reqSheetName);
  if (!reqSheet) {
    Logger.log(`⚠ 「${reqSheetName}」が見つからないため過去の資料請求を判定できません → 無し扱い`);
    return 0;
  }
  const reqData = reqSheet.getDataRange().getValues();
  if (reqData.length === 0) return 0;
  const reqHeaderMap = buildHeaderMap_(reqData[0]);
  return countNameMatches_(reqData, reqHeaderMap[WEBHOOK_CONFIG.COL_PROJECT], name);
}

/** data の nameCol 列（0始まり）で、name と一致する行数を数える（ヘッダー行は除外） */
function countNameMatches_(data, nameCol, name) {
  if (nameCol == null) return 0;
  const target = String(name).trim();
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][nameCol]).trim() === target) count++;
  }
  return count;
}

/** ヘッダー配列 → { ヘッダー名(トリム済): 列インデックス(0始まり) } */
function buildHeaderMap_(headers) {
  const map = {};
  headers.forEach(function (h, i) { map[String(h).trim()] = i; });
  return map;
}

/** Webhook パラメータ → { ヘッダー名: 値 } */
function buildWriteMap_(p) {
  const clean = function (v) { return String(v || "").trim(); };
  const map = {};

  // 直接マッピング項目
  Object.keys(WEBHOOK_FIELD_MAP).forEach(function (header) {
    map[header] = clean(p[WEBHOOK_FIELD_MAP[header]]);
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
 * 物件ナンバーが一致する行(1始まり)を返す。無ければ null。
 *   （以前の scanRows_ は案件名一致数も返していたが、過去の資料請求の判定は
 *     countPastRequests_ に分離したため、ここは更新行の特定のみに専念する）
 */
function findUpdateRow_(data, headerMap, writeMap) {
  const propCol = headerMap[WEBHOOK_CONFIG.COL_MATCH];
  const inProp  = writeMap[WEBHOOK_CONFIG.COL_MATCH] || "";
  if (propCol == null || inProp === "") return null;

  let updateRow = null;
  for (let i = 1; i < data.length; i++) {   // 0行目＝ヘッダーは除外
    if (String(data[i][propCol]).trim() === inProp) {
      updateRow = i + 1;
    }
  }
  return updateRow;
}

/** 指定列がすべて空白の最初の行(1始まり)。無ければ最終行の次 */
function findFirstEmptyRow_(data, headerMap, checkCols) {
  const idxs = checkCols
    .map(function (h) { return headerMap[h]; })
    .filter(function (i) { return i != null; });

  for (let i = 1; i < data.length; i++) {
    const allEmpty = idxs.every(function (c) {
      const v = data[i][c];
      return v === "" || v === null || v === undefined;
    });
    if (allEmpty) return i + 1;
  }
  return data.length + 1;
}

/** writeMap の各項目を、ヘッダー名で見つけた列に書き込む（空値はスキップ＝既存値を温存） */
function writeMappedCells_(sheet, row, headerMap, writeMap) {
  Object.keys(writeMap).forEach(function (header) {
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
  let s = dateStr.replace(/[０-９]/g, function (ch) {
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