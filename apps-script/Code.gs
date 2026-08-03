/**
 * eBay価格管理表 自動更新スクリプト
 * =================================================================
 * 【週次の作業はこれだけ】
 *   ① eBay Seller Hubで以下の2つのレポートをダウンロードする
 *      - Traffic Report              → Views・Impressions・CTR
 *      - All active listings report  → Watchers
 *   ② それぞれ対応するシートに貼り付ける（またはFile→Importで取り込む）
 *      - 「CSVインポート(Traffic Report)」
 *      - 「CSVインポート(Active Listings)」
 *   ③ メニュー「eBay価格管理」→「① 更新する」を実行する
 *
 * 手入力は一切不要。どちらのレポートも eBay Item ID（Traffic Reportでは
 * "eBay item ID"、Active Listingsでは "Item number"）で商品マスターと突き合わせる。
 *
 * 【自動反映される項目と、その情報源】
 *   - Views / Impressions / CTR … Traffic Report CSV から自動反映
 *   - Watchers                   … Active Listings CSV から自動反映
 *   - Watch率                    … Watchers ÷ Views で自動計算
 *   - 前週比（各指標）            … Performance DB の前回スナップショットと比較して自動計算
 *   - 優先順位                   … Watch率が高い順に自動採番
 *   - 残り日数(90日)             … 出品日から自動計算
 *   - 今週やること               … 判断・Watch率から自動生成
 *   - Dashboardの今週の順位      … Watch率上位3件を自動生成
 *
 * 「更新履歴」タブは廃止。更新するたびに Performance DB にその週のスナップショットを
 * 自動追加し、商品マスターには常に最新値だけを表示する。
 * =================================================================
 */

// ------------------------------------------------------------------
// 設定
// ------------------------------------------------------------------
const CSV_TRAFFIC_SHEET_NAME = 'CSVインポート(Traffic Report)';
const CSV_ACTIVE_SHEET_NAME = 'CSVインポート(Active Listings)';
const MASTER_SHEET_NAME = '商品マスター';
const PERF_DB_SHEET_NAME = 'Performance DB';
const DASHBOARD_SHEET_NAME = 'Dashboard';

const REVIEW_WINDOW_DAYS = 90; // 90日ルール
const DASHBOARD_TOP_N = 3;
const DASHBOARD_WATCH_DROP_N = 5;

// 商品マスターの列見出し（既存 + このスクリプトが自動追加する列）
const M = {
  SKU: 'SKU',
  NAME: '商品名',
  LIST_DATE: '出品日',
  ELAPSED: '経過日数',
  REMAIN: '残り日数(90日)',
  PRICE: '出品価格(USD)',
  VIEWS: 'Views',
  VIEWS_WOW: 'Views前週比',
  IMPRESSIONS: 'Impressions',
  IMPRESSIONS_WOW: 'Impressions前週比',
  CTR: 'CTR',
  CTR_WOW: 'CTR前週比',
  WATCHERS: 'Watchers',
  WATCHERS_WOW: 'Watchers前週比',
  WATCH_RATE: 'Watch率',
  WATCH_RATE_WOW: 'Watch率前週比',
  JUDGE: '判断',
  TODO: '今週やること',
  PRIORITY: '優先順位',
  ITEM_ID: 'eBay Item ID',
};

// 商品マスターに必須の既存列（無ければエラーで止める）
const REQUIRED_EXISTING_HEADERS = [
  M.SKU, M.NAME, M.LIST_DATE, M.PRICE, M.VIEWS, M.IMPRESSIONS,
  M.CTR, M.WATCHERS, M.JUDGE, M.ITEM_ID,
];

// 無ければ末尾に自動追加する列（この順番で）
const AUTO_APPEND_HEADERS = [
  M.REMAIN, M.VIEWS_WOW, M.IMPRESSIONS_WOW, M.CTR_WOW,
  M.WATCHERS_WOW, M.WATCH_RATE, M.WATCH_RATE_WOW, M.PRIORITY, M.TODO,
];

const PERF_DB_HEADERS = [
  '記録日', 'eBay Item ID', 'SKU', '商品名', '出品価格(USD)',
  'Views', 'Impressions', 'CTR', 'Watchers', 'Watch率',
];

// ------------------------------------------------------------------
// メニュー
// ------------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('eBay価格管理')
    .addItem('① 更新する（2つのCSVを商品マスターに反映）', 'runUpdate')
    .addSeparator()
    .addItem('初期セットアップ（シート/列を確認・作成）', 'setupSheets')
    .addToUi();
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureTrafficCsvSheet_(ss);
  ensureActiveCsvSheet_(ss);
  ensurePerfDbSheet_(ss);
  ensureDashboardSheet_(ss);
  const master = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!master) {
    throw new Error(`「${MASTER_SHEET_NAME}」シートが見つかりません。既存の商品マスターのシート名を確認してください。`);
  }
  ensureMasterHeaders_(master);
  SpreadsheetApp.getUi().alert('セットアップ完了：CSVインポート(2種) / Performance DB / Dashboard を確認・作成し、商品マスターの列を確認しました。');
}

// ------------------------------------------------------------------
// メイン処理
// ------------------------------------------------------------------
function runUpdate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const trafficSheet = ensureTrafficCsvSheet_(ss);
  const activeSheet = ensureActiveCsvSheet_(ss);
  const perfDbSheet = ensurePerfDbSheet_(ss);
  ensureDashboardSheet_(ss);
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    ui.alert(`「${MASTER_SHEET_NAME}」シートが見つかりません。`);
    return;
  }

  const headerCols = ensureMasterHeaders_(masterSheet);
  const trafficRecords = parseTrafficCsv_(trafficSheet);
  const activeRecords = parseActiveCsv_(activeSheet);
  const perfDbLatest = loadLatestPerfDbSnapshots_(perfDbSheet);

  const result = updateMasterRows_(masterSheet, headerCols, trafficRecords, activeRecords, perfDbLatest);
  applyPriority_(masterSheet, headerCols, result.rows);
  writePerfDbSnapshots_(perfDbSheet, result.rows);
  rebuildDashboard_(ss, result.rows);

  ui.alert(
    '更新完了\n' +
    `・Traffic Reportから Views/Impressions/CTR を反映: ${result.trafficMatchedCount}件\n` +
    `・Active Listingsから Watchers を反映: ${result.activeMatchedCount}件\n` +
    `・タイトル類似度で新規に eBay Item ID を自動紐づけ: ${result.autoLinkedCount}件\n` +
    `・Performance DB にスナップショットを追加: ${result.rows.length}件\n` +
    (result.unmatchedTitles.length
      ? `・どちらのCSVにも対応が見つからなかった商品（要確認）:\n  - ${result.unmatchedTitles.join('\n  - ')}`
      : '・CSV未対応の商品はありません')
  );
}

// ------------------------------------------------------------------
// シート準備
// ------------------------------------------------------------------
function ensureTrafficCsvSheet_(ss) {
  let sheet = ss.getSheetByName(CSV_TRAFFIC_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CSV_TRAFFIC_SHEET_NAME);
    sheet.getRange(1, 1).setValue('ここに Traffic Report を貼り付け（またはFile→Importで取り込み）してください');
  }
  return sheet;
}

function ensureActiveCsvSheet_(ss) {
  let sheet = ss.getSheetByName(CSV_ACTIVE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CSV_ACTIVE_SHEET_NAME);
    sheet.getRange(1, 1).setValue('ここに All active listings report を貼り付け（またはFile→Importで取り込み）してください');
  }
  return sheet;
}

function ensurePerfDbSheet_(ss) {
  let sheet = ss.getSheetByName(PERF_DB_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PERF_DB_SHEET_NAME);
    sheet.getRange(1, 1, 1, PERF_DB_HEADERS.length).setValues([PERF_DB_HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function ensureDashboardSheet_(ss) {
  let sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DASHBOARD_SHEET_NAME);
  }
  return sheet;
}

/**
 * 商品マスターの見出し行を確認し、無ければ末尾に自動追加する。
 * 戻り値: { ヘッダー名: 列番号(1始まり) } のマップ
 */
function ensureMasterHeaders_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const colByHeader = {};
  headerRow.forEach((h, i) => {
    const name = String(h || '').trim();
    if (name) colByHeader[name] = i + 1;
  });

  const missingRequired = REQUIRED_EXISTING_HEADERS.filter(h => !colByHeader[h]);
  if (missingRequired.length) {
    throw new Error(
      `商品マスターに必要な列が見つかりません: ${missingRequired.join(', ')}\n` +
      '列の見出し（1行目）の表記を確認してください。'
    );
  }

  let nextCol = lastCol;
  AUTO_APPEND_HEADERS.forEach(h => {
    if (!colByHeader[h]) {
      nextCol += 1;
      sheet.getRange(1, nextCol).setValue(h);
      colByHeader[h] = nextCol;
    }
  });

  return colByHeader;
}

// ------------------------------------------------------------------
// CSV読み込み共通ユーティリティ
// ------------------------------------------------------------------
/**
 * 指定した列名（複数候補）を含む行を見出し行として先頭20行から自動検出する。
 * Traffic Reportは見出し行の上に「Disclaimers」等の説明行が入ることがあるため、
 * 単純に1行目を見出しとせず、目印列を含む行を探す。
 */
function findHeaderRow_(values, markerHeaderLower) {
  for (let r = 0; r < Math.min(values.length, 20); r++) {
    const row = values[r];
    for (let c = 0; c < row.length; c++) {
      if (String(row[c] || '').trim().toLowerCase() === markerHeaderLower) {
        return r;
      }
    }
  }
  return -1;
}

function parseNumber_(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (typeof v === 'number') return v;
  const s = String(v).trim();
  if (s === '-' || s === '') return null;
  const n = parseFloat(s.replace(/,/g, ''));
  return isNaN(n) ? null : n;
}

function parsePercent_(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (typeof v === 'number') return v; // 既に 0.003 のような割合として保存されているケース
  const s = String(v).trim();
  if (s === '-' || s === '') return null;
  if (s.endsWith('%')) {
    const n = parseFloat(s.slice(0, -1));
    return isNaN(n) ? null : n / 100;
  }
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

// ------------------------------------------------------------------
// Traffic Report CSV（Views / Impressions / CTR）
// ------------------------------------------------------------------
function parseTrafficCsv_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headerRowIdx = findHeaderRow_(values, 'ebay item id');
  if (headerRowIdx === -1) {
    throw new Error(`「${CSV_TRAFFIC_SHEET_NAME}」で見出し行（eBay item ID を含む行）が見つかりません。Traffic Reportをそのまま貼り付けているか確認してください。`);
  }

  const headerMap = { itemId: -1, title: -1, impressions: -1, ctr: -1, views: -1 };
  values[headerRowIdx].forEach((cell, i) => {
    const h = String(cell || '').trim().toLowerCase();
    if (!h) return;
    if (h === 'ebay item id') headerMap.itemId = i;
    else if (h === 'listing title') headerMap.title = i;
    else if (h === 'total impressions') headerMap.impressions = i;
    else if (h.startsWith('click-through rate')) headerMap.ctr = i;
    else if (h === 'total page views') headerMap.views = i;
  });

  const records = [];
  for (let r = headerRowIdx + 1; r < values.length; r++) {
    const row = values[r];
    const itemId = String(row[headerMap.itemId] || '').trim();
    if (!itemId) continue;
    records.push({
      itemId,
      title: String(row[headerMap.title] || '').trim(),
      impressions: headerMap.impressions !== -1 ? parseNumber_(row[headerMap.impressions]) : null,
      ctr: headerMap.ctr !== -1 ? parsePercent_(row[headerMap.ctr]) : null,
      views: headerMap.views !== -1 ? parseNumber_(row[headerMap.views]) : null,
    });
  }
  return records;
}

// ------------------------------------------------------------------
// All active listings report CSV（Watchers）
// ------------------------------------------------------------------
function parseActiveCsv_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headerRowIdx = findHeaderRow_(values, 'item number');
  if (headerRowIdx === -1) {
    throw new Error(`「${CSV_ACTIVE_SHEET_NAME}」で見出し行（Item number を含む行）が見つかりません。All active listings reportをそのまま貼り付けているか確認してください。`);
  }

  const headerMap = { itemId: -1, title: -1, watchers: -1 };
  values[headerRowIdx].forEach((cell, i) => {
    const h = String(cell || '').trim().toLowerCase();
    if (!h) return;
    if (h === 'item number') headerMap.itemId = i;
    else if (h === 'title') headerMap.title = i;
    else if (h === 'watchers') headerMap.watchers = i;
  });

  const records = [];
  for (let r = headerRowIdx + 1; r < values.length; r++) {
    const row = values[r];
    const itemId = String(row[headerMap.itemId] || '').trim();
    if (!itemId) continue;
    records.push({
      itemId,
      title: String(row[headerMap.title] || '').trim(),
      watchers: headerMap.watchers !== -1 ? parseNumber_(row[headerMap.watchers]) : null,
    });
  }
  return records;
}

// ------------------------------------------------------------------
// 商品マスター更新（Views/Impressions/CTR/Watchersを2つのCSVから自動反映）
// ------------------------------------------------------------------
function updateMasterRows_(sheet, col, trafficRecords, activeRecords, perfDbLatest) {
  const lastRow = sheet.getLastRow(); // ← 固定行数ではなく毎回最終行まで自動取得（下段が反映されないバグの修正点）
  if (lastRow < 2) {
    return { rows: [], trafficMatchedCount: 0, activeMatchedCount: 0, autoLinkedCount: 0, unmatchedTitles: [] };
  }

  const numCols = sheet.getLastColumn();
  const range = sheet.getRange(2, 1, lastRow - 1, numCols);
  const values = range.getValues();

  const trafficById = new Map(trafficRecords.map(r => [r.itemId, r]));
  const activeById = new Map(activeRecords.map(r => [r.itemId, r]));
  const usedIds = new Set();

  // 既にeBay Item IDが入っている行は、そのIDをリンク済みとして扱う
  values.forEach(row => {
    const itemId = String(row[col[M.ITEM_ID] - 1] || '').trim();
    if (itemId) usedIds.add(itemId);
  });

  const unusedTraffic = () => trafficRecords.filter(r => !usedIds.has(r.itemId));
  const unusedActive = () => activeRecords.filter(r => !usedIds.has(r.itemId));

  const today = new Date();
  const rows = [];
  const unmatchedTitles = [];
  let trafficMatchedCount = 0;
  let activeMatchedCount = 0;
  let autoLinkedCount = 0;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const sku = String(row[col[M.SKU] - 1] || '').trim();
    const name = String(row[col[M.NAME] - 1] || '').trim();
    if (!sku && !name) continue; // 空行はスキップ

    let itemId = String(row[col[M.ITEM_ID] - 1] || '').trim();

    // ID未紐づけの新規商品は、タイトル類似度でCSVの行を自動で探して紐づける
    if (!itemId && name) {
      let best = findBestTitleMatch_(name, unusedTraffic());
      if (!best) best = findBestTitleMatch_(name, unusedActive());
      if (best) {
        itemId = best.itemId;
        row[col[M.ITEM_ID] - 1] = itemId; // 商品マスターに自動で書き戻す（次回からIDで直接一致する）
        usedIds.add(itemId);
        autoLinkedCount++;
      }
    }

    const trafficRecord = itemId ? trafficById.get(itemId) : null;
    const activeRecord = itemId ? activeById.get(itemId) : null;

    if (trafficRecord) {
      if (trafficRecord.views !== null) row[col[M.VIEWS] - 1] = trafficRecord.views;
      if (trafficRecord.impressions !== null) row[col[M.IMPRESSIONS] - 1] = trafficRecord.impressions;
      if (trafficRecord.ctr !== null) row[col[M.CTR] - 1] = trafficRecord.ctr;
      trafficMatchedCount++;
    }
    if (activeRecord) {
      if (activeRecord.watchers !== null) row[col[M.WATCHERS] - 1] = activeRecord.watchers;
      activeMatchedCount++;
    }
    if (!trafficRecord && !activeRecord && name) {
      unmatchedTitles.push(name);
    }

    const views = parseNumber_(row[col[M.VIEWS] - 1]);
    const watchers = parseNumber_(row[col[M.WATCHERS] - 1]);
    const impressions = parseNumber_(row[col[M.IMPRESSIONS] - 1]);
    const ctr = parsePercent_(row[col[M.CTR] - 1]);

    // Watch率 = Watchers ÷ Views で自動計算
    const watchRate = (views && views > 0 && watchers !== null) ? watchers / views : null;
    row[col[M.WATCH_RATE] - 1] = watchRate === null ? '' : watchRate;

    // 経過日数 / 残り日数(90日) を出品日から自動計算
    const listDate = row[col[M.LIST_DATE] - 1];
    if (listDate instanceof Date) {
      const elapsed = Math.floor((today - listDate) / 86400000);
      row[col[M.ELAPSED] - 1] = elapsed;
      row[col[M.REMAIN] - 1] = Math.max(REVIEW_WINDOW_DAYS - elapsed, 0);
    }

    // 前週比（Performance DBの前回スナップショットと比較）
    const prev = perfDbLatest.get(itemId);
    const watchRateDiff = diffOrNull_(watchRate, prev && prev.watchRate);
    row[col[M.VIEWS_WOW] - 1] = formatWow_(diffOrNull_(views, prev && prev.views), false);
    row[col[M.IMPRESSIONS_WOW] - 1] = formatWow_(diffOrNull_(impressions, prev && prev.impressions), false);
    row[col[M.WATCHERS_WOW] - 1] = formatWow_(diffOrNull_(watchers, prev && prev.watchers), false);
    row[col[M.CTR_WOW] - 1] = formatWow_(diffOrNull_(ctr, prev && prev.ctr), true);
    row[col[M.WATCH_RATE_WOW] - 1] = formatWow_(watchRateDiff, true);

    // 今週やること
    const judge = String(row[col[M.JUDGE] - 1] || '').trim();
    row[col[M.TODO] - 1] = decideTodo_(judge, watchRate, !prev);

    rows.push({
      rowIndex: i,
      sku, name, itemId,
      price: row[col[M.PRICE] - 1],
      views, impressions, ctr, watchers, watchRate, watchRateDiff,
      judge,
    });
  }

  range.setValues(values);
  // 表示フォーマット
  sheet.getRange(2, col[M.CTR], values.length, 1).setNumberFormat('0.00%');
  sheet.getRange(2, col[M.WATCH_RATE], values.length, 1).setNumberFormat('0.0%');

  return { rows, trafficMatchedCount, activeMatchedCount, autoLinkedCount, unmatchedTitles };
}

function diffOrNull_(current, previous) {
  if (current === null || current === undefined) return null;
  if (previous === null || previous === undefined) return null;
  return current - previous;
}

function formatWow_(diff, isPercentPoint) {
  if (diff === null || diff === undefined || isNaN(diff)) return ''; // 前回データなし（新規商品）
  if (isPercentPoint) {
    const pt = diff * 100;
    if (Math.abs(pt) < 0.05) return '→ 0';
    const sign = pt > 0 ? '⬆ +' : '⬇ ';
    return `${sign}${pt.toFixed(1)}pt`;
  }
  if (diff === 0) return '→ 0';
  const sign = diff > 0 ? '⬆ +' : '⬇ ';
  return `${sign}${diff}`;
}

function decideTodo_(judge, watchRate, isNew) {
  if (isNew) return '初回データ収集待ち';
  if (judge.indexOf('値下げ') !== -1) return '価格を見直す';
  if (judge.indexOf('相場') !== -1) return '相場を再調査する';
  if (judge.indexOf('様子見') !== -1) return '来週も様子見';
  if (watchRate !== null && watchRate >= 0.20) return '好調のため現状維持';
  return '現状維持';
}

// ------------------------------------------------------------------
// タイトル類似度による自動リンク（下段の新規商品対策）
// ------------------------------------------------------------------
function normalizeTitle_(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .split(/\s+/)
    .filter(t => t.length >= 2);
}

function jaccard_(aTokens, bTokens) {
  const a = new Set(aTokens);
  const b = new Set(bTokens);
  let intersection = 0;
  a.forEach(t => { if (b.has(t)) intersection++; });
  const union = new Set([...a, ...b]).size;
  return union === 0 ? 0 : intersection / union;
}

function findBestTitleMatch_(masterName, candidates) {
  const masterTokens = normalizeTitle_(masterName);
  if (!masterTokens.length || !candidates.length) return null;

  let best = null, bestScore = 0, secondScore = 0;
  candidates.forEach(c => {
    const score = jaccard_(masterTokens, normalizeTitle_(c.title));
    if (score > bestScore) {
      secondScore = bestScore;
      bestScore = score;
      best = c;
    } else if (score > secondScore) {
      secondScore = score;
    }
  });

  const MIN_SCORE = 0.3;
  const MIN_MARGIN = 0.1;
  if (best && bestScore >= MIN_SCORE && (bestScore - secondScore) >= MIN_MARGIN) {
    return best;
  }
  return null;
}

// ------------------------------------------------------------------
// 優先順位（Watch率が高い順）
// ------------------------------------------------------------------
function applyPriority_(sheet, col, rows) {
  const ranked = rows
    .filter(r => r.watchRate !== null)
    .sort((a, b) => (b.watchRate - a.watchRate) || ((b.views || 0) - (a.views || 0)));

  ranked.forEach((r, i) => { r.priority = i + 1; });

  const priorityByRowIndex = new Map(rows.map(r => [r.rowIndex, r.priority || '']));
  const numDataRows = sheet.getLastRow() - 1;
  const priorityCol = sheet.getRange(2, col[M.PRIORITY], numDataRows, 1);
  const priorityValues = [];
  for (let i = 0; i < numDataRows; i++) {
    priorityValues.push([priorityByRowIndex.has(i) ? priorityByRowIndex.get(i) : '']);
  }
  priorityCol.setValues(priorityValues);
}

// ------------------------------------------------------------------
// Performance DB（週次スナップショットの蓄積）
// ------------------------------------------------------------------
function loadLatestPerfDbSnapshots_(sheet) {
  const lastRow = sheet.getLastRow();
  const map = new Map();
  if (lastRow < 2) return map;

  const values = sheet.getRange(2, 1, lastRow - 1, PERF_DB_HEADERS.length).getValues();
  // 同一IDが複数あれば、シート下方（＝より新しい記録）で上書きすることで最新値を保持
  values.forEach(row => {
    const itemId = String(row[1] || '').trim();
    if (!itemId) return;
    map.set(itemId, {
      views: row[5] === '' ? null : row[5],
      impressions: row[6] === '' ? null : row[6],
      ctr: row[7] === '' ? null : row[7],
      watchers: row[8] === '' ? null : row[8],
      watchRate: row[9] === '' ? null : row[9],
    });
  });
  return map;
}

function writePerfDbSnapshots_(sheet, rows) {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy-MM-dd');
  const lastRow = sheet.getLastRow();

  // 同日・同IDの重複スナップショットを防ぐため、既存の当日分は上書き対象として位置を記録
  const existingTodayRowByItemId = new Map();
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, PERF_DB_HEADERS.length).getValues();
    values.forEach((row, i) => {
      const recordDate = row[0] instanceof Date
        ? Utilities.formatDate(row[0], Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy-MM-dd')
        : String(row[0]);
      if (recordDate === today) {
        existingTodayRowByItemId.set(String(row[1] || '').trim(), i + 2);
      }
    });
  }

  const newRows = [];
  rows.forEach(r => {
    if (!r.itemId) return; // IDが紐づかなかった商品は蓄積対象外（要確認リストで通知済み）
    const record = [
      today, r.itemId, r.sku, r.name, r.price,
      r.views === null ? '' : r.views,
      r.impressions === null ? '' : r.impressions,
      r.ctr === null ? '' : r.ctr,
      r.watchers === null ? '' : r.watchers,
      r.watchRate === null ? '' : r.watchRate,
    ];
    const existingRow = existingTodayRowByItemId.get(r.itemId);
    if (existingRow) {
      sheet.getRange(existingRow, 1, 1, record.length).setValues([record]);
    } else {
      newRows.push(record);
    }
  });

  if (newRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, PERF_DB_HEADERS.length).setValues(newRows);
  }
}

// ------------------------------------------------------------------
// Dashboard 再生成
// ------------------------------------------------------------------
function rebuildDashboard_(ss, rows) {
  const sheet = ensureDashboardSheet_(ss);
  sheet.clear();

  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
  sheet.getRange(1, 1).setValue('eBay価格管理 Dashboard');
  sheet.getRange(2, 1).setValue(`最終更新: ${now}`);

  const ranked = rows
    .filter(r => r.watchRate !== null)
    .sort((a, b) => (b.watchRate - a.watchRate) || ((b.views || 0) - (a.views || 0)));

  sheet.getRange(4, 1).setValue('今週の順位');
  sheet.getRange(5, 1, 1, 3).setValues([['順位', '商品', '理由']]);
  const top = ranked.slice(0, DASHBOARD_TOP_N);
  const topTable = top.map((r, i) => [i + 1, r.name, dashboardReason_(r)]);
  if (topTable.length) {
    sheet.getRange(6, 1, topTable.length, 3).setValues(topTable);
  }

  const dropStart = 6 + topTable.length + 2;
  sheet.getRange(dropStart, 1).setValue('⚠ Watch率が下がった商品（要チェック）');
  sheet.getRange(dropStart + 1, 1, 1, 3).setValues([['商品', 'Watch率', '前週比']]);
  const dropping = rows
    .filter(r => r.watchRateDiff !== null && r.watchRateDiff < 0)
    .sort((a, b) => a.watchRateDiff - b.watchRateDiff)
    .slice(0, DASHBOARD_WATCH_DROP_N);
  const dropTable = dropping.map(r => [
    r.name,
    `${(r.watchRate * 100).toFixed(1)}%`,
    formatWow_(r.watchRateDiff, true),
  ]);
  if (dropTable.length) {
    sheet.getRange(dropStart + 2, 1, dropTable.length, 3).setValues(dropTable);
  }

  sheet.setColumnWidths(1, 3, 220);
}

function dashboardReason_(r) {
  if (r.judge && r.judge.indexOf('値下げ') !== -1) return '価格見直し候補';
  const pct = (r.watchRate * 100).toFixed(1);
  return `Watch率${pct}%`;
}
