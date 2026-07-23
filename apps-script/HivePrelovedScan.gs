/**
 * HIVE PRELOVED (hivepreloved.com) 週次スキャン → 仕入れ判断シート
 *
 * 【重要な前提】このスクリプトはShopifyストアが標準で公開している
 * /collections/{handle}/products.json エンドポイントを使う設計になっている。
 * 開発時点でHIVE PRELOVEDの実際のページ構造を確認できなかったため、
 * 以下は "Shopifyの一般的な慣習" に基づく推測であり、初回実行後に
 * 実際の結果を見て調整が必要になる可能性が高い:
 *   - COLLECTIONS の handle (URLの一部。実際にブラウザで
 *     https://hivepreloved.com/collections/bags を開いて確認・修正すること)
 *   - ブランド = product.vendor という前提(店名が入っていた場合は要修正)
 *   - 素材・カラーは tags/title からキーワード一致で推測しているだけで、
 *     専用フィールドではない(MATERIAL_KEYWORDS / COLOR_KEYWORDSで調整)
 *   - 価格の通貨は未確認(variants[].priceをそのまま数値化しているだけ)
 *
 * 使い方:
 *   1. 検品マスターのスプレッドシートを開く
 *   2. 拡張機能 > Apps Script を開く（売上シート用とは別のスクリプトプロジェクト）
 *   3. このファイルの内容を貼り付けて保存
 *   4. scanHivePreloved を一度手動実行して結果を確認・調整する
 *   5. 問題なければメニュー「HIVE PRELOVED調査 > 週次自動実行を設定」を実行
 *      (以後は毎週自動でscanHivePrelovedが走り、シートに追記される)
 *
 * 「未仕入れフラグ」はブランド名・商品名の単語一致による目安であり、
 * 最終判断は人間の確認が必要(型番のような構造化データが双方に無いため)。
 * 「推定利益率20-30%」はHIVE側のデータから計算した数字ではなく、
 * 固定の参考値としてそのまま書き出しているだけ。
 */

const HIVE_STORE_DOMAIN = 'hivepreloved.com';

// 実際のコレクションURL(https://hivepreloved.com/collections/xxx)のxxx部分に
// 合わせて要確認・修正。
const HIVE_COLLECTIONS = [
  { handle: 'bags', label: 'Bags' },
  { handle: 'shoes', label: 'Shoes' },
];

const HIVE_SHEET_NAME = '仕入れ判断シート';
const HIVE_MASTER_SHEET_NAME = '検品マスター';
// 検品マスターの実データは行3以降(行1=簡易ヘッダー, 行2=正式ヘッダー)。
// SalesSheetSync.gs側の調査で判明済み: BRANDはCol7(G列)、ITEM NAMEはCol8(H列)。
const HIVE_MASTER_DATA_START_ROW = 3;
const HIVE_MASTER_BRAND_COL = 7;
const HIVE_MASTER_TITLE_COL = 8;

const HIVE_OUTPUT_HEADERS = [
  'ブランド', '型番/アイテム名', 'カテゴリ', '素材(推定)', 'カラー(推定)',
  '価格', 'Sold Out確認日', '未仕入れフラグ', '推定利益率(参考値・確定値ではない)',
  '商品ページURL', '_handle(内部キー・触らない)',
];

const HIVE_MATERIAL_KEYWORDS = [
  'Leather', 'Suede', 'Canvas', 'Cotton', 'Wool', 'Silk', 'Nylon', 'Denim',
  'Cashmere', 'Linen', 'Polyester', 'PVC', 'Velvet', 'Satin', 'Tweed', 'Fur',
];
const HIVE_COLOR_KEYWORDS = [
  'Black', 'Brown', 'Red', 'Blue', 'Green', 'White', 'Beige', 'Pink', 'Yellow',
  'Grey', 'Gray', 'Navy', 'Gold', 'Silver', 'Purple', 'Orange', 'Tan', 'Cream',
  'Multicolor', 'Ivory', 'Khaki',
];
const HIVE_STOPWORDS = new Set([
  'the', 'and', 'with', 'for', 'of', 'in', 'a', 'an', 'size', 'us', 'eu', 'cm',
  'new', 'preloved', 'pre-owned', 'vintage',
]);

// 未仕入れ判定の類似度しきい値(Jaccard係数)。これ未満なら「似た商品が
// 検品マスターに無い」とみなす。0〜1の値で、低いほどフラグが立ちやすくなる。
const HIVE_SIMILARITY_THRESHOLD = 0.3;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('HIVE PRELOVED調査')
    .addItem('今すぐ調査を実行', 'scanHivePreloved')
    .addItem('週次自動実行を設定', 'installHiveWeeklyTrigger')
    .addToUi();
}

function scanHivePreloved() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const outSheet = ensureHiveOutputSheet_(ss);
  const masterIndex = buildHiveMasterIndex_(ss);
  const existingHandles = readHiveExistingHandles_(outSheet);

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const rowsToAppend = [];
  const seenInThisRun = new Set();

  HIVE_COLLECTIONS.forEach((col) => {
    let products;
    try {
      products = fetchHiveCollectionProducts_(col.handle);
    } catch (e) {
      Logger.log('コレクション取得失敗: ' + col.handle + ' / ' + e);
      return;
    }
    products.forEach((p) => {
      if (!isHiveSoldOut_(p)) return;
      if (existingHandles.has(p.handle) || seenInThisRun.has(p.handle)) return;
      seenInThisRun.add(p.handle);

      const price = p.variants && p.variants[0] ? p.variants[0].price : '';
      const flag = computeHiveFlag_(p, masterIndex);
      rowsToAppend.push([
        p.vendor || '',
        p.title || '',
        col.label,
        extractHiveMaterial_(p),
        extractHiveColor_(p),
        price,
        today,
        flag,
        '20-30%(推測値・HIVE側データから確定できる数字ではない)',
        `https://${HIVE_STORE_DOMAIN}/products/${p.handle}`,
        p.handle,
      ]);
    });
  });

  if (rowsToAppend.length === 0) {
    Logger.log('新しいSold Out商品は見つかりませんでした。');
    return;
  }

  const startRow = outSheet.getLastRow() + 1;
  outSheet.getRange(startRow, 1, rowsToAppend.length, HIVE_OUTPUT_HEADERS.length)
    .setValues(rowsToAppend);
}

function installHiveWeeklyTrigger() {
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === 'scanHivePreloved') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('scanHivePreloved')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(6)
    .everyWeeks(1)
    .create();
  SpreadsheetApp.getUi().alert('毎週月曜 朝6時頃に自動実行するよう設定しました。');
}

function ensureHiveOutputSheet_(ss) {
  let sheet = ss.getSheetByName(HIVE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(HIVE_SHEET_NAME);
  }
  const header = sheet.getRange(1, 1, 1, HIVE_OUTPUT_HEADERS.length).getValues()[0];
  const matches = HIVE_OUTPUT_HEADERS.every((h, i) => header[i] === h);
  if (!matches) {
    sheet.getRange(1, 1, 1, HIVE_OUTPUT_HEADERS.length)
      .setValues([HIVE_OUTPUT_HEADERS])
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readHiveExistingHandles_(sheet) {
  const handles = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return handles;
  const handleCol = HIVE_OUTPUT_HEADERS.length; // _handle は最終列
  sheet.getRange(2, handleCol, lastRow - 1, 1).getValues().forEach((r) => {
    if (r[0] !== '') handles.add(r[0]);
  });
  return handles;
}

function buildHiveMasterIndex_(ss) {
  const sheet = ss.getSheetByName(HIVE_MASTER_SHEET_NAME);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < HIVE_MASTER_DATA_START_ROW) return [];
  const numRows = lastRow - HIVE_MASTER_DATA_START_ROW + 1;
  const brands = sheet.getRange(HIVE_MASTER_DATA_START_ROW, HIVE_MASTER_BRAND_COL, numRows, 1).getValues();
  const titles = sheet.getRange(HIVE_MASTER_DATA_START_ROW, HIVE_MASTER_TITLE_COL, numRows, 1).getValues();
  const index = [];
  for (let i = 0; i < numRows; i++) {
    const brand = brands[i][0];
    const title = titles[i][0];
    if (!brand && !title) continue;
    index.push({ brand: String(brand || ''), title: String(title || '') });
  }
  return index;
}

// Shopifyの /collections/{handle}/products.json をページングしながら全件取得する。
function fetchHiveCollectionProducts_(handle) {
  const all = [];
  let page = 1;
  while (true) {
    const url = `https://${HIVE_STORE_DOMAIN}/collections/${handle}/products.json?limit=250&page=${page}`;
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) break;
    const data = JSON.parse(res.getContentText());
    const products = data.products || [];
    all.push(...products);
    if (products.length < 250) break;
    page++;
  }
  return all;
}

function isHiveSoldOut_(product) {
  if (!product.variants || product.variants.length === 0) return false;
  return product.variants.every((v) => v.available === false);
}

function extractHiveMaterial_(product) {
  return findHiveKeyword_(product, HIVE_MATERIAL_KEYWORDS);
}

function extractHiveColor_(product) {
  // "Color"/"Colour" というoptionがあれば、そちらを優先的に使う方が
  // キーワード推測より信頼できる。
  if (product.options) {
    const idx = product.options.findIndex((o) =>
      /color|colour/i.test(o.name || '')
    );
    if (idx !== -1 && product.variants && product.variants[0]) {
      const key = 'option' + (idx + 1);
      const val = product.variants[0][key];
      if (val) return val;
    }
  }
  return findHiveKeyword_(product, HIVE_COLOR_KEYWORDS);
}

function findHiveKeyword_(product, keywords) {
  const haystack = [
    (product.tags || []).join(' '),
    product.title || '',
  ].join(' ').toLowerCase();
  const hit = keywords.find((k) => haystack.indexOf(k.toLowerCase()) !== -1);
  return hit || '';
}

function computeHiveFlag_(hiveProduct, masterIndex) {
  const brand = (hiveProduct.vendor || '').trim().toLowerCase();
  if (!brand) return '要確認(ブランド不明)';

  const brandMatches = masterIndex.filter((m) => m.brand.trim().toLowerCase() === brand);
  if (brandMatches.length === 0) return '未仕入れ(ブランド自体が検品マスターに無い)';

  const hiveTokens = hiveTokenize_(hiveProduct.title || '');
  let best = 0;
  brandMatches.forEach((m) => {
    const score = hiveOverlapScore_(hiveTokens, hiveTokenize_(m.title));
    if (score > best) best = score;
  });

  return best < HIVE_SIMILARITY_THRESHOLD
    ? '未仕入れ(類似デザインが見当たらない・要確認)'
    : '取扱いあり(類似商品が既にある可能性)';
}

function hiveTokenize_(str) {
  return String(str)
    .toLowerCase()
    .split(/[^a-z0-9]+/)
    .filter((t) => t.length > 1 && !HIVE_STOPWORDS.has(t));
}

function hiveOverlapScore_(tokensA, tokensB) {
  if (tokensA.length === 0 || tokensB.length === 0) return 0;
  const setA = new Set(tokensA);
  const setB = new Set(tokensB);
  let intersection = 0;
  setA.forEach((t) => { if (setB.has(t)) intersection++; });
  const union = new Set([...setA, ...setB]).size;
  return union === 0 ? 0 : intersection / union;
}
