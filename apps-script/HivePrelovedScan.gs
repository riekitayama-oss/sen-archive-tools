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
 *   - 素材・カラー・ライン名・柄名は tags/title からキーワード一致で推測しているだけで、
 *     専用フィールドではない(HIVE_MATERIAL_KEYWORDS等の各リストで調整)
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
 *
 * 「仕入れ値上限」について: HIVEの販売価格そのものが実際の落札可能額の保証では
 * ないため、これは「HIVEと同程度の価格で売れると仮定した場合、粗利率20%/30%を
 * 確保するにはいくら以下で仕入れる必要があるか」という参考の目安。想定手数料率・
 * 想定送料実費は仕入れ判断シートのP2/P3セルで調整できる(初期値: eBay手数料14.4%、
 * 送料1500円)。
 *
 * 「型番タグ」について: "FENDI Baguette Zucca" と "FENDI Baguette Leopard" を
 * 別物として集計したい、という要望に対応するため、既知のライン名(Baguette,
 * Speedy等)と柄名(Zucca, Monogram, Leopard等)をタイトルから検出し、
 * 「ライン名+柄/素材」の組み合わせをタグとして作る。HIVE_KNOWN_LINES /
 * HIVE_PATTERN_KEYWORDS は網羅的ではないので、集計結果を見ながら追加すること。
 */

const HIVE_STORE_DOMAIN = 'hivepreloved.com';

// 実際のコレクションURL(https://hivepreloved.com/collections/xxx)のxxx部分に
// 合わせて要確認・修正。
const HIVE_COLLECTIONS = [
  { handle: 'bags', label: 'Bags' },
  { handle: 'shoes', label: 'Shoes' },
];

const HIVE_SHEET_NAME = '仕入れ判断シート';
const HIVE_SUMMARY_SHEET_NAME = '集計サマリー';
const HIVE_MASTER_SHEET_NAME = '検品マスター';
// 検品マスターの実データは行3以降(行1=簡易ヘッダー, 行2=正式ヘッダー)。
// SalesSheetSync.gs側の調査で判明済み: BRANDはCol7(G列)、ITEM NAMEはCol8(H列)。
const HIVE_MASTER_DATA_START_ROW = 3;
const HIVE_MASTER_BRAND_COL = 7;
const HIVE_MASTER_TITLE_COL = 8;

const HIVE_OUTPUT_HEADERS = [
  'ブランド', '型番/アイテム名', 'カテゴリ', '素材(推定)', 'カラー(推定)',
  '型番タグ(ライン+柄/素材・推定)',
  '価格(HIVE表示・通貨未確認)', 'Sold Out確認日', '未仕入れフラグ',
  'HIVE価格(JPY換算)', '仕入れ値上限(粗利20%狙い)', '仕入れ値上限(粗利30%狙い)',
  '商品ページURL', '_handle(内部キー・触らない)',
];

// 仕入れ値上限の計算に使う前提値を置くセル。P2/P3は初期値を入れるだけで、
// 以後はユーザーが自由に上書きしてよい(スクリプトは空でない限り再設定しない)。
const HIVE_CONFIG_FX_LABEL_CELL = 'O1';
const HIVE_CONFIG_FX_RATE_CELL = 'P1';
const HIVE_CONFIG_FX_RATE_ABS = '$P$1';
const HIVE_CONFIG_FEE_LABEL_CELL = 'O2';
const HIVE_CONFIG_FEE_RATE_CELL = 'P2';
const HIVE_CONFIG_FEE_RATE_ABS = '$P$2';
const HIVE_CONFIG_SHIP_LABEL_CELL = 'O3';
const HIVE_CONFIG_SHIP_COST_CELL = 'P3';
const HIVE_CONFIG_SHIP_COST_ABS = '$P$3';

const HIVE_MATERIAL_KEYWORDS = [
  'Leather', 'Suede', 'Canvas', 'Cotton', 'Wool', 'Silk', 'Nylon', 'Denim',
  'Cashmere', 'Linen', 'Polyester', 'PVC', 'Velvet', 'Satin', 'Tweed', 'Fur',
];
const HIVE_COLOR_KEYWORDS = [
  'Black', 'Brown', 'Red', 'Blue', 'Green', 'White', 'Beige', 'Pink', 'Yellow',
  'Grey', 'Gray', 'Navy', 'Gold', 'Silver', 'Purple', 'Orange', 'Tan', 'Cream',
  'Multicolor', 'Ivory', 'Khaki',
];

// 柄・モノグラム名。素材ではないが、同じ素材でも柄違いで人気が大きく変わることが
// 多いブランド品では重要な区別軸(例: FENDIのZucca vs Leopard)。
const HIVE_PATTERN_KEYWORDS = [
  'Zucca', 'Monogram', 'Damier', 'GG', 'Toile', 'Python', 'Snake', 'Croc',
  'Crocodile', 'Leopard', 'Cheetah', 'Zebra', 'Houndstooth', 'Tie-Dye',
  'Ombre', 'Sequin', 'Metallic', 'Quilted', 'Chevron', 'Herringbone',
  'Tartan', 'Plaid',
];

// 主要ブランドの代表的なライン名(非網羅的・要追加)。"FENDI Baguette Zucca" の
// ように、ブランド名の直後に来ることが多い型番/ライン名をタイトルから検出する。
const HIVE_KNOWN_LINES = [
  'Baguette', 'Peekaboo', 'Mon Tresor', // Fendi
  'Speedy', 'Neverfull', 'Alma', 'Keepall', 'Twist', 'Capucines', 'Petite Malle', // Louis Vuitton
  'Classic Flap', 'Boy', '2.55', '19 Bag', 'Coco Handle', // Chanel
  'Kelly', 'Birkin', 'Constance', 'Evelyne', 'Picotin', 'Garden Party', // Hermes
  'Lady Dior', 'Saddle', 'Book Tote', 'Caro', // Dior
  'GG Marmont', 'Dionysus', 'Jackie', 'Soho', 'Sylvie', // Gucci
  'Antigona', 'Nightingale', // Givenchy
  'Trapeze', 'Bayswater', // Mulberry
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
  ensureHiveConfigCells_(outSheet);
  const masterIndex = buildHiveMasterIndex_(ss);
  const existingHandles = readHiveExistingHandles_(outSheet);

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const rows = [];
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
      rows.push({
        brand: p.vendor || '',
        title: p.title || '',
        category: col.label,
        material: extractHiveMaterial_(p),
        color: extractHiveColor_(p),
        tag: extractHiveTag_(p),
        price: price,
        date: today,
        flag: computeHiveFlag_(p, masterIndex),
        url: `https://${HIVE_STORE_DOMAIN}/products/${p.handle}`,
        handle: p.handle,
      });
    });
  });

  if (rows.length === 0) {
    Logger.log('新しいSold Out商品は見つかりませんでした。');
    return;
  }

  const startRow = outSheet.getLastRow() + 1;
  const values = rows.map((r, idx) => {
    const rowNum = startRow + idx;
    return [
      r.brand, r.title, r.category, r.material, r.color, r.tag, r.price, r.date, r.flag,
      // HIVE価格(JPY換算) : G列(価格)
      `=IF(G${rowNum}="","",G${rowNum}*${HIVE_CONFIG_FX_RATE_ABS})`,
      // 仕入れ値上限(粗利20%狙い) = JPY価格×(1-想定手数料率-0.20) - 想定送料実費
      `=IF(J${rowNum}="","",ROUNDDOWN(J${rowNum}*(1-${HIVE_CONFIG_FEE_RATE_ABS}-0.2)-${HIVE_CONFIG_SHIP_COST_ABS},0))`,
      // 仕入れ値上限(粗利30%狙い)
      `=IF(J${rowNum}="","",ROUNDDOWN(J${rowNum}*(1-${HIVE_CONFIG_FEE_RATE_ABS}-0.3)-${HIVE_CONFIG_SHIP_COST_ABS},0))`,
      r.url,
      r.handle,
    ];
  });

  outSheet.getRange(startRow, 1, values.length, HIVE_OUTPUT_HEADERS.length).setValues(values);
  ensureHiveSummarySheet_(ss);
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

// 仕入れ値上限の計算に使う前提値。P1は常にライブレートで上書き、
// P2/P3は空の時だけ初期値を入れる(ユーザーが調整した値は壊さない)。
function ensureHiveConfigCells_(sheet) {
  sheet.getRange(HIVE_CONFIG_FX_LABEL_CELL).setValue('想定為替レート(USD/JPY)').setFontWeight('bold');
  sheet.getRange(HIVE_CONFIG_FX_RATE_CELL).setFormula('=GOOGLEFINANCE("CURRENCY:USDJPY")').setNumberFormat('0.00');

  sheet.getRange(HIVE_CONFIG_FEE_LABEL_CELL).setValue('想定販売手数料率(例: eBay 14.4%)').setFontWeight('bold');
  if (sheet.getRange(HIVE_CONFIG_FEE_RATE_CELL).getValue() === '') {
    sheet.getRange(HIVE_CONFIG_FEE_RATE_CELL).setValue(0.144);
  }
  sheet.getRange(HIVE_CONFIG_FEE_RATE_CELL).setNumberFormat('0.0%');

  sheet.getRange(HIVE_CONFIG_SHIP_LABEL_CELL).setValue('想定送料実費(円)').setFontWeight('bold');
  if (sheet.getRange(HIVE_CONFIG_SHIP_COST_CELL).getValue() === '') {
    sheet.getRange(HIVE_CONFIG_SHIP_COST_CELL).setValue(1500);
  }
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

// "ライン名 + 柄/素材" の組み合わせタグを作る(例: "Baguette Zucca",
// "Baguette Leopard", "Baguette Wool")。柄名が見つからない時は素材名、
// それも無ければ色名にフォールバックする。ラインが見つからない場合は
// 柄/素材/色だけのタグになる。
function extractHiveTag_(product) {
  const line = findHiveKeyword_(product, HIVE_KNOWN_LINES);
  const modifier =
    findHiveKeyword_(product, HIVE_PATTERN_KEYWORDS) ||
    findHiveKeyword_(product, HIVE_MATERIAL_KEYWORDS) ||
    findHiveKeyword_(product, HIVE_COLOR_KEYWORDS);
  return [line, modifier].filter(Boolean).join(' ');
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

// 仕入れ判断シートの蓄積データから、ブランド×カテゴリ×型番タグ／素材／カラー、
// および ブランド×カテゴリ単体で Sold Out件数と平均価格・平均仕入れ上限額を
// 集計するタブ。データが増えるほど「このライン+柄はよく売れる」といった
// 傾向が見えてくる想定。
function ensureHiveSummarySheet_(ss) {
  let sheet = ss.getSheetByName(HIVE_SUMMARY_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(HIVE_SUMMARY_SHEET_NAME);

  const countLabel = "count(A) '件数'";
  const avgLabels = "avg(J) '平均HIVE価格(JPY)', avg(K) '平均仕入上限(粗利20%)'";
  const range = `'${HIVE_SHEET_NAME}'!A2:L`;

  // 最優先: FENDI Baguette Zucca / Leopard / Wool のように、ライン+柄/素材の
  // 組み合わせ単位で見たいという要望に応える集計。
  sheet.getRange('A1').setValue('ブランド×カテゴリ×型番タグ 別 Sold Out集計').setFontWeight('bold');
  sheet.getRange('A2').setFormula(
    `=IFERROR(QUERY(${range}, "select A, C, F, count(A), avg(J), avg(K) ` +
    `where A is not null and F is not null and F <> '' group by A, C, F order by count(A) desc ` +
    `label ${countLabel}, ${avgLabels}", 0), "データがまだありません")`
  );

  sheet.getRange('H1').setValue('ブランド×カテゴリ×素材 別 Sold Out集計').setFontWeight('bold');
  sheet.getRange('H2').setFormula(
    `=IFERROR(QUERY(${range}, "select A, C, D, count(A), avg(J), avg(K) ` +
    `where A is not null and D is not null and D <> '' group by A, C, D order by count(A) desc ` +
    `label ${countLabel}, ${avgLabels}", 0), "データがまだありません")`
  );

  sheet.getRange('O1').setValue('ブランド×カテゴリ×カラー 別 Sold Out集計').setFontWeight('bold');
  sheet.getRange('O2').setFormula(
    `=IFERROR(QUERY(${range}, "select A, C, E, count(A), avg(J), avg(K) ` +
    `where A is not null and E is not null and E <> '' group by A, C, E order by count(A) desc ` +
    `label ${countLabel}, ${avgLabels}", 0), "データがまだありません")`
  );

  sheet.getRange('V1').setValue('ブランド×カテゴリ 別 Sold Out集計').setFontWeight('bold');
  sheet.getRange('V2').setFormula(
    `=IFERROR(QUERY(${range}, "select A, C, count(A), avg(J), avg(K) ` +
    `where A is not null group by A, C order by count(A) desc ` +
    `label ${countLabel}, ${avgLabels}", 0), "データがまだありません")`
  );

  sheet.getRange('D3:F500').setNumberFormat('#,##0');
  sheet.getRange('K3:M500').setNumberFormat('#,##0');
  sheet.getRange('R3:T500').setNumberFormat('#,##0');
  sheet.getRange('Y3:Z500').setNumberFormat('#,##0');
}
