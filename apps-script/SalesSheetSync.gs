/**
 * SEN ARCHIVE 売上シート ⇔ 検品マスター 自動連携セットアップ
 *
 * 使い方:
 *   1. 「SEN ARCHIVE 売上シート」を開く
 *   2. 拡張機能 > Apps Script を開く
 *   3. このファイルの内容を貼り付けて保存
 *   4. エディタ上部の関数選択で setupSalesSheet を選び「実行」
 *      (初回はIMPORTRANGEの権限承認が必要。一度スプレッドシートに戻り、
 *       R2セルをクリックして「アクセスを許可」を押してから再実行する)
 *   5. 以降はスプレッドシートのメニュー「SEN ARCHIVE > 売上シートを再同期」からいつでも再実行可能
 *
 * 設計メモ: SKU(A列)は一度書き込んだら行の位置を固定し、検品マスターに新しいSKUが
 * 増えた時だけ末尾に追記する。Item Name(B列)・出品日(C列)はR:T列に置いた参照テーブルへの
 * VLOOKUPで取得する。こうすることで検品マスター側の行の並び替え・削除が起きても、
 * 既存の手入力データ(Channel/販売日/販売価格/Ad Fee/値引き額/送料実費/仕入れ値)が
 * 別のSKUの行にズレて紐付いてしまう事故を防ぐ。
 */

const MASTER_SPREADSHEET_URL =
  'https://docs.google.com/spreadsheets/d/166BARD0R6QVWFO7uwBQvnZsvmntC9W362UmVdXSXa0Y/edit';

// 検品マスターは 行1=簡易ヘッダー, 行2=正式ヘッダー, 行3以降=実データ という3段構成。
// SKUはCol6(F列)、ITEM NAMEはCol8(H列)、出品日はCol73(BU列 / "eBay出品日"だが
// メルカリ出品行にも日付が入っており全チャネル共通の出品日として使えることを確認済み)。
const MASTER_RANGE = '検品マスター!A3:BU';
const MASTER_QUERY = 'select Col6, Col8, Col73 where Col6 is not null';

const NEW_HEADERS = [
  'SKU', 'Item Name', '出品日', 'Channel', '販売日', '販売価格',
  '滞留日数', 'Transaction fee', 'Ad Fee', '値引き額(クーポン等)',
  '送料実費', 'Order earnings', '仕入れ値', '実質粗利', '粗利率',
];

// 旧レイアウトからの一度きりの移行時にSKUをキーにして復元する手入力項目
const MANUAL_FIELDS = [
  'Channel', '販売日', '販売価格', 'Ad Fee', '値引き額(クーポン等)', '送料実費', '仕入れ値',
];

const CHANNEL_OPTIONS = ['eBay', 'Mercari', 'Shopify', 'VC'];

// USD建ての Order earnings / 販売価格 をJPYの送料実費・仕入れ値と合算するための為替レート。
// P1にラベル、P2にGOOGLEFINANCEのライブレートを置き、各行の計算式から$P$2として参照する。
const FX_LABEL_CELL = 'P1';
const FX_RATE_CELL = 'P2';
const FX_RATE_ABS = '$P$2';

// 検品マスターの参照テーブル(内部用)。R2にQUERYを置き、SKU/Item Name/出品日をR:T列に展開する。
// B列・C列はここへVLOOKUPするので、検品マスターの行順が変わってもSKUで正しく引ける。
const HELPER_LABEL_CELL = 'R1';
const HELPER_FORMULA_CELL = 'R2';
const HELPER_COL = 18; // R列
const HELPER_RANGE_ABS = '$R$2:$T';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SEN ARCHIVE')
    .addItem('売上シートを再同期', 'setupSalesSheet')
    .addToUi();
}

function setupSalesSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

  const currentHeader = sheet.getRange(1, 1, 1, NEW_HEADERS.length).getValues()[0];
  const alreadyMigrated = NEW_HEADERS.every((h, i) => currentHeader[i] === h);

  const preserved = alreadyMigrated ? {} : snapshotManualData_(sheet);

  if (!alreadyMigrated) {
    sheet.clear();
    sheet.getRange(1, 1, 1, NEW_HEADERS.length)
      .setValues([NEW_HEADERS])
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else {
    detachOldQuerySpill_(sheet);
  }

  refreshHelperTable_(sheet);
  SpreadsheetApp.flush();
  Utilities.sleep(2000);

  const masterSkus = readHelperSkus_(sheet);
  if (masterSkus.length === 0) {
    throw new Error(
      '検品マスターからデータを取得できませんでした。スプレッドシートに戻りR2セルを' +
      'クリックして「アクセスを許可」を選択してから、もう一度 setupSalesSheet を実行してください。'
    );
  }

  appendNewSkuRows_(sheet, masterSkus);

  const numRows = lastRowInColumn_(sheet, 1) - 1;
  restoreManualData_(sheet, numRows, preserved);
  applyFxRateCell_(sheet);
  applyFormulaColumns_(sheet);
  applyFormatting_(sheet, numRows);
  applyChannelValidation_(sheet, numRows);
  protectAutoColumns_(sheet, numRows);

  SpreadsheetApp.flush();
}

// 列内で実際に値が入っている最終行を返す(getLastRow()はシート全体基準のためR:T列の
// 参照テーブルに引っ張られてしまうので、A列だけを見てSKUの終端行を判定する)
function lastRowInColumn_(sheet, col) {
  const maxRows = sheet.getMaxRows();
  const values = sheet.getRange(1, col, maxRows, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== '') return i + 1;
  }
  return 0;
}

// 旧バージョン(A2に3列分spillするQUERYを直接置いていた設計)からの移行:
// spillされている計算済みのSKU値を確定値として書き戻し、以後はA列を script が
// 追記管理するプレーンな値の列にする。
function detachOldQuerySpill_(sheet) {
  const a2Formula = sheet.getRange('A2').getFormula();
  if (!a2Formula) return;
  const lastRow = lastRowInColumn_(sheet, 1);
  if (lastRow < 2) {
    sheet.getRange('A2').clearContent();
    return;
  }
  const skuValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  sheet.getRange('A2').clearContent();
  SpreadsheetApp.flush();
  sheet.getRange(2, 1, skuValues.length, 1).setValues(skuValues);
}

function refreshHelperTable_(sheet) {
  sheet.getRange(HELPER_LABEL_CELL)
    .setValue('(内部用・触らない: 検品マスター参照テーブル)')
    .setFontColor('#999999');
  const formula = `=QUERY(IMPORTRANGE("${MASTER_SPREADSHEET_URL}", "${MASTER_RANGE}"), "${MASTER_QUERY}", 0)`;
  sheet.getRange(HELPER_FORMULA_CELL).setFormula(formula);
}

function readHelperSkus_(sheet) {
  const maxRows = sheet.getMaxRows();
  const values = sheet.getRange(2, HELPER_COL, maxRows - 1, 1).getValues();
  return values.map((r) => r[0]).filter((v) => v !== '' && v !== null && v !== '#REF!');
}

function appendNewSkuRows_(sheet, masterSkus) {
  const lastSkuRow = lastRowInColumn_(sheet, 1);
  const existing = new Set();
  if (lastSkuRow >= 2) {
    sheet.getRange(2, 1, lastSkuRow - 1, 1).getValues().forEach((r) => {
      if (r[0] !== '') existing.add(r[0]);
    });
  }
  const newSkus = masterSkus.filter((sku) => !existing.has(sku));
  if (newSkus.length === 0) return;
  const startRow = Math.max(lastSkuRow, 1) + 1;
  sheet.getRange(startRow, 1, newSkus.length, 1).setValues(newSkus.map((sku) => [sku]));
}

function snapshotManualData_(sheet) {
  const preserved = {};
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return preserved;

  const oldHeader = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const skuColIdx = oldHeader.indexOf('SKU');
  if (skuColIdx === -1) return preserved;

  const oldData = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  oldData.forEach((row) => {
    const sku = row[skuColIdx];
    if (!sku || String(sku).indexOf('(例)') === 0) return;
    const rec = {};
    MANUAL_FIELDS.forEach((field) => {
      const idx = oldHeader.indexOf(field);
      if (idx !== -1 && row[idx] !== '') rec[field] = row[idx];
    });
    preserved[sku] = rec;
  });
  return preserved;
}

function restoreManualData_(sheet, numRows, preserved) {
  if (numRows <= 0) return;
  const skuValues = sheet.getRange(2, 1, numRows, 1).getValues();
  for (let i = 0; i < numRows; i++) {
    const rec = preserved[skuValues[i][0]];
    if (!rec) continue;
    MANUAL_FIELDS.forEach((field) => {
      if (rec[field] === undefined) return;
      const colIdx = NEW_HEADERS.indexOf(field) + 1;
      sheet.getRange(i + 2, colIdx).setValue(rec[field]);
    });
  }
}

function applyFormulaColumns_(sheet) {
  const col = (name) => NEW_HEADERS.indexOf(name) + 1;

  // Item Name/出品日は検品マスターの行順ではなく、自分のSKU(A列)をキーに
  // 参照テーブル(R:T)からVLOOKUPする。検品マスター側で行が並び替わっても
  // このSKUに紐づく値だけが変わり、他の行には影響しない。
  sheet.getRange(2, col('Item Name')).setFormula(
    `=ARRAYFORMULA(IF($A2:A="","",IFERROR(VLOOKUP($A2:A,${HELPER_RANGE_ABS},2,FALSE),"")))`
  );
  sheet.getRange(2, col('出品日')).setFormula(
    `=ARRAYFORMULA(IF($A2:A="","",IFERROR(VLOOKUP($A2:A,${HELPER_RANGE_ABS},3,FALSE),"")))`
  );

  sheet.getRange(2, col('滞留日数')).setFormula(
    '=ARRAYFORMULA(IF($A2:A="","",IF($C2:C="","",IF($E2:E="",TODAY()-$C2:C,$E2:E-$C2:C))))'
  );
  sheet.getRange(2, col('Transaction fee')).setFormula(
    '=ARRAYFORMULA(IF($A2:A="","",IF($D2:D="Mercari",$F2:F*0.1,IF($D2:D="eBay",$F2:F*0.144,""))))'
  );
  sheet.getRange(2, col('Order earnings')).setFormula(
    '=ARRAYFORMULA(IF($A2:A="","",$F2:F-$H2:H-$I2:I-$J2:J))'
  );
  // Mercariは販売価格・Order earningsが元々JPYなので為替換算せず、
  // eBay/Shopify/VC(USD建て)の行だけ$P$2のレートをかけてJPYに揃える。1円未満は切り捨て。
  sheet.getRange(2, col('実質粗利')).setFormula(
    `=ARRAYFORMULA(IF($A2:A="","",ROUNDDOWN(IF($D2:D="Mercari",$L2:L-$K2:K-$M2:M,($L2:L*${FX_RATE_ABS})-$K2:K-$M2:M),0)))`
  );
  sheet.getRange(2, col('粗利率')).setFormula(
    `=ARRAYFORMULA(IF($A2:A="","",IF($F2:F="","",IF($D2:D="Mercari",$N2:N/$F2:F,$N2:N/($F2:F*${FX_RATE_ABS})))))`
  );
}

function applyFxRateCell_(sheet) {
  sheet.getRange(FX_LABEL_CELL).setValue('USD/JPY').setFontWeight('bold');
  sheet.getRange(FX_RATE_CELL).setFormula('=GOOGLEFINANCE("CURRENCY:USDJPY")').setNumberFormat('0.00');
}

function applyFormatting_(sheet, numRows) {
  const col = (name) => NEW_HEADERS.indexOf(name) + 1;
  sheet.getRange(2, col('出品日'), numRows, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, col('販売日'), numRows, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, col('粗利率'), numRows, 1).setNumberFormat('0.0%');
}

function applyChannelValidation_(sheet, numRows) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CHANNEL_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, NEW_HEADERS.indexOf('Channel') + 1, numRows, 1).setDataValidation(rule);
}

function protectAutoColumns_(sheet, numRows) {
  const existing = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  existing.forEach((p) => {
    if (p.getDescription() === 'auto-sync (検品マスターから自動反映・編集不可)') p.remove();
  });

  const range = sheet.getRange(2, 1, numRows, 3); // A:C = SKU, Item Name, 出品日
  const protection = range.protect().setDescription('auto-sync (検品マスターから自動反映・編集不可)');
  protection.setWarningOnly(true); // 警告のみ。誤って上書きしそうになった時にダイアログで注意喚起する
}
