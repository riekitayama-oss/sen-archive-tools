/**
 * SEN ARCHIVE 売上シート ⇔ 検品マスター 自動連携セットアップ
 *
 * 使い方:
 *   1. 「SEN ARCHIVE 売上シート」を開く
 *   2. 拡張機能 > Apps Script を開く
 *   3. このファイルの内容を貼り付けて保存
 *   4. エディタ上部の関数選択で setupSalesSheet を選び「実行」
 *      (初回はIMPORTRANGEの権限承認が必要。一度スプレッドシートに戻り、
 *       A2セルをクリックして「アクセスを許可」を押してから再実行する)
 *   5. 以降はスプレッドシートのメニュー「SEN ARCHIVE > 売上シートを再同期」からいつでも再実行可能
 *
 * 何度実行しても安全（冪等）。既存の手入力データ(Channel/販売日/販売価格/Ad Fee/
 * 値引き額/送料実費/仕入れ値)はSKUをキーに読み込み直してから新しい列配置に復元する。
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

// 移行時にSKUをキーにして復元する手入力項目
const MANUAL_FIELDS = [
  'Channel', '販売日', '販売価格', 'Ad Fee', '値引き額(クーポン等)', '送料実費', '仕入れ値',
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SEN ARCHIVE')
    .addItem('売上シートを再同期', 'setupSalesSheet')
    .addToUi();
}

function setupSalesSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

  const preserved = snapshotManualData_(sheet);

  sheet.clear();
  sheet.clearNote();
  sheet.getRange(1, 1, 1, NEW_HEADERS.length)
    .setValues([NEW_HEADERS])
    .setFontWeight('bold');
  sheet.setFrozenRows(1);

  const formula = `=QUERY(IMPORTRANGE("${MASTER_SPREADSHEET_URL}", "${MASTER_RANGE}"), "${MASTER_QUERY}", 0)`;
  sheet.getRange('A2').setFormula(formula);
  SpreadsheetApp.flush();
  Utilities.sleep(2000);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || sheet.getRange('A2').getValue() === '#REF!') {
    throw new Error(
      '検品マスターからデータを取得できませんでした。スプレッドシートに戻りA2セルを' +
      'クリックして「アクセスを許可」を選択してから、もう一度 setupSalesSheet を実行してください。'
    );
  }
  const numRows = lastRow - 1;

  restoreManualData_(sheet, numRows, preserved);
  applyFormulaColumns_(sheet);
  applyFormatting_(sheet, numRows);
  protectAutoColumns_(sheet, numRows);

  SpreadsheetApp.flush();
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

  sheet.getRange(2, col('滞留日数')).setFormula(
    '=ARRAYFORMULA(IF($A2:A="","",IF($E2:E="",TODAY()-$C2:C,$E2:E-$C2:C)))'
  );
  sheet.getRange(2, col('Transaction fee')).setFormula(
    '=ARRAYFORMULA(IF($A2:A="","",IF($D2:D="Mercari",$F2:F*0.1,IF($D2:D="eBay",$F2:F*0.144,""))))'
  );
  sheet.getRange(2, col('Order earnings')).setFormula(
    '=ARRAYFORMULA(IF($A2:A="","",$F2:F-$H2:H-$I2:I-$J2:J))'
  );
  sheet.getRange(2, col('実質粗利')).setFormula(
    '=ARRAYFORMULA(IF($A2:A="","",$L2:L-$K2:K-$M2:M))'
  );
  sheet.getRange(2, col('粗利率')).setFormula(
    '=ARRAYFORMULA(IF($A2:A="","",IF($F2:F="","",$N2:N/$F2:F)))'
  );
}

function applyFormatting_(sheet, numRows) {
  const col = (name) => NEW_HEADERS.indexOf(name) + 1;
  sheet.getRange(2, col('出品日'), numRows, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, col('販売日'), numRows, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, col('粗利率'), numRows, 1).setNumberFormat('0.0%');
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
