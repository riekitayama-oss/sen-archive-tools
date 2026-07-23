# SEN ARCHIVE 売上シート ⇔ 検品マスター 自動連携

売上シートに検品マスター(Item Master)のSKU・商品名・出品日を自動反映し、滞留日数や粗利などの計算式をセットアップするツール。

## ファイル

- `apps-script/SalesSheetSync.gs` — 売上シートにセットする Google Apps Script

## 使い方

1. [売上シート](https://docs.google.com/spreadsheets/d/1m0IxyLZN9DC_YuLZ9JexcBQc9-wUzBFDk9mVx7w544Y) を開く
2. 拡張機能 > Apps Script
3. `apps-script/SalesSheetSync.gs` の内容を貼り付けて保存
4. エディタ上部の関数選択で `setupSalesSheet` を選び「実行」
   - 初回はIMPORTRANGEの権限承認ダイアログが出るので許可する
   - `検品マスター`側で未承認の場合は一度スプレッドシートに戻ってA2セルをクリックし
     「アクセスを許可」を押してから再実行する
5. 以降はスプレッドシートを開くと出てくるメニュー「SEN ARCHIVE > 売上シートを再同期」から
   いつでも再実行できる（既存の手入力データはSKUをキーに保持されるので、検品マスターに
   新しい商品が増えるたびに再実行してよい）

## 検品マスター側で見つかった問題（元の依頼内容からの修正点）

元の依頼にあった数式 `select Col1, Col3, Col73` には列ズレがあったため修正した。

検品マスターは3段ヘッダー構成になっている:

| 行 | 内容 |
|---|---|
| 行1 | 簡易ラベル（SKU, TITLE, ... の短縮名。多くの列は空欄） |
| 行2 | 正式ヘッダー（実際に使われている列名） |
| 行3以降 | 実データ |

行2を実データ範囲に含めると「行2＝ヘッダー行」がデータの1行目として紛れ込んでしまうため、
`IMPORTRANGE` の範囲は `検品マスター!A2:BU` ではなく **`検品マスター!A3:BU`** にする必要がある。

また列の対応も実データで確認したところ以下の通り（Col1やCol3ではない）:

| 項目 | 依頼にあった列 | 実際の列 | 実際の列の内容(行2ヘッダー) |
|---|---|---|---|
| SKU | Col1 | **Col6** | Col1は"SOURCE"（担当者名） |
| Item Name (TITLE) | Col3 | **Col8** | Col3は"Auction Receipt Day" |
| 出品日 | Col73 (BU列) | Col73 (BU列) | 列位置は依頼通り。ヘッダー名は"eBay出品日"だが、メルカリ出品中の行にも日付が入っているため全チャネル共通の出品日として使えることをデータで確認済み |

修正後の数式（`SalesSheetSync.gs` の `MASTER_RANGE` / `MASTER_QUERY` に反映済み）:

```
=QUERY(IMPORTRANGE("https://docs.google.com/spreadsheets/d/166BARD0R6QVWFO7uwBQvnZsvmntC9W362UmVdXSXa0Y/edit", "検品マスター!A3:BU"), "select Col6, Col8, Col73 where Col6 is not null", 0)
```

**要確認**: 検品マスターのBU列(出品日)がテキストではなく日付型で入力されているか。テキストのままだと
売上シート側の「滞留日数」計算(`販売日-出品日`)が正しく計算されない場合がある。

## 売上シートの新しいレイアウト

| 列 | 項目 | 入力方法 |
|---|---|---|
| A | SKU | 自動（検品マスターから） |
| B | Item Name | 自動（検品マスターから） |
| C | 出品日 | 自動（検品マスターから・編集不可の警告つき保護あり） |
| D | Channel | 手入力 |
| E | 販売日 | 手入力 |
| F | 販売価格 | 手入力 |
| G | 滞留日数 | 自動計算 |
| H | Transaction fee | 自動計算 |
| I | Ad Fee | 手入力 |
| J | 値引き額(クーポン等) | 手入力 |
| K | 送料実費 | 手入力（都度異なるため自動化しない） |
| L | Order earnings | 自動計算 |
| M | 仕入れ値 | 手入力 |
| N | 実質粗利 | 自動計算 |
| O | 粗利率 | 自動計算 |

## 計算式（すべてARRAYFORMULAで自動的に行を追従）

- 滞留日数: `IF(販売日="", TODAY()-出品日, 販売日-出品日)`
- Transaction fee: `IF(Channel="Mercari", 販売価格*0.1, IF(Channel="eBay", 販売価格*0.144, ""))`
  - Shopify/VCは手数料率が未確定のため空欄のまま。判明次第 `SalesSheetSync.gs` の
    `applyFormulaColumns_` 内のIF文に追加する
- Order earnings: `販売価格 - Transaction fee - Ad Fee - 値引き額`
- 実質粗利: `Order earnings - 送料実費 - 仕入れ値`
- 粗利率: `実質粗利 / 販売価格`
