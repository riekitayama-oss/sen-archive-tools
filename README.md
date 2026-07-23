# SEN ARCHIVE 売上シート ⇔ 検品マスター 自動連携

売上シートに検品マスター(Item Master)のSKU・商品名・出品日を自動反映し、滞留日数や粗利などの計算式をセットアップするツール。

## ファイル

- `apps-script/SalesSheetSync.gs` — 売上シートにセットする Google Apps Script
- `apps-script/HivePrelovedScan.gs` — 検品マスターにセットする、HIVE PRELOVED週次調査用 Google Apps Script（詳細は下部「HIVE PRELOVED週次調査」参照）

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
| A | SKU | 自動（検品マスターから。一度書き込んだ行は固定され、新しいSKUは末尾に追記される） |
| B | Item Name | 自動（A列のSKUでR:T参照テーブルをVLOOKUP） |
| C | 出品日 | 自動（同上・編集不可の警告つき保護あり） |
| D | Channel | 手入力（`eBay`/`Mercari`/`Shopify`/`VC`のプルダウン） |
| E | 販売日 | 手入力（実際の日付として入力する。例: `2026/7/2`） |
| F | 販売価格 | 手入力（Channelの通貨建てのまま。eBay/Shopify/VCはUSD、MercariはJPY） |
| G | 滞留日数 | 自動計算 |
| H | Transaction fee | 自動計算 |
| I | Ad Fee | 手入力 |
| J | 値引き額(クーポン等) | 手入力 |
| K | 送料実費 | 手入力（JPY。都度異なるため自動化しない） |
| L | Order earnings | 自動計算（F列と同じ通貨） |
| M | 仕入れ値 | 手入力（JPY） |
| N | 実質粗利 | 自動計算（JPY。1円未満切り捨て） |
| O | 粗利率 | 自動計算 |
| P1/P2 | USD/JPYレート | 自動（GOOGLEFINANCEのライブレート） |
| R:T | 検品マスター参照テーブル(内部用) | 自動。触らない |

### SKU行が固定される仕組み

以前のバージョンはA2に置いた1つのQUERY数式がA:C列全体にspillする設計だったため、
検品マスター側で行の削除・並べ替えが起きると、下の行が全部ズレて既存の手入力データが
別のSKUに紐付いてしまう問題があった。

現在はSKU(A列)を「一度書いたら動かさない」プレーンな値として管理し、検品マスターに新しい
SKUが増えたときだけ末尾に追記する。Item Name(B列)・出品日(C列)は自分の行のSKUをキーに
R:T列の参照テーブルへVLOOKUPするため、検品マスターの行順が変わっても既存行のSKUとの
対応関係は崩れない。

## 為替（USD/JPY）の扱い

- eBay・Shopify・VCの売上（販売価格・Order earnings）はUSD建て、Mercariの売上はJPY建てで入力する
- P1に「USD/JPY」ラベル、P2にGOOGLEFINANCEのライブレートが自動で入る（都度の実際のレートとは
  ズレる可能性がある概算値）
- 実質粗利・粗利率は、ChannelがMercari以外の行だけP2のレートを掛けてJPYに揃えてから
  JPY建ての送料実費・仕入れ値と差し引きする。Mercariの行は元々JPYなので為替換算しない

## 計算式（すべてARRAYFORMULAで自動的に行を追従）

- 滞留日数: `IF(出品日="", "", IF(販売日="", TODAY()-出品日, 販売日-出品日))`
- Transaction fee: `IF(Channel="Mercari", 販売価格*0.1, IF(Channel="eBay", 販売価格*0.144, ""))`
  - Shopify/VCは手数料率が未確定のため空欄のまま。判明次第 `SalesSheetSync.gs` の
    `applyFormulaColumns_` 内のIF文に追加する
- Order earnings: `販売価格 - Transaction fee - Ad Fee - 値引き額`
- 実質粗利: `ROUNDDOWN(IF(Channel="Mercari", Order earnings-送料実費-仕入れ値, Order earnings×USD/JPYレート-送料実費-仕入れ値), 0)`
- 粗利率: `実質粗利 / (Channel="Mercari" ? 販売価格 : 販売価格×USD/JPYレート)`

## HIVE PRELOVED週次調査

HIVE PRELOVED (hivepreloved.com) のShopifyストアを週次で調査し、Sold Outになった商品を
検品マスターのBRAND/ITEM NAME列と突き合わせて、SENでまだ扱っていなさそうなブランド・
デザインを検品マスターの新規タブ「仕入れ判断シート」にリストアップするツール。

### 前提と制限（要検証）

このツールを作った環境からはネットワークポリシーによりHIVE PRELOVEDのページを
実際に開いて確認することができなかった。そのため、Shopifyストアが一般的に公開している
`/collections/{handle}/products.json` エンドポイントが使える前提でコードを書いており、
以下は未検証の推測が含まれる。**初回実行後、実際の出力を見て調整すること。**

- `HIVE_COLLECTIONS`のhandle（`bags`/`shoes`）: 実際に `https://hivepreloved.com/collections/bags`
  のようなURLをブラウザで開いて、正しいhandleか確認する
- ブランド名は`vendor`フィールドから取得する前提。HIVE側で店名などが入っている場合は
  `HivePrelovedScan.gs`のロジックを調整する必要がある
- 素材・カラーはShopifyの標準フィールドではないため、タグやタイトルからキーワードで
  推測しているだけ(`HIVE_MATERIAL_KEYWORDS`/`HIVE_COLOR_KEYWORDS`で調整可能)
- 価格の通貨は未確認

### 使い方

1. [検品マスター](https://docs.google.com/spreadsheets/d/166BARD0R6QVWFO7uwBQvnZsvmntC9W362UmVdXSXa0Y) を開く
2. 拡張機能 > Apps Script（売上シート用とは別のスクリプトプロジェクトになる）
3. `apps-script/HivePrelovedScan.gs` の内容を貼り付けて保存
4. `scanHivePreloved` を一度手動実行し、新規タブ「仕入れ判断シート」に出力された内容を確認する
5. 問題なければスプレッドシートのメニュー「HIVE PRELOVED調査 > 週次自動実行を設定」を実行
   （毎週月曜朝に自動実行されるようになる。以後はメニューの「今すぐ調査を実行」でいつでも
   手動実行も可能）

### 「未仕入れフラグ」の考え方

検品マスターとHIVE PRELOVED側には型番のような構造化された共通キーが無いため、
以下の目安で判定している（最終判断は人間の確認が必要）:

1. 検品マスターにそのブランド自体の取り扱いが無ければ「未仕入れ(ブランド自体が無い)」
2. ブランドはあるが、商品名の単語一致率（Jaccard係数）が閾値未満なら
   「未仕入れ(類似デザインが見当たらない)」
3. 閾値以上の一致があれば「取扱いあり」

「Sold Out確認日」は商品ごとに初めて検出した日付を保持し、2回目以降の実行では
同じ商品(`handle`で識別)を重複して追加しない。

### 推定利益率について

「推定利益率(参考値)」列には固定で `20-30%(推測値・HIVE側データから確定できる数字ではない)`
と書き出している。これはHIVE側の販売データから計算した数値ではなく、あくまで参考の目安。
