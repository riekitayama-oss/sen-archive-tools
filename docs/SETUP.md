# セットアップ手順

## 1. スクリプトの導入

1. 対象のスプレッドシートを開く
2. メニュー `拡張機能 → Apps Script` を開く
3. デフォルトで開かれる `Code.gs`（またはコード.gs）の中身を全て削除し、このリポジトリの
   `apps-script/Code.gs` の内容を貼り付ける
4. 保存（フロッピーアイコン、またはCtrl/Cmd+S）
5. スプレッドシートのタブに戻ってページをリロードする
   → メニューバーに **「eBay価格管理」** が追加される
6. 初回はどこかの操作でGoogleの権限承認ダイアログが出るので、自分のアカウントで許可する
7. メニュー **「eBay価格管理」→「初期セットアップ（シート/列を確認・作成）」** を一度実行する
   → `CSVインポート(Traffic Report)` / `CSVインポート(Active Listings)` / `Performance DB` / `Dashboard`
      シートが無ければ自動作成され、`商品マスター` に必要な列が無ければ末尾に自動追加される

## 2. シート構成

| シート名 | 役割 |
| --- | --- |
| `商品マスター` | 最新の状態を表示するメインシート。既存のシートをそのまま使う |
| `CSVインポート(Traffic Report)` | Traffic Reportを貼り付ける場所（Views / Impressions / CTRの元） |
| `CSVインポート(Active Listings)` | All active listings reportを貼り付ける場所（Watchersの元） |
| `Performance DB` | 更新のたびに追加される週次スナップショット（旧「更新履歴」の代わり） |
| `Dashboard` | 今週の順位・注意商品を自動生成 |

シート名はこの通りである必要があります（`Code.gs` 冒頭の定数で変更可能）。

## 3. 2つのCSVの取得場所

| レポート | 取得場所 | 使う列 |
| --- | --- | --- |
| Traffic Report | eBay Seller Hub → Business Reports → Traffic Report | eBay item ID / Listing title / Total page views / Total impressions / Click-through rate |
| All active listings report | eBay Seller Hub → Reports → Downloads → Download report → Report source: `Listings` → Report type: `All active listings` | Item number / Title / Watchers |

貼り付け方法は次のどちらでもOK:
- レポート画面からコピーして、対応するシートにそのまま貼り付け
- ダウンロードした`.csv`ファイルを `ファイル → インポート → アップロード` で選択し、
  対応するシートに「現在のシートを置換」で取り込む

見出し行の上に説明文（Disclaimers等）が入っていても、スクリプトが自動でヘッダー行を検出します。

## 4. 商品マスターの列

### 既存必須列（無いとエラーで停止する）

`SKU` / `商品名` / `出品日` / `出品価格(USD)` / `Views` / `Impressions` / `CTR` / `Watchers` / `判断` / `eBay Item ID`

### スクリプトが自動追加する列（末尾に追加される。既にあれば追加しない）

| 列名 | 内容 |
| --- | --- |
| `残り日数(90日)` | 90 − 経過日数 |
| `Views前週比` | 例: `⬆ +8` / `⬇ -2` / `→ 0` |
| `Impressions前週比` | 同上 |
| `CTR前週比` | 例: `⬆ +2.1pt` |
| `Watchers前週比` | 例: `⬆ +3` |
| `Watch率` | Watchers ÷ Views（`0.0%`表示） |
| `Watch率前週比` | 例: `⬆ +4.5pt` |
| `優先順位` | Watch率が高い順の順位（1が最優先） |
| `今週やること` | 判断・Watch率から自動生成される短いアクション |

新規商品を追加するときは、`SKU` と `商品名` と `出品日` だけ入れておけば、`eBay Item ID` は
CSV側タイトルとの類似度照合で自動的に埋まります（万一見つからなければ手動で入力してください）。

## 5. Performance DB の列

`記録日 / eBay Item ID / SKU / 商品名 / 出品価格(USD) / Views / Impressions / CTR / Watchers / Watch率`

同じ日に複数回「更新する」を実行しても、同日・同IDの行は上書きされ、重複した行は増えません。

## 6. 週次オペレーション

1. Traffic Reportをダウンロードして `CSVインポート(Traffic Report)` に貼り付け
2. All active listings reportをダウンロードして `CSVインポート(Active Listings)` に貼り付け
3. メニュー `eBay価格管理 → ① 更新する` を実行
4. 完了ダイアログで反映件数・自動リンク件数・どちらのCSVにも対応が見つからなかった商品を確認する
