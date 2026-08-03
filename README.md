# eBay価格管理表 自動化スクリプト

Googleスプレッドシート「[eBay価格管理表](https://docs.google.com/spreadsheets/d/1mqUP2njh5OwJ8m9wsr_lIxhR8InyaQkzhL63oqn_x1c/edit)」用の
Google Apps Script（GAS）です。「CSVインポート → 更新履歴 → 商品マスター」という多段フローをやめ、
**CSVを貼って更新ボタンを押すだけ**で商品マスター・Dashboard・Performance DBが自動更新される構成にしています。

## 毎週やること

1. eBay Seller Hub で **Traffic Report** をダウンロードする
2. **「CSVインポート」** シートに、そのまま貼り付ける（見出し行の上に説明文が付いていてもOK）
3. eBayの **Active Listings** 画面を見て、商品マスターの **Views / Watchers**（必要なら Promoted有無・広告レート）を入力する
4. メニュー **「eBay価格管理」→「① 更新する」** を実行する

これで完了。

## 情報源の切り分け（重要）

| 項目 | 情報源 | 反映方法 |
| --- | --- | --- |
| Impressions / CTR | Traffic Report CSV | 自動反映（スクリプトがCSVから読み取る） |
| Views / Watchers | Active Listings 画面 | 手入力（人が商品マスターに直接入力） |
| Watch率 | Watchers ÷ Views | 自動計算 |
| Promoted有無 / 広告レート | Active Listings 画面 | 手入力 |
| 前週比（各指標） | Performance DB の前回スナップショット | 自動計算 |
| 優先順位 | Watch率の高い順 | 自動採番 |
| 残り日数(90日) | 出品日からの経過日数 | 自動計算 |
| 今週やること | 判断・Watch率 | 自動生成 |
| Dashboardの今週の順位 | Watch率上位3件 | 自動生成 |

「更新履歴」タブは廃止しました。更新のたびに **Performance DB** にその週のスナップショット（Views / Impressions / CTR / Watchers / Watch率など）を自動追加し、商品マスターには常に最新値だけを表示します。過去推移の分析はPerformance DBを見れば追えます。

## 今回のバグ修正（新規商品が下段で反映されない問題）

- 更新範囲を `getRange(2, 1, 固定行数, ...)` ではなく `sheet.getLastRow()` で毎回取得するように変更し、行数が増えても最終行まで確実に更新されるようにしました。
- 商品照合は **eBay Item ID を最優先** で行い、eBay Item IDが未入力の新規商品は **商品名とTraffic ReportのListing titleの類似度** で自動的に紐づけ、見つかったIDを商品マスターに自動で書き戻します（次回以降はIDで直接一致するようになります）。
- CSV側に対応が見つからなかった商品は、更新後のアラートに一覧表示されるので、SKU・商品名の表記ゆれを確認できます。

## セットアップ

詳細は [`docs/SETUP.md`](docs/SETUP.md) を参照してください。概要:

1. スプレッドシートを開き、`拡張機能 → Apps Script` を開く
2. `apps-script/Code.gs` の内容をコピーして貼り付け、保存する
3. スプレッドシートをリロードすると、メニューに **「eBay価格管理」** が追加される
4. メニューから **「初期セットアップ」** を実行し、必要なシート・列を確認/作成する

## ファイル構成

- `apps-script/Code.gs` — 自動更新スクリプト本体
- `docs/SETUP.md` — シート構成・列定義・導入手順の詳細
