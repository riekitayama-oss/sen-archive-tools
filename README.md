# eBay価格管理表 自動化スクリプト

Googleスプレッドシート「[eBay価格管理表](https://docs.google.com/spreadsheets/d/1mqUP2njh5OwJ8m9wsr_lIxhR8InyaQkzhL63oqn_x1c/edit)」用の
Google Apps Script（GAS）です。「CSVインポート → 更新履歴 → 商品マスター」という多段フローをやめ、
**2つのCSVを貼って更新ボタンを押すだけ**（手入力ゼロ）で商品マスター・Dashboard・Performance DBが自動更新される構成にしています。

## 毎週やること

1. eBay Seller Hub で以下の **2つのレポート** をダウンロードする
   - **Traffic Report**（Business Reports）→ Views・Impressions・CTR
   - **All active listings report**（Reports → Downloads → Report source: Listings）→ Watchers
2. それぞれ対応するシートに貼り付ける（またはFile→Importで取り込む）
   - **「CSVインポート(Traffic Report)」**
   - **「CSVインポート(Active Listings)」**
3. メニュー **「eBay価格管理」→「① 更新する」** を実行する

これで完了。Views・Watchersも含めて手入力は不要。

## 情報源の切り分け

| 項目 | 情報源 | 反映方法 |
| --- | --- | --- |
| Views / Impressions / CTR | Traffic Report CSV | 自動反映 |
| Watchers | All active listings report CSV | 自動反映 |
| Watch率 | Watchers ÷ Views | 自動計算 |
| 前週比（各指標） | Performance DB の前回スナップショット | 自動計算 |
| 優先順位 | Watch率の高い順 | 自動採番 |
| 残り日数(90日) | 出品日からの経過日数 | 自動計算 |
| 今週やること | 判断・Watch率 | 自動生成 |
| Dashboardの今週の順位 | Watch率上位3件 | 自動生成 |

Traffic Reportの「Total page views」列が、これまで商品マスターのViewsに手入力していた値と一致することを確認済みです（例: CHANEL CC Ballet Flats = 24件で一致）。Active Listings画面での確認作業は不要になりました。

Promoted有無・広告レートは現時点でCSV化できるレポートが確認できていないため、引き続きActive Listings画面を見ての手入力が必要です。

「更新履歴」タブは廃止しました。更新のたびに **Performance DB** にその週のスナップショット（Views / Impressions / CTR / Watchers / Watch率など）を自動追加し、商品マスターには常に最新値だけを表示します。過去推移の分析はPerformance DBを見れば追えます。

## 今回のバグ修正（新規商品が下段で反映されない問題）

- 更新範囲を `getRange(2, 1, 固定行数, ...)` ではなく `sheet.getLastRow()` で毎回取得するように変更し、行数が増えても最終行まで確実に更新されるようにしました。
- 商品照合は **eBay Item ID を最優先** で行い、eBay Item IDが未入力の新規商品は **商品名とCSV側のタイトルの類似度** で自動的に紐づけ、見つかったIDを商品マスターに自動で書き戻します（次回以降はIDで直接一致するようになります）。
- どちらのCSVにも対応が見つからなかった商品は、更新後のアラートに一覧表示されるので、SKU・商品名の表記ゆれを確認できます。

## セットアップ

詳細は [`docs/SETUP.md`](docs/SETUP.md) を参照してください。概要:

1. スプレッドシートを開き、`拡張機能 → Apps Script` を開く
2. `apps-script/Code.gs` の内容をコピーして貼り付け、保存する
3. スプレッドシートをリロードすると、メニューに **「eBay価格管理」** が追加される
4. メニューから **「初期セットアップ」** を実行し、必要なシート・列を確認/作成する

## ファイル構成

- `apps-script/Code.gs` — 自動更新スクリプト本体
- `docs/SETUP.md` — シート構成・列定義・導入手順の詳細
