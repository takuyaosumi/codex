# RMA 管理 GAS

ブラウザで開いている Google スプレッドシートに Apps Script として `Code.gs` を貼り付けて使います。

## できること

- `NewEntry` シートを手入力用に初期作成
- カスタムメニュー `💎Jitbit > フォーム生成` を追加
- `RMA-yyyyMMdd-0001` 形式で Entry ID を発番
- `管理台帳` シートへ転記
- 添付の RMA Registration Form に準じた Google Docs フォームを生成
- 生成フォームの `Your ref.` に Entry ID を自動記入
- 生成フォーム URL と Document ID を台帳へ保存

## 導入手順

1. 対象スプレッドシートを開く
2. `拡張機能 > Apps Script` を開く
3. 既存の `Code.gs` を削除または置き換える
4. このフォルダの `Code.gs` の内容を貼り付ける
5. 保存する
6. スプレッドシートを再読み込みする
7. メニュー `💎Jitbit > 初期セットアップ` を実行する
8. `NewEntry` の 2 行目に 1 件分を入力する
9. `💎Jitbit > フォーム生成` を実行する

初回実行時は Google の認証画面が出ます。未確認アプリの警告が表示された場合は、内容を確認したうえで `詳細 > プロジェクトに移動 > 許可` を選んでください。

## 入力ルール

`NewEntry` は 1 行目が見出し、2 行目が入力欄です。必須項目は以下です。

- `Company or institution`
- `Contact person`
- `Email`
- 返送品情報 `Goods 1` 以降のいずれか 1 行

返送品は最大 5 行まで対応しています。

## 生成物

Google Drive の `RMA Generated Forms` フォルダに Google Docs が作成されます。作成されたフォーム URL は `管理台帳` の `Form URL` に記録されます。

## 設定変更

`Code.gs` 冒頭の `CONFIG` で変更できます。

- `idPrefix`: Entry ID の接頭辞
- `maxGoodsRows`: 返送品の最大行数
- `clearAfterSubmit`: フォーム生成後に `NewEntry` 2 行目を消すかどうか

