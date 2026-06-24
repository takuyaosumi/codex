# RMA 管理 GAS

Google Apps Script project `1RWF1L7jGGXs28r32oP6Ae65ziOwTglOC0v5hSBoktSy5vHx3UyQtp6h9` に `clasp push` して使います。

## できること

- `NewEntry` シートを手入力用に初期作成
- カスタムメニュー `💎Jitbit > フォーム生成` を追加
- `RMA-yyyyMMdd-0001` 形式で Entry ID を発番
- `管理台帳` シートへ転記
- Google Docs template `1CSSxjyQbFN4HYw4uHJnPS9NqOTRvIsc9` をコピーして RMA フォームを生成
- 生成フォームの `Your ref.` に Entry ID を自動記入
- 生成フォーム URL と Document ID を台帳へ保存

## Push 手順

1. `clasp login` で対象 Google アカウントにログインする
2. このフォルダで `clasp push` を実行する
3. 対象スプレッドシートを再読み込みする
4. メニュー `💎Jitbit > 初期セットアップ` を実行する
5. `NewEntry` の 2 行目に 1 件分を入力する
6. `💎Jitbit > フォーム生成` を実行する

初回実行時は Google の認証画面が出ます。未確認アプリの警告が表示された場合は、内容を確認したうえで `詳細 > プロジェクトに移動 > 許可` を選んでください。

## Target IDs

- Spreadsheet ID: `1uGWkEmfg0B00mGANWSVLLMgMUWKEiKIbHncDZhRzBzs`
- Apps Script project ID: `1RWF1L7jGGXs28r32oP6Ae65ziOwTglOC0v5hSBoktSy5vHx3UyQtp6h9`
- RMA Registration Form template document ID: `1CSSxjyQbFN4HYw4uHJnPS9NqOTRvIsc9`

## 入力ルール

`NewEntry` は 1 行目が見出し、2 行目が入力欄です。必須項目は以下です。

- `Company or institution`
- `Contact person`
- `Email`
- 返送品情報 `Goods 1` 以降のいずれか 1 行

返送品は最大 5 行まで対応しています。

## 生成物

Google Drive の `RMA Generated Forms` フォルダに Google Docs が作成されます。作成されたフォーム URL は `RMA Ledger` の `Form URL` に記録されます。

## 設定変更

`Code.gs` 冒頭の `CONFIG` で変更できます。

- `idPrefix`: Entry ID の接頭辞
- `maxGoodsRows`: 返送品の最大行数
- `clearAfterSubmit`: フォーム生成後に `NewEntry` 2 行目を消すかどうか
