# 対応している言語
- **[日本語](./README_ja.md)** ←
- [English](./README_en.md)

# このプログラムについて
このツールは、PowerPointファイルに対して以下の処理をローカル環境で行います。

- PPTX/PPSX のパスワードロック解除
- PPSX を PPTX に変換

# 注意事項
パスワード解除は、作成者の意図に反する利用となる可能性があります。対象ファイルは、権利と利用範囲を確認したうえで、自己責任で扱ってください。  
本プログラムの利用によって生じた損害について、作成者は責任を負いません。

# 使い方
1. このリポジトリを clone するか、ダウンロードする。
2. プロジェクト直下でターミナルを開く。
3. `npm i` を実行する。
4. `node index.js` または `npx nodemon index.js` を実行する。
5. 表示されたURL（通常は http://localhost:6490）にアクセスする。
6. 以下のどちらかを選ぶ。
	- パスワード削除: 「パスワードを削除したいPPTX/PPSX」を選択して GO を押す。
	- PPSX→PPTX変換: 「PPTXに変換したいPPSX」を選択して GO を押す。
7. 処理後のファイルがダウンロードされる。

# API 仕様
画面からは以下のPOSTが送信されます。

- パスワード解除: `/api/progress?type=removePassword`
- PPSX→PPTX変換: `/api/progress?type=convertToPPTX`

フォームデータ:

- `file` (必須): `.pptx` または `.ppsx`
- ただし `convertToPPTX` は `.ppsx` のみ受け付け

# エラー仕様（主なもの）
- `No files were uploaded.`
- `Only .pptx or .ppsx files are supported.`
- `convertToPPTX mode only supports .ppsx files.`
- `Uploaded PowerPoint file does not appear to be password-protected.`
	- パスワード解除モードで `p:modifyVerifier` タグが存在しない場合

# 技術的な処理概要
1. `express` でローカルWebサーバーを起動
2. アップロードされたファイルを一時保存
3. `jszip` でPowerPoint内部XMLを読み込む
4. 処理モードごとにXMLを書き換える
5. ファイルを再生成してダウンロードさせる
6. デバッグモードが無効なら一時ファイルを削除

# 使用パッケージ
- `express`
- `express-fileupload`
- `jszip`
- `nodemon` (開発時)
