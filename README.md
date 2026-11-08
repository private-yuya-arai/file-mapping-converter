# file-mapping-converter

fileA（CSV / Excel、複数ファイル）を、設定 Excel のマッピングに従って fileB 形式（CSV / Excel）に変換し、ファイル名の日付ごとに1ファイルにまとめて出力するツール（VBScript）。

## フォルダ
| フォルダ | 内容 |
|---|---|
| `v1.1/` | 最新版。ファイル名の日付チェックを追加、テスト A6・E14、サンプル（`sample/`） |
| `v1.0/` | コメント追加版。タスクスケジューラ運用検討資料 |
| 直下 | 最初の版 |

## 使い方（v1.1）
- 実行: `DataMapping.vbs` をダブルクリック、または `cscript //nologo DataMapping.vbs "設定.xlsx"`
- テスト: `test\テスト実行.bat`（Excel を裏で起動します）
- サンプル: `cscript //nologo DataMapping.vbs "sample\DataMapping_サンプル.xlsx"`
- タスクスケジューラ: `タスクスケジューラ運用検討.md`

## 注意
- `.vbs` / `.bat` は Shift-JIS・CRLF です。GitHub の画面では日本語が文字化けして表示されますが、ダウンロードしてそのまま使えます。
- Windows と Microsoft Excel が必要です。
