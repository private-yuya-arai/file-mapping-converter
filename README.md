# file-mapping-converter

fileA（CSV / Excel、複数ファイル）を、設定 Excel のマッピングに従って fileB 形式（CSV / Excel）に変換し、ファイル名の日付ごとに1ファイルにまとめて出力するツール（VBScript）。

## フォルダ
| フォルダ | 内容 |
|---|---|
| `v1.2/` | 最新版。変換する日付を選べる（設定「対象日付」：当日 / 範囲 / すべて）、範囲を指定して一括変換（`DataMapping_範囲実行.bat`、`/from:` `/to:`）、出力ファイル名に変換時刻（`{time}`）、テスト I1〜I12・J1〜J7・K1〜K2（100項目）・P2 |
| `v1.1/` | ファイル名の日付チェックを追加、テスト A6・E14、サンプル（`sample/`） |
| `v1.0/` | コメント追加版。タスクスケジューラ運用検討資料 |
| 直下 | 最初の版 |

## 使い方（v1.2）
- 実行: `DataMapping.vbs` をダブルクリック、または `cscript //nologo DataMapping.vbs "設定.xlsx"`
- 毎日の実行: 設定「対象日付」を「当日」にすると、実行した日の日付の fileA だけを変換
- 最初の一括作成: `DataMapping_範囲実行.bat` をダブルクリックして開始日・終了日を入力（または `cscript //nologo DataMapping.vbs "設定.xlsx" /from:20261101 /to:20261107`）
- fileB の名前に変換した時刻を付ける: 設定「出力ファイル名」に `{time}` を書く（例 `1_EL_出庫検索_{date}_{time}.csv` → `1_EL_出庫検索_20261108_160835.csv`）
- テスト: `test\テスト実行.bat`（Excel を裏で起動します）
- サンプル: `cscript //nologo DataMapping.vbs "sample\DataMapping_サンプル.xlsx"`
- タスクスケジューラ: `タスクスケジューラ運用検討.md`

## 注意
- `.vbs` / `.bat` は Shift-JIS・CRLF です。GitHub の画面では日本語が文字化けして表示されますが、ダウンロードしてそのまま使えます。
- Windows と Microsoft Excel が必要です。
