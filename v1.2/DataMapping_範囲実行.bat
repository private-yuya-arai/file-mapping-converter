@echo off
setlocal
rem ==============================================================
rem  日付の範囲を指定して、まとめて変換する（最初の一括作成・後から届いた fileA の再出力用）
rem   ・設定ファイルの「対象日付」（当日など）は書き換えずに、この実行だけ範囲で変換します
rem   ・ダブルクリック … 開始日・終了日を聞かれるので yyyymmdd で入力（終了日が空欄なら開始日の1日だけ）
rem   ・コマンドから   … DataMapping_範囲実行.bat 20261101 20261106 [設定ファイル]
rem   ・出力先に同じ日付の fileB があると、設定「上書き」がなしならその日付はスキップされます
rem   ・終了コード 0 = エラーなし / 1 = エラーあり
rem ==============================================================
set "FROM=%~1"
set "TO=%~2"
set "CONFIG=%~3"
if "%FROM%"=="" set /p "FROM=開始日（yyyymmdd）: "
if "%TO%"=="" if "%~1"=="" set /p "TO=終了日（yyyymmdd、空欄なら開始日の1日だけ）: "
if "%TO%"=="" set "TO=%FROM%"
if "%CONFIG%"=="" set "CONFIG=%~dp0DataMapping.xlsx"

cscript //nologo "%~dp0DataMapping.vbs" "%CONFIG%" /from:%FROM% /to:%TO%
set RC=%ERRORLEVEL%
if "%~1"=="" pause
exit /b %RC%
