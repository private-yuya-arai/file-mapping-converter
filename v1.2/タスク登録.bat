@echo off
setlocal
rem ==============================================================
rem  DataMapping をタスクスケジューラに登録する
rem   ・下の「設定」を必要に応じて書き換えてから実行してください
rem   ・ログオンしているときだけ実行されます（Excel を使うため）
rem   ・同じ名前のタスクがあれば上書きします
rem ==============================================================

rem ---- 設定 ----
set TASK_NAME=DataMapping_日次変換
set SCHEDULE=DAILY
set START_TIME=06:00
set CONFIG=%~dp0DataMapping.xlsx
rem ---------------

set TOOL=%~dp0DataMapping.vbs
if not exist "%TOOL%" (echo DataMapping.vbs が見つかりません: %TOOL% & pause & exit /b 1)
if not exist "%CONFIG%" (echo 設定ファイルが見つかりません: %CONFIG% & pause & exit /b 1)

schtasks /Create /F /IT /TN "%TASK_NAME%" /SC %SCHEDULE% /ST %START_TIME% ^
  /TR "\"%SystemRoot%\System32\cscript.exe\" //nologo \"%TOOL%\" \"%CONFIG%\""
if errorlevel 1 (echo 登録に失敗しました & pause & exit /b 1)

echo.
echo 登録しました: %TASK_NAME%（%SCHEDULE% %START_TIME%）
echo.
choice /M "今すぐ1回実行して動作を確認しますか"
if errorlevel 2 goto END
schtasks /Run /TN "%TASK_NAME%"
echo 実行を開始しました。終了後にタスクスケジューラの「前回の実行結果」とログを確認してください（0x0 = 成功 / 0x1 = エラーあり）。
:END
pause
