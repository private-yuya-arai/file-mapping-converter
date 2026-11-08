@echo off
rem コマンドプロンプト・タスクスケジューラ用（ログを画面に表示、終了コード 0=成功 / 1=エラーあり）
rem 引数に設定ファイルを渡すと、その設定で実行します
cscript //nologo "%~dp0DataMapping.vbs" %*
exit /b %ERRORLEVEL%
