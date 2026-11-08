@echo off
rem DataMapping.vbs の自動テスト（結果はこのフォルダの テスト結果.csv）
rem テスト中は Excel を裏で使います。Excel の操作は控えてください。
cscript //nologo "%~dp0RunTests.vbs"
echo.
if %ERRORLEVEL%==0 (echo すべてのテストに合格しました) else (echo 不合格のテストがあります。テスト結果.csv を確認してください)
pause
