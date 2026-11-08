'==============================================================
' RunTests.vbs  DataMapping.vbs の自動テスト
'
'   実行: テスト実行.bat をダブルクリック（または cscript //nologo RunTests.vbs）
'   一部だけ: cscript //nologo RunTests.vbs <DataMapping.vbs> <DataMapping.xlsx> A1,B2a
'   結果: このフォルダの「テスト結果.csv」と画面表示
'   終了コード: 0 = 全件PASS / 1 = FAILあり
'
'   作業フォルダ（work）にテスト用の入力ファイル・設定ファイルを毎回作り直し、
'   パターンごとに DataMapping.vbs を実行して、出力・ログ・終了コードを検証する。
'   ※ テスト中は Excel を裏で起動します。テスト中は Excel を操作しないでください。
'
' フォルダ構成（引数なしで実行する場合）
'   <ツールのフォルダ>\DataMapping.vbs, DataMapping.xlsx   … テスト対象と設定のひな形
'   <ツールのフォルダ>\test\RunTests.vbs, テスト実行.bat   … このファイル
'   <ツールのフォルダ>\test\work\<ID>\                     … テストごとの作業フォルダ（毎回削除して作り直す）
'       in\  … テスト用の fileA / out\ … ツールの出力 / log\ … ツールのログ / config.xlsx … 設定
'
' 1件のテストの流れ（例 T_A1）
'   BeginCase   … 作業フォルダを作り、設定・マッピングを標準値に戻す
'   MakeCsv など … 入力ファイルを作る。SetV / gMap などで設定をパターンに合わせて変える
'   RunTool     … ひな形をコピーして設定を書き込み、DataMapping.vbs を cscript で実行
'   Expect○○    … 終了コード・出力ファイル・ログの内容を期待値と比べる（Check で記録）
'
' テストIDの分類
'   A … 基本動作（日付ごとのまとめ・ファイル名順・列の並び・固定値・空行・番号つきのファイル名）
'   B … 文字コード・改行・区切り文字・ダブルクォート・ヘッダーなし
'   C … Excel 入出力（先頭ゼロ・日付セル・数値・エラー値・数式・13桁・混在）
'   D … 値の変換（TRIM・日付・数値・置換・左右・ゼロ埋め・コード変換・変換エラー）
'   E … 異常系（項目不足・既存ファイル・設定不正・0行・ファイルのロック・壊れたCSV・ファイル名の日付不正）
'   F … 実行方法（引数なし・パスに空白や日本語）
'   G … マッピングの空列
'   H … 抽出条件・項目の連結・先頭/末尾レコード・末尾行の除外・末尾改行
'   P … 性能（10万行）
'==============================================================
Option Explicit

' 定数：文字コード名・ADODB.Stream / Excel の定数・作業用の名前など
'   PROCESS_WAIT_SEC … ツール終了後、EXCEL.EXE が消えるまで待つ最大秒数
'   MAP_CLEAR_ROWS   … 設定のひな形の表（マッピングなど）を消去する行数
'   MAP_COLS / SIMPLE_COLS … マッピング表の列数 / コード変換・抽出条件表の列数
Const CS_SJIS = "Shift_JIS"
Const CS_UTF8 = "UTF-8"
Const CS_UTF8BOM = "UTF-8-BOM"
Const AD_TYPE_BINARY = 1
Const AD_TYPE_TEXT = 2
Const AD_SAVE_OVERWRITE = 2
Const UTF8_BOM_BYTES = 3
Const XL_WBAT_WORKSHEET = -4167
Const XL_NO_UPDATE_LINKS = 0
Const XL_FMT_XLSX = 51
Const PROCESS_WAIT_SEC = 15
Const RESULT_FILE = "テスト結果.csv"
Const WORK_DIR = "work"
Const MAP_CLEAR_ROWS = 60
Const MAP_COLS = 6
Const SIMPLE_COLS = 4

' 全体で使う値
'   gXL       … テスト側で使う Excel（入力用 xlsx の作成、設定の書き込み、出力 xlsx の確認）
'   gBase     … このファイルのフォルダ / gWork … 作業フォルダ（work）
'   gTool     … テスト対象の DataMapping.vbs / gTemplate … 設定ファイルのひな形（DataMapping.xlsx）
'   gResults  … テスト結果.csv に書く行 / gPass・gFail・gSkip … 確認項目ごとの件数
'   gCase・gTitle・gDir … 実行中のテストID・題名・作業フォルダ
'   gSet      … 「設定」シートに書く値（キー → 値）/ gSys … 「システム設定」シートに書く値
'   gMap / gCodes / gFilters … マッピング・コード変換・抽出条件の表（行の配列）
'   gRemoveKeys … 設定シートから消すキー（キー不足のテスト用）
Dim gFso, gSh, gXL, gBase, gWork, gTool, gTemplate
Dim gResults, gResN, gPass, gFail, gSkip
Dim gCase, gTitle, gDir, gCaseFail
Dim gSet, gSys, gMap, gCodes, gFilters, gRemoveKeys

' 入口。RunAllTests の戻り値（0 = 全件PASS / 1 = FAILあり）を終了コードにする
WScript.Quit RunAllTests()

'==============================================================
' 環境依存部分（wine 検証時はここを差し替える）
'   NewObject … COM オブジェクトの生成（差し替えでモックにできる）
'   IsMock    … True にすると、実機の Excel でしか確認できないテスト・確認項目を SKIP する
'==============================================================
Function NewObject(progId)
    Set NewObject = CreateObject(progId)
End Function

Function IsMock()
    IsMock = False
End Function

'==============================================================
' 全体の流れ
'==============================================================
' テスト全体を実行する
'   1. テスト対象とひな形の場所を決める（引数 1・2 で指定できる。省略時は1つ上のフォルダ）
'   2. 作業フォルダ work を空にして作り直す
'   3. テスト側の Excel を起動し、指定IDのテストを順に実行
'   4. Excel を終了し、テスト結果.csv を保存して件数を表示
Function RunAllTests()
    Dim ids, i, t0

    Set gFso = NewObject("Scripting.FileSystemObject")
    Set gSh = NewObject("WScript.Shell")
    gBase = gFso.GetParentFolderName(WScript.ScriptFullName)
    If WScript.Arguments.Count > 0 Then
        gTool = gFso.GetAbsolutePathName(WScript.Arguments(0))
    Else
        gTool = gFso.BuildPath(gFso.GetParentFolderName(gBase), "DataMapping.vbs")
    End If
    If WScript.Arguments.Count > 1 Then
        gTemplate = gFso.GetAbsolutePathName(WScript.Arguments(1))
    Else
        gTemplate = gFso.BuildPath(gFso.GetParentFolderName(gBase), "DataMapping.xlsx")
    End If
    If Not gFso.FileExists(gTool) Then
        WScript.Echo "DataMapping.vbs が見つかりません: " & gTool
        RunAllTests = 1
        Exit Function
    End If
    If Not gFso.FileExists(gTemplate) Then
        WScript.Echo "設定ファイルのひな形が見つかりません: " & gTemplate
        RunAllTests = 1
        Exit Function
    End If

    ' 前回の作業フォルダは消してから作り直す（前回の出力が残って結果に影響しないように）
    gWork = gFso.BuildPath(gBase, WORK_DIR)
    If gFso.FolderExists(gWork) Then gFso.DeleteFolder gWork, True
    gFso.CreateFolder gWork

    Set gXL = NewObject("Excel.Application")
    gXL.Visible = False
    gXL.DisplayAlerts = False

    ReDim gResults(255)
    gResN = 0: gPass = 0: gFail = 0: gSkip = 0
    AddResult "ID", "パターン", "確認項目", "結果", "詳細"
    t0 = Timer

    If WScript.Arguments.Count > 2 Then
        ids = Split(WScript.Arguments(2), ",")      ' 指定したIDだけ実行（例: A1,B2a）
    Else
        ids = Split("A1 A2 A3 A4 A5 A6 B1a B1b B2a B2b B2c B3 B4 B5a B5b B6 B7 " & _
                "C1 C2 C3 C4 C5 C6 C7 C8 C9 C10 C11 C13 C14 C15 " & _
                "D1 D2 D3 D4a D4b D5 D6 " & _
                "E1 E2a E2b E3 E4 E5a E5b E5c E5d E5e E6a E6b E7 E8a E8b E9 E10 E11a E11b E12 E13 E14 " & _
                "F2 F5 G1a G1b G1c G1d G2 G3 H1 H2 H3 H4 H5a H5b H6 H7 H8 P1", " ")
    End If
    ' テストIDを追加したときは、上の一覧と RunCase の Select Case の両方に書く
    For i = 0 To UBound(ids)
        RunCase ids(i)
    Next

    gXL.Quit
    Set gXL = Nothing
    WriteResults
    WScript.Echo ""
    WScript.Echo "PASS: " & gPass & "  FAIL: " & gFail & "  SKIP: " & gSkip & "  （" & Round(Timer - t0) & " 秒）"
    WScript.Echo "結果: " & gFso.BuildPath(gBase, RESULT_FILE)
    If gFail > 0 Then RunAllTests = 1 Else RunAllTests = 0
End Function

' テストIDに対応するテストを1件実行し、画面に ok / FAIL を表示する
'   テストコード自体が例外で止まっても FAIL として記録し、次のテストへ進む
'   1件の中に確認項目が複数あり、1つでも FAIL があればそのテストは FAIL と表示する
Sub RunCase(id)
    Dim before
    before = gFail
    On Error Resume Next
    Select Case id
        Case "A1": T_A1
        Case "A2": T_A2
        Case "A3": T_A3
        Case "A4": T_A4
        Case "A5": T_A5
        Case "A6": T_A6
        Case "B1a": T_B1 "B1a", CS_UTF8
        Case "B1b": T_B1 "B1b", CS_UTF8BOM
        Case "B2a": T_B2 "B2a", CS_SJIS, "9358"
        Case "B2b": T_B2 "B2b", CS_UTF8, "E5BA97"
        Case "B2c": T_B2 "B2c", CS_UTF8BOM, "EFBBBFE5BA97"
        Case "B3": T_B3
        Case "B4": T_B4
        Case "B5a": T_B5a
        Case "B5b": T_B5b
        Case "B6": T_B6
        Case "B7": T_B7
        Case "C1": T_C1
        Case "C2": T_C2
        Case "C3": T_C3
        Case "C4": T_C4
        Case "C5": T_C5
        Case "C6": T_C6
        Case "C7": T_C7
        Case "C8": T_C8
        Case "C9": T_C9
        Case "C10": T_C10
        Case "C11": T_C11
        Case "C13": T_C13
        Case "C14": T_C14
        Case "C15": T_C15
        Case "D1": T_D1
        Case "D2": T_D2
        Case "D3": T_D3
        Case "D4a": T_D4a
        Case "D4b": T_D4b
        Case "D5": T_D5
        Case "D6": T_D6
        Case "E1": T_E1
        Case "E2a": T_E2 "E2a", False
        Case "E2b": T_E2 "E2b", True
        Case "E3": T_E3
        Case "E4": T_E4
        Case "E5a": T_E5a
        Case "E5b": T_E5b
        Case "E5c": T_E5c
        Case "E5d": T_E5d
        Case "E5e": T_E5e
        Case "E6a": T_E6a
        Case "E6b": T_E6b
        Case "E7": T_E7
        Case "E8a": T_E8a
        Case "E8b": T_E8b
        Case "E9": T_E9
        Case "E10": T_E10
        Case "E11a": T_E11a
        Case "E11b": T_E11b
        Case "E12": T_E12
        Case "E13": T_E13
        Case "E14": T_E14
        Case "F2": T_F2
        Case "F5": T_F5
        Case "G1a": T_G1a
        Case "G1b": T_G1b
        Case "G1c": T_G1c
        Case "G1d": T_G1d
        Case "G2": T_G2
        Case "G3": T_G3
        Case "H1": T_H1
        Case "H2": T_H2
        Case "H3": T_H3
        Case "H4": T_H4
        Case "H5a": T_H5a
        Case "H5b": T_H5b
        Case "H6": T_H6
        Case "H7": T_H7
        Case "H8": T_H8
        Case "P1": T_P1
    End Select
    If Err.Number <> 0 Then
        Check "テストコードの実行", False, "例外: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    If gFail > before Then
        WScript.Echo "FAIL  " & id & "  " & gTitle
    Else
        WScript.Echo "ok    " & id & "  " & gTitle
    End If
End Sub


'==============================================================
' テストケース
'   1つの Sub が1つのテスト（T_ + テストID）。BeginCase の第2引数が「何を確認するか」の題名
'   設定・マッピングは BeginCase で標準値になり、各テストでは変えたいところだけを書き換える
'   標準のマッピング（BeginCase 参照）
'     店舗コード ← 店舗CD（TRIM|ゼロ埋め:5）/ 取引日 ← 売上日（日付:yyyymmdd）
'     数量 ← 販売数（数値:0）/ 区分 ← 区分（コード:区分 で 1→01, 2→02）/ 連携区分 = 固定値 1
'==============================================================
' 標準の入力・期待値
Function StdHdr(): StdHdr = "店舗CD,売上日,販売数,区分": End Function
Function OutHdr(): OutHdr = "店舗コード,取引日,数量,区分,連携区分": End Function

' 標準の入力：1106 が2ファイル（計3行）、1107 が1ファイル（1行）
'   Exp1106 / Exp1107 がそれぞれを標準マッピングで変換したときの期待値
Sub MakeStdInputs()
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr(), "12,20261106,5,1", "345,20261106,10,2")), CS_SJIS
    MakeCsv "in\fileA_20261106_02.csv", L(Array(StdHdr(), "7,20261106,1,1")), CS_SJIS
    MakeCsv "in\fileA_20261107_01.csv", L(Array(StdHdr(), "12,20261107,3,2")), CS_SJIS
End Sub

Function Exp1106(): Exp1106 = Array(OutHdr(), "00012,20261106,5,01,1", "00345,20261106,10,02,1", "00007,20261106,1,01,1"): End Function
Function Exp1107(): Exp1107 = Array(OutHdr(), "00012,20261107,3,02,1"): End Function

'---------------- A: 基本動作 ----------------
' A1: 標準の入力で、日付ごとの出力内容・ログの OK 行・ログが BOM 付き UTF-8 であることを確認
Sub T_A1()
    Dim rc
    BeginCase "A1", "基本：同じ日付の2ファイルを結合、別日付は別ファイル"
    MakeStdInputs
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Exp1106())
    ExpectText "out\fileB_20261107.csv", CS_SJIS, L(Exp1107())
    ExpectLog "OK", "fileB_20261106.csv"
    ExpectLog "OK", "2 ファイル → 3 行 出力"
    ExpectLogHex "EFBBBF"
End Sub

Sub T_A2()
    Dim rc
    BeginCase "A2", "結合順：ファイル名順（作成順ではない）"
    MakeCsv "in\fileA_20261106_02.csv", L(Array(StdHdr(), "7,20261106,1,1")), CS_SJIS
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr(), "12,20261106,5,1")), CS_SJIS
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array(OutHdr(), "00012,20261106,5,01,1", "00007,20261106,1,01,1"))
End Sub

Sub T_A3()
    Dim rc
    BeginCase "A3", "ファイルごとに列の並びが違っても項目名で対応"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr(), "12,20261106,5,1")), CS_SJIS
    MakeCsv "in\fileA_20261106_02.csv", L(Array("区分,販売数,売上日,店舗CD", "2,10,20261106,345")), CS_SJIS
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array(OutHdr(), "00012,20261106,5,01,1", "00345,20261106,10,02,1"))
End Sub

Sub T_A4()
    Dim rc
    BeginCase "A4", "固定値：fileA項目が空欄の列だけ固定値を使う"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr(), "12,20261106,5,1")), CS_SJIS
    gMap = Array(Array("店舗", "店舗CD", "X", "", "", ""), Array("固定", "", "ABC", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("店舗,固定", "12,ABC"))
End Sub

Sub T_A5()
    Dim rc
    BeginCase "A5", "空行・区切りだけの行・末尾の空行は出力しない"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr(), "12,20261106,5,1", "", ",,,", "345,20261106,10,2", "", "")), CS_SJIS
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array(OutHdr(), "00012,20261106,5,01,1", "00345,20261106,10,02,1"))
End Sub

'---------------- B: 文字コード・改行・区切り・クォート ----------------
' 文字コード系で使う、備考列つきの入力
'   備考に全角文字（テスト　漢字 など）を入れて、文字化けしないかを見る
' A6: 送り元の番号が付いたファイル名（17000_20261106.csv など）
'   番号が違っても同じ日付なら1つの fileB にまとまり、日付が変われば別の fileB になる。
'   日付の後ろに時刻が付いた名前（17002_20261106093000.csv）も、先頭8桁を日付として扱う
Sub T_A6()
    Dim rc
    BeginCase "A6", "番号つきのファイル名：番号が違っても同じ日付は1ファイルに結合"
    MakeCsv "in\17000_20261106.csv", L(Array(StdHdr(), "12,20261106,5,1", "345,20261106,10,2")), CS_SJIS
    MakeCsv "in\17002_20261106093000.csv", L(Array(StdHdr(), "7,20261106,1,1")), CS_SJIS
    MakeCsv "in\17000_20261107.csv", L(Array(StdHdr(), "12,20261107,3,2")), CS_SJIS
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Exp1106())
    ExpectText "out\fileB_20261107.csv", CS_SJIS, L(Exp1107())
    ExpectLog "OK", "2 ファイル → 3 行 出力"
End Sub

Function BikoHdr(): BikoHdr = "店舗CD,売上日,販売数,区分,備考": End Function
Sub UseBikoMap()
    gMap = Array(Array("店舗コード", "店舗CD", "", "TRIM|ゼロ埋め:5", "", ""), Array("備考", "備考", "", "", "", ""))
End Sub

' B1a / B1b: 入力が UTF-8（BOMなし・あり）でも読めること。BOM が先頭の項目名にくっつかないこと
Sub T_B1(id, cs)
    Dim rc
    BeginCase id, "入力の文字コード：" & cs & "（BOMが項目名に混ざらない）"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(BikoHdr(), "12,20261106,5,1,テスト①髙")), cs
    SetV "入力文字コード", cs
    UseBikoMap
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("店舗コード,備考", "00012,テスト①髙"))
End Sub

' B2a～c: 出力の文字コード。内容に加え、ファイル先頭のバイト列（hexPrefix）で実際の文字コードを確認
'   9358 = Shift_JIS の「店」/ E5BA97 = UTF-8 の「店」/ EFBBBF = UTF-8 の BOM
Sub T_B2(id, cs, hexPrefix)
    Dim rc
    BeginCase id, "出力の文字コード：" & cs
    MakeCsv "in\fileA_20261106_01.csv", L(Array(BikoHdr(), "12,20261106,5,1,テスト①髙")), CS_SJIS
    gMap = Array(Array("店舗", "店舗CD", "", "", "", ""), Array("備考", "備考", "", "", "", ""))
    SetV "出力文字コード", cs
    rc = RunTool()
    ExpectExit rc, 0
    If cs = CS_SJIS Then
        ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("店舗,備考", "12,テスト①髙"))
    Else
        ExpectText "out\fileB_20261106.csv", CS_UTF8, L(Array("店舗,備考", "12,テスト①髙"))
    End If
    ExpectHex "out\fileB_20261106.csv", hexPrefix
End Sub

Sub T_B3()
    Dim rc
    BeginCase "B3", "改行コード LF"
    MakeStdInputs
    SetV "改行コード", "LF"
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261107.csv", CS_SJIS, Join(Exp1107(), vbLf) & vbLf
End Sub

Sub T_B4()
    Dim rc
    BeginCase "B4", "タブ区切り（入力・出力）"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(Replace(StdHdr(), ",", vbTab), "12" & vbTab & "20261106" & vbTab & "5" & vbTab & "1")), CS_SJIS
    SetV "区切り文字", "TAB"
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array(Replace(OutHdr(), ",", vbTab), Replace("00012,20261106,5,01,1", ",", vbTab)))
End Sub

Sub T_B5a()
    Dim rc, q
    q = """"
    BeginCase "B5a", "カンマ・改行・ダブルクォートを含む値（入力も出力も正しく囲む）"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(BikoHdr(), _
        "12,20261106,5,1," & q & "a,b" & q, _
        "13,20261106,5,1," & q & "1行目" & vbCrLf & "2行目" & q, _
        "14,20261106,5,1," & q & "say " & q & q & "hi" & q & q & q)), CS_SJIS
    UseBikoMap
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("店舗コード,備考", _
        "00012," & q & "a,b" & q, _
        "00013," & q & "1行目" & vbCrLf & "2行目" & q, _
        "00014," & q & "say " & q & q & "hi" & q & q & q))
End Sub

Sub T_B5b()
    Dim rc
    BeginCase "B5b", "全項目をダブルクォートで囲む（空の値は """"）"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(BikoHdr(), "12,20261106,5,1,")), CS_SJIS
    UseBikoMap
    SetV "全項目をダブルクォートで囲む", "あり"
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("""店舗コード"",""備考""", """00012"","""""))
End Sub

Sub T_B6()
    Dim rc
    BeginCase "B6", "ヘッダーなし（列番号で指定）"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("12,20261106,5,1", "345,20261106,10,2")), CS_SJIS
    SetV "ヘッダー行番号", "0"
    gMap = Array(Array("店舗コード", "1", "", "ゼロ埋め:5", "", ""), Array("区分", "4", "", "コード:区分", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("店舗コード,区分", "00012,01", "00345,02"))
End Sub

Sub T_B7()
    Dim rc
    BeginCase "B7", "機種依存文字・半角カナ（Shift_JIS 入出力）"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(BikoHdr(), "12,20261106,5,1,①Ⅱ㈱髙﨑ｱｲｳﾞ")), CS_SJIS
    UseBikoMap
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("店舗コード,備考", "00012,①Ⅱ㈱髙﨑ｱｲｳﾞ"))
End Sub

'---------------- C: Excel 入出力 ----------------
' Excel 入力系（対象ファイルを *.xlsx にする）
Sub UseXlsxInput()
    SetV "対象ファイル", "*.xlsx"
End Sub

Sub T_C1()
    Dim rc
    BeginCase "C1", "CSVの 0001 は先頭ゼロのまま"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("商品CD", "0001")), CS_SJIS
    gMap = Array(Array("商品コード", "商品CD", "", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("商品コード", "0001"))
End Sub

Sub T_C2()
    Dim rc
    BeginCase "C2", "Excelの数値セル 1 に ゼロ埋め:4 → 0001"
    MakeXlsx "in\fileA_20261106_01.xlsx", Array(Array("商品CD"), Array(1)), 1, 1, Empty
    UseXlsxInput
    gMap = Array(Array("商品コード", "商品CD", "", "ゼロ埋め:4", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("商品コード", "0001"))
End Sub

Sub T_C3()
    Dim rc
    BeginCase "C3", "Excelの文字列セル 0001 はそのまま"
    MakeXlsx "in\fileA_20261106_01.xlsx", Array(Array("商品CD"), Array("0001")), 1, 1, Empty
    UseXlsxInput
    gMap = Array(Array("商品コード", "商品CD", "", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("商品コード", "0001"))
End Sub

Sub T_C4()
    Dim rc, d
    d = DateSerial(2026, 11, 6) + TimeSerial(10, 30, 0)
    BeginCase "C4", "Excelの日付セル（時刻付き）を日付変換"
    MakeXlsx "in\fileA_20261106_01.xlsx", Array(Array("売上日"), Array(Array(d, "yyyy/m/d h:mm"))), 1, 1, Empty
    UseXlsxInput
    gMap = Array(Array("取引日", "売上日", "", "日付:yyyymmdd", "", ""), Array("取引日時", "売上日", "", "日付:yyyy/mm/dd hh:mm", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("取引日,取引日時", "20261106,2026/11/06 10:30"))
End Sub

Sub T_C5()
    Dim rc
    BeginCase "C5", "CSVの日付文字列（複数の書き方）を日付変換"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("売上日", "2026/11/06", "20261106", "2026-11-06", "2026/11/6", "２０２６１１０６")), CS_SJIS
    gMap = Array(Array("取引日", "売上日", "", "日付:yyyymmdd", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("取引日", "20261106", "20261106", "20261106", "20261106", "20261106"))
End Sub

Sub T_C6()
    Dim rc
    BeginCase "C6", "Excelの数値は表示形式ではなく値で変換"
    MakeXlsx "in\fileA_20261106_01.xlsx", Array(Array("金額"), Array(Array(1234.5, "#,##0"))), 1, 1, Empty
    UseXlsxInput
    gMap = Array(Array("金額2桁", "金額", "", "数値:0.00", "", ""), Array("金額そのまま", "金額", "", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("金額2桁,金額そのまま", "1234.50,1234.5"))
End Sub

Sub T_C7()
    Dim rc
    BeginCase "C7", "Excelのエラーセル（#N/A）はその日付をエラーにする"
    If IsMock() Then
        SkipCheck "Excel実機のみ（エラー値はモックで再現できない）"
        Exit Sub
    End If
    MakeXlsx "in\fileA_20261106_01.xlsx", Array(Array("商品CD"), Array(Array("=NA()", "FORMULA"))), 1, 1, Empty
    MakeXlsx "in\fileA_20261107_01.xlsx", Array(Array("商品CD"), Array("A")), 1, 1, Empty
    UseXlsxInput
    gMap = Array(Array("商品コード", "商品CD", "", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 1
    ExpectNoFile "out\fileB_20261106.csv"
    ExpectText "out\fileB_20261107.csv", CS_SJIS, L(Array("商品コード", "A"))
    ExpectLog "ERROR", "エラー値"
    ExpectLog "ERROR", "2行目"
End Sub

Sub T_C8()
    Dim rc
    BeginCase "C8", "表がA1から始まらないExcel・2枚目のシートは読まない"
    MakeXlsx "in\fileA_20261106_01.xlsx", Array(Array("商品CD", "数量"), Array("A1", 5)), 3, 2, _
             Array(Array("商品CD", "数量"), Array("別シート", 99))
    UseXlsxInput
    SetV "ヘッダー行番号", "3"
    gMap = Array(Array("商品コード", "商品CD", "", "", "", ""), Array("数量", "数量", "", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("商品コード,数量", "A1,5"))
End Sub

Sub T_C9()
    Dim rc
    BeginCase "C9", "同じ日付にCSVとExcelが混在しても1ファイルに結合"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr(), "12,20261106,5,1")), CS_SJIS
    MakeXlsx "in\fileA_20261106_02.xlsx", Array(Array("店舗CD", "売上日", "販売数", "区分"), Array(7, "20261106", 1, 1)), 1, 1, Empty
    SetV "対象ファイル", "fileA_*.*"
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array(OutHdr(), "00012,20261106,5,01,1", "00007,20261106,1,01,1"))
End Sub

Sub T_C10()
    Dim rc
    BeginCase "C10", "Excelの一時ファイル（~$）は対象外"
    MakeXlsx "in\fileA_20261106_01.xlsx", Array(Array("商品CD"), Array("A")), 1, 1, Empty
    MakeCsv "in\~$fileA_20261106_01.xlsx", "壊れたファイル", CS_SJIS
    UseXlsxInput
    gMap = Array(Array("商品コード", "商品CD", "", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("商品コード", "A"))
End Sub

Sub T_C11()
    Dim rc, exp, i
    BeginCase "C11", "Excel出力：CSV出力と同じ値・全セル文字列・シート名・拡張子・=で始まる値"
    MakeStdInputs
    SetV "出力形式", "Excel"
    SetV "出力シート名", "連携データ"
    gMap = Array(Array("店舗コード", "店舗CD", "", "TRIM|ゼロ埋め:5", "", ""), Array("取引日", "売上日", "", "日付:yyyymmdd", "", ""), _
                 Array("数量", "販売数", "", "数値:0", "", ""), Array("区分", "区分", "", "コード:区分", "", ""), _
                 Array("連携区分", "", "1", "", "", ""), Array("式", "", "=1+1", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    exp = Exp1106()
    For i = 0 To UBound(exp)
        If i = 0 Then exp(i) = exp(i) & ",式" Else exp(i) = exp(i) & ",=1+1"
    Next
    ExpectXlsx "out\fileB_20261106.xlsx", "連携データ", RowsOf(exp), True
    ExpectNoFile "out\fileB_20261106.csv"
End Sub

Sub T_C13()
    Dim rc
    BeginCase "C13", "Excelの13桁数値（JANコード）が指数表記にならない"
    MakeXlsx "in\fileA_20261106_01.xlsx", Array(Array("JAN"), Array(4912345678901)), 1, 1, Empty
    UseXlsxInput
    gMap = Array(Array("JANコード", "JAN", "", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("JANコード", "4912345678901"))
End Sub

Sub T_C14()
    Dim rc
    BeginCase "C14", "Excelの数式セルは計算結果で読む"
    MakeXlsx "in\fileA_20261106_01.xlsx", Array(Array("数量"), Array(Array("=1+2", "FORMULA"))), 1, 1, Empty
    UseXlsxInput
    gMap = Array(Array("数量", "数量", "", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("数量", "3"))
End Sub

Sub T_C15()
    Dim rc
    BeginCase "C15", "負の数・全角数字・全角カンマの数値変換"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("数量", "-5", "１２", """１，２００""")), CS_SJIS
    gMap = Array(Array("数量", "数量", "", "数値:0", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("数量", "-5", "12", "1200"))
End Sub

'---------------- D: 値の変換 ----------------
Sub T_D1()
    Dim rc
    BeginCase "D1", "TRIM：全角スペース・タブも除去"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("店舗CD", "　12" & vbTab)), CS_SJIS
    gMap = Array(Array("店舗コード", "店舗CD", "", "TRIM|ゼロ埋め:5", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("店舗コード", "00012"))
End Sub

Sub T_D2()
    Dim rc
    BeginCase "D2", "各変換：置換・空なら・左・右・桁区切り・日付書式"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("電話,備考,名称,金額,売上日", "03-1234-5678,,ABCDE,1234567,20261106")), CS_SJIS
    gMap = Array(Array("置換", "電話", "", "置換:-=>", "", ""), Array("空なら", "備考", "", "空なら:未設定", "", ""), _
                 Array("左", "名称", "", "左:3", "", ""), Array("右", "名称", "", "右:2", "", ""), _
                 Array("桁区切り", "金額", "", "数値:#,##0", "", ""), Array("日付", "売上日", "", "日付:yyyy-mm-dd", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("置換,空なら,左,右,桁区切り,日付", "0312345678,未設定,ABC,DE,""1,234,567"",2026-11-06"))
End Sub

Sub T_D3()
    Dim rc
    BeginCase "D3", "変換の連結は左から順に適用"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("電話", "03-1234")), CS_SJIS
    gMap = Array(Array("A", "電話", "", "置換:-=>|左:4", "", ""), Array("B", "電話", "", "左:4|置換:-=>", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("A,B", "0312,031"))
End Sub

Sub T_D4a()
    Dim rc
    BeginCase "D4a", "コード変換：表にない値は * の既定値"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("区分", "1", "2", "3")), CS_SJIS
    gMap = Array(Array("区分", "区分", "", "コード:区分", "", ""))
    gCodes = Array(Array("区分", "1", "01"), Array("区分", "2", "02"), Array("区分", "*", "99"))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("区分", "01", "02", "99"))
End Sub

Sub T_D4b()
    Dim rc
    BeginCase "D4b", "コード変換：既定値がなく表にない値はその日付をエラー"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("区分", "1", "3")), CS_SJIS
    MakeCsv "in\fileA_20261107_01.csv", L(Array("区分", "2")), CS_SJIS
    gMap = Array(Array("区分", "区分", "", "コード:区分", "", ""))
    rc = RunTool()
    ExpectExit rc, 1
    ExpectNoFile "out\fileB_20261106.csv"
    ExpectText "out\fileB_20261107.csv", CS_SJIS, L(Array("区分", "02"))
    ExpectLog "ERROR", "値「3」がありません"
End Sub

Sub T_D5()
    Dim rc
    BeginCase "D5", "不正な日付・数値はその日付だけエラー（場所をログに出す）"
    MakeCsv "in\fileA_20261101_01.csv", L(Array(StdHdr(), "12,20261332,5,1")), CS_SJIS
    MakeCsv "in\fileA_20261102_01.csv", L(Array(StdHdr(), "12,20261102,abc,1")), CS_SJIS
    MakeCsv "in\fileA_20261103_01.csv", L(Array(StdHdr(), "12,20261103,5,1")), CS_SJIS
    rc = RunTool()
    ExpectExit rc, 1
    ExpectNoFile "out\fileB_20261101.csv"
    ExpectNoFile "out\fileB_20261102.csv"
    ExpectText "out\fileB_20261103.csv", CS_SJIS, L(Array(OutHdr(), "00012,20261103,5,01,1"))
    ExpectLog "ERROR", "日付として不正です"
    ExpectLog "ERROR", "数値に変換できません"
    ExpectLog "ERROR", "2行目 / 数量"
End Sub

Sub T_D6()
    Dim rc
    BeginCase "D6", "不明な変換名はエラー"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("店舗CD", "12")), CS_SJIS
    gMap = Array(Array("店舗", "店舗CD", "", "XYZ:1", "", ""))
    rc = RunTool()
    ExpectExit rc, 1
    ExpectNoFile "out\fileB_20261106.csv"
    ExpectLog "ERROR", "不明な変換指定"
End Sub

'---------------- E: 異常系 ----------------
' 多くのテストで「終了コード 1」「その日付のファイルが作られない」「ログに ERROR と理由」を確認する
Sub T_E1()
    Dim rc
    BeginCase "E1", "項目が足りない日付だけ出力しない（他の日付は出力）"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("店舗CD,売上日,販売数", "12,20261106,5")), CS_SJIS
    MakeCsv "in\fileA_20261107_01.csv", L(Array(StdHdr(), "12,20261107,3,2")), CS_SJIS
    rc = RunTool()
    ExpectExit rc, 1
    ExpectNoFile "out\fileB_20261106.csv"
    ExpectText "out\fileB_20261107.csv", CS_SJIS, L(Exp1107())
    ExpectLog "ERROR", "fileAに項目が見つかりません: [区分]"
End Sub

Sub T_E2(id, overwrite)
    Dim rc
    If overwrite Then
        BeginCase id, "出力先に同名ファイル：上書き あり → 置き換える"
        SetV "上書き", "あり"
    Else
        BeginCase id, "出力先に同名ファイル：上書き なし → スキップして既存ファイルを変更しない"
    End If
    MakeStdInputs
    MakeCsv "out\fileB_20261106.csv", "OLD" & vbCrLf, CS_SJIS
    rc = RunTool()
    ExpectExit rc, 0
    If overwrite Then
        ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Exp1106())
    Else
        ExpectText "out\fileB_20261106.csv", CS_SJIS, "OLD" & vbCrLf
        ExpectLog "SKIP", "既にある"
    End If
    ExpectText "out\fileB_20261107.csv", CS_SJIS, L(Exp1107())
End Sub

Sub T_E3()
    Dim rc
    BeginCase "E3", "ファイル名から日付が取れないファイルはSKIP"
    MakeStdInputs
    MakeCsv "in\fileA_latest.csv", L(Array(StdHdr(), "1,20261106,1,1")), CS_SJIS
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Exp1106())
    ExpectLog "SKIP", "fileA_latest.csv"
End Sub

Sub T_E4()
    Dim rc
    BeginCase "E4", "対象ファイル0件はエラー終了"
    rc = RunTool()
    ExpectExit rc, 1
    ExpectLog "ERROR", "対象ファイルが見つかりません"
End Sub

Sub T_E5a()
    Dim rc
    BeginCase "E5a", "設定不備：入力フォルダがない"
    SetV "入力フォルダ", gFso.BuildPath(gDir, "nothing")
    rc = RunTool()
    ExpectExit rc, 1
End Sub

Sub T_E5b()
    Dim rc
    BeginCase "E5b", "設定不備：入力と出力が同じフォルダ"
    MakeStdInputs
    SetV "出力フォルダ", gFso.BuildPath(gDir, "in")
    rc = RunTool()
    ExpectExit rc, 1
    ExpectNoFile "in\fileB_20261106.csv"
End Sub

Sub T_E5c()
    Dim rc
    BeginCase "E5c", "設定不備：設定キーの欠落"
    MakeStdInputs
    gRemoveKeys = Array("上書き")
    rc = RunTool()
    ExpectExit rc, 1
    ExpectNoFile "out\fileB_20261106.csv"
End Sub

Sub T_E5d()
    Dim rc
    BeginCase "E5d", "設定不備：数値の欄に文字"
    MakeStdInputs
    SetV "ヘッダー行番号", "abc"
    rc = RunTool()
    ExpectExit rc, 1
End Sub

Sub T_E5e()
    Dim rc
    BeginCase "E5e", "設定不備：出力形式が不正"
    MakeStdInputs
    SetV "出力形式", "PDF"
    rc = RunTool()
    ExpectExit rc, 1
End Sub

Sub T_E6a()
    Dim rc
    BeginCase "E6a", "0バイトのファイルは項目不足でエラー"
    MakeCsv "in\fileA_20261106_01.csv", "", CS_SJIS
    rc = RunTool()
    ExpectExit rc, 1
    ExpectNoFile "out\fileB_20261106.csv"
    ExpectLog "ERROR", "見つかりません"
End Sub

Sub T_E6b()
    Dim rc
    BeginCase "E6b", "見出しだけのファイルは0行（見出しのみ出力＋WARN）"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr())), CS_SJIS
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array(OutHdr()))
    ExpectLog "WARN", "0行"
End Sub

Sub T_E7()
    Dim rc
    BeginCase "E7", "fileAに同じ項目名が2つあればエラー"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("店舗CD,売上日,販売数,区分,店舗CD", "12,20261106,5,1,99")), CS_SJIS
    rc = RunTool()
    ExpectExit rc, 1
    ExpectLog "ERROR", "同じ項目名が複数あります: [店舗CD]"
End Sub

Sub T_E8a()
    Dim rc
    BeginCase "E8a", "見出しの前後の空白は無視"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(" 店舗CD ,売上日,販売数,区分", "12,20261106,5,1")), CS_SJIS
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array(OutHdr(), "00012,20261106,5,01,1"))
End Sub

Sub T_E8b()
    Dim rc
    BeginCase "E8b", "全角・半角の違う見出しは別の項目（エラー）"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("店舗ＣＤ,売上日,販売数,区分", "12,20261106,5,1")), CS_SJIS
    rc = RunTool()
    ExpectExit rc, 1
    ExpectLog "ERROR", "[店舗CD]"
End Sub

' E9: 出力先の xlsx をテスト側の Excel で開いたまま（ロックした状態で）ツールを実行し、
'     その日付だけエラーになり、他の日付は出力されることを確認する
Sub T_E9()
    Dim rc, wb, p
    BeginCase "E9", "出力先ファイルが開かれている日付はエラー（他は続行）"
    If IsMock() Then
        SkipCheck "Excel実機のみ（ファイルのロック）"
        Exit Sub
    End If
    MakeStdInputs
    SetV "出力形式", "Excel"
    SetV "上書き", "あり"
    MakeXlsx "out\fileB_20261106.xlsx", Array(Array("OLD")), 1, 1, Empty
    p = gFso.BuildPath(gDir, "out\fileB_20261106.xlsx")
    Set wb = gXL.Workbooks.Open(p, XL_NO_UPDATE_LINKS, False)
    rc = RunTool()
    wb.Close False
    ExpectExit rc, 1
    ExpectLog "ERROR", "20261106"
    ExpectLog "OK", "fileB_20261107.xlsx"
End Sub

Sub T_E10()
    Dim rc
    BeginCase "E10", "末尾に改行なし・EOF文字（0x1A）があっても正しく読む"
    MakeCsv "in\fileA_20261106_01.csv", StdHdr() & vbCrLf & "12,20261106,5,1" & Chr(26), CS_SJIS
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array(OutHdr(), "00012,20261106,5,01,1"))
End Sub

Sub T_E11a()
    Dim rc
    BeginCase "E11a", "Excel出力：シート名に使えない文字は設定エラー"
    MakeStdInputs
    SetV "出力形式", "Excel"
    SetV "出力シート名", "a/b"
    rc = RunTool()
    ExpectExit rc, 1
    ExpectNoFile "out\fileB_20261106.xlsx"
End Sub

' E11b: 32768 文字（Excel のセル上限 32767 を1文字超える）の値でエラーになることを確認
Sub T_E11b()
    Dim rc
    BeginCase "E11b", "Excel出力：セルの最大文字数を超える値はその日付をエラー"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("備考", String(32768, "x"))), CS_SJIS
    MakeCsv "in\fileA_20261107_01.csv", L(Array("備考", "ok")), CS_SJIS
    SetV "出力形式", "Excel"
    gMap = Array(Array("備考", "備考", "", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 1
    ExpectNoFile "out\fileB_20261106.xlsx"
    ExpectXlsx "out\fileB_20261107.xlsx", "Sheet1", RowsOf(Array("備考", "ok")), True
    ExpectLog "ERROR", "最大文字数"
End Sub

Sub T_E12()
    Dim rc
    BeginCase "E12", "データ0行の日付を出力しない設定"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr())), CS_SJIS
    SetV "データ0行の日付も出力", "なし"
    rc = RunTool()
    ExpectExit rc, 0
    ExpectNoFile "out\fileB_20261106.csv"
    ExpectLog "SKIP", "0行"
End Sub

Sub T_E13()
    Dim rc
    BeginCase "E13", "閉じていないダブルクォートはエラー"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("備考", """abc")), CS_SJIS
    gMap = Array(Array("備考", "備考", "", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 1
    ExpectNoFile "out\fileB_20261106.csv"
    ExpectLog "ERROR", "ダブルクォートが閉じられていません"
End Sub

' E14: ファイル名から取った8桁が日付として正しくない
'   1700020261106.csv（番号と日付がつながっている）は先頭8桁「17000202」を取ってしまうが、
'   年が範囲外なので SKIP。fileA_20261332.csv（13月）も SKIP。正しい日付のファイルは出力される
Sub T_E14()
    Dim rc
    BeginCase "E14", "ファイル名の日付が正しくないファイルはSKIP（不正な名前のfileBを作らない）"
    MakeCsv "in\1700020261106.csv", L(Array(StdHdr(), "1,20261106,1,1")), CS_SJIS
    MakeCsv "in\fileA_20261332.csv", L(Array(StdHdr(), "1,20261106,1,1")), CS_SJIS
    MakeCsv "in\fileA_20261107_01.csv", L(Array(StdHdr(), "12,20261107,3,2")), CS_SJIS
    rc = RunTool()
    ExpectExit rc, 0
    ExpectNoFile "out\fileB_17000202.csv"
    ExpectNoFile "out\fileB_20261332.csv"
    ExpectText "out\fileB_20261107.csv", CS_SJIS, L(Exp1107())
    ExpectLog "SKIP", "ファイル名の日付が正しくないため対象外: 17000202"
    ExpectLog "SKIP", "ファイル名の日付が正しくないため対象外: 20261332"
End Sub

'---------------- F: 実行方法 ----------------
' F2: ツールを Tool.vbs という名前でコピーし、同じフォルダに Tool.xlsx を置いて引数なしで実行
Sub T_F2()
    Dim rc, tool
    BeginCase "F2", "引数なし実行：スクリプトと同名の設定ファイルを使う"
    MakeStdInputs
    tool = gFso.BuildPath(gDir, "Tool.vbs")
    gFso.CopyFile gTool, tool, True
    WriteConfig gFso.BuildPath(gDir, "Tool.xlsx")
    rc = gSh.Run("cscript //nologo """ & tool & """", 0, True)
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Exp1106())
End Sub

Sub T_F5()
    Dim rc
    BeginCase "F5", "パスにスペース・日本語を含む"
    MakeCsv "入力 フォルダ\fileA_20261106_01.csv", L(Array(StdHdr(), "12,20261106,5,1")), CS_SJIS
    EnsureDir "出力 フォルダ"
    SetV "入力フォルダ", gFso.BuildPath(gDir, "入力 フォルダ")
    SetV "出力フォルダ", gFso.BuildPath(gDir, "出力 フォルダ")
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "出力 フォルダ\fileB_20261106.csv", CS_SJIS, L(Array(OutHdr(), "00012,20261106,5,01,1"))
End Sub

'---------------- G: マッピングの空列 ----------------
' fileA項目・固定値・変換のどれもない列（常に空）が、CSV / Excel でどう出力されるか
Sub T_G1a()
    Dim rc
    BeginCase "G1a", "常に空の列（途中）：列がずれず ,, になる＋WARN"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr(), "12,20261106,5,1")), CS_SJIS
    gMap = Array(Array("店舗", "店舗CD", "", "", "", ""), Array("備考", "", "", "", "", ""), Array("数量", "販売数", "", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("店舗,備考,数量", "12,,5"))
    ExpectLog "WARN", "[備考]"
End Sub

Sub T_G1b()
    Dim rc
    BeginCase "G1b", "常に空の列＋全項目囲む：空は """""
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr(), "12,20261106,5,1")), CS_SJIS
    gMap = Array(Array("店舗", "店舗CD", "", "", "", ""), Array("備考", "", "", "", "", ""), Array("数量", "販売数", "", "", "", ""))
    SetV "全項目をダブルクォートで囲む", "あり"
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("""店舗"",""備考"",""数量""", """12"","""",""5"""))
End Sub

Sub T_G1c()
    Dim rc
    BeginCase "G1c", "常に空の列＋全項目囲む＋空は囲まない：,, になる"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr(), "12,20261106,5,1")), CS_SJIS
    gMap = Array(Array("店舗", "店舗CD", "", "", "", ""), Array("備考", "", "", "", "", ""), Array("数量", "販売数", "", "", "", ""))
    SetV "全項目をダブルクォートで囲む", "あり"
    SetV "空の値もダブルクォートで囲む", "なし"
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("""店舗"",""備考"",""数量""", """12"",,""5"""))
End Sub

Sub T_G1d()
    Dim rc
    BeginCase "G1d", "常に空の列＋Excel出力：空のセル"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr(), "12,20261106,5,1")), CS_SJIS
    gMap = Array(Array("店舗", "店舗CD", "", "", "", ""), Array("備考", "", "", "", "", ""), Array("数量", "販売数", "", "", "", ""))
    SetV "出力形式", "Excel"
    rc = RunTool()
    ExpectExit rc, 0
    ExpectXlsx "out\fileB_20261106.xlsx", "Sheet1", RowsOf(Array("店舗,備考,数量", "12,,5")), True
End Sub

Sub T_G2()
    Dim rc
    BeginCase "G2", "一部の行だけ空の値：空のまま／空なら で補完"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(BikoHdr(), "12,20261106,5,1,あり", "13,20261106,5,1,")), CS_SJIS
    gMap = Array(Array("店舗", "店舗CD", "", "", "", ""), Array("備考", "備考", "", "", "", ""), Array("備考2", "備考", "", "空なら:なし", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("店舗,備考,備考2", "12,あり,あり", "13,,なし"))
End Sub

Sub T_G3()
    Dim rc
    BeginCase "G3", "先頭・末尾の列が空でも列数がずれない"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr(), "12,20261106,5,1")), CS_SJIS
    gMap = Array(Array("先頭空", "", "", "", "", ""), Array("店舗", "店舗CD", "", "", "", ""), Array("末尾空", "", "", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("先頭空,店舗,末尾空", ",12,"))
    ExpectLog "WARN", "[先頭空] [末尾空]"
End Sub

'---------------- H: 抽出条件・連結・レコード・フッター ----------------
' 抽出条件・連結・レコード・フッター
'   gFilters の各行は Array(種別, fileA項目名, 比較, 値)
Function FilHdr(): FilHdr = "店舗CD,区分,備考": End Function
Sub UseFilMap()
    gMap = Array(Array("店舗", "店舗CD", "", "", "", ""), Array("区分", "区分", "", "", "", ""))
End Sub

Sub T_H1()
    Dim rc
    BeginCase "H1", "抽出条件：含める（区分 = 1 の行だけ）"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(FilHdr(), "12,1,", "13,2,", "14,1,")), CS_SJIS
    UseFilMap
    gFilters = Array(Array("含める", "区分", "等しい", "1"))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("店舗,区分", "12,1", "14,1"))
    ExpectLog "INFO", "1 行除外"
End Sub

Sub T_H2()
    Dim rc
    BeginCase "H2", "抽出条件：除外する（どれか1つに当てはまれば除外）"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(FilHdr(), "12,1,", ",1,", "13,2,合計行", "14,2,")), CS_SJIS
    UseFilMap
    gFilters = Array(Array("除外する", "店舗CD", "空", ""), Array("除外する", "備考", "含む", "合計"))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("店舗,区分", "12,1", "14,2"))
End Sub

Sub T_H3()
    Dim rc
    BeginCase "H3", "抽出条件：正規表現と含める・除外の組み合わせ"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(FilHdr(), "12,1,", "A3,1,", "14,2,", "15,1,テスト")), CS_SJIS
    UseFilMap
    gFilters = Array(Array("含める", "店舗CD", "正規表現", "^\d+$"), Array("含める", "区分", "等しくない", "2"), Array("除外する", "備考", "空でない", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("店舗,区分", "12,1"))
End Sub

Sub T_H4()
    Dim rc
    BeginCase "H4", "複数項目の連結（空の項目は詰める）"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("姓,名", "山田,太郎", "佐藤,", ",花子")), CS_SJIS
    gMap = Array(Array("氏名", "姓|名", "", "", "　", ""), Array("氏名続け", "姓|名", "", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("氏名,氏名続け", "山田　太郎,山田太郎", "佐藤,佐藤", "花子,花子"))
End Sub

Sub T_H5a()
    Dim rc
    BeginCase "H5a", "先頭・末尾レコード（日付・件数・合計）：CSV"
    MakeStdInputs
    SetV "先頭レコード", "H,{date}"
    SetV "末尾レコード", "T,{count},{sum:数量}"
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("H,20261106", OutHdr(), "00012,20261106,5,01,1", "00345,20261106,10,02,1", "00007,20261106,1,01,1", "T,3,16"))
End Sub

Sub T_H5b()
    Dim rc
    BeginCase "H5b", "先頭・末尾レコード：Excel出力はセルに分割"
    MakeStdInputs
    SetV "出力形式", "Excel"
    SetV "先頭レコード", "H,{date}"
    SetV "末尾レコード", "T,{count},{sum:数量}"
    rc = RunTool()
    ExpectExit rc, 0
    ExpectXlsx "out\fileB_20261107.xlsx", "Sheet1", RowsOf(Array("H,20261107", OutHdr(), "00012,20261107,3,02,1", "T,1,3")), True
End Sub

Sub T_H6()
    Dim rc
    BeginCase "H6", "fileA末尾の合計行を除外（後ろの空行は数えない）"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr(), "12,20261106,5,1", "345,20261106,10,2", "合計,,15,", "", "")), CS_SJIS
    MakeCsv "in\fileA_20261106_02.csv", L(Array(StdHdr(), "7,20261106,1,1", "合計,,1,")), CS_SJIS
    SetV "fileA末尾の除外行数", "1"
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Exp1106())
End Sub

Sub T_H7()
    Dim rc
    BeginCase "H7", "ファイル末尾の改行なし"
    MakeStdInputs
    SetV "ファイル末尾に改行", "なし"
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261107.csv", CS_SJIS, Join(Exp1107(), vbCrLf)
End Sub

Sub T_H8()
    Dim rc
    BeginCase "H8", "抽出条件の項目がfileAにない場合はエラー"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(FilHdr(), "12,1,")), CS_SJIS
    UseFilMap
    gFilters = Array(Array("含める", "存在しない項目", "等しい", "1"))
    rc = RunTool()
    ExpectExit rc, 1
    ExpectNoFile "out\fileB_20261106.csv"
    ExpectLog "ERROR", "[存在しない項目]"
End Sub

'---------------- P: 性能 ----------------
' P1: 10万行の CSV を処理できること。処理時間は合否ではなく参考として結果に残す
Sub T_P1()
    Dim rc, i, lines, t0, n
    n = 100000
    If IsMock() Then n = 2000
    BeginCase "P1", "大量データ（" & n & " 行）の処理時間"
    ReDim lines(n)
    lines(0) = StdHdr()
    For i = 1 To n
        lines(i) = i & ",20261106," & (i Mod 100) & "," & (1 + (i Mod 2))
    Next
    MakeCsv "in\fileA_20261106_01.csv", L(lines), CS_SJIS
    t0 = Timer
    rc = RunTool()
    ExpectExit rc, 0
    ExpectLog "OK", n & " 行 出力"
    Check "処理時間（参考）", True, Round(Timer - t0, 1) & " 秒"
End Sub


'==============================================================
' テストの部品
'==============================================================
' テストの準備：作業フォルダ work\<ID>\（in / out / log）を作り、
' 設定（gSet）・マッピング（gMap）・コード表（gCodes）を標準値に、抽出条件を「なし」に戻す
'   入力・出力・ログフォルダはこのテストの作業フォルダを指すので、テスト同士は影響し合わない
Sub BeginCase(id, title)
    gCase = id
    gTitle = title
    gDir = gFso.BuildPath(gWork, id)
    gFso.CreateFolder gDir
    EnsureDir "in"
    EnsureDir "out"
    EnsureDir "log"
    Set gSet = NewObject("Scripting.Dictionary")
    gSet("入力フォルダ") = gFso.BuildPath(gDir, "in")
    gSet("対象ファイル") = "*.csv"
    gSet("日付抽出パターン") = "(\d{8})"
    gSet("入力文字コード") = CS_SJIS
    gSet("ヘッダー行番号") = "1"
    gSet("fileA末尾の除外行数") = "0"
    gSet("区切り文字") = ","
    gSet("出力フォルダ") = gFso.BuildPath(gDir, "out")
    gSet("出力ファイル名") = "fileB_{date}.csv"
    gSet("出力形式") = "CSV"
    gSet("出力シート名") = "Sheet1"
    gSet("出力文字コード") = CS_SJIS
    gSet("出力ヘッダー") = "あり"
    gSet("全項目をダブルクォートで囲む") = "なし"
    gSet("空の値もダブルクォートで囲む") = "あり"
    gSet("改行コード") = "CRLF"
    gSet("ファイル末尾に改行") = "あり"
    gSet("先頭レコード") = ""
    gSet("末尾レコード") = ""
    gSet("データ0行の日付も出力") = "あり"
    gSet("上書き") = "なし"
    gSet("ログフォルダ") = gFso.BuildPath(gDir, "log")
    Set gSys = NewObject("Scripting.Dictionary")
    gMap = Array(Array("店舗コード", "店舗CD", "", "TRIM|ゼロ埋め:5", "", ""), _
                 Array("取引日", "売上日", "", "日付:yyyymmdd", "", ""), _
                 Array("数量", "販売数", "", "数値:0", "", ""), _
                 Array("区分", "区分", "", "コード:区分", "", ""), _
                 Array("連携区分", "", "1", "", "", ""))
    gCodes = Array(Array("区分", "1", "01"), Array("区分", "2", "02"))
    gFilters = Array()
    gRemoveKeys = Array()
End Sub

' 設定シートの値を1つ変える（例 SetV "出力形式", "Excel"）
Sub SetV(k, v)
    gSet(k) = v
End Sub

' 作業フォルダの中にフォルダを作る（既にあれば何もしない）
Sub EnsureDir(rel)
    Dim p
    p = gFso.BuildPath(gDir, rel)
    If Not gFso.FolderExists(p) Then gFso.CreateFolder p
End Sub

' 行の配列 → CRLF 区切り＋末尾改行
Function L(arr)
    L = Join(arr, vbCrLf) & vbCrLf
End Function

' "a,b" 形式の行の配列 → 2次元の期待値（配列の配列）
Function RowsOf(lines)
    Dim r, i
    ReDim r(UBound(lines))
    For i = 0 To UBound(lines)
        r(i) = Split(lines(i), ",")
    Next
    RowsOf = r
End Function

' 作業フォルダの中にテキストファイル（入力 CSV など）を作る。rel は作業フォルダからの相対パス
Sub MakeCsv(rel, txt, cs)
    Dim p
    p = gFso.BuildPath(gDir, rel)
    If Not gFso.FolderExists(gFso.GetParentFolderName(p)) Then gFso.CreateFolder gFso.GetParentFolderName(p)
    WriteTextFile p, txt, cs
End Sub

' 文字コードを指定してテキストを保存する（ツールの WriteText と同じ方法）
'   CS_UTF8 は BOM なし（先頭3バイトを除いて保存）、CS_UTF8BOM は BOM 付き、空文字なら0バイトのファイル
Sub WriteTextFile(path, txt, cs)
    Dim st, bin
    If Len(txt) = 0 Then
        gFso.CreateTextFile(path, True).Close
        Exit Sub
    End If
    Set st = NewObject("ADODB.Stream")
    st.Type = AD_TYPE_TEXT
    If cs = CS_UTF8BOM Then st.Charset = CS_UTF8 Else st.Charset = cs
    st.Open
    st.WriteText txt
    If cs = CS_UTF8 Then
        st.Position = 0
        st.Type = AD_TYPE_BINARY
        st.Position = UTF8_BOM_BYTES
        Set bin = NewObject("ADODB.Stream")
        bin.Type = AD_TYPE_BINARY
        bin.Open
        st.CopyTo bin
        bin.SaveToFile path, AD_SAVE_OVERWRITE
        bin.Close
    Else
        st.SaveToFile path, AD_SAVE_OVERWRITE
    End If
    st.Close
End Sub

' 文字コードを指定してテキストを読む（出力内容・ログの確認用）
Function ReadTextFile(path, cs)
    Dim st
    Set st = NewObject("ADODB.Stream")
    st.Type = AD_TYPE_TEXT
    If cs = CS_UTF8BOM Then st.Charset = CS_UTF8 Else st.Charset = cs
    st.Open
    st.LoadFromFile path
    ReadTextFile = st.ReadText(-1)
    st.Close
End Function

' ファイルの先頭 n バイトを16進文字列（大文字）で返す（例 BOM なら "EFBBBF"）
Function ReadHex(path, n)
    Dim st, bytes, el
    Set st = NewObject("ADODB.Stream")
    st.Type = AD_TYPE_BINARY
    st.Open
    st.LoadFromFile path
    If st.Size < n Then n = st.Size
    bytes = st.Read(n)
    st.Close
    Set el = NewObject("MSXML2.DOMDocument").createElement("b")
    el.DataType = "bin.hex"
    el.NodeTypedValue = bytes
    ReadHex = UCase(el.Text)
End Function

' Excel の入力ファイルを作る
'   rows : 行の配列。値は 文字列（文字列セル）/ 数値 / 日付 / Array(値, 表示形式) / Array("=式", "FORMULA")
'   r0,c0: 表の左上の位置　other: 2枚目のシートの内容（Empty なら作らない）
Sub MakeXlsx(rel, rows, r0, c0, other)
    Dim wb, p
    p = gFso.BuildPath(gDir, rel)
    Set wb = gXL.Workbooks.Add(XL_WBAT_WORKSHEET)
    If Not IsEmpty(other) Then
        wb.Worksheets.Add
        FillSheet wb.Worksheets(2), other, 1, 1
    End If
    FillSheet wb.Worksheets(1), rows, r0, c0
    If gFso.FileExists(p) Then gFso.DeleteFile p, True
    wb.SaveAs p, XL_FMT_XLSX
    wb.Close False
End Sub

' MakeXlsx の下請け：rows の値の種類に応じてセルに書く
'   文字列 → 書式「@（文字列）」で書く / 数値・日付 → そのまま / Array(値, 書式) → 表示形式つき
'   Array("=式", "FORMULA") → 数式として書く
Sub FillSheet(ws, rows, r0, c0)
    Dim r, c, v, cell
    For r = 0 To UBound(rows)
        For c = 0 To UBound(rows(r))
            v = rows(r)(c)
            Set cell = ws.Cells(r0 + r, c0 + c)
            If IsArray(v) Then
                If v(1) = "FORMULA" Then
                    cell.Formula = v(0)
                Else
                    cell.NumberFormat = v(1)
                    cell.Value = v(0)
                End If
            ElseIf VarType(v) = vbString Then
                cell.NumberFormat = "@"
                cell.Value = v
            Else
                cell.Value = v
            End If
        Next
    Next
End Sub

' 設定ファイルを作る（ひな形をコピーして値を書き換える）
'   本物の DataMapping.xlsx をひな形にするので、シート構成やシステム設定は本番と同じになる
'   設定シート：gSet のキーの行の B 列を書き換え（ひな形にないキーはテストの誤りとして例外）
'   表（マッピング・コード変換・抽出条件）：2～60行目を消してから gMap などを書き込む
Sub WriteConfig(path)
    Dim wb, k, i
    gFso.CopyFile gTemplate, path, True
    Set wb = gXL.Workbooks.Open(path, XL_NO_UPDATE_LINKS, False)
    For Each k In gSet.Keys
        SetKeyValue wb.Worksheets("設定"), k, gSet(k)
    Next
    For Each k In gSys.Keys
        SetKeyValue wb.Worksheets("システム設定"), k, gSys(k)
    Next
    For i = 0 To UBound(gRemoveKeys)
        RemoveKey wb.Worksheets("設定"), gRemoveKeys(i)
    Next
    FillTable wb.Worksheets("マッピング"), gMap, MAP_COLS
    FillTable wb.Worksheets("コード変換"), gCodes, SIMPLE_COLS
    FillTable wb.Worksheets("抽出条件"), gFilters, SIMPLE_COLS
    wb.Save
    wb.Close False
End Sub

' 設定シートで A 列がキー k の行番号（なければ 0）
Function FindKeyRow(ws, k)
    Dim r
    FindKeyRow = 0
    For r = 2 To ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        If CStr(ws.Cells(r, 1).Value) = k Then
            FindKeyRow = r
            Exit Function
        End If
    Next
End Function

' キー k の値を文字列として書く（"0001" などが数値に変わらないよう書式を @ にする）
Sub SetKeyValue(ws, k, v)
    Dim r
    r = FindKeyRow(ws, k)
    If r = 0 Then Err.Raise vbObjectError + 1, "RunTests", "ひな形に設定キーがありません: " & k
    ws.Cells(r, 2).NumberFormat = "@"
    ws.Cells(r, 2).Value = v
End Sub

' キー k の行を空にする（設定キー不足のテスト用）
Sub RemoveKey(ws, k)
    Dim r
    r = FindKeyRow(ws, k)
    If r > 0 Then
        ws.Cells(r, 1).Value = Empty
        ws.Cells(r, 2).Value = Empty
    End If
End Sub

' 表のシートの2行目以降を消し、rows（行の配列）を2行目から文字列として書き込む
Sub FillTable(ws, rows, nCols)
    Dim r, c
    For r = 2 To MAP_CLEAR_ROWS
        For c = 1 To nCols
            ws.Cells(r, c).Value = Empty
        Next
    Next
    For r = 0 To UBound(rows)
        For c = 0 To UBound(rows(r))
            ws.Cells(r + 2, c + 1).NumberFormat = "@"
            ws.Cells(r + 2, c + 1).Value = rows(r)(c)
        Next
    Next
End Sub

' ツールを実行して終了コードを返す（Excel プロセスの残りも確認）
'   作業フォルダに config.xlsx を作り、cscript で DataMapping.vbs を実行して終了まで待つ
'   （ウィンドウ非表示）。実行後、ツールが起動した EXCEL.EXE が消えたことも確認項目にする
Function RunTool()
    Dim cfg, before
    cfg = gFso.BuildPath(gDir, "config.xlsx")
    WriteConfig cfg
    before = ExcelCount()
    RunTool = gSh.Run("cscript //nologo """ & gTool & """ """ & cfg & """", 0, True)
    If Not IsMock() Then
        If WaitExcelCount(before) Then
            Check "Excelのプロセスが残らない", True, ""
        Else
            Check "Excelのプロセスが残らない", False, "実行前 " & before & " 個 → 実行後 " & ExcelCount() & " 個"
        End If
    End If
End Function

' 動いている EXCEL.EXE の数（WMI で数える）
Function ExcelCount()
    If IsMock() Then
        ExcelCount = 0
    Else
        ExcelCount = GetObject("winmgmts:\\.\root\cimv2").ExecQuery("Select * From Win32_Process Where Name = 'EXCEL.EXE'").Count
    End If
End Function

' EXCEL.EXE の数が実行前（before）以下に戻るまで、最大 PROCESS_WAIT_SEC 秒待つ
'   Excel の終了は少し遅れることがあるので、0.5 秒ごとに数え直す
Function WaitExcelCount(before)
    Dim t
    t = Timer
    Do
        If ExcelCount() <= before Then
            WaitExcelCount = True
            Exit Function
        End If
        WScript.Sleep 500
    Loop While Timer - t < PROCESS_WAIT_SEC
    WaitExcelCount = False
End Function


'==============================================================
' 検証
'==============================================================
' 判定の部品は、どれも最後に Check（PASS / FAIL の記録）を呼ぶ。
' FAIL のときは「何が違ったか」を詳細に残すので、テスト結果.csv の「詳細」列を見れば原因の見当がつく
'   ExpectExit   … 終了コード
'   ExpectText   … 出力ファイルの全文（改行コード・末尾改行まで一致）。違えば最初に違う位置を表示
'   ExpectHex    … ファイル先頭のバイト列（文字コード・BOM の確認）
'   ExpectLogHex … ログファイル先頭のバイト列
'   ExpectNoFile … ファイルが作られていないこと
'   ExpectXlsx   … Excel 出力のシート名・各セルの値・文字列セルか・余分なセルがないか
'   ExpectLog    … ログに「区分 kind を含み、text を含む行」があるか
Sub ExpectExit(rc, expected)
    Check "終了コード " & expected, (rc = expected), "実際: " & rc
End Sub

Sub ExpectText(rel, cs, expected)
    Dim p, actual
    p = gFso.BuildPath(gDir, rel)
    If Not gFso.FileExists(p) Then
        Check "出力内容 " & rel, False, "ファイルがありません"
        Exit Sub
    End If
    actual = ReadTextFile(p, cs)
    If actual = expected Then
        Check "出力内容 " & rel, True, ""
    Else
        Check "出力内容 " & rel, False, DiffText(expected, actual)
    End If
End Sub

Sub ExpectHex(rel, hexPrefix)
    Dim p, h
    If IsMock() Then
        SkipCheck "バイト確認 " & rel & "（Windows実機のみ）"
        Exit Sub
    End If
    p = gFso.BuildPath(gDir, rel)
    h = ReadHex(p, Len(hexPrefix) / 2)
    Check "先頭バイト " & rel, (h = hexPrefix), "期待: " & hexPrefix & " / 実際: " & h
End Sub

Sub ExpectLogHex(hexPrefix)
    Dim p
    If IsMock() Then
        SkipCheck "ログのBOM（Windows実機のみ）"
        Exit Sub
    End If
    p = LogPath()
    If p = "" Then
        Check "ログのBOM", False, "ログがありません"
    Else
        Check "ログのBOM", (ReadHex(p, Len(hexPrefix) / 2) = hexPrefix), ""
    End If
End Sub

Sub ExpectNoFile(rel)
    Check "出力されない " & rel, Not gFso.FileExists(gFso.BuildPath(gDir, rel)), "ファイルが作られています"
End Sub

' rows : 期待値（行ごとの配列。RowsOf で "a,b" 形式から作れる）
' requireText = True なら、値のあるセルが文字列型かつ書式 @ であることも確認する
'   （先頭ゼロや日付が Excel に数値として解釈されていないことの確認）
Sub ExpectXlsx(rel, sheetName, rows, requireText)
    Dim p, wb, ws, r, c, v, s, e, ok, detail, lastR, lastC, nCols
    p = gFso.BuildPath(gDir, rel)
    If Not gFso.FileExists(p) Then
        Check "Excel出力 " & rel, False, "ファイルがありません"
        Exit Sub
    End If
    Set wb = gXL.Workbooks.Open(p, XL_NO_UPDATE_LINKS, True)
    Set ws = wb.Worksheets(1)
    Check "シート名 " & sheetName, (ws.Name = sheetName), "実際: " & ws.Name
    ok = True
    detail = ""
    nCols = 0
    For r = 0 To UBound(rows)
        If UBound(rows(r)) + 1 > nCols Then nCols = UBound(rows(r)) + 1
        For c = 0 To UBound(rows(r))
            v = ws.Cells(r + 1, c + 1).Value
            If IsEmpty(v) Then s = "" Else s = CStr(v)
            e = rows(r)(c)
            If s <> e Then
                ok = False
                If detail = "" Then detail = (r + 1) & "行" & (c + 1) & "列 期待[" & e & "] 実際[" & s & "]"
            ElseIf requireText And e <> "" Then
                If VarType(v) <> vbString Or ws.Cells(r + 1, c + 1).NumberFormat <> "@" Then
                    ok = False
                    If detail = "" Then detail = (r + 1) & "行" & (c + 1) & "列 が文字列セルではありません"
                End If
            End If
        Next
    Next
    ' 期待値の範囲の外に値が入っていないか（余分な行・列が出力されていないか）
    lastR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    lastC = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If lastR > UBound(rows) + 1 Or lastC > nCols Then
        For r = 1 To lastR
            For c = 1 To lastC
                If r > UBound(rows) + 1 Or c > nCols Then
                    If Not IsEmpty(ws.Cells(r, c).Value) Then
                        ok = False
                        If detail = "" Then detail = "余分なデータ: " & r & "行" & c & "列"
                    End If
                End If
            Next
        Next
    End If
    wb.Close False
    Check "Excel出力 " & rel, ok, detail
End Sub

' このテストの log フォルダにあるログファイルのパス（なければ ""）
'   ツールは1回の実行でログを1つ作り、テストごとに作業フォルダが新しいので、通常は1つだけ
Function LogPath()
    Dim f, best
    best = ""
    LogPath = ""
    If Not gFso.FolderExists(gFso.BuildPath(gDir, "log")) Then Exit Function
    For Each f In gFso.GetFolder(gFso.BuildPath(gDir, "log")).Files
        best = f.Path
    Next
    LogPath = best
End Function

Sub ExpectLog(kind, text)
    Dim p, lines, i, found
    p = LogPath()
    If p = "" Then
        Check "ログ " & kind & "「" & text & "」", False, "ログファイルがありません"
        Exit Sub
    End If
    lines = Split(ReadTextFile(p, CS_UTF8), vbCrLf)
    found = False
    For i = 0 To UBound(lines)
        If InStr(lines(i), "," & kind & ",") > 0 And InStr(lines(i), text) > 0 Then found = True
    Next
    Check "ログ " & kind & "「" & text & "」", found, "該当する行がありません"
End Sub

' 期待値と実際の文字列で最初に違う位置を探し、その位置から30文字ずつを並べて返す
Function DiffText(expected, actual)
    Dim i, n
    n = Len(expected)
    If Len(actual) < n Then n = Len(actual)
    For i = 1 To n
        If Mid(expected, i, 1) <> Mid(actual, i, 1) Then Exit For
    Next
    DiffText = i & "文字目から違います　期待[" & Visible(Mid(expected, i, 30)) & "] 実際[" & Visible(Mid(actual, i, 30)) & "]"
End Function

' 改行・タブを \r \n \t と表示する（違いが改行コードだけのときに見分けられるように）
Function Visible(s)
    Visible = Replace(Replace(Replace(s, vbCr, "\r"), vbLf, "\n"), vbTab, "\t")
End Function


'==============================================================
' 結果の記録
'==============================================================
' 確認項目1つの結果を記録する。FAIL は画面にも「NG ID 項目 : 詳細」と表示する
Sub Check(name, ok, detail)
    If ok Then
        gPass = gPass + 1
        AddResult gCase, gTitle, name, "PASS", ""
    Else
        gFail = gFail + 1
        AddResult gCase, gTitle, name, "FAIL", detail
        WScript.Echo "  NG " & gCase & " " & name & " : " & detail
    End If
End Sub

' 確認できない項目（モック環境など）を SKIP として記録する
Sub SkipCheck(name)
    gSkip = gSkip + 1
    AddResult gCase, gTitle, name, "SKIP", ""
End Sub

' テスト結果.csv の1行（ID, パターン, 確認項目, 結果, 詳細）を作ってためる。全項目を "" で囲む
Sub AddResult(a, b, c, d, e)
    Dim v, i
    v = Array(a, b, c, d, e)
    For i = 0 To UBound(v)
        v(i) = """" & Replace(CStr(v(i)), """", """""") & """"
    Next
    If gResN > UBound(gResults) Then ReDim Preserve gResults(UBound(gResults) * 2 + 1)
    gResults(gResN) = Join(v, ",")
    gResN = gResN + 1
End Sub

' ためた結果をテスト結果.csv（UTF-8 BOM 付き。Excel でそのまま開ける）に保存する
Sub WriteResults()
    Dim lines, i
    ReDim lines(gResN - 1)
    For i = 0 To gResN - 1
        lines(i) = gResults(i)
    Next
    WriteTextFile gFso.BuildPath(gBase, RESULT_FILE), Join(lines, vbCrLf) & vbCrLf, CS_UTF8BOM
End Sub
