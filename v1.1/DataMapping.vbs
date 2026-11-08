'==============================================================
' DataMapping.vbs  データマッピング変換ツール（VBScript版）  v1.1
'
'   fileA（複数ファイル）を取り込み、設定Excelのマッピング定義に従って
'   fileB形式（CSV または Excel）へ変換し、日付ごとに1ファイルにまとめて出力する。
'
' 処理の流れ（呼び出し順）
'   Main
'    ├ ResolveConfigPath ……… 使う設定ファイル（.xlsx）を決める
'    ├ Excel を裏で起動（画面非表示・確認ダイアログなし）
'    └ RunAll
'       ├ 設定の読み込み ……… LoadConfig → LoadSystem → LoadSettings
'       │                        → LoadMapping → LoadFilters → LoadCodes
'       ├ InitLog ……………… ログをメモリ上に用意（保存は最後にまとめて）
'       ├ 入力フォルダの対象ファイルを、ファイル名の日付ごとにグループ化
'       └ 日付ごとに ProcessDate（1日付がエラーでも次の日付へ進む）
'          ├ ReadTable …………… CSV / Excel を2次元配列に読む
'          ├ ResolveMapping …… fileB の各項目が fileA の何列目かを決める
'          ├ PassFilters ……… 抽出条件で行を残すか判定
'          ├ ApplyConv ………… 値の変換（日付・数値・ゼロ埋め・コード変換など）
'          └ BuildCsv + WriteText / WriteExcel …… 1日付分を1ファイルに保存
'    最後に FlushLog（ログCSVを保存）→ Excel を終了 → 終了コードを返す
'
' エラーの扱い
'   ・Fail（下の「共通部品」）でエラーを発生させ、呼び出し元の On Error で受け止める。
'   ・設定の誤りなど全体に関わるエラー … Main で受け止めて全体を停止（終了コード 1）。
'   ・データの誤りなど1日付だけのエラー … RunAll で受け止め、その日付は出力せずに次へ。
'     ファイルは全行の変換が終わってから書くので、中途半端な出力ファイルは残らない。
'
' 変更履歴
'   v1.1  ファイル名から取り出した日付が正しい日付か確かめるようにした
'         （例 番号と日付がつながった 1700020261106.csv から「17000202」を日付と誤って
'           fileB_17000202.csv を作っていた → 正しくない日付のファイルは SKIP にしてログに残す）
'   v1.0  処理内容が分かるようにコメントを追加（処理ロジックの変更なし）
'         セル文字数超過のエラー文言を「Excelのセルの最大文字数（32767）を超えています」に統一
'
' 役割分担
'   設定Excel（.xlsx・マクロなし）… 設定値・マッピング・抽出条件・コード表・内部パラメータ
'   このファイル（.vbs）           … 処理ロジックのみ
'
' 実行方法
'   1) このファイルと同じフォルダ・同じ名前の .xlsx を設定として使う
'        DataMapping.vbs ＋ DataMapping.xlsx → DataMapping.vbs をダブルクリック
'   2) 別の設定ファイルを使う
'        設定ファイル(.xlsx)を DataMapping.vbs にドラッグ＆ドロップ
'        または  cscript //nologo DataMapping.vbs "C:\path\設定.xlsx"
'
' 終了コード: 0 = エラーなし / 1 = エラーあり（タスクスケジューラ等で判定可）
' ログ: 設定の「ログフォルダ」にCSVで出力（cscript 実行時は画面にも表示）
'
' 想定外のデータは誤った出力をせず、その日付をエラーにしてログに場所を残す。
'==============================================================
Option Explicit

'---------------- 起点（ここだけはコード側で決める） ----------------
' 設定シートの名前と位置だけはコードに固定する。それ以外の記号・キーワード・
' 列位置などはすべて設定Excelの「システム設定」シートから読む。
Const SH_SET = "設定"
Const SH_SYS = "システム設定"
Const CFG_KEY_COL = 1          ' 設定シートのキー列（A）
Const CFG_VAL_COL = 2          ' 設定シートの値列（B）
Const CFG_START_ROW = 2        ' 設定シートの1行目は見出し
Const CONFIG_EXT = ".xlsx"     ' 引数なしのとき「スクリプト名＋この拡張子」を設定ファイルとする

'---------------- API・仕様上の定数 ----------------
' AD_* は ADODB.Stream（文字コードを指定したファイル読み書き）、XL_* は Excel の定数。
' RC_* は1日付分の処理結果（ProcessDate の戻り値）、EXIT_* はスクリプトの終了コード。
Const AD_TYPE_BINARY = 1
Const AD_TYPE_TEXT = 2
Const AD_READ_ALL = -1
Const AD_SAVE_OVERWRITE = 2
Const UTF8_BOM_BYTES = 3
Const XL_NO_UPDATE_LINKS = 0
Const XL_WBAT_WORKSHEET = -4167
Const ERR_OFFSET = 513
Const QUOTE = """"
Const PAD_CHAR = "0"
Const EOF_CHAR_CODE = 26       ' テキスト末尾の EOF 文字（Ctrl+Z）
Const FW_FIRST = 65281         ' 全角英数記号の範囲 U+FF01（！）～ U+FF5E（～）
Const FW_LAST = 65374
Const FW_OFFSET = 65248        ' 全角 → 半角 のコード差（U+FF01 - U+0021）
Const UNICODE_RANGE = 65536    ' AscW が負数を返す場合の補正
Const NO_COL = -1
Const SM_YEAR = 0              ' 8桁日付正規表現のサブマッチ位置
Const SM_MONTH = 1
Const SM_DAY = 2
Const FILE_DATE_MIN_YEAR = 2000 ' ファイル名の日付として認める年の範囲（番号を日付と誤らないため）
Const FILE_DATE_MAX_YEAR = 2099
Const RC_ERROR = -1
Const RC_SKIP = 0
Const RC_OK = 1
Const EXIT_OK = 0
Const EXIT_ERROR = 1
Const APP_TITLE = "DataMapping"

'---------------- 設定キー：設定シート ----------------
' 設定シートA列のキー名。値（B列）は LoadSettings で読み、g で始まる変数に入れる。
Const K_IN_DIR = "入力フォルダ"
Const K_PATTERN = "対象ファイル"
Const K_DATE_RE = "日付抽出パターン"
Const K_IN_CS = "入力文字コード"
Const K_HDR_ROW = "ヘッダー行番号"
Const K_FOOTER = "fileA末尾の除外行数"
Const K_DLM = "区切り文字"
Const K_OUT_DIR = "出力フォルダ"
Const K_OUT_NAME = "出力ファイル名"
Const K_OUT_FMT = "出力形式"
Const K_OUT_SHEET = "出力シート名"
Const K_OUT_CS = "出力文字コード"
Const K_OUT_HDR = "出力ヘッダー"
Const K_QUOTE_ALL = "全項目をダブルクォートで囲む"
Const K_QUOTE_EMPTY = "空の値もダブルクォートで囲む"
Const K_NL = "改行コード"
Const K_TRAIL_NL = "ファイル末尾に改行"
Const K_HEAD_REC = "先頭レコード"
Const K_TAIL_REC = "末尾レコード"
Const K_OUT_ZERO = "データ0行の日付も出力"
Const K_OVERWRITE = "上書き"
Const K_LOG_DIR = "ログフォルダ"

'---------------- 設定キー：システム設定シート ----------------
' 内部パラメータ（シート名・列位置・変換名・記号など）のキー名。LoadSystem で読む。
Const S_SH_MAP = "マッピングシート名"
Const S_SH_CODE = "コード変換シート名"
Const S_SH_FILTER = "抽出条件シート名"
Const S_MAP_ROW = "マッピング開始行"
Const S_MAP_COL_OUT = "マッピング列_fileB項目"
Const S_MAP_COL_IN = "マッピング列_fileA項目"
Const S_MAP_COL_FIX = "マッピング列_固定値"
Const S_MAP_COL_CONV = "マッピング列_変換"
Const S_MAP_COL_JOIN = "マッピング列_連結文字"
Const S_MULTI_SEP = "fileA項目の複数指定区切り"
Const S_CODE_ROW = "コード変換開始行"
Const S_CODE_COL_NAME = "コード変換列_変換名"
Const S_CODE_COL_FROM = "コード変換列_変換前"
Const S_CODE_COL_TO = "コード変換列_変換後"
Const S_CODE_DEFAULT = "コード変換既定値キー"
Const S_FIL_ROW = "抽出条件開始行"
Const S_FIL_COL_KIND = "抽出条件列_種別"
Const S_FIL_COL_NAME = "抽出条件列_fileA項目"
Const S_FIL_COL_OP = "抽出条件列_比較"
Const S_FIL_COL_VAL = "抽出条件列_値"
Const S_FIL_INCLUDE = "抽出種別_含める"
Const S_FIL_EXCLUDE = "抽出種別_除外する"
Const S_OP_EQ = "比較_等しい"
Const S_OP_NE = "比較_等しくない"
Const S_OP_HAS = "比較_含む"
Const S_OP_NOT_HAS = "比較_含まない"
Const S_OP_EMPTY = "比較_空"
Const S_OP_NOT_EMPTY = "比較_空でない"
Const S_OP_REGEX = "比較_正規表現"
Const S_COUNT_TOKEN = "件数置換文字"
Const S_SUM_OPEN = "合計置換文字_開始"
Const S_SUM_CLOSE = "合計置換文字_終了"
Const S_LOG_HEADER = "ログ見出し"
Const S_LOG_TIME_FMT = "ログ日時書式"
Const S_LOG_NAME = "ログファイル名"
Const S_LOG_TIME_TOKEN = "ログファイル名_日時置換文字"
Const S_LOG_FILE_TIME_FMT = "ログファイル名_日時書式"
Const S_LOG_CS = "ログ文字コード"
Const S_LOG_DLM = "ログ区切り文字"
Const S_LIST_SEP = "設定リスト区切り"
Const S_DATE_TOKEN = "日付置換文字"
Const S_TAB_WORD = "タブ指定文字"
Const S_LF_WORD = "改行LF指定値"
Const S_YES_WORDS = "あり判定値"
Const S_FMT_CSV = "出力形式_CSV指定値"
Const S_FMT_EXCEL = "出力形式_Excel指定値"
Const S_EXCLUDE_PREFIX = "除外ファイル名接頭辞"
Const S_XL_IN_EXT = "Excel入力拡張子"
Const S_XL_OUT_EXT = "Excel出力拡張子"
Const S_XL_OUT_FORMAT = "Excel出力ファイル形式"
Const S_XL_CELL_FMT = "Excel出力セル書式"
Const S_XL_SHEET_BAD = "Excelシート名の禁止文字"
Const S_XL_SHEET_MAX = "Excelシート名の最大文字数"
Const S_XL_CELL_MAX = "Excelセルの最大文字数"
Const S_UTF8_NOBOM = "UTF8_BOMなし指定値"
Const S_UTF8_BOM = "UTF8_BOMあり指定値"
Const S_UTF8_CHARSET = "UTF8文字コード名"
Const S_CONV_SEP = "変換区切り"
Const S_ARG_SEP = "変換引数区切り"
Const S_ARG_SEP_WIDE = "変換引数区切り_全角"
Const S_REPL_SEP = "置換区切り"
Const S_CV_TRIM = "変換名_TRIM"
Const S_CV_DATE = "変換名_日付"
Const S_CV_NUM = "変換名_数値"
Const S_CV_ZERO = "変換名_ゼロ埋め"
Const S_CV_REPL = "変換名_置換"
Const S_CV_EMPTY = "変換名_空なら"
Const S_CV_LEFT = "変換名_左"
Const S_CV_RIGHT = "変換名_右"
Const S_CV_CODE = "変換名_コード"
Const S_DATE8_RE = "8桁日付正規表現"
Const S_THOUSANDS = "数値の桁区切り"
Const S_BUF_ROWS = "行バッファ初期サイズ"
Const S_BUF_FIELDS = "項目バッファ初期サイズ"

'---------------- 実行時の値 ----------------
' gXL        … 裏で起動した Excel（設定・Excel入力の読み込み、Excel出力、TEXT関数に使う）
' gFso       … ファイル操作（FileSystemObject）
' gCfg       … 設定シート＋システム設定シートの「キー → 値」をまとめた辞書
' gIsCscript … cscript で実行中なら True（画面に逐次表示。wscript ならメッセージボックス）
' gCtx       … いま処理している場所（ファイル名 / 行 / 項目）。エラー時にログへ書く
' gNOk / gNSkip / gNNg … 成功・スキップ・エラーになった日付の数
Dim gXL, gFso, gCfg, gIsCscript, gTrimChars, gEofChar, gCtx
Dim gNOk, gNSkip, gNNg
' ログ
Dim gLogLines, gLogCount, gLogPath, gLogDir, gLogName, gLogCs, gLogDlm, gLogTimeFmt
' 業務設定
Dim gInDir, gPattern, gReDate, gInCs, gHdrRow, gFooterSkip, gDlm
Dim gOutDir, gOutName, gOutSheet, gOutCs, gIsExcelOut, gOutHdr, gQuoteAll, gQuoteEmpty
Dim gNL, gTrailNL, gHeadRec, gTailRec, gOutZero, gOverwrite
' システム設定
Dim gShMap, gShCode, gShFilter, gDateToken, gListSep, gExcludePrefix, gMultiSep
Dim gConvSep, gArgSep, gArgSepWide, gReplSep
Dim gCvTrim, gCvDate, gCvNum, gCvZero, gCvRepl, gCvEmpty, gCvLeft, gCvRight, gCvCode
Dim gCodeDefault, gReDate8, gThousands, gCountToken, gSumOpen, gSumClose
Dim gFilInclude, gFilExclude, gOpEq, gOpNe, gOpHas, gOpNotHas, gOpEmpty, gOpNotEmpty, gOpRegex
Dim gXlInExt, gXlOutExt, gXlOutFormat, gXlCellFmt, gXlSheetBad, gXlSheetMax, gXlCellMax
Dim gUtf8NoBom, gUtf8Bom, gUtf8Charset
Dim gBufRows, gBufFields
' マッピング（mParts(j) は fileA項目名の配列。空配列なら固定値）
'   j 番目の fileB 項目について：mOut=項目名 / mParts=元の fileA 項目名 / mFix=固定値
'                                mConv=変換指定 / mJoin=複数項目をつなぐ文字。mCnt=項目数
Dim mCnt, mOut, mParts, mFix, mConv, mJoin
' 抽出条件
'   k 番目の条件：fInclude=True なら「含める」・False なら「除外する」/ fName=fileA項目名
'                 fOp=比較方法 / fVal=比較する値 / fRe=正規表現（比較が正規表現のときだけ）。fCnt=条件数
Dim fCnt, fInclude, fName, fOp, fVal, fRe
' コード表
'   gCodes(変換名) が「変換前 → 変換後」の辞書（例 gCodes("区分")("1") = "01"）
Dim gCodes

' スクリプトの入口。Main の戻り値（0 / 1）を終了コードにする
WScript.Quit Main()


'==============================================================
' メイン
'   Excel の起動から終了までを受け持つ。途中で何が起きても、
'   ログの保存と Excel の終了（プロセスを残さない）は必ず行う。
'   戻り値: EXIT_OK（0）= エラーの日付なし / EXIT_ERROR（1）= 停止またはエラーの日付あり
'==============================================================
Function Main()
    Dim cfgPath, errMsg, summary

    ' 最後まで成功したときだけ EXIT_OK に変える
    Main = EXIT_ERROR
    gIsCscript = (InStr(LCase(WScript.FullName), "cscript") > 0)
    gTrimChars = " 　" & vbTab
    gEofChar = Chr(EOF_CHAR_CODE)
    gNOk = 0: gNSkip = 0: gNNg = 0
    gLogPath = ""
    errMsg = ""
    Set gFso = NewObject("Scripting.FileSystemObject")

    ' ここから下はエラーで止めずに Err を見て判断する（後片付けを必ず実行するため）
    On Error Resume Next
    cfgPath = ResolveConfigPath()
    If Err.Number = 0 Then
        Set gXL = NewObject("Excel.Application")
        If Err.Number = 0 Then
            gXL.Visible = False
            gXL.DisplayAlerts = False
            gXL.ScreenUpdating = False
            summary = RunAll(cfgPath)
        End If
    End If

    ' 全体を止めるエラー（設定ファイルがない・設定の誤り・Excel が起動できない など）
    ' ログの準備（InitLog）より前のエラーはログに書けないので、画面表示だけになる
    If Err.Number <> 0 Then
        errMsg = Err.Description
        Err.Clear
        If gLogPath <> "" Then WriteLog "ERROR", "", gCtx, errMsg
    End If
    ' ログはメモリにためてあるので、ここでまとめてファイルに保存する
    If gLogPath <> "" Then
        FlushLog
        If Err.Number <> 0 Then
            errMsg = errMsg & vbLf & "ログの保存に失敗しました: " & Err.Description
            Err.Clear
        End If
    End If
    ' 裏で起動した Excel を終了する（EXCEL.EXE が残らないように）
    If IsObject(gXL) Then gXL.Quit
    Set gXL = Nothing
    On Error GoTo 0

    ' 結果の表示。終了コード 0 になるのは「停止せず、エラーの日付も0件」のときだけ
    If errMsg <> "" Then
        ShowMsg "エラーで停止しました。" & vbLf & errMsg
    Else
        ShowMsg summary
        If gNNg = 0 Then Main = EXIT_OK
    End If
End Function

' COM オブジェクトの生成はすべてここを通す
' （テストや検証環境で、生成するオブジェクトを差し替えられるようにするため）
Function NewObject(progId)
    Set NewObject = CreateObject(progId)
End Function

' 使う設定ファイルのフルパスを返す
'   引数あり（ドラッグ＆ドロップ・コマンド指定）… 1つ目の引数のファイル
'   引数なし（ダブルクリック）……………………… このスクリプトと同じフォルダの「同じ名前.xlsx」
'   相対パスは「カレントフォルダ」基準で解決される点に注意（タスクスケジューラでは
'   開始フォルダを指定しないと C:\Windows\System32 になる）
Function ResolveConfigPath()
    Dim p
    If WScript.Arguments.Count > 0 Then
        p = WScript.Arguments(0)
    Else
        p = gFso.BuildPath(gFso.GetParentFolderName(WScript.ScriptFullName), _
                           gFso.GetBaseName(WScript.ScriptFullName) & CONFIG_EXT)
    End If
    p = gFso.GetAbsolutePathName(p)
    If Not gFso.FileExists(p) Then Fail "設定ファイルが見つかりません: " & p
    ResolveConfigPath = p
End Function

' 設定を読み込み、対象ファイルを日付ごとに変換・出力する
'   戻り値: 完了メッセージ（成功・スキップ・エラーの日付数とログの場所）
'   設定の誤りはここから Fail で呼び出し元（Main）へ伝わり、全体が止まる
Function RunAll(cfgPath)
    Dim wb, groups, d, f, dt, keys, files, i, rc, reFile

    '--- 設定の読み込み（設定ファイルは読み取り専用で開く）
    '    システム設定（記号・列位置など）を先に読むのは、業務設定やマッピングの
    '    解釈にシステム設定の値を使うため
    gCtx = cfgPath
    Set wb = gXL.Workbooks.Open(cfgPath, XL_NO_UPDATE_LINKS, True)
    LoadConfig wb
    LoadSystem
    LoadSettings
    LoadMapping wb
    LoadFilters wb
    LoadCodes wb
    wb.Close False
    InitLog
    WriteLog "INFO", "", cfgPath, "設定ファイルを読み込みました"
    WarnEmptyColumns

    '--- 対象ファイルを日付ごとにグループ化
    '    groups(日付 "20261106") = そのファイル名の辞書
    '    「対象ファイル」のワイルドカードに合い、Excel の一時ファイル（~$）でないものが対象。
    '    ファイル名から日付を取り出せないファイル、取り出した日付が正しくないファイルは
    '    SKIP としてログに残す（例 17000_20261106.csv と 17002_20261106.csv は同じ日付として1つにまとまる）
    Set reFile = NewRegExp(WildToRegex(gPattern))
    Set groups = NewObject("Scripting.Dictionary")
    For Each f In gFso.GetFolder(gInDir).Files
        If reFile.Test(f.Name) And Not IsExcluded(f.Name) Then
            dt = ExtractDate(f.Name)
            If dt = "" Then
                WriteLog "SKIP", "", f.Name, "ファイル名から日付を取り出せないため対象外"
            ElseIf Not IsValidFileDate(dt) Then
                WriteLog "SKIP", "", f.Name, "ファイル名の日付が正しくないため対象外: " & dt
            Else
                If Not groups.Exists(dt) Then groups.Add dt, NewObject("Scripting.Dictionary")
                Set d = groups(dt)
                d.Add f.Name, True
            End If
        End If
    Next
    If groups.Count = 0 Then Fail "対象ファイルが見つかりません: " & gFso.BuildPath(gInDir, gPattern)

    '--- 日付ごとに変換・出力（1日付のエラーで全体は止めない）
    '    日付は古い順、同じ日付のファイルはファイル名順に処理する（作成日時の順ではない）。
    '    ProcessDate の中で起きたエラーはここで受け止め、ERROR としてログに残して次の日付へ
    keys = SortArr(groups.Keys)
    For i = 0 To UBound(keys)
        If gIsCscript Then WScript.Echo "変換中: " & keys(i) & "（" & (i + 1) & " / " & groups.Count & "）"
        Set d = groups(keys(i))
        files = SortArr(d.Keys)
        rc = RC_ERROR
        On Error Resume Next
        rc = ProcessDate(keys(i), files)
        If Err.Number <> 0 Then
            WriteLog "ERROR", keys(i), gCtx, Err.Description & "（この日付は出力していません）"
            Err.Clear
            rc = RC_ERROR
        End If
        On Error GoTo 0
        Select Case rc
            Case RC_OK: gNOk = gNOk + 1
            Case RC_SKIP: gNSkip = gNSkip + 1
            Case Else: gNNg = gNNg + 1
        End Select
    Next

    RunAll = "完了しました。" & vbLf & vbLf & _
             "成功: " & gNOk & " 日付" & vbLf & _
             "スキップ: " & gNSkip & " 日付" & vbLf & _
             "エラー: " & gNNg & " 日付" & vbLf & vbLf & _
             "ログ: " & gLogPath
End Function

' 常に空で出力される列を警告（対応付けの設定漏れに気づけるように）
'   fileA項目・固定値・変換のどれも指定されていないマッピング行が対象
Sub WarnEmptyColumns()
    Dim j, names
    names = ""
    For j = 0 To mCnt - 1
        If UBound(mParts(j)) < 0 And mFix(j) = "" And mConv(j) = "" Then names = names & " [" & mOut(j) & "]"
    Next
    If names <> "" Then WriteLog "WARN", "", gShMap, "常に空で出力される列があります:" & names
End Sub


'==============================================================
' 1日付分の処理（エラーは呼び出し元で記録。中途半端なファイルは作らない）
'   dt    : 日付（例 "20261106"）
'   files : その日付のファイル名の配列（ファイル名順）
'   戻り値: RC_OK = 出力した / RC_SKIP = 出力しなかった（既存ファイル・0行）
'   出力する全行をいったんメモリ（rowBuf）にため、最後に1回だけファイルに書く
'==============================================================
Function ProcessDate(dt, files)
    Dim outPath, rowBuf, nRow, dataStart, data, idx, fidx, rows, nRows, vals
    Dim i, j, k, r, fRows, fExcl, total

    '--- 出力ファイル名を決め、既にあれば（上書き＝なしのとき）何もせずスキップ
    outPath = gFso.BuildPath(gOutDir, Replace(gOutName, gDateToken, dt))
    gCtx = outPath
    If gFso.FileExists(outPath) And Not gOverwrite Then
        WriteLog "SKIP", dt, outPath, "出力ファイルが既にあるためスキップ（上書きする場合は設定の「" & K_OVERWRITE & "」を変更）"
        ProcessDate = RC_SKIP
        Exit Function
    End If

    '--- 出力行のバッファ。rowBuf(0..nRow-1) に「先頭レコード → 見出し → データ行 → 末尾レコード」
    '    の順でためる。dataStart はデータ行の始まりの位置（合計の計算に使う）
    ReDim rowBuf(gBufRows - 1)
    nRow = 0
    total = 0
    ReDim vals(mCnt - 1)

    If gHeadRec <> "" Then AddItem rowBuf, nRow, Empty          ' 先頭レコードの場所（後で置き換え）
    If gOutHdr Then
        For j = 0 To mCnt - 1
            vals(j) = mOut(j)
        Next
        AddItem rowBuf, nRow, vals
    End If
    dataStart = nRow

    '--- その日付のファイルを1つずつ読み、1行ずつ変換してバッファに追加する
    '    列の位置はファイルごとに見出しから探すので、ファイルによって列の並びが違っても良い
    For i = 0 To UBound(files)
        gCtx = files(i)
        data = ReadTable(gFso.BuildPath(gInDir, files(i)))
        idx = ResolveMapping(data)
        fidx = ResolveFilters(data)
        rows = DataRows(data, nRows)
        fRows = 0
        fExcl = 0
        ' gCtx には「ファイル / 行 / 項目」を入れておき、変換エラー時にどこで起きたかをログに出す
        For k = 0 To nRows - 1
            r = rows(k)
            gCtx = files(i) & " / " & (r - LBound(data, 1) + 1) & "行目"
            If PassFilters(data, r, fidx) Then
                For j = 0 To mCnt - 1
                    gCtx = files(i) & " / " & (r - LBound(data, 1) + 1) & "行目 / " & mOut(j)
                    vals(j) = ApplyConv(SourceValue(data, r, idx(j), j), mConv(j))
                Next
                AddItem rowBuf, nRow, vals
                fRows = fRows + 1
            Else
                fExcl = fExcl + 1
            End If
        Next
        If fExcl > 0 Then
            WriteLog "INFO", dt, files(i), fRows & " 行 取込（抽出条件で " & fExcl & " 行除外）"
        Else
            WriteLog "INFO", dt, files(i), fRows & " 行 取込"
        End If
        total = total + fRows
    Next

    '--- データが0行の日付：設定により「見出しだけ出力（WARN）」か「出力しない（SKIP）」
    gCtx = outPath
    If total = 0 Then
        If Not gOutZero Then
            WriteLog "SKIP", dt, outPath, "データが0行のため出力しません（設定「" & K_OUT_ZERO & "」）"
            ProcessDate = RC_SKIP
            Exit Function
        End If
        WriteLog "WARN", dt, outPath, "データが0行です（見出しのみ出力）"
    End If

    '--- 先頭・末尾レコードは件数・合計が確定してから作る（{count} {sum:項目} を置き換え）
    If gHeadRec <> "" Then rowBuf(0) = ExpandRecord(gHeadRec, dt, total, rowBuf, dataStart, nRow)
    If gTailRec <> "" Then AddItem rowBuf, nRow, ExpandRecord(gTailRec, dt, total, rowBuf, dataStart, nRow)

    '--- ここで初めてファイルに書く（ここまでにエラーがあればファイルは作られない）
    If gIsExcelOut Then
        WriteExcel outPath, rowBuf, nRow
    Else
        WriteText outPath, BuildCsv(rowBuf, nRow), gOutCs
    End If
    WriteLog "OK", dt, outPath, (UBound(files) + 1) & " ファイル → " & total & " 行 出力"
    ProcessDate = RC_OK
End Function

' 見出し行の次から、空行を除いたデータ行の位置一覧（末尾の除外行数を反映）
'   data : ReadTable で読んだ2次元配列
'   n    : （戻り）データ行の数。list(0..n-1) が data の行位置
'   「fileA末尾の除外行数」は空行を数えずに、最後のデータ行から数える（合計行の除外用）
Function DataRows(data, n)
    Dim list, r, startRow
    ReDim list(UBound(data, 1) - LBound(data, 1))
    n = 0
    If gHdrRow = 0 Then startRow = LBound(data, 1) Else startRow = LBound(data, 1) + gHdrRow
    For r = startRow To UBound(data, 1)
        If Not IsBlankRow(data, r) Then
            list(n) = r
            n = n + 1
        End If
    Next
    n = n - gFooterSkip
    If n < 0 Then n = 0
    DataRows = list
End Function

' マッピング1列分の元の値（複数項目なら空でない値を連結文字でつなぐ）
'   cols : この fileB 項目の元になる fileA の列位置の配列
'          空配列 → 固定値 / 1つ → そのセルの値 / 複数 → 連結（例 姓|名 → 「山田　太郎」）
Function SourceValue(data, r, cols, j)
    Dim k, v, s, parts, n
    If UBound(cols) < 0 Then
        SourceValue = mFix(j)
        Exit Function
    End If
    If UBound(cols) = 0 Then
        SourceValue = data(r, cols(0))      ' 単一項目は型（日付など）を保ったまま渡す
        Exit Function
    End If
    ReDim parts(UBound(cols))
    n = 0
    For k = 0 To UBound(cols)
        v = data(r, cols(k))
        If VarType(v) = vbError Then Fail "セルがエラー値です"
        s = CellText(v)
        If s <> "" Then
            parts(n) = s
            n = n + 1
        End If
    Next
    If n = 0 Then
        SourceValue = ""
    Else
        ReDim Preserve parts(n - 1)
        SourceValue = Join(parts, mJoin(j))
    End If
End Function

' fileA の見出しから項目名の列位置を探す（重複・不足はエラー）
'   見出しは前後の空白（半角・全角・タブ）を除いて比べる。全角と半角は別の文字として扱う
'   ヘッダー行番号 = 0（ヘッダーなし）のときは、name を列番号（1始まり）として扱う
Function FindCol(data, name)
    Dim c, hdrIdx, n, found
    If gHdrRow = 0 Then
        n = CLng(name)
        c = LBound(data, 2) + n - 1
        If n < 1 Or c > UBound(data, 2) Then Fail "fileAに列がありません: [列" & name & "]"
        FindCol = c
        Exit Function
    End If
    hdrIdx = LBound(data, 1) + gHdrRow - 1
    If hdrIdx > UBound(data, 1) Then Fail "ヘッダー行（" & gHdrRow & "行目）がありません"
    found = NO_COL
    For c = LBound(data, 2) To UBound(data, 2)
        If VarType(data(hdrIdx, c)) <> vbError Then
            If JTrim(CellText(data(hdrIdx, c))) = name Then
                If found <> NO_COL Then Fail "fileAに同じ項目名が複数あります: [" & name & "]"
                found = c
            End If
        End If
    Next
    If found = NO_COL Then Fail "fileAに項目が見つかりません: [" & name & "]"
    FindCol = found
End Function

' マッピングの各列 → 列位置の配列（固定値の列は空配列）
'   戻り値 idx(j) が fileB の j 番目の項目の元になる列位置。足りない項目はまとめて1回でエラーにする
Function ResolveMapping(data)
    Dim idx, j, k, cols, miss
    ReDim idx(mCnt - 1)
    miss = ""
    For j = 0 To mCnt - 1
        If UBound(mParts(j)) < 0 Then
            idx(j) = Array()
        Else
            ReDim cols(UBound(mParts(j)))
            For k = 0 To UBound(mParts(j))
                cols(k) = TryFindCol(data, mParts(j)(k), miss)
            Next
            idx(j) = cols
        End If
    Next
    If miss <> "" Then Fail miss
    ResolveMapping = idx
End Function

' 見つからない場合はエラー内容を miss にためて続ける（不足項目をまとめて報告するため）
Function TryFindCol(data, name, miss)
    On Error Resume Next
    TryFindCol = FindCol(data, name)
    If Err.Number <> 0 Then
        If InStr(miss, Err.Description) = 0 Then
            If miss <> "" Then miss = miss & " / "
            miss = miss & Err.Description
        End If
        Err.Clear
        TryFindCol = NO_COL
    End If
    On Error GoTo 0
End Function

' 抽出条件の各条件 → fileA の列位置の配列（条件の項目がなければエラー）
Function ResolveFilters(data)
    Dim fidx, k
    If fCnt = 0 Then
        ResolveFilters = Array()
        Exit Function
    End If
    ReDim fidx(fCnt - 1)
    For k = 0 To fCnt - 1
        fidx(k) = FindCol(data, fName(k))
    Next
    ResolveFilters = fidx
End Function

' 抽出条件：「含める」はすべて満たすこと（AND）、「除外する」はどれか1つでも満たせば除外（OR）
Function PassFilters(data, r, fidx)
    Dim k, hit
    PassFilters = True
    For k = 0 To fCnt - 1
        hit = MatchFilter(data(r, fidx(k)), k)
        If fInclude(k) And Not hit Then PassFilters = False: Exit Function
        If Not fInclude(k) And hit Then PassFilters = False: Exit Function
    Next
End Function

' セルの値 v が k 番目の条件に当てはまるか（値は前後の空白を除いた変換前の値で比べる）
Function MatchFilter(v, k)
    Dim s
    If VarType(v) = vbError Then Fail "セルがエラー値です（抽出条件 [" & fName(k) & "]）"
    s = JTrim(CellText(v))
    Select Case fOp(k)
        Case gOpEq: MatchFilter = (s = fVal(k))
        Case gOpNe: MatchFilter = (s <> fVal(k))
        Case gOpHas: MatchFilter = (InStr(s, fVal(k)) > 0)
        Case gOpNotHas: MatchFilter = (InStr(s, fVal(k)) = 0)
        Case gOpEmpty: MatchFilter = (s = "")
        Case gOpNotEmpty: MatchFilter = (s <> "")
        Case gOpRegex: MatchFilter = fRe(k).Test(s)
        Case Else: Fail "不明な比較: " & fOp(k)
    End Select
End Function

' 先頭・末尾レコードの置換文字を展開
'   例 "T,{count},{sum:数量}" → "T,3,16"
'   {date} = 日付 / {count} = データ行数 / {sum:fileB項目名} = その列の合計（変換後の値で計算）
Function ExpandRecord(tpl, dt, total, rowBuf, dataStart, nRow)
    Dim s, p, q, name
    s = Replace(tpl, gDateToken, dt)
    s = Replace(s, gCountToken, CStr(total))
    Do
        p = InStr(s, gSumOpen)
        If p = 0 Then Exit Do
        q = InStr(p + Len(gSumOpen), s, gSumClose)
        If q = 0 Then Fail "合計の指定が閉じられていません: " & tpl
        name = Mid(s, p + Len(gSumOpen), q - p - Len(gSumOpen))
        s = Left(s, p - 1) & CStr(SumColumn(name, rowBuf, dataStart, nRow)) & Mid(s, q + Len(gSumClose))
    Loop
    ExpandRecord = s
End Function

' fileB 項目 name の列を、データ行（dataStart ～ nRow-1）だけ合計する。空の値は数えない
Function SumColumn(name, rowBuf, dataStart, nRow)
    Dim j, c, i, total, s
    c = NO_COL
    For j = 0 To mCnt - 1
        If mOut(j) = name Then c = j
    Next
    If c = NO_COL Then Fail "合計の対象がfileBの項目にありません: [" & name & "]"
    total = 0
    For i = dataStart To nRow - 1
        s = rowBuf(i)(c)
        If s <> "" Then total = total + ToNum(s)
    Next
    SumColumn = total
End Function


'==============================================================
' 値の変換（変換名・記号はシステム設定シートで定義）
'   日付・数値の書式は Excel の TEXT 関数と同じ書式（yyyymmdd、0.00 など）
'
'   v    : 元の値（セルの値。Excel 入力なら日付型・数値型のこともある）
'   conv : 変換指定。「|」区切りで左から順に適用する（例 "TRIM|ゼロ埋め:5"）
'          各ステップは「変換名:引数」。引数の区切りは全角の「：」でも良い
'   変換の種類（名前はシステム設定で変更可）
'     TRIM      … 前後の半角・全角スペース、タブを除く
'     日付:書式 … 日付に変換して書式化（例 日付:yyyymmdd）
'     数値:書式 … 数値に変換して書式化（例 数値:#,##0）。全角数字・桁区切りも可
'     ゼロ埋め:n… n 桁になるまで先頭に 0 を付ける（空の値はそのまま）
'     置換:旧=>新… 文字列の置き換え
'     空なら:値 … 空のとき指定の値にする
'     左:n / 右:n … 先頭・末尾から n 文字
'     コード:表名 … コード変換シートの表で読み替え（表にない値は既定値 * かエラー）
'   変換できない値は Fail でエラーにする（その日付は出力されない）
'==============================================================
Function ApplyConv(v, conv)
    Dim s, steps, k, stp, kind, arg, p, d, n

    If VarType(v) = vbError Then Fail "セルがエラー値です"
    s = CellText(v)
    If conv = "" Then
        ApplyConv = s
        Exit Function
    End If

    ' 変換指定をステップに分け、「変換名」と「引数」に分解して順に適用する
    steps = Split(conv, gConvSep)
    For k = 0 To UBound(steps)
        stp = Trim(steps(k))
        p = InStr(stp, gArgSep)
        If p = 0 And gArgSepWide <> "" Then p = InStr(stp, gArgSepWide)
        If p > 0 Then
            kind = UCase(Left(stp, p - 1))
            arg = Mid(stp, p + 1)
        Else
            kind = UCase(stp)
            arg = ""
        End If

        Select Case kind
            Case ""
                ' 何もしない
            Case gCvTrim
                s = JTrim(s)
            Case gCvDate
                ' 最初のステップで元が日付型（Excel の日付セル）なら、文字列にせず日付のまま書式化する
                ' （時刻を含む値も正しく扱うため）。それ以外は文字列から日付を読み取る
                If arg = "" Then Fail "日付の書式が指定されていません: " & stp
                If Len(s) > 0 Then
                    If VarType(v) = vbDate And k = 0 Then
                        s = gXL.WorksheetFunction.Text(v, arg)   ' Excelの日付セルはそのまま使う
                    Else
                        s = gXL.WorksheetFunction.Text(ToDate(s), arg)
                    End If
                End If
            Case gCvNum
                If arg = "" Then Fail "数値の書式が指定されていません: " & stp
                If Len(s) > 0 Then s = gXL.WorksheetFunction.Text(ToNum(s), arg)
            Case gCvZero
                n = ArgLng(arg, stp)
                If Len(s) > 0 And Len(s) < n Then s = String(n - Len(s), PAD_CHAR) & s
            Case gCvRepl
                p = InStr(arg, gReplSep)
                If p = 0 Then Fail "置換は「旧" & gReplSep & "新」の形式で指定してください: " & stp
                s = Replace(s, Left(arg, p - 1), Mid(arg, p + Len(gReplSep)))
            Case gCvEmpty
                If Len(s) = 0 Then s = arg
            Case gCvLeft
                s = Left(s, ArgLng(arg, stp))
            Case gCvRight
                s = Right(s, ArgLng(arg, stp))
            Case gCvCode
                ' 表にある値 → 変換後の値 / ない値 → 既定値キー（*）の行があればその値、なければエラー
                If Not gCodes.Exists(arg) Then Fail "コード変換表「" & arg & "」がありません"
                Set d = gCodes(arg)
                If d.Exists(s) Then
                    s = d(s)
                ElseIf d.Exists(gCodeDefault) Then
                    s = d(gCodeDefault)
                Else
                    Fail "コード変換表「" & arg & "」に値「" & s & "」がありません"
                End If
            Case Else
                Fail "不明な変換指定: " & stp
        End Select
    Next
    ApplyConv = s
End Function

' セル値 → 文字列（空・Null は空文字）
Function CellText(v)
    If IsEmpty(v) Then
        CellText = ""
    ElseIf IsNull(v) Then
        CellText = ""
    Else
        CellText = CStr(v)
    End If
End Function

' 変換の引数（桁数・文字数）を数値として取り出す
Function ArgLng(arg, stp)
    If Not IsNumeric(arg) Then Fail "変換の引数は数値で指定してください: " & stp
    ArgLng = CLng(arg)
End Function

' 文字列を日付に変換する。受け付ける形
'   8桁（20261106）… 年月日に分けて組み立て、存在しない日付（20261332 など）はエラー
'                     （DateSerial は 13月を翌年1月に繰り上げてしまうため、組み立て後に一致を確認）
'   2026/11/06・2026-11-06 など … VBScript の日付として読めればそのまま
'   数値 … Excel のシリアル値として扱う
'   全角数字（２０２６１１０６）も半角にしてから判定する
Function ToDate(ByVal s)
    Dim m, y, mo, dd, d
    s = ToHalfWidth(JTrim(s))
    Set m = gReDate8.Execute(s)
    If m.Count > 0 Then
        y = CInt(m(0).SubMatches(SM_YEAR))
        mo = CInt(m(0).SubMatches(SM_MONTH))
        dd = CInt(m(0).SubMatches(SM_DAY))
        d = DateSerial(y, mo, dd)
        If Year(d) <> y Or Month(d) <> mo Or Day(d) <> dd Then Fail "日付として不正です: " & s
        ToDate = d
    ElseIf IsDate(s) Then
        ToDate = CDate(s)
    ElseIf IsNumeric(s) Then
        ToDate = CDate(CDbl(s))   ' Excelのシリアル値
    Else
        Fail "日付に変換できません: " & s
    End If
End Function

' 文字列を数値に変換する（全角数字・全角カンマ・桁区切りのカンマを許す。例 「１，２００」→ 1200）
Function ToNum(ByVal s)
    s = Replace(ToHalfWidth(JTrim(s)), gThousands, "")
    If Not IsNumeric(s) Then Fail "数値に変換できません: " & s
    ToNum = CDbl(s)
End Function

' 全角英数記号（！～～）を半角に
Function ToHalfWidth(s)
    Dim i, c, w, r
    r = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        w = AscW(c)
        If w < 0 Then w = w + UNICODE_RANGE
        If w >= FW_FIRST And w <= FW_LAST Then c = ChrW(w - FW_OFFSET)
        r = r & c
    Next
    ToHalfWidth = r
End Function


'==============================================================
' 読み込み
'   どの形式も「2次元配列 data(行, 列)」にそろえて返す。以降の処理は形式を意識しない
'   （Excel から読んだ配列は 1 始まり、CSV から作った配列は 0 始まり。
'     そのため位置の計算には必ず LBound を使っている）
'==============================================================
' 拡張子が Excel（xlsx, xlsm, xls, xlsb）なら Excel で、それ以外は CSV として読む
Function ReadTable(path)
    If IsExcelExt(path) Then
        ReadTable = ReadExcel(path)
    Else
        ReadTable = ParseCsv(ReadText(path, gInCs), gDlm)
    End If
End Function

' Excelファイル：1枚目のシートをA1から読む（失敗しても必ず閉じる）
'   読み取り専用・リンク更新なしで開く。セルの値は表示形式ではなく実際の値（数値・日付・計算結果）
Function ReadExcel(path)
    Dim wb, v, en, ed
    Set wb = gXL.Workbooks.Open(path, XL_NO_UPDATE_LINKS, True)
    On Error Resume Next
    v = SheetArray(wb.Worksheets(1))
    en = Err.Number: ed = Err.Description
    wb.Close False
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, APP_TITLE, ed
    ReadExcel = v
End Function

' シートの A1 から使用範囲の最終セルまでを2次元配列で返す
'   表が A1 から始まっていなくても、A1 起点で読むので行・列番号がシート上の番号と一致する
'   セルが1つだけのときは Value が配列にならないので、1×1 の配列に包んで返す
Function SheetArray(ws)
    Dim ur, lastR, lastC, v, one
    Set ur = ws.UsedRange
    lastR = ur.Row + ur.Rows.Count - 1
    lastC = ur.Column + ur.Columns.Count - 1
    v = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value
    If IsArray(v) Then
        SheetArray = v
    Else
        ReDim one(0, 0)
        one(0, 0) = v
        SheetArray = one
    End If
End Function

' テキストファイルを文字コード cs で読む（ADODB.Stream。UTF-8 の BOM は自動で除かれる）
Function ReadText(path, cs)
    Dim st
    Set st = NewObject("ADODB.Stream")
    st.Type = AD_TYPE_TEXT
    st.Charset = StreamCharset(cs)
    st.Open
    st.LoadFromFile path
    ReadText = st.ReadText(AD_READ_ALL)
    st.Close
End Function

' CSV → 2次元配列（0始まり）。ダブルクォート・項目内改行・末尾EOF文字に対応
Function ParseCsv(ByVal txt, dlm)
    Dim rows, nRows, fields, fCount, lines, k, i, n, c, fld, inQ

    ReDim rows(gBufRows - 1)
    nRows = 0
    ReDim fields(gBufFields - 1)
    fCount = 0
    If Len(txt) > 0 Then
        If Right(txt, 1) = gEofChar Then txt = Left(txt, Len(txt) - 1)
    End If

    ' 2通りの読み方を使い分ける
    '   ダブルクォートが1つもない … 改行と区切り文字で Split するだけ（大量データでも速い）
    '   ダブルクォートがある ……… 1文字ずつ読み、クォート内の区切り文字・改行・"" を正しく扱う
    If InStr(txt, QUOTE) = 0 Then
        ' クォートなし：高速処理
        txt = Replace(Replace(txt, vbCrLf, vbLf), vbCr, vbLf)
        lines = Split(txt, vbLf)
        For k = 0 To UBound(lines)
            If k < UBound(lines) Or Len(lines(k)) > 0 Then
                If Len(lines(k)) = 0 Then
                    PushArr rows, nRows, Array()
                Else
                    PushArr rows, nRows, Split(lines(k), dlm)
                End If
            End If
        Next
    Else
        ' クォートあり：1文字ずつ読む
        '   fld = 読みかけの項目 / fields = 読みかけの行の項目 / inQ = クォートの内側にいるか
        n = Len(txt)
        i = 1
        fld = ""
        inQ = False
        Do While i <= n
            c = Mid(txt, i, 1)
            If inQ Then
                If c = QUOTE Then
                    If Mid(txt, i + 1, 1) = QUOTE Then
                        fld = fld & QUOTE
                        i = i + 1
                    Else
                        inQ = False
                    End If
                Else
                    fld = fld & c
                End If
            ElseIf c = QUOTE Then
                inQ = True
            ElseIf c = dlm Then
                AddField fields, fCount, fld
                fld = ""
            ElseIf c = vbCr Or c = vbLf Then
                If c = vbCr And Mid(txt, i + 1, 1) = vbLf Then i = i + 1
                AddField fields, fCount, fld
                fld = ""
                PushRow rows, nRows, fields, fCount
            Else
                fld = fld & c
            End If
            i = i + 1
        Loop
        ' 最後までクォートが閉じない＝壊れたCSV。誤った出力を避けるためエラーにする
        If inQ Then Fail "ダブルクォートが閉じられていません"
        If fCount > 0 Or Len(fld) > 0 Then
            AddField fields, fCount, fld
            PushRow rows, nRows, fields, fCount
        End If
    End If

    ParseCsv = ToMatrix(rows, nRows)
End Function

' 以下3つは「足りなくなったら2倍に広げる」可変長配列の部品（ReDim Preserve の回数を減らすため）
' AddField: 読みかけの行に項目を1つ追加
Sub AddField(fields, fCount, v)
    If fCount > UBound(fields) Then ReDim Preserve fields(UBound(fields) * 2 + 1)
    fields(fCount) = v
    fCount = fCount + 1
End Sub

' PushRow: 読みかけの行を確定して rows に追加し、次の行の読み込みに備える
Sub PushRow(rows, nRows, fields, fCount)
    Dim r, k
    ReDim r(fCount - 1)
    For k = 0 To fCount - 1
        r(k) = fields(k)
    Next
    PushArr rows, nRows, r
    fCount = 0
End Sub

' PushArr: rows の末尾に1行（項目の配列）を追加
Sub PushArr(rows, nRows, arr)
    If nRows > UBound(rows) Then ReDim Preserve rows(UBound(rows) * 2 + 1)
    rows(nRows) = arr
    nRows = nRows + 1
End Sub

' 行ごとに項目数が違う「配列の配列」を、Excel から読んだときと同じ2次元配列にそろえる
'   列数は一番長い行に合わせ、足りない項目は Empty（空）になる
Function ToMatrix(rows, nRows)
    Dim m, maxC, r, k
    maxC = 0
    For r = 0 To nRows - 1
        If UBound(rows(r)) + 1 > maxC Then maxC = UBound(rows(r)) + 1
    Next
    If nRows = 0 Or maxC = 0 Then
        ReDim m(0, 0)
        ToMatrix = m
        Exit Function
    End If

    ReDim m(nRows - 1, maxC - 1)
    For r = 0 To nRows - 1
        For k = 0 To UBound(rows(r))
            m(r, k) = rows(r)(k)
        Next
    Next
    ToMatrix = m
End Function

' 空行か（全項目が空白のみ。「,,,」だけの行も空行）。エラー値のセルがあれば空行ではない
Function IsBlankRow(data, r)
    Dim c
    IsBlankRow = False
    For c = LBound(data, 2) To UBound(data, 2)
        If VarType(data(r, c)) = vbError Then Exit Function
        If Len(JTrim(CellText(data(r, c)))) > 0 Then Exit Function
    Next
    IsBlankRow = True
End Function


'==============================================================
' 書き込み
'   rowBuf の各要素：配列 = 見出し・データ行 / 文字列 = 先頭・末尾レコード（そのまま出力）
'==============================================================
' 出力バッファの末尾に1行追加（足りなければ2倍に広げる）
Sub AddItem(rowBuf, nRow, item)
    If nRow > UBound(rowBuf) Then ReDim Preserve rowBuf(UBound(rowBuf) * 2 + 1)
    rowBuf(nRow) = item          ' 配列はコピーされる
    nRow = nRow + 1
End Sub

' CSV の1項目を出力用に整える
'   区切り文字・ダブルクォート・改行を含む値は必ず "..." で囲み、中の " は "" にする
'   quoteAll = 全項目を囲む / quoteEmpty = 全項目を囲むときに空の値も "" にする
Function Esc(s, dlm, quoteAll, quoteEmpty)
    If s = "" Then
        If quoteAll And quoteEmpty Then Esc = QUOTE & QUOTE Else Esc = ""
    ElseIf quoteAll Or InStr(s, dlm) > 0 Or InStr(s, QUOTE) > 0 Or InStr(s, vbLf) > 0 Or InStr(s, vbCr) > 0 Then
        Esc = QUOTE & Replace(s, QUOTE, QUOTE & QUOTE) & QUOTE
    Else
        Esc = s
    End If
End Function

' 出力バッファ全体を CSV の文字列にする（改行コード・末尾改行は設定に従う）
'   配列の要素 = 見出し・データ行（項目ごとに Esc）/ 文字列の要素 = 先頭・末尾レコード（そのまま）
Function BuildCsv(rowBuf, nRow)
    Dim lines, r, i, j
    BuildCsv = ""
    If nRow = 0 Then Exit Function
    ReDim lines(nRow - 1)
    ReDim r(mCnt - 1)
    For i = 0 To nRow - 1
        If IsArray(rowBuf(i)) Then
            For j = 0 To mCnt - 1
                r(j) = Esc(CStr(rowBuf(i)(j)), gDlm, gQuoteAll, gQuoteEmpty)
            Next
            lines(i) = Join(r, gDlm)
        Else
            lines(i) = rowBuf(i)
        End If
    Next
    BuildCsv = Join(lines, gNL)
    If gTrailNL Then BuildCsv = BuildCsv & gNL
End Function

' 文字列 txt を文字コード cs でファイルに保存する（同名ファイルは上書き）
'   ADODB.Stream は UTF-8 で保存すると必ず BOM を付けるので、「UTF-8」（BOMなし）指定のときは
'   先頭3バイトを飛ばして別のストリームにコピーしてから保存する
'   「UTF-8-BOM」指定のときはそのまま保存（BOM 付き）。Shift_JIS などもそのまま
Sub WriteText(path, txt, cs)
    Dim st, bin
    If Len(txt) = 0 Then
        gFso.CreateTextFile(path, True).Close
        Exit Sub
    End If

    Set st = NewObject("ADODB.Stream")
    st.Type = AD_TYPE_TEXT
    st.Charset = StreamCharset(cs)
    st.Open
    st.WriteText txt

    If UCase(cs) = UCase(gUtf8NoBom) Then
        ' BOMを除去して保存
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

' 設定の文字コード名（UTF-8 / UTF-8-BOM / Shift_JIS など）→ ADODB.Stream に渡す名前
Function StreamCharset(cs)
    If UCase(cs) = UCase(gUtf8NoBom) Or UCase(cs) = UCase(gUtf8Bom) Then
        StreamCharset = gUtf8Charset
    Else
        StreamCharset = cs
    End If
End Function

' Excel で保存（失敗しても新規ブックは必ず閉じる）
Sub WriteExcel(path, rowBuf, nRow)
    Dim wb, en, ed
    Set wb = gXL.Workbooks.Add(XL_WBAT_WORKSHEET)
    On Error Resume Next
    FillAndSave wb, path, rowBuf, nRow
    en = Err.Number: ed = Err.Description
    wb.Close False
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, APP_TITLE, ed
End Sub

' 値は文字列のまま書き込む（先頭ゼロや日付書式を崩さないため）
'   全セルの書式を「@（文字列）」にしてから、2次元配列で一括代入する（1セルずつより大幅に速い）
'   先頭・末尾レコードは区切り文字で分けて別々のセルに入れる
'   Excel の上限（行数・セルの文字数）を超える場合はエラーにする
Sub FillAndSave(wb, path, rowBuf, nRow)
    Dim ws, m, i, j, item, nCol, s
    Set ws = wb.Worksheets(1)
    If nRow > ws.Rows.Count Then Fail "Excelの最大行数を超えています（" & nRow & " 行）"
    If gOutSheet <> "" Then ws.Name = gOutSheet
    If nRow > 0 Then
        nCol = mCnt
        For i = 0 To nRow - 1
            If Not IsArray(rowBuf(i)) Then
                If UBound(Split(rowBuf(i), gDlm)) + 1 > nCol Then nCol = UBound(Split(rowBuf(i), gDlm)) + 1
            End If
        Next
        ReDim m(nRow - 1, nCol - 1)
        For i = 0 To nRow - 1
            If IsArray(rowBuf(i)) Then item = rowBuf(i) Else item = Split(rowBuf(i), gDlm)
            For j = 0 To UBound(item)
                s = CStr(item(j))
                If Len(s) > gXlCellMax Then Fail "Excelのセルの最大文字数（" & gXlCellMax & "）を超えています: " & (i + 1) & "行目 " & (j + 1) & "列目"
                m(i, j) = s
            Next
        Next
        With ws.Range(ws.Cells(1, 1), ws.Cells(nRow, nCol))
            .NumberFormat = gXlCellFmt
            .Value = m
        End With
        If gOutHdr Then
            If gHeadRec <> "" Then ws.Rows(2).Font.Bold = True Else ws.Rows(1).Font.Bold = True
        End If
    End If
    wb.SaveAs path, gXlOutFormat
End Sub


'==============================================================
' 設定の読み込み
'   設定ファイルの値はここで一度だけ読み、g / m / f で始まる変数に入れる。
'   読み込み中の誤り（キーの不足・数値でない値・存在しないフォルダなど）は
'   すべて Fail で全体停止にする（誤った設定のまま変換しないため）
'==============================================================
' 設定シートとシステム設定シートの「A列キー → B列値」を1つの辞書 gCfg にまとめる
' （同じキーが両方のシートにあるとエラー）
Sub LoadConfig(wb)
    Set gCfg = NewObject("Scripting.Dictionary")
    ReadCfgSheet wb, SH_SYS
    ReadCfgSheet wb, SH_SET
End Sub

' シート nm の2行目以降を読み、A列が空でない行を gCfg に追加する（見出し行・区切り行は無視）
Sub ReadCfgSheet(wb, nm)
    Dim a, r, k
    a = SheetArray(GetSheet(wb, nm))
    For r = RowIdx(a, CFG_START_ROW) To UBound(a, 1)
        k = CellStr(a, r, CFG_KEY_COL)
        If k <> "" Then
            If gCfg.Exists(k) Then Fail "設定キーが重複しています: " & k & "（" & nm & "シート）"
            gCfg.Add k, CellStr(a, r, CFG_VAL_COL)
        End If
    Next
End Sub

' 設定値の取り出し
'   Cfg     … 値（キーがなければエラー、空欄は許す）
'   CfgReq  … 値（空欄もエラー＝必須項目）
'   CfgLng  … 数値（数値でなければエラー）
'   CfgList … 「設定リスト区切り（,）」で区切った配列
'   InList  … v がリストのどれかと等しいか（大文字小文字・前後空白を無視）
Function Cfg(k)
    If Not gCfg.Exists(k) Then Fail "設定が見つかりません:「" & k & "」（" & SH_SET & " / " & SH_SYS & " シートのA列）"
    Cfg = gCfg(k)
End Function

Function CfgReq(k)
    CfgReq = Cfg(k)
    If CfgReq = "" Then Fail "設定「" & k & "」が未入力です"
End Function

Function CfgLng(k)
    Dim s
    s = Cfg(k)
    If Not IsNumeric(s) Then Fail "設定「" & k & "」は数値で入力してください: " & s
    CfgLng = CLng(s)
End Function

Function CfgList(k)
    CfgList = Split(Cfg(k), gListSep)
End Function

Function InList(v, k)
    Dim item
    InList = False
    For Each item In CfgList(k)
        If UCase(Trim(item)) = UCase(Trim(v)) Then
            InList = True
            Exit Function
        End If
    Next
End Function

' 大文字小文字を区別しない正規表現オブジェクトを作る
Function NewRegExp(pattern)
    Dim re
    Set re = NewObject("VBScript.RegExp")
    re.Pattern = pattern
    re.IgnoreCase = True
    Set NewRegExp = re
End Function

' システム設定シートの値（内部パラメータ）を読む。通常は変更しない値
'   変換名（TRIM・日付 など）は大文字にそろえ、変換指定の大文字小文字を区別しないようにする
Sub LoadSystem()
    gListSep = CfgReq(S_LIST_SEP)
    gShMap = CfgReq(S_SH_MAP)
    gShCode = CfgReq(S_SH_CODE)
    gShFilter = CfgReq(S_SH_FILTER)
    gMultiSep = CfgReq(S_MULTI_SEP)
    gDateToken = CfgReq(S_DATE_TOKEN)
    gCountToken = CfgReq(S_COUNT_TOKEN)
    gSumOpen = CfgReq(S_SUM_OPEN)
    gSumClose = CfgReq(S_SUM_CLOSE)
    gExcludePrefix = Cfg(S_EXCLUDE_PREFIX)
    gConvSep = CfgReq(S_CONV_SEP)
    gArgSep = CfgReq(S_ARG_SEP)
    gArgSepWide = Cfg(S_ARG_SEP_WIDE)
    gReplSep = CfgReq(S_REPL_SEP)
    gCvTrim = UCase(CfgReq(S_CV_TRIM))
    gCvDate = UCase(CfgReq(S_CV_DATE))
    gCvNum = UCase(CfgReq(S_CV_NUM))
    gCvZero = UCase(CfgReq(S_CV_ZERO))
    gCvRepl = UCase(CfgReq(S_CV_REPL))
    gCvEmpty = UCase(CfgReq(S_CV_EMPTY))
    gCvLeft = UCase(CfgReq(S_CV_LEFT))
    gCvRight = UCase(CfgReq(S_CV_RIGHT))
    gCvCode = UCase(CfgReq(S_CV_CODE))
    gCodeDefault = CfgReq(S_CODE_DEFAULT)
    gFilInclude = CfgReq(S_FIL_INCLUDE)
    gFilExclude = CfgReq(S_FIL_EXCLUDE)
    gOpEq = CfgReq(S_OP_EQ)
    gOpNe = CfgReq(S_OP_NE)
    gOpHas = CfgReq(S_OP_HAS)
    gOpNotHas = CfgReq(S_OP_NOT_HAS)
    gOpEmpty = CfgReq(S_OP_EMPTY)
    gOpNotEmpty = CfgReq(S_OP_NOT_EMPTY)
    gOpRegex = CfgReq(S_OP_REGEX)
    Set gReDate8 = NewRegExp(CfgReq(S_DATE8_RE))
    gThousands = CfgReq(S_THOUSANDS)
    gXlInExt = CfgList(S_XL_IN_EXT)
    gXlOutExt = CfgReq(S_XL_OUT_EXT)
    gXlOutFormat = CfgLng(S_XL_OUT_FORMAT)
    gXlCellFmt = CfgReq(S_XL_CELL_FMT)
    gXlSheetBad = CfgReq(S_XL_SHEET_BAD)
    gXlSheetMax = CfgLng(S_XL_SHEET_MAX)
    gXlCellMax = CfgLng(S_XL_CELL_MAX)
    gUtf8NoBom = CfgReq(S_UTF8_NOBOM)
    gUtf8Bom = CfgReq(S_UTF8_BOM)
    gUtf8Charset = CfgReq(S_UTF8_CHARSET)
    gLogName = CfgReq(S_LOG_NAME)
    gLogCs = CfgReq(S_LOG_CS)
    gLogDlm = CfgReq(S_LOG_DLM)
    gLogTimeFmt = CfgReq(S_LOG_TIME_FMT)
    gBufRows = CfgLng(S_BUF_ROWS)
    gBufFields = CfgLng(S_BUF_FIELDS)
    If gBufRows < 1 Or gBufFields < 1 Then Fail "バッファ初期サイズは1以上にしてください"
End Sub

' 設定シート（業務設定）の値を読み、組み合わせとして正しいかを確認する
'   あり／なしの項目は「あり判定値」（あり,する,はい,TRUE…）に含まれれば あり
'   ログフォルダが空欄なら出力フォルダに保存する
Sub LoadSettings()
    Dim fmt, i

    gInDir = CfgReq(K_IN_DIR)
    gPattern = CfgReq(K_PATTERN)
    Set gReDate = NewRegExp(CfgReq(K_DATE_RE))
    gInCs = CfgReq(K_IN_CS)
    gHdrRow = CfgLng(K_HDR_ROW)
    gFooterSkip = CfgLng(K_FOOTER)
    gDlm = CfgReq(K_DLM)
    If UCase(gDlm) = UCase(Cfg(S_TAB_WORD)) Then gDlm = vbTab
    gOutDir = CfgReq(K_OUT_DIR)
    gOutName = CfgReq(K_OUT_NAME)
    gOutSheet = Cfg(K_OUT_SHEET)
    gOutCs = CfgReq(K_OUT_CS)
    gOutHdr = IsYes(Cfg(K_OUT_HDR))
    gQuoteAll = IsYes(Cfg(K_QUOTE_ALL))
    gQuoteEmpty = IsYes(Cfg(K_QUOTE_EMPTY))
    If UCase(Cfg(K_NL)) = UCase(Cfg(S_LF_WORD)) Then gNL = vbLf Else gNL = vbCrLf
    gTrailNL = IsYes(Cfg(K_TRAIL_NL))
    gHeadRec = Cfg(K_HEAD_REC)
    gTailRec = Cfg(K_TAIL_REC)
    gOutZero = IsYes(Cfg(K_OUT_ZERO))
    gOverwrite = IsYes(Cfg(K_OVERWRITE))
    gLogDir = Cfg(K_LOG_DIR)
    If gLogDir = "" Then gLogDir = gOutDir

    ' 出力形式の判定。Excel 出力なら出力ファイル名の拡張子を .xlsx に置き換え、
    ' シート名が Excel の制約（31文字以内・禁止文字なし）を満たすか確認する
    fmt = CfgReq(K_OUT_FMT)
    If InList(fmt, S_FMT_EXCEL) Then
        gIsExcelOut = True
        gOutName = ChangeExt(gOutName, gXlOutExt)
        If gOutSheet <> "" Then
            If Len(gOutSheet) > gXlSheetMax Then Fail "「" & K_OUT_SHEET & "」は " & gXlSheetMax & " 文字以内にしてください"
            For i = 1 To Len(gXlSheetBad)
                If InStr(gOutSheet, Mid(gXlSheetBad, i, 1)) > 0 Then Fail "「" & K_OUT_SHEET & "」に使えない文字があります: " & Mid(gXlSheetBad, i, 1)
            Next
        End If
    ElseIf InList(fmt, S_FMT_CSV) Then
        gIsExcelOut = False
        If IsExcelExt(gOutName) Then Fail "出力形式が CSV なのに出力ファイル名が Excel の拡張子です: " & gOutName
    Else
        Fail "「" & K_OUT_FMT & "」の値が不正です: " & fmt
    End If

    ' 値の範囲とフォルダの確認。入力と出力が同じフォルダだと、出力ファイルを次回の入力として
    ' 読んでしまうおそれがあるので禁止している
    If gHdrRow < 0 Then Fail "「" & K_HDR_ROW & "」は0以上にしてください"
    If gFooterSkip < 0 Then Fail "「" & K_FOOTER & "」は0以上にしてください"
    If Not gFso.FolderExists(gInDir) Then Fail "入力フォルダが存在しません: " & gInDir
    If Not gFso.FolderExists(gOutDir) Then Fail "出力フォルダが存在しません: " & gOutDir
    If Not gFso.FolderExists(gLogDir) Then Fail "ログフォルダが存在しません: " & gLogDir
    If LCase(gFso.GetAbsolutePathName(gInDir)) = LCase(gFso.GetAbsolutePathName(gOutDir)) Then
        Fail "入力フォルダと出力フォルダは別にしてください"
    End If
    If InStr(gOutName, gDateToken) = 0 Then Fail "出力ファイル名に " & gDateToken & " を含めてください"
End Sub

' マッピングシートを読み、m で始まる配列に入れる（行の順番 = fileB の出力列の順番）
'   fileB項目名が入っている最後の行までを対象にする（途中の空行はエラー）
'   fileA項目名は「|」で複数指定できる（連結）。ヘッダーなしの場合は列番号で指定
'   固定値・連結文字は前後の空白も意味があるので Trim しない（CellRaw）
Sub LoadMapping(wb)
    Dim a, r, j, k, startR, last, cOut, cIn, cFix, cConv, cJoin, src, parts

    a = SheetArray(GetSheet(wb, gShMap))
    startR = RowIdx(a, CfgLng(S_MAP_ROW))
    cOut = CfgLng(S_MAP_COL_OUT)
    cIn = CfgLng(S_MAP_COL_IN)
    cFix = CfgLng(S_MAP_COL_FIX)
    cConv = CfgLng(S_MAP_COL_CONV)
    cJoin = CfgLng(S_MAP_COL_JOIN)

    last = startR - 1
    For r = startR To UBound(a, 1)
        If CellStr(a, r, cOut) <> "" Then last = r
    Next
    If last < startR Then Fail "マッピングが入力されていません"

    mCnt = last - startR + 1
    ReDim mOut(mCnt - 1), mParts(mCnt - 1), mFix(mCnt - 1), mConv(mCnt - 1), mJoin(mCnt - 1)
    For r = startR To last
        j = r - startR
        mOut(j) = CellStr(a, r, cOut)
        mFix(j) = CellRaw(a, r, cFix)
        mConv(j) = CellStr(a, r, cConv)
        mJoin(j) = CellRaw(a, r, cJoin)
        src = CellStr(a, r, cIn)
        If src = "" Then
            mParts(j) = Array()
        Else
            parts = Split(src, gMultiSep)
            For k = 0 To UBound(parts)
                parts(k) = Trim(parts(k))
                If parts(k) = "" Then Fail gShMap & " " & (r - LBound(a, 1) + 1) & "行目: fileA項目名の指定が不正です: " & src
                If gHdrRow = 0 And Not IsNumeric(parts(k)) Then
                    Fail gShMap & " " & (r - LBound(a, 1) + 1) & "行目: ヘッダーなしの場合、fileA項目には列番号を入れてください"
                End If
            Next
            mParts(j) = parts
        End If
        If mOut(j) = "" Then Fail gShMap & " " & (r - LBound(a, 1) + 1) & "行目: fileB項目名が空です"
    Next
End Sub

' 抽出条件シートを読み、f で始まる配列に入れる（シートがない・行がないなら条件なし＝全行出力）
'   種別列が空の行は説明行として読み飛ばす。正規表現はここで1回だけ作っておく
Sub LoadFilters(wb)
    Dim a, r, n, kind, cKind, cName, cOp, cVal

    fCnt = 0
    If Not SheetExists(wb, gShFilter) Then Exit Sub
    a = SheetArray(GetSheet(wb, gShFilter))
    cKind = CfgLng(S_FIL_COL_KIND)
    cName = CfgLng(S_FIL_COL_NAME)
    cOp = CfgLng(S_FIL_COL_OP)
    cVal = CfgLng(S_FIL_COL_VAL)

    n = UBound(a, 1) - LBound(a, 1) + 1
    ReDim fInclude(n), fName(n), fOp(n), fVal(n), fRe(n)
    For r = RowIdx(a, CfgLng(S_FIL_ROW)) To UBound(a, 1)
        kind = CellStr(a, r, cKind)
        If kind <> "" Then
            If kind = gFilInclude Then
                fInclude(fCnt) = True
            ElseIf kind = gFilExclude Then
                fInclude(fCnt) = False
            Else
                Fail gShFilter & " " & (r - LBound(a, 1) + 1) & "行目: 種別は「" & gFilInclude & "」か「" & gFilExclude & "」にしてください"
            End If
            fName(fCnt) = CellStr(a, r, cName)
            fOp(fCnt) = CellStr(a, r, cOp)
            fVal(fCnt) = CellStr(a, r, cVal)
            If fName(fCnt) = "" Then Fail gShFilter & " " & (r - LBound(a, 1) + 1) & "行目: fileA項目名が空です"
            Select Case fOp(fCnt)
                Case gOpEq, gOpNe, gOpHas, gOpNotHas, gOpEmpty, gOpNotEmpty
                Case gOpRegex
                    Set fRe(fCnt) = NewObject("VBScript.RegExp")
                    fRe(fCnt).Pattern = fVal(fCnt)
                Case Else
                    Fail gShFilter & " " & (r - LBound(a, 1) + 1) & "行目: 不明な比較: " & fOp(fCnt)
            End Select
            fCnt = fCnt + 1
        End If
    Next
End Sub

' コード変換シートを読み、gCodes(変換名)(変換前) = 変換後 の辞書を作る
'   同じ変換名の行は1つの表にまとまる。変換前・変換後は文字列のまま比較する（"1" と "01" は別）
Sub LoadCodes(wb)
    Dim a, r, nm, d, cName, cFrom, cTo

    Set gCodes = NewObject("Scripting.Dictionary")
    If Not SheetExists(wb, gShCode) Then Exit Sub
    a = SheetArray(GetSheet(wb, gShCode))
    cName = CfgLng(S_CODE_COL_NAME)
    cFrom = CfgLng(S_CODE_COL_FROM)
    cTo = CfgLng(S_CODE_COL_TO)

    For r = RowIdx(a, CfgLng(S_CODE_ROW)) To UBound(a, 1)
        nm = CellStr(a, r, cName)
        If nm <> "" Then
            If Not gCodes.Exists(nm) Then gCodes.Add nm, NewObject("Scripting.Dictionary")
            Set d = gCodes(nm)
            d(CellRaw(a, r, cFrom)) = CellRaw(a, r, cTo)
        End If
    Next
End Sub


'==============================================================
' ログ（メモリに溜めて最後にCSVで保存。cscript 実行時は画面にも表示）
'   1行の形式: 日時,区分,日付,ファイル,内容
'   区分: INFO = 経過 / OK = 出力した / SKIP = 対象外・出力しない / WARN = 注意 / ERROR = エラー
'   ファイル名は「変換ログ_yyyymmdd_hhmmss.csv」（実行ごとに新しいファイル）。文字コードは UTF-8（BOM付き）
'==============================================================
' ログファイルの場所と名前を決め、見出し行を用意する（設定を読み終えてから呼ぶ）
Sub InitLog()
    Dim nm
    nm = Replace(gLogName, CfgReq(S_LOG_TIME_TOKEN), _
                 gXL.WorksheetFunction.Text(Now, CfgReq(S_LOG_FILE_TIME_FMT)))
    ReDim gLogLines(gBufRows - 1)
    gLogCount = 0
    AddLogLine CfgList(S_LOG_HEADER)
    gLogPath = gFso.BuildPath(gLogDir, nm)
End Sub

' ログを1行追加する（kind = 区分 / dt = 日付 / target = ファイル・場所 / msg = 内容）
Sub WriteLog(kind, dt, target, msg)
    AddLogLine Array(gXL.WorksheetFunction.Text(Now, gLogTimeFmt), kind, dt, target, msg)
    If gIsCscript Then WScript.Echo kind & vbTab & dt & vbTab & target & vbTab & msg
End Sub

' 値の配列を CSV の1行にしてメモリにためる
Sub AddLogLine(vals)
    Dim a, i
    ReDim a(UBound(vals))
    For i = 0 To UBound(vals)
        a(i) = Esc(Trim(CStr(vals(i))), gLogDlm, False, False)
    Next
    If gLogCount > UBound(gLogLines) Then ReDim Preserve gLogLines(UBound(gLogLines) * 2 + 1)
    gLogLines(gLogCount) = Join(a, gLogDlm)
    gLogCount = gLogCount + 1
End Sub

' ためたログをファイルに保存する（Main の最後に1回だけ呼ぶ）
Sub FlushLog()
    Dim lines, i
    If gLogCount = 0 Then Exit Sub
    ReDim lines(gLogCount - 1)
    For i = 0 To gLogCount - 1
        lines(i) = gLogLines(i)
    Next
    WriteText gLogPath, Join(lines, vbCrLf) & vbCrLf, gLogCs
End Sub


'==============================================================
' 共通部品
'==============================================================
' エラーを発生させる。処理はここで中断し、直近の On Error Resume Next の場所
' （1日付分なら RunAll、全体なら Main）まで戻って、msg がログに書かれる
Sub Fail(msg)
    Err.Raise vbObjectError + ERR_OFFSET, APP_TITLE, msg
End Sub

' 結果の表示。cscript（コマンド・タスクスケジューラ）なら画面に出力、
' wscript（ダブルクリック）ならメッセージボックス（OK を押すまで待つ）
Sub ShowMsg(msg)
    If gIsCscript Then
        WScript.Echo msg
    Else
        MsgBox msg, vbOKOnly, APP_TITLE
    End If
End Sub

' ブック wb にシート nm があるか / シートを取り出す（なければエラー）
Function SheetExists(wb, nm)
    Dim i
    SheetExists = False
    For i = 1 To wb.Worksheets.Count
        If wb.Worksheets(i).Name = nm Then
            SheetExists = True
            Exit Function
        End If
    Next
End Function

Function GetSheet(wb, nm)
    If Not SheetExists(wb, nm) Then Fail "シート「" & nm & "」がありません"
    Set GetSheet = wb.Worksheets(nm)
End Function

' Excel の行番号・列番号（1始まり）→ 配列の位置
Function RowIdx(a, rowNo)
    RowIdx = LBound(a, 1) + rowNo - 1
End Function

' 配列 a の r 行・colNo 列（1始まり）の値を文字列で返す。範囲外・空・エラー値は ""
'   CellRaw はそのまま、CellStr は前後の半角スペースを除く
Function CellRaw(a, r, colNo)
    Dim c, v
    CellRaw = ""
    c = LBound(a, 2) + colNo - 1
    If r > UBound(a, 1) Or c > UBound(a, 2) Then Exit Function
    v = a(r, c)
    If VarType(v) = vbError Or IsEmpty(v) Then Exit Function
    CellRaw = CStr(v)
End Function

Function CellStr(a, r, colNo)
    CellStr = Trim(CellRaw(a, r, colNo))
End Function

' ファイル名から日付を取り出す（日付抽出パターンの最初の ( ) の部分。なければ一致した全体）
'   例 "fileA_20261106_01.csv" + (\d{8}) → "20261106"。取り出せなければ ""
Function ExtractDate(fn)
    Dim m
    ExtractDate = ""
    Set m = gReDate.Execute(fn)
    If m.Count = 0 Then Exit Function
    If m(0).SubMatches.Count > 0 Then
        ExtractDate = m(0).SubMatches(0)
    Else
        ExtractDate = m(0).Value
    End If
End Function

' ファイル名から取り出した日付が正しいか
'   8桁の数字（yyyymmdd）なら、年が FILE_DATE_MIN_YEAR～FILE_DATE_MAX_YEAR の範囲で、
'   実際にある日付（月 1～12、その月にある日）のときだけ True。
'     例 "20261106" → True / "20261332" → False
'        "17000202"（番号と日付がつながった 1700020261106.csv から取った値）→ 年が 1700 なので False
'   8桁の数字でない形（「日付抽出パターン」を変えて別の形を取り出す場合）は、これまでどおり True
Function IsValidFileDate(dt)
    Dim m, y, mo, dd, d
    IsValidFileDate = True
    Set m = gReDate8.Execute(dt)
    If m.Count = 0 Or Len(dt) <> 8 Then Exit Function
    y = CInt(m(0).SubMatches(SM_YEAR))
    mo = CInt(m(0).SubMatches(SM_MONTH))
    dd = CInt(m(0).SubMatches(SM_DAY))
    If y < FILE_DATE_MIN_YEAR Or y > FILE_DATE_MAX_YEAR Or mo < 1 Or mo > 12 Or dd < 1 Then
        IsValidFileDate = False
        Exit Function
    End If
    ' DateSerial は 13月や32日を繰り上げるので、組み立てた日付が元と一致するかで判定する
    d = DateSerial(y, mo, dd)
    IsValidFileDate = (Year(d) = y And Month(d) = mo And Day(d) = dd)
End Function

' ワイルドカード（* ?）→ 正規表現
'   例 "fileA_*.csv" → "^fileA_.*\.csv$"（正規表現の記号は \ を付けて普通の文字として扱う）
Function WildToRegex(p)
    Dim i, ch, s
    s = "^"
    For i = 1 To Len(p)
        ch = Mid(p, i, 1)
        Select Case ch
            Case "*": s = s & ".*"
            Case "?": s = s & "."
            Case Else
                If InStr("\.+^$()[]{}|", ch) > 0 Then s = s & "\"
                s = s & ch
        End Select
    Next
    WildToRegex = s & "$"
End Function

' 対象外のファイル名か（Excel で開いているときにできる一時ファイル ~$xxx.xlsx など）
Function IsExcluded(fn)
    IsExcluded = False
    If gExcludePrefix = "" Then Exit Function
    IsExcluded = (Left(fn, Len(gExcludePrefix)) = gExcludePrefix)
End Function

' 拡張子が Excel 入力拡張子（xlsx,xlsm,xls,xlsb）のどれかか
Function IsExcelExt(fn)
    Dim ext, item
    IsExcelExt = False
    ext = LCase(gFso.GetExtensionName(fn))
    For Each item In gXlInExt
        If LCase(Trim(item)) = ext Then
            IsExcelExt = True
            Exit Function
        End If
    Next
End Function

' ファイル名の拡張子を ext に置き換える（{date} より後ろにある「.」だけを拡張子とみなす）
Function ChangeExt(fn, ext)
    Dim p, s
    s = fn
    p = InStrRev(s, ".")
    If p > 0 And p > InStrRev(s, gDateToken) Then s = Left(s, p - 1)
    ChangeExt = s & ext
End Function

' 設定値が「あり」系の言葉か
Function IsYes(s)
    IsYes = InList(s, S_YES_WORDS)
End Function

' 前後の半角スペース・全角スペース・タブを除く（VBScript の Trim は半角スペースしか除かないため）
Function JTrim(ByVal s)
    Do While Len(s) > 0
        If InStr(gTrimChars, Left(s, 1)) = 0 Then Exit Do
        s = Mid(s, 2)
    Loop
    Do While Len(s) > 0
        If InStr(gTrimChars, Right(s, 1)) = 0 Then Exit Do
        s = Left(s, Len(s) - 1)
    Loop
    JTrim = s
End Function

' 配列を昇順に並べ替えたコピーを返す（挿入ソート。日付・ファイル名の数は少ないので十分速い）
Function SortArr(ByVal a)
    Dim i, j, t
    For i = LBound(a) + 1 To UBound(a)
        t = a(i)
        j = i - 1
        Do While j >= LBound(a)
            If a(j) <= t Then Exit Do
            a(j + 1) = a(j)
            j = j - 1
        Loop
        a(j + 1) = t
    Next
    SortArr = a
End Function
