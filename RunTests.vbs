'==============================================================
' RunTests.vbs  DataMapping.vbs ‚Ì©“®ƒeƒXƒg
'
'   Às: ƒeƒXƒgÀs.bat ‚ğƒ_ƒuƒ‹ƒNƒŠƒbƒNi‚Ü‚½‚Í cscript //nologo RunTests.vbsj
'   ˆê•”‚¾‚¯: cscript //nologo RunTests.vbs <DataMapping.vbs> <DataMapping.xlsx> A1,B2a
'   Œ‹‰Ê: ‚±‚ÌƒtƒHƒ‹ƒ_‚ÌuƒeƒXƒgŒ‹‰Ê.csvv‚Æ‰æ–Ê•\¦
'   I—¹ƒR[ƒh: 0 = ‘SŒPASS / 1 = FAIL‚ ‚è
'
'   ì‹ÆƒtƒHƒ‹ƒ_iworkj‚ÉƒeƒXƒg—p‚Ì“ü—Íƒtƒ@ƒCƒ‹Eİ’èƒtƒ@ƒCƒ‹‚ğ–ˆ‰ñì‚è’¼‚µA
'   ƒpƒ^[ƒ“‚²‚Æ‚É DataMapping.vbs ‚ğÀs‚µ‚ÄAo—ÍEƒƒOEI—¹ƒR[ƒh‚ğŒŸØ‚·‚éB
'   ¦ ƒeƒXƒg’†‚Í Excel ‚ğ— ‚Å‹N“®‚µ‚Ü‚·BƒeƒXƒg’†‚Í Excel ‚ğ‘€ì‚µ‚È‚¢‚Å‚­‚¾‚³‚¢B
'==============================================================
Option Explicit

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
Const RESULT_FILE = "ƒeƒXƒgŒ‹‰Ê.csv"
Const WORK_DIR = "work"
Const MAP_CLEAR_ROWS = 60
Const MAP_COLS = 6
Const SIMPLE_COLS = 4

Dim gFso, gSh, gXL, gBase, gWork, gTool, gTemplate
Dim gResults, gResN, gPass, gFail, gSkip
Dim gCase, gTitle, gDir, gCaseFail
Dim gSet, gSys, gMap, gCodes, gFilters, gRemoveKeys

WScript.Quit RunAllTests()

'==============================================================
' ŠÂ‹«ˆË‘¶•”•ªiwine ŒŸØ‚Í‚±‚±‚ğ·‚µ‘Ö‚¦‚éj
'==============================================================
Function NewObject(progId)
    Set NewObject = CreateObject(progId)
End Function

Function IsMock()
    IsMock = False
End Function

'==============================================================
' ‘S‘Ì‚Ì—¬‚ê
'==============================================================
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
        WScript.Echo "DataMapping.vbs ‚ªŒ©‚Â‚©‚è‚Ü‚¹‚ñ: " & gTool
        RunAllTests = 1
        Exit Function
    End If
    If Not gFso.FileExists(gTemplate) Then
        WScript.Echo "İ’èƒtƒ@ƒCƒ‹‚Ì‚Ğ‚ÈŒ`‚ªŒ©‚Â‚©‚è‚Ü‚¹‚ñ: " & gTemplate
        RunAllTests = 1
        Exit Function
    End If

    gWork = gFso.BuildPath(gBase, WORK_DIR)
    If gFso.FolderExists(gWork) Then gFso.DeleteFolder gWork, True
    gFso.CreateFolder gWork

    Set gXL = NewObject("Excel.Application")
    gXL.Visible = False
    gXL.DisplayAlerts = False

    ReDim gResults(255)
    gResN = 0: gPass = 0: gFail = 0: gSkip = 0
    AddResult "ID", "ƒpƒ^[ƒ“", "Šm”F€–Ú", "Œ‹‰Ê", "Ú×"
    t0 = Timer

    If WScript.Arguments.Count > 2 Then
        ids = Split(WScript.Arguments(2), ",")      ' w’è‚µ‚½ID‚¾‚¯Àsi—á: A1,B2aj
    Else
        ids = Split("A1 A2 A3 A4 A5 B1a B1b B2a B2b B2c B3 B4 B5a B5b B6 B7 " & _
                "C1 C2 C3 C4 C5 C6 C7 C8 C9 C10 C11 C13 C14 C15 " & _
                "D1 D2 D3 D4a D4b D5 D6 " & _
                "E1 E2a E2b E3 E4 E5a E5b E5c E5d E5e E6a E6b E7 E8a E8b E9 E10 E11a E11b E12 E13 " & _
                "F2 F5 G1a G1b G1c G1d G2 G3 H1 H2 H3 H4 H5a H5b H6 H7 H8 P1", " ")
    End If
    For i = 0 To UBound(ids)
        RunCase ids(i)
    Next

    gXL.Quit
    Set gXL = Nothing
    WriteResults
    WScript.Echo ""
    WScript.Echo "PASS: " & gPass & "  FAIL: " & gFail & "  SKIP: " & gSkip & "  i" & Round(Timer - t0) & " •bj"
    WScript.Echo "Œ‹‰Ê: " & gFso.BuildPath(gBase, RESULT_FILE)
    If gFail > 0 Then RunAllTests = 1 Else RunAllTests = 0
End Function

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
        Check "ƒeƒXƒgƒR[ƒh‚ÌÀs", False, "—áŠO: " & Err.Description
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
' ƒeƒXƒgƒP[ƒX
'==============================================================
' •W€‚Ì“ü—ÍEŠú‘Ò’l
Function StdHdr(): StdHdr = "“X•ÜCD,”„ã“ú,”Ì”„”,‹æ•ª": End Function
Function OutHdr(): OutHdr = "“X•ÜƒR[ƒh,æˆø“ú,”—Ê,‹æ•ª,˜AŒg‹æ•ª": End Function

Sub MakeStdInputs()
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr(), "12,20261106,5,1", "345,20261106,10,2")), CS_SJIS
    MakeCsv "in\fileA_20261106_02.csv", L(Array(StdHdr(), "7,20261106,1,1")), CS_SJIS
    MakeCsv "in\fileA_20261107_01.csv", L(Array(StdHdr(), "12,20261107,3,2")), CS_SJIS
End Sub

Function Exp1106(): Exp1106 = Array(OutHdr(), "00012,20261106,5,01,1", "00345,20261106,10,02,1", "00007,20261106,1,01,1"): End Function
Function Exp1107(): Exp1107 = Array(OutHdr(), "00012,20261107,3,02,1"): End Function

Sub T_A1()
    Dim rc
    BeginCase "A1", "Šî–{F“¯‚¶“ú•t‚Ì2ƒtƒ@ƒCƒ‹‚ğŒ‹‡A•Ê“ú•t‚Í•Êƒtƒ@ƒCƒ‹"
    MakeStdInputs
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Exp1106())
    ExpectText "out\fileB_20261107.csv", CS_SJIS, L(Exp1107())
    ExpectLog "OK", "fileB_20261106.csv"
    ExpectLog "OK", "2 ƒtƒ@ƒCƒ‹ ¨ 3 s o—Í"
    ExpectLogHex "EFBBBF"
End Sub

Sub T_A2()
    Dim rc
    BeginCase "A2", "Œ‹‡‡Fƒtƒ@ƒCƒ‹–¼‡iì¬‡‚Å‚Í‚È‚¢j"
    MakeCsv "in\fileA_20261106_02.csv", L(Array(StdHdr(), "7,20261106,1,1")), CS_SJIS
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr(), "12,20261106,5,1")), CS_SJIS
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array(OutHdr(), "00012,20261106,5,01,1", "00007,20261106,1,01,1"))
End Sub

Sub T_A3()
    Dim rc
    BeginCase "A3", "ƒtƒ@ƒCƒ‹‚²‚Æ‚É—ñ‚Ì•À‚Ñ‚ªˆá‚Á‚Ä‚à€–Ú–¼‚Å‘Î‰"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr(), "12,20261106,5,1")), CS_SJIS
    MakeCsv "in\fileA_20261106_02.csv", L(Array("‹æ•ª,”Ì”„”,”„ã“ú,“X•ÜCD", "2,10,20261106,345")), CS_SJIS
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array(OutHdr(), "00012,20261106,5,01,1", "00345,20261106,10,02,1"))
End Sub

Sub T_A4()
    Dim rc
    BeginCase "A4", "ŒÅ’è’lFfileA€–Ú‚ª‹ó—“‚Ì—ñ‚¾‚¯ŒÅ’è’l‚ğg‚¤"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr(), "12,20261106,5,1")), CS_SJIS
    gMap = Array(Array("“X•Ü", "“X•ÜCD", "X", "", "", ""), Array("ŒÅ’è", "", "ABC", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("“X•Ü,ŒÅ’è", "12,ABC"))
End Sub

Sub T_A5()
    Dim rc
    BeginCase "A5", "‹ósE‹æØ‚è‚¾‚¯‚ÌsE––”ö‚Ì‹ós‚Ío—Í‚µ‚È‚¢"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr(), "12,20261106,5,1", "", ",,,", "345,20261106,10,2", "", "")), CS_SJIS
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array(OutHdr(), "00012,20261106,5,01,1", "00345,20261106,10,02,1"))
End Sub

' •¶šƒR[ƒhŒn‚Åg‚¤A”õl—ñ‚Â‚«‚Ì“ü—Í
Function BikoHdr(): BikoHdr = "“X•ÜCD,”„ã“ú,”Ì”„”,‹æ•ª,”õl": End Function
Sub UseBikoMap()
    gMap = Array(Array("“X•ÜƒR[ƒh", "“X•ÜCD", "", "TRIM|ƒ[ƒ–„‚ß:5", "", ""), Array("”õl", "”õl", "", "", "", ""))
End Sub

Sub T_B1(id, cs)
    Dim rc
    BeginCase id, "“ü—Í‚Ì•¶šƒR[ƒhF" & cs & "iBOM‚ª€–Ú–¼‚É¬‚´‚ç‚È‚¢j"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(BikoHdr(), "12,20261106,5,1,ƒeƒXƒg‡@îà")), cs
    SetV "“ü—Í•¶šƒR[ƒh", cs
    UseBikoMap
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("“X•ÜƒR[ƒh,”õl", "00012,ƒeƒXƒg‡@îà"))
End Sub

Sub T_B2(id, cs, hexPrefix)
    Dim rc
    BeginCase id, "o—Í‚Ì•¶šƒR[ƒhF" & cs
    MakeCsv "in\fileA_20261106_01.csv", L(Array(BikoHdr(), "12,20261106,5,1,ƒeƒXƒg‡@îà")), CS_SJIS
    gMap = Array(Array("“X•Ü", "“X•ÜCD", "", "", "", ""), Array("”õl", "”õl", "", "", "", ""))
    SetV "o—Í•¶šƒR[ƒh", cs
    rc = RunTool()
    ExpectExit rc, 0
    If cs = CS_SJIS Then
        ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("“X•Ü,”õl", "12,ƒeƒXƒg‡@îà"))
    Else
        ExpectText "out\fileB_20261106.csv", CS_UTF8, L(Array("“X•Ü,”õl", "12,ƒeƒXƒg‡@îà"))
    End If
    ExpectHex "out\fileB_20261106.csv", hexPrefix
End Sub

Sub T_B3()
    Dim rc
    BeginCase "B3", "‰üsƒR[ƒh LF"
    MakeStdInputs
    SetV "‰üsƒR[ƒh", "LF"
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261107.csv", CS_SJIS, Join(Exp1107(), vbLf) & vbLf
End Sub

Sub T_B4()
    Dim rc
    BeginCase "B4", "ƒ^ƒu‹æØ‚èi“ü—ÍEo—Íj"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(Replace(StdHdr(), ",", vbTab), "12" & vbTab & "20261106" & vbTab & "5" & vbTab & "1")), CS_SJIS
    SetV "‹æØ‚è•¶š", "TAB"
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array(Replace(OutHdr(), ",", vbTab), Replace("00012,20261106,5,01,1", ",", vbTab)))
End Sub

Sub T_B5a()
    Dim rc, q
    q = """"
    BeginCase "B5a", "ƒJƒ“ƒ}E‰üsEƒ_ƒuƒ‹ƒNƒH[ƒg‚ğŠÜ‚Ş’li“ü—Í‚ào—Í‚à³‚µ‚­ˆÍ‚Şj"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(BikoHdr(), _
        "12,20261106,5,1," & q & "a,b" & q, _
        "13,20261106,5,1," & q & "1s–Ú" & vbCrLf & "2s–Ú" & q, _
        "14,20261106,5,1," & q & "say " & q & q & "hi" & q & q & q)), CS_SJIS
    UseBikoMap
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("“X•ÜƒR[ƒh,”õl", _
        "00012," & q & "a,b" & q, _
        "00013," & q & "1s–Ú" & vbCrLf & "2s–Ú" & q, _
        "00014," & q & "say " & q & q & "hi" & q & q & q))
End Sub

Sub T_B5b()
    Dim rc
    BeginCase "B5b", "‘S€–Ú‚ğƒ_ƒuƒ‹ƒNƒH[ƒg‚ÅˆÍ‚Şi‹ó‚Ì’l‚Í """"j"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(BikoHdr(), "12,20261106,5,1,")), CS_SJIS
    UseBikoMap
    SetV "‘S€–Ú‚ğƒ_ƒuƒ‹ƒNƒH[ƒg‚ÅˆÍ‚Ş", "‚ ‚è"
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("""“X•ÜƒR[ƒh"",""”õl""", """00012"","""""))
End Sub

Sub T_B6()
    Dim rc
    BeginCase "B6", "ƒwƒbƒ_[‚È‚µi—ñ”Ô†‚Åw’èj"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("12,20261106,5,1", "345,20261106,10,2")), CS_SJIS
    SetV "ƒwƒbƒ_[s”Ô†", "0"
    gMap = Array(Array("“X•ÜƒR[ƒh", "1", "", "ƒ[ƒ–„‚ß:5", "", ""), Array("‹æ•ª", "4", "", "ƒR[ƒh:‹æ•ª", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("“X•ÜƒR[ƒh,‹æ•ª", "00012,01", "00345,02"))
End Sub

Sub T_B7()
    Dim rc
    BeginCase "B7", "‹@íˆË‘¶•¶šE”¼ŠpƒJƒiiShift_JIS “üo—Íj"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(BikoHdr(), "12,20261106,5,1,‡@‡U‡Šîàí•±²³Ş")), CS_SJIS
    UseBikoMap
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("“X•ÜƒR[ƒh,”õl", "00012,‡@‡U‡Šîàí•±²³Ş"))
End Sub

' Excel “ü—ÍŒn
Sub UseXlsxInput()
    SetV "‘ÎÛƒtƒ@ƒCƒ‹", "*.xlsx"
End Sub

Sub T_C1()
    Dim rc
    BeginCase "C1", "CSV‚Ì 0001 ‚Íæ“ªƒ[ƒ‚Ì‚Ü‚Ü"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("¤•iCD", "0001")), CS_SJIS
    gMap = Array(Array("¤•iƒR[ƒh", "¤•iCD", "", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("¤•iƒR[ƒh", "0001"))
End Sub

Sub T_C2()
    Dim rc
    BeginCase "C2", "Excel‚Ì”’lƒZƒ‹ 1 ‚É ƒ[ƒ–„‚ß:4 ¨ 0001"
    MakeXlsx "in\fileA_20261106_01.xlsx", Array(Array("¤•iCD"), Array(1)), 1, 1, Empty
    UseXlsxInput
    gMap = Array(Array("¤•iƒR[ƒh", "¤•iCD", "", "ƒ[ƒ–„‚ß:4", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("¤•iƒR[ƒh", "0001"))
End Sub

Sub T_C3()
    Dim rc
    BeginCase "C3", "Excel‚Ì•¶š—ñƒZƒ‹ 0001 ‚Í‚»‚Ì‚Ü‚Ü"
    MakeXlsx "in\fileA_20261106_01.xlsx", Array(Array("¤•iCD"), Array("0001")), 1, 1, Empty
    UseXlsxInput
    gMap = Array(Array("¤•iƒR[ƒh", "¤•iCD", "", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("¤•iƒR[ƒh", "0001"))
End Sub

Sub T_C4()
    Dim rc, d
    d = DateSerial(2026, 11, 6) + TimeSerial(10, 30, 0)
    BeginCase "C4", "Excel‚Ì“ú•tƒZƒ‹i•t‚«j‚ğ“ú•t•ÏŠ·"
    MakeXlsx "in\fileA_20261106_01.xlsx", Array(Array("”„ã“ú"), Array(Array(d, "yyyy/m/d h:mm"))), 1, 1, Empty
    UseXlsxInput
    gMap = Array(Array("æˆø“ú", "”„ã“ú", "", "“ú•t:yyyymmdd", "", ""), Array("æˆø“ú", "”„ã“ú", "", "“ú•t:yyyy/mm/dd hh:mm", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("æˆø“ú,æˆø“ú", "20261106,2026/11/06 10:30"))
End Sub

Sub T_C5()
    Dim rc
    BeginCase "C5", "CSV‚Ì“ú•t•¶š—ñi•¡”‚Ì‘‚«•ûj‚ğ“ú•t•ÏŠ·"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("”„ã“ú", "2026/11/06", "20261106", "2026-11-06", "2026/11/6", "‚Q‚O‚Q‚U‚P‚P‚O‚U")), CS_SJIS
    gMap = Array(Array("æˆø“ú", "”„ã“ú", "", "“ú•t:yyyymmdd", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("æˆø“ú", "20261106", "20261106", "20261106", "20261106", "20261106"))
End Sub

Sub T_C6()
    Dim rc
    BeginCase "C6", "Excel‚Ì”’l‚Í•\¦Œ`®‚Å‚Í‚È‚­’l‚Å•ÏŠ·"
    MakeXlsx "in\fileA_20261106_01.xlsx", Array(Array("‹àŠz"), Array(Array(1234.5, "#,##0"))), 1, 1, Empty
    UseXlsxInput
    gMap = Array(Array("‹àŠz2Œ…", "‹àŠz", "", "”’l:0.00", "", ""), Array("‹àŠz‚»‚Ì‚Ü‚Ü", "‹àŠz", "", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("‹àŠz2Œ…,‹àŠz‚»‚Ì‚Ü‚Ü", "1234.50,1234.5"))
End Sub

Sub T_C7()
    Dim rc
    BeginCase "C7", "Excel‚ÌƒGƒ‰[ƒZƒ‹i#N/Aj‚Í‚»‚Ì“ú•t‚ğƒGƒ‰[‚É‚·‚é"
    If IsMock() Then
        SkipCheck "ExcelÀ‹@‚Ì‚İiƒGƒ‰[’l‚Íƒ‚ƒbƒN‚ÅÄŒ»‚Å‚«‚È‚¢j"
        Exit Sub
    End If
    MakeXlsx "in\fileA_20261106_01.xlsx", Array(Array("¤•iCD"), Array(Array("=NA()", "FORMULA"))), 1, 1, Empty
    MakeXlsx "in\fileA_20261107_01.xlsx", Array(Array("¤•iCD"), Array("A")), 1, 1, Empty
    UseXlsxInput
    gMap = Array(Array("¤•iƒR[ƒh", "¤•iCD", "", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 1
    ExpectNoFile "out\fileB_20261106.csv"
    ExpectText "out\fileB_20261107.csv", CS_SJIS, L(Array("¤•iƒR[ƒh", "A"))
    ExpectLog "ERROR", "ƒGƒ‰[’l"
    ExpectLog "ERROR", "2s–Ú"
End Sub

Sub T_C8()
    Dim rc
    BeginCase "C8", "•\‚ªA1‚©‚çn‚Ü‚ç‚È‚¢ExcelE2–‡–Ú‚ÌƒV[ƒg‚Í“Ç‚Ü‚È‚¢"
    MakeXlsx "in\fileA_20261106_01.xlsx", Array(Array("¤•iCD", "”—Ê"), Array("A1", 5)), 3, 2, _
             Array(Array("¤•iCD", "”—Ê"), Array("•ÊƒV[ƒg", 99))
    UseXlsxInput
    SetV "ƒwƒbƒ_[s”Ô†", "3"
    gMap = Array(Array("¤•iƒR[ƒh", "¤•iCD", "", "", "", ""), Array("”—Ê", "”—Ê", "", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("¤•iƒR[ƒh,”—Ê", "A1,5"))
End Sub

Sub T_C9()
    Dim rc
    BeginCase "C9", "“¯‚¶“ú•t‚ÉCSV‚ÆExcel‚ª¬İ‚µ‚Ä‚à1ƒtƒ@ƒCƒ‹‚ÉŒ‹‡"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr(), "12,20261106,5,1")), CS_SJIS
    MakeXlsx "in\fileA_20261106_02.xlsx", Array(Array("“X•ÜCD", "”„ã“ú", "”Ì”„”", "‹æ•ª"), Array(7, "20261106", 1, 1)), 1, 1, Empty
    SetV "‘ÎÛƒtƒ@ƒCƒ‹", "fileA_*.*"
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array(OutHdr(), "00012,20261106,5,01,1", "00007,20261106,1,01,1"))
End Sub

Sub T_C10()
    Dim rc
    BeginCase "C10", "Excel‚Ìˆêƒtƒ@ƒCƒ‹i~$j‚Í‘ÎÛŠO"
    MakeXlsx "in\fileA_20261106_01.xlsx", Array(Array("¤•iCD"), Array("A")), 1, 1, Empty
    MakeCsv "in\~$fileA_20261106_01.xlsx", "‰ó‚ê‚½ƒtƒ@ƒCƒ‹", CS_SJIS
    UseXlsxInput
    gMap = Array(Array("¤•iƒR[ƒh", "¤•iCD", "", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("¤•iƒR[ƒh", "A"))
End Sub

Sub T_C11()
    Dim rc, exp, i
    BeginCase "C11", "Excelo—ÍFCSVo—Í‚Æ“¯‚¶’lE‘SƒZƒ‹•¶š—ñEƒV[ƒg–¼EŠg’£qE=‚Ån‚Ü‚é’l"
    MakeStdInputs
    SetV "o—ÍŒ`®", "Excel"
    SetV "o—ÍƒV[ƒg–¼", "˜AŒgƒf[ƒ^"
    gMap = Array(Array("“X•ÜƒR[ƒh", "“X•ÜCD", "", "TRIM|ƒ[ƒ–„‚ß:5", "", ""), Array("æˆø“ú", "”„ã“ú", "", "“ú•t:yyyymmdd", "", ""), _
                 Array("”—Ê", "”Ì”„”", "", "”’l:0", "", ""), Array("‹æ•ª", "‹æ•ª", "", "ƒR[ƒh:‹æ•ª", "", ""), _
                 Array("˜AŒg‹æ•ª", "", "1", "", "", ""), Array("®", "", "=1+1", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    exp = Exp1106()
    For i = 0 To UBound(exp)
        If i = 0 Then exp(i) = exp(i) & ",®" Else exp(i) = exp(i) & ",=1+1"
    Next
    ExpectXlsx "out\fileB_20261106.xlsx", "˜AŒgƒf[ƒ^", RowsOf(exp), True
    ExpectNoFile "out\fileB_20261106.csv"
End Sub

Sub T_C13()
    Dim rc
    BeginCase "C13", "Excel‚Ì13Œ…”’liJANƒR[ƒhj‚ªw”•\‹L‚É‚È‚ç‚È‚¢"
    MakeXlsx "in\fileA_20261106_01.xlsx", Array(Array("JAN"), Array(4912345678901)), 1, 1, Empty
    UseXlsxInput
    gMap = Array(Array("JANƒR[ƒh", "JAN", "", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("JANƒR[ƒh", "4912345678901"))
End Sub

Sub T_C14()
    Dim rc
    BeginCase "C14", "Excel‚Ì”®ƒZƒ‹‚ÍŒvZŒ‹‰Ê‚Å“Ç‚Ş"
    MakeXlsx "in\fileA_20261106_01.xlsx", Array(Array("”—Ê"), Array(Array("=1+2", "FORMULA"))), 1, 1, Empty
    UseXlsxInput
    gMap = Array(Array("”—Ê", "”—Ê", "", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("”—Ê", "3"))
End Sub

Sub T_C15()
    Dim rc
    BeginCase "C15", "•‰‚Ì”E‘SŠp”šE‘SŠpƒJƒ“ƒ}‚Ì”’l•ÏŠ·"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("”—Ê", "-5", "‚P‚Q", """‚PC‚Q‚O‚O""")), CS_SJIS
    gMap = Array(Array("”—Ê", "”—Ê", "", "”’l:0", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("”—Ê", "-5", "12", "1200"))
End Sub

Sub T_D1()
    Dim rc
    BeginCase "D1", "TRIMF‘SŠpƒXƒy[ƒXEƒ^ƒu‚àœ‹"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("“X•ÜCD", "@12" & vbTab)), CS_SJIS
    gMap = Array(Array("“X•ÜƒR[ƒh", "“X•ÜCD", "", "TRIM|ƒ[ƒ–„‚ß:5", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("“X•ÜƒR[ƒh", "00012"))
End Sub

Sub T_D2()
    Dim rc
    BeginCase "D2", "Še•ÏŠ·F’uŠ·E‹ó‚È‚çE¶E‰EEŒ…‹æØ‚èE“ú•t‘®"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("“d˜b,”õl,–¼Ì,‹àŠz,”„ã“ú", "03-1234-5678,,ABCDE,1234567,20261106")), CS_SJIS
    gMap = Array(Array("’uŠ·", "“d˜b", "", "’uŠ·:-=>", "", ""), Array("‹ó‚È‚ç", "”õl", "", "‹ó‚È‚ç:–¢İ’è", "", ""), _
                 Array("¶", "–¼Ì", "", "¶:3", "", ""), Array("‰E", "–¼Ì", "", "‰E:2", "", ""), _
                 Array("Œ…‹æØ‚è", "‹àŠz", "", "”’l:#,##0", "", ""), Array("“ú•t", "”„ã“ú", "", "“ú•t:yyyy-mm-dd", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("’uŠ·,‹ó‚È‚ç,¶,‰E,Œ…‹æØ‚è,“ú•t", "0312345678,–¢İ’è,ABC,DE,""1,234,567"",2026-11-06"))
End Sub

Sub T_D3()
    Dim rc
    BeginCase "D3", "•ÏŠ·‚Ì˜AŒ‹‚Í¶‚©‚ç‡‚É“K—p"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("“d˜b", "03-1234")), CS_SJIS
    gMap = Array(Array("A", "“d˜b", "", "’uŠ·:-=>|¶:4", "", ""), Array("B", "“d˜b", "", "¶:4|’uŠ·:-=>", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("A,B", "0312,031"))
End Sub

Sub T_D4a()
    Dim rc
    BeginCase "D4a", "ƒR[ƒh•ÏŠ·F•\‚É‚È‚¢’l‚Í * ‚ÌŠù’è’l"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("‹æ•ª", "1", "2", "3")), CS_SJIS
    gMap = Array(Array("‹æ•ª", "‹æ•ª", "", "ƒR[ƒh:‹æ•ª", "", ""))
    gCodes = Array(Array("‹æ•ª", "1", "01"), Array("‹æ•ª", "2", "02"), Array("‹æ•ª", "*", "99"))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("‹æ•ª", "01", "02", "99"))
End Sub

Sub T_D4b()
    Dim rc
    BeginCase "D4b", "ƒR[ƒh•ÏŠ·FŠù’è’l‚ª‚È‚­•\‚É‚È‚¢’l‚Í‚»‚Ì“ú•t‚ğƒGƒ‰["
    MakeCsv "in\fileA_20261106_01.csv", L(Array("‹æ•ª", "1", "3")), CS_SJIS
    MakeCsv "in\fileA_20261107_01.csv", L(Array("‹æ•ª", "2")), CS_SJIS
    gMap = Array(Array("‹æ•ª", "‹æ•ª", "", "ƒR[ƒh:‹æ•ª", "", ""))
    rc = RunTool()
    ExpectExit rc, 1
    ExpectNoFile "out\fileB_20261106.csv"
    ExpectText "out\fileB_20261107.csv", CS_SJIS, L(Array("‹æ•ª", "02"))
    ExpectLog "ERROR", "’lu3v‚ª‚ ‚è‚Ü‚¹‚ñ"
End Sub

Sub T_D5()
    Dim rc
    BeginCase "D5", "•s³‚È“ú•tE”’l‚Í‚»‚Ì“ú•t‚¾‚¯ƒGƒ‰[iêŠ‚ğƒƒO‚Éo‚·j"
    MakeCsv "in\fileA_20261101_01.csv", L(Array(StdHdr(), "12,20261332,5,1")), CS_SJIS
    MakeCsv "in\fileA_20261102_01.csv", L(Array(StdHdr(), "12,20261102,abc,1")), CS_SJIS
    MakeCsv "in\fileA_20261103_01.csv", L(Array(StdHdr(), "12,20261103,5,1")), CS_SJIS
    rc = RunTool()
    ExpectExit rc, 1
    ExpectNoFile "out\fileB_20261101.csv"
    ExpectNoFile "out\fileB_20261102.csv"
    ExpectText "out\fileB_20261103.csv", CS_SJIS, L(Array(OutHdr(), "00012,20261103,5,01,1"))
    ExpectLog "ERROR", "“ú•t‚Æ‚µ‚Ä•s³‚Å‚·"
    ExpectLog "ERROR", "”’l‚É•ÏŠ·‚Å‚«‚Ü‚¹‚ñ"
    ExpectLog "ERROR", "2s–Ú / ”—Ê"
End Sub

Sub T_D6()
    Dim rc
    BeginCase "D6", "•s–¾‚È•ÏŠ·–¼‚ÍƒGƒ‰["
    MakeCsv "in\fileA_20261106_01.csv", L(Array("“X•ÜCD", "12")), CS_SJIS
    gMap = Array(Array("“X•Ü", "“X•ÜCD", "", "XYZ:1", "", ""))
    rc = RunTool()
    ExpectExit rc, 1
    ExpectNoFile "out\fileB_20261106.csv"
    ExpectLog "ERROR", "•s–¾‚È•ÏŠ·w’è"
End Sub

Sub T_E1()
    Dim rc
    BeginCase "E1", "€–Ú‚ª‘«‚è‚È‚¢“ú•t‚¾‚¯o—Í‚µ‚È‚¢i‘¼‚Ì“ú•t‚Ío—Íj"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("“X•ÜCD,”„ã“ú,”Ì”„”", "12,20261106,5")), CS_SJIS
    MakeCsv "in\fileA_20261107_01.csv", L(Array(StdHdr(), "12,20261107,3,2")), CS_SJIS
    rc = RunTool()
    ExpectExit rc, 1
    ExpectNoFile "out\fileB_20261106.csv"
    ExpectText "out\fileB_20261107.csv", CS_SJIS, L(Exp1107())
    ExpectLog "ERROR", "fileA‚É€–Ú‚ªŒ©‚Â‚©‚è‚Ü‚¹‚ñ: [‹æ•ª]"
End Sub

Sub T_E2(id, overwrite)
    Dim rc
    If overwrite Then
        BeginCase id, "o—Íæ‚É“¯–¼ƒtƒ@ƒCƒ‹Fã‘‚« ‚ ‚è ¨ ’u‚«Š·‚¦‚é"
        SetV "ã‘‚«", "‚ ‚è"
    Else
        BeginCase id, "o—Íæ‚É“¯–¼ƒtƒ@ƒCƒ‹Fã‘‚« ‚È‚µ ¨ ƒXƒLƒbƒv‚µ‚ÄŠù‘¶ƒtƒ@ƒCƒ‹‚ğ•ÏX‚µ‚È‚¢"
    End If
    MakeStdInputs
    MakeCsv "out\fileB_20261106.csv", "OLD" & vbCrLf, CS_SJIS
    rc = RunTool()
    ExpectExit rc, 0
    If overwrite Then
        ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Exp1106())
    Else
        ExpectText "out\fileB_20261106.csv", CS_SJIS, "OLD" & vbCrLf
        ExpectLog "SKIP", "Šù‚É‚ ‚é"
    End If
    ExpectText "out\fileB_20261107.csv", CS_SJIS, L(Exp1107())
End Sub

Sub T_E3()
    Dim rc
    BeginCase "E3", "ƒtƒ@ƒCƒ‹–¼‚©‚ç“ú•t‚ªæ‚ê‚È‚¢ƒtƒ@ƒCƒ‹‚ÍSKIP"
    MakeStdInputs
    MakeCsv "in\fileA_latest.csv", L(Array(StdHdr(), "1,20261106,1,1")), CS_SJIS
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Exp1106())
    ExpectLog "SKIP", "fileA_latest.csv"
End Sub

Sub T_E4()
    Dim rc
    BeginCase "E4", "‘ÎÛƒtƒ@ƒCƒ‹0Œ‚ÍƒGƒ‰[I—¹"
    rc = RunTool()
    ExpectExit rc, 1
    ExpectLog "ERROR", "‘ÎÛƒtƒ@ƒCƒ‹‚ªŒ©‚Â‚©‚è‚Ü‚¹‚ñ"
End Sub

Sub T_E5a()
    Dim rc
    BeginCase "E5a", "İ’è•s”õF“ü—ÍƒtƒHƒ‹ƒ_‚ª‚È‚¢"
    SetV "“ü—ÍƒtƒHƒ‹ƒ_", gFso.BuildPath(gDir, "nothing")
    rc = RunTool()
    ExpectExit rc, 1
End Sub

Sub T_E5b()
    Dim rc
    BeginCase "E5b", "İ’è•s”õF“ü—Í‚Æo—Í‚ª“¯‚¶ƒtƒHƒ‹ƒ_"
    MakeStdInputs
    SetV "o—ÍƒtƒHƒ‹ƒ_", gFso.BuildPath(gDir, "in")
    rc = RunTool()
    ExpectExit rc, 1
    ExpectNoFile "in\fileB_20261106.csv"
End Sub

Sub T_E5c()
    Dim rc
    BeginCase "E5c", "İ’è•s”õFİ’èƒL[‚ÌŒ‡—"
    MakeStdInputs
    gRemoveKeys = Array("ã‘‚«")
    rc = RunTool()
    ExpectExit rc, 1
    ExpectNoFile "out\fileB_20261106.csv"
End Sub

Sub T_E5d()
    Dim rc
    BeginCase "E5d", "İ’è•s”õF”’l‚Ì—“‚É•¶š"
    MakeStdInputs
    SetV "ƒwƒbƒ_[s”Ô†", "abc"
    rc = RunTool()
    ExpectExit rc, 1
End Sub

Sub T_E5e()
    Dim rc
    BeginCase "E5e", "İ’è•s”õFo—ÍŒ`®‚ª•s³"
    MakeStdInputs
    SetV "o—ÍŒ`®", "PDF"
    rc = RunTool()
    ExpectExit rc, 1
End Sub

Sub T_E6a()
    Dim rc
    BeginCase "E6a", "0ƒoƒCƒg‚Ìƒtƒ@ƒCƒ‹‚Í€–Ú•s‘«‚ÅƒGƒ‰["
    MakeCsv "in\fileA_20261106_01.csv", "", CS_SJIS
    rc = RunTool()
    ExpectExit rc, 1
    ExpectNoFile "out\fileB_20261106.csv"
    ExpectLog "ERROR", "Œ©‚Â‚©‚è‚Ü‚¹‚ñ"
End Sub

Sub T_E6b()
    Dim rc
    BeginCase "E6b", "Œ©o‚µ‚¾‚¯‚Ìƒtƒ@ƒCƒ‹‚Í0siŒ©o‚µ‚Ì‚İo—Í{WARNj"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr())), CS_SJIS
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array(OutHdr()))
    ExpectLog "WARN", "0s"
End Sub

Sub T_E7()
    Dim rc
    BeginCase "E7", "fileA‚É“¯‚¶€–Ú–¼‚ª2‚Â‚ ‚ê‚ÎƒGƒ‰["
    MakeCsv "in\fileA_20261106_01.csv", L(Array("“X•ÜCD,”„ã“ú,”Ì”„”,‹æ•ª,“X•ÜCD", "12,20261106,5,1,99")), CS_SJIS
    rc = RunTool()
    ExpectExit rc, 1
    ExpectLog "ERROR", "“¯‚¶€–Ú–¼‚ª•¡”‚ ‚è‚Ü‚·: [“X•ÜCD]"
End Sub

Sub T_E8a()
    Dim rc
    BeginCase "E8a", "Œ©o‚µ‚Ì‘OŒã‚Ì‹ó”’‚Í–³‹"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(" “X•ÜCD ,”„ã“ú,”Ì”„”,‹æ•ª", "12,20261106,5,1")), CS_SJIS
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array(OutHdr(), "00012,20261106,5,01,1"))
End Sub

Sub T_E8b()
    Dim rc
    BeginCase "E8b", "‘SŠpE”¼Šp‚Ìˆá‚¤Œ©o‚µ‚Í•Ê‚Ì€–ÚiƒGƒ‰[j"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("“X•Ü‚b‚c,”„ã“ú,”Ì”„”,‹æ•ª", "12,20261106,5,1")), CS_SJIS
    rc = RunTool()
    ExpectExit rc, 1
    ExpectLog "ERROR", "[“X•ÜCD]"
End Sub

Sub T_E9()
    Dim rc, wb, p
    BeginCase "E9", "o—Íæƒtƒ@ƒCƒ‹‚ªŠJ‚©‚ê‚Ä‚¢‚é“ú•t‚ÍƒGƒ‰[i‘¼‚Í‘±sj"
    If IsMock() Then
        SkipCheck "ExcelÀ‹@‚Ì‚İiƒtƒ@ƒCƒ‹‚ÌƒƒbƒNj"
        Exit Sub
    End If
    MakeStdInputs
    SetV "o—ÍŒ`®", "Excel"
    SetV "ã‘‚«", "‚ ‚è"
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
    BeginCase "E10", "––”ö‚É‰üs‚È‚µEEOF•¶ši0x1Aj‚ª‚ ‚Á‚Ä‚à³‚µ‚­“Ç‚Ş"
    MakeCsv "in\fileA_20261106_01.csv", StdHdr() & vbCrLf & "12,20261106,5,1" & Chr(26), CS_SJIS
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array(OutHdr(), "00012,20261106,5,01,1"))
End Sub

Sub T_E11a()
    Dim rc
    BeginCase "E11a", "Excelo—ÍFƒV[ƒg–¼‚Ég‚¦‚È‚¢•¶š‚Íİ’èƒGƒ‰["
    MakeStdInputs
    SetV "o—ÍŒ`®", "Excel"
    SetV "o—ÍƒV[ƒg–¼", "a/b"
    rc = RunTool()
    ExpectExit rc, 1
    ExpectNoFile "out\fileB_20261106.xlsx"
End Sub

Sub T_E11b()
    Dim rc
    BeginCase "E11b", "Excelo—ÍFƒZƒ‹‚ÌÅ‘å•¶š”‚ğ’´‚¦‚é’l‚Í‚»‚Ì“ú•t‚ğƒGƒ‰["
    MakeCsv "in\fileA_20261106_01.csv", L(Array("”õl", String(32768, "x"))), CS_SJIS
    MakeCsv "in\fileA_20261107_01.csv", L(Array("”õl", "ok")), CS_SJIS
    SetV "o—ÍŒ`®", "Excel"
    gMap = Array(Array("”õl", "”õl", "", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 1
    ExpectNoFile "out\fileB_20261106.xlsx"
    ExpectXlsx "out\fileB_20261107.xlsx", "Sheet1", RowsOf(Array("”õl", "ok")), True
    ExpectLog "ERROR", "Å‘å•¶š”"
End Sub

Sub T_E12()
    Dim rc
    BeginCase "E12", "ƒf[ƒ^0s‚Ì“ú•t‚ğo—Í‚µ‚È‚¢İ’è"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr())), CS_SJIS
    SetV "ƒf[ƒ^0s‚Ì“ú•t‚ào—Í", "‚È‚µ"
    rc = RunTool()
    ExpectExit rc, 0
    ExpectNoFile "out\fileB_20261106.csv"
    ExpectLog "SKIP", "0s"
End Sub

Sub T_E13()
    Dim rc
    BeginCase "E13", "•Â‚¶‚Ä‚¢‚È‚¢ƒ_ƒuƒ‹ƒNƒH[ƒg‚ÍƒGƒ‰["
    MakeCsv "in\fileA_20261106_01.csv", L(Array("”õl", """abc")), CS_SJIS
    gMap = Array(Array("”õl", "”õl", "", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 1
    ExpectNoFile "out\fileB_20261106.csv"
    ExpectLog "ERROR", "ƒ_ƒuƒ‹ƒNƒH[ƒg‚ª•Â‚¶‚ç‚ê‚Ä‚¢‚Ü‚¹‚ñ"
End Sub

Sub T_F2()
    Dim rc, tool
    BeginCase "F2", "ˆø”‚È‚µÀsFƒXƒNƒŠƒvƒg‚Æ“¯–¼‚Ìİ’èƒtƒ@ƒCƒ‹‚ğg‚¤"
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
    BeginCase "F5", "ƒpƒX‚ÉƒXƒy[ƒXE“ú–{Œê‚ğŠÜ‚Ş"
    MakeCsv "“ü—Í ƒtƒHƒ‹ƒ_\fileA_20261106_01.csv", L(Array(StdHdr(), "12,20261106,5,1")), CS_SJIS
    EnsureDir "o—Í ƒtƒHƒ‹ƒ_"
    SetV "“ü—ÍƒtƒHƒ‹ƒ_", gFso.BuildPath(gDir, "“ü—Í ƒtƒHƒ‹ƒ_")
    SetV "o—ÍƒtƒHƒ‹ƒ_", gFso.BuildPath(gDir, "o—Í ƒtƒHƒ‹ƒ_")
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "o—Í ƒtƒHƒ‹ƒ_\fileB_20261106.csv", CS_SJIS, L(Array(OutHdr(), "00012,20261106,5,01,1"))
End Sub

Sub T_G1a()
    Dim rc
    BeginCase "G1a", "í‚É‹ó‚Ì—ñi“r’†jF—ñ‚ª‚¸‚ê‚¸ ,, ‚É‚È‚é{WARN"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr(), "12,20261106,5,1")), CS_SJIS
    gMap = Array(Array("“X•Ü", "“X•ÜCD", "", "", "", ""), Array("”õl", "", "", "", "", ""), Array("”—Ê", "”Ì”„”", "", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("“X•Ü,”õl,”—Ê", "12,,5"))
    ExpectLog "WARN", "[”õl]"
End Sub

Sub T_G1b()
    Dim rc
    BeginCase "G1b", "í‚É‹ó‚Ì—ñ{‘S€–ÚˆÍ‚ŞF‹ó‚Í """""
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr(), "12,20261106,5,1")), CS_SJIS
    gMap = Array(Array("“X•Ü", "“X•ÜCD", "", "", "", ""), Array("”õl", "", "", "", "", ""), Array("”—Ê", "”Ì”„”", "", "", "", ""))
    SetV "‘S€–Ú‚ğƒ_ƒuƒ‹ƒNƒH[ƒg‚ÅˆÍ‚Ş", "‚ ‚è"
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("""“X•Ü"",""”õl"",""”—Ê""", """12"","""",""5"""))
End Sub

Sub T_G1c()
    Dim rc
    BeginCase "G1c", "í‚É‹ó‚Ì—ñ{‘S€–ÚˆÍ‚Ş{‹ó‚ÍˆÍ‚Ü‚È‚¢F,, ‚É‚È‚é"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr(), "12,20261106,5,1")), CS_SJIS
    gMap = Array(Array("“X•Ü", "“X•ÜCD", "", "", "", ""), Array("”õl", "", "", "", "", ""), Array("”—Ê", "”Ì”„”", "", "", "", ""))
    SetV "‘S€–Ú‚ğƒ_ƒuƒ‹ƒNƒH[ƒg‚ÅˆÍ‚Ş", "‚ ‚è"
    SetV "‹ó‚Ì’l‚àƒ_ƒuƒ‹ƒNƒH[ƒg‚ÅˆÍ‚Ş", "‚È‚µ"
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("""“X•Ü"",""”õl"",""”—Ê""", """12"",,""5"""))
End Sub

Sub T_G1d()
    Dim rc
    BeginCase "G1d", "í‚É‹ó‚Ì—ñ{Excelo—ÍF‹ó‚ÌƒZƒ‹"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr(), "12,20261106,5,1")), CS_SJIS
    gMap = Array(Array("“X•Ü", "“X•ÜCD", "", "", "", ""), Array("”õl", "", "", "", "", ""), Array("”—Ê", "”Ì”„”", "", "", "", ""))
    SetV "o—ÍŒ`®", "Excel"
    rc = RunTool()
    ExpectExit rc, 0
    ExpectXlsx "out\fileB_20261106.xlsx", "Sheet1", RowsOf(Array("“X•Ü,”õl,”—Ê", "12,,5")), True
End Sub

Sub T_G2()
    Dim rc
    BeginCase "G2", "ˆê•”‚Ìs‚¾‚¯‹ó‚Ì’lF‹ó‚Ì‚Ü‚Ü^‹ó‚È‚ç ‚Å•âŠ®"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(BikoHdr(), "12,20261106,5,1,‚ ‚è", "13,20261106,5,1,")), CS_SJIS
    gMap = Array(Array("“X•Ü", "“X•ÜCD", "", "", "", ""), Array("”õl", "”õl", "", "", "", ""), Array("”õl2", "”õl", "", "‹ó‚È‚ç:‚È‚µ", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("“X•Ü,”õl,”õl2", "12,‚ ‚è,‚ ‚è", "13,,‚È‚µ"))
End Sub

Sub T_G3()
    Dim rc
    BeginCase "G3", "æ“ªE––”ö‚Ì—ñ‚ª‹ó‚Å‚à—ñ”‚ª‚¸‚ê‚È‚¢"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr(), "12,20261106,5,1")), CS_SJIS
    gMap = Array(Array("æ“ª‹ó", "", "", "", "", ""), Array("“X•Ü", "“X•ÜCD", "", "", "", ""), Array("––”ö‹ó", "", "", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("æ“ª‹ó,“X•Ü,––”ö‹ó", ",12,"))
    ExpectLog "WARN", "[æ“ª‹ó] [––”ö‹ó]"
End Sub

' ’ŠoğŒE˜AŒ‹EƒŒƒR[ƒhEƒtƒbƒ^[
Function FilHdr(): FilHdr = "“X•ÜCD,‹æ•ª,”õl": End Function
Sub UseFilMap()
    gMap = Array(Array("“X•Ü", "“X•ÜCD", "", "", "", ""), Array("‹æ•ª", "‹æ•ª", "", "", "", ""))
End Sub

Sub T_H1()
    Dim rc
    BeginCase "H1", "’ŠoğŒFŠÜ‚ß‚éi‹æ•ª = 1 ‚Ìs‚¾‚¯j"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(FilHdr(), "12,1,", "13,2,", "14,1,")), CS_SJIS
    UseFilMap
    gFilters = Array(Array("ŠÜ‚ß‚é", "‹æ•ª", "“™‚µ‚¢", "1"))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("“X•Ü,‹æ•ª", "12,1", "14,1"))
    ExpectLog "INFO", "1 sœŠO"
End Sub

Sub T_H2()
    Dim rc
    BeginCase "H2", "’ŠoğŒFœŠO‚·‚éi‚Ç‚ê‚©1‚Â‚É“–‚Ä‚Í‚Ü‚ê‚ÎœŠOj"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(FilHdr(), "12,1,", ",1,", "13,2,‡Œvs", "14,2,")), CS_SJIS
    UseFilMap
    gFilters = Array(Array("œŠO‚·‚é", "“X•ÜCD", "‹ó", ""), Array("œŠO‚·‚é", "”õl", "ŠÜ‚Ş", "‡Œv"))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("“X•Ü,‹æ•ª", "12,1", "14,2"))
End Sub

Sub T_H3()
    Dim rc
    BeginCase "H3", "’ŠoğŒF³‹K•\Œ»‚ÆŠÜ‚ß‚éEœŠO‚Ì‘g‚İ‡‚í‚¹"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(FilHdr(), "12,1,", "A3,1,", "14,2,", "15,1,ƒeƒXƒg")), CS_SJIS
    UseFilMap
    gFilters = Array(Array("ŠÜ‚ß‚é", "“X•ÜCD", "³‹K•\Œ»", "^\d+$"), Array("ŠÜ‚ß‚é", "‹æ•ª", "“™‚µ‚­‚È‚¢", "2"), Array("œŠO‚·‚é", "”õl", "‹ó‚Å‚È‚¢", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("“X•Ü,‹æ•ª", "12,1"))
End Sub

Sub T_H4()
    Dim rc
    BeginCase "H4", "•¡”€–Ú‚Ì˜AŒ‹i‹ó‚Ì€–Ú‚Í‹l‚ß‚éj"
    MakeCsv "in\fileA_20261106_01.csv", L(Array("©,–¼", "R“c,‘¾˜Y", "²“¡,", ",‰Ôq")), CS_SJIS
    gMap = Array(Array("–¼", "©|–¼", "", "", "@", ""), Array("–¼‘±‚¯", "©|–¼", "", "", "", ""))
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("–¼,–¼‘±‚¯", "R“c@‘¾˜Y,R“c‘¾˜Y", "²“¡,²“¡", "‰Ôq,‰Ôq"))
End Sub

Sub T_H5a()
    Dim rc
    BeginCase "H5a", "æ“ªE––”öƒŒƒR[ƒhi“ú•tEŒ”E‡ŒvjFCSV"
    MakeStdInputs
    SetV "æ“ªƒŒƒR[ƒh", "H,{date}"
    SetV "––”öƒŒƒR[ƒh", "T,{count},{sum:”—Ê}"
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Array("H,20261106", OutHdr(), "00012,20261106,5,01,1", "00345,20261106,10,02,1", "00007,20261106,1,01,1", "T,3,16"))
End Sub

Sub T_H5b()
    Dim rc
    BeginCase "H5b", "æ“ªE––”öƒŒƒR[ƒhFExcelo—Í‚ÍƒZƒ‹‚É•ªŠ„"
    MakeStdInputs
    SetV "o—ÍŒ`®", "Excel"
    SetV "æ“ªƒŒƒR[ƒh", "H,{date}"
    SetV "––”öƒŒƒR[ƒh", "T,{count},{sum:”—Ê}"
    rc = RunTool()
    ExpectExit rc, 0
    ExpectXlsx "out\fileB_20261107.xlsx", "Sheet1", RowsOf(Array("H,20261107", OutHdr(), "00012,20261107,3,02,1", "T,1,3")), True
End Sub

Sub T_H6()
    Dim rc
    BeginCase "H6", "fileA––”ö‚Ì‡Œvs‚ğœŠOiŒã‚ë‚Ì‹ós‚Í”‚¦‚È‚¢j"
    MakeCsv "in\fileA_20261106_01.csv", L(Array(StdHdr(), "12,20261106,5,1", "345,20261106,10,2", "‡Œv,,15,", "", "")), CS_SJIS
    MakeCsv "in\fileA_20261106_02.csv", L(Array(StdHdr(), "7,20261106,1,1", "‡Œv,,1,")), CS_SJIS
    SetV "fileA––”ö‚ÌœŠOs”", "1"
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261106.csv", CS_SJIS, L(Exp1106())
End Sub

Sub T_H7()
    Dim rc
    BeginCase "H7", "ƒtƒ@ƒCƒ‹––”ö‚Ì‰üs‚È‚µ"
    MakeStdInputs
    SetV "ƒtƒ@ƒCƒ‹––”ö‚É‰üs", "‚È‚µ"
    rc = RunTool()
    ExpectExit rc, 0
    ExpectText "out\fileB_20261107.csv", CS_SJIS, Join(Exp1107(), vbCrLf)
End Sub

Sub T_H8()
    Dim rc
    BeginCase "H8", "’ŠoğŒ‚Ì€–Ú‚ªfileA‚É‚È‚¢ê‡‚ÍƒGƒ‰["
    MakeCsv "in\fileA_20261106_01.csv", L(Array(FilHdr(), "12,1,")), CS_SJIS
    UseFilMap
    gFilters = Array(Array("ŠÜ‚ß‚é", "‘¶İ‚µ‚È‚¢€–Ú", "“™‚µ‚¢", "1"))
    rc = RunTool()
    ExpectExit rc, 1
    ExpectNoFile "out\fileB_20261106.csv"
    ExpectLog "ERROR", "[‘¶İ‚µ‚È‚¢€–Ú]"
End Sub

Sub T_P1()
    Dim rc, i, lines, t0, n
    n = 100000
    If IsMock() Then n = 2000
    BeginCase "P1", "‘å—Êƒf[ƒ^i" & n & " sj‚Ìˆ—ŠÔ"
    ReDim lines(n)
    lines(0) = StdHdr()
    For i = 1 To n
        lines(i) = i & ",20261106," & (i Mod 100) & "," & (1 + (i Mod 2))
    Next
    MakeCsv "in\fileA_20261106_01.csv", L(lines), CS_SJIS
    t0 = Timer
    rc = RunTool()
    ExpectExit rc, 0
    ExpectLog "OK", n & " s o—Í"
    Check "ˆ—ŠÔiQlj", True, Round(Timer - t0, 1) & " •b"
End Sub


'==============================================================
' ƒeƒXƒg‚Ì•”•i
'==============================================================
Sub BeginCase(id, title)
    gCase = id
    gTitle = title
    gDir = gFso.BuildPath(gWork, id)
    gFso.CreateFolder gDir
    EnsureDir "in"
    EnsureDir "out"
    EnsureDir "log"
    Set gSet = NewObject("Scripting.Dictionary")
    gSet("“ü—ÍƒtƒHƒ‹ƒ_") = gFso.BuildPath(gDir, "in")
    gSet("‘ÎÛƒtƒ@ƒCƒ‹") = "*.csv"
    gSet("“ú•t’Šoƒpƒ^[ƒ“") = "(\d{8})"
    gSet("“ü—Í•¶šƒR[ƒh") = CS_SJIS
    gSet("ƒwƒbƒ_[s”Ô†") = "1"
    gSet("fileA––”ö‚ÌœŠOs”") = "0"
    gSet("‹æØ‚è•¶š") = ","
    gSet("o—ÍƒtƒHƒ‹ƒ_") = gFso.BuildPath(gDir, "out")
    gSet("o—Íƒtƒ@ƒCƒ‹–¼") = "fileB_{date}.csv"
    gSet("o—ÍŒ`®") = "CSV"
    gSet("o—ÍƒV[ƒg–¼") = "Sheet1"
    gSet("o—Í•¶šƒR[ƒh") = CS_SJIS
    gSet("o—Íƒwƒbƒ_[") = "‚ ‚è"
    gSet("‘S€–Ú‚ğƒ_ƒuƒ‹ƒNƒH[ƒg‚ÅˆÍ‚Ş") = "‚È‚µ"
    gSet("‹ó‚Ì’l‚àƒ_ƒuƒ‹ƒNƒH[ƒg‚ÅˆÍ‚Ş") = "‚ ‚è"
    gSet("‰üsƒR[ƒh") = "CRLF"
    gSet("ƒtƒ@ƒCƒ‹––”ö‚É‰üs") = "‚ ‚è"
    gSet("æ“ªƒŒƒR[ƒh") = ""
    gSet("––”öƒŒƒR[ƒh") = ""
    gSet("ƒf[ƒ^0s‚Ì“ú•t‚ào—Í") = "‚ ‚è"
    gSet("ã‘‚«") = "‚È‚µ"
    gSet("ƒƒOƒtƒHƒ‹ƒ_") = gFso.BuildPath(gDir, "log")
    Set gSys = NewObject("Scripting.Dictionary")
    gMap = Array(Array("“X•ÜƒR[ƒh", "“X•ÜCD", "", "TRIM|ƒ[ƒ–„‚ß:5", "", ""), _
                 Array("æˆø“ú", "”„ã“ú", "", "“ú•t:yyyymmdd", "", ""), _
                 Array("”—Ê", "”Ì”„”", "", "”’l:0", "", ""), _
                 Array("‹æ•ª", "‹æ•ª", "", "ƒR[ƒh:‹æ•ª", "", ""), _
                 Array("˜AŒg‹æ•ª", "", "1", "", "", ""))
    gCodes = Array(Array("‹æ•ª", "1", "01"), Array("‹æ•ª", "2", "02"))
    gFilters = Array()
    gRemoveKeys = Array()
End Sub

Sub SetV(k, v)
    gSet(k) = v
End Sub

Sub EnsureDir(rel)
    Dim p
    p = gFso.BuildPath(gDir, rel)
    If Not gFso.FolderExists(p) Then gFso.CreateFolder p
End Sub

' s‚Ì”z—ñ ¨ CRLF ‹æØ‚è{––”ö‰üs
Function L(arr)
    L = Join(arr, vbCrLf) & vbCrLf
End Function

' "a,b" Œ`®‚Ìs‚Ì”z—ñ ¨ 2ŸŒ³‚ÌŠú‘Ò’li”z—ñ‚Ì”z—ñj
Function RowsOf(lines)
    Dim r, i
    ReDim r(UBound(lines))
    For i = 0 To UBound(lines)
        r(i) = Split(lines(i), ",")
    Next
    RowsOf = r
End Function

Sub MakeCsv(rel, txt, cs)
    Dim p
    p = gFso.BuildPath(gDir, rel)
    If Not gFso.FolderExists(gFso.GetParentFolderName(p)) Then gFso.CreateFolder gFso.GetParentFolderName(p)
    WriteTextFile p, txt, cs
End Sub

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

' Excel ‚Ì“ü—Íƒtƒ@ƒCƒ‹‚ğì‚é
'   rows : s‚Ì”z—ñB’l‚Í •¶š—ñi•¶š—ñƒZƒ‹j/ ”’l / “ú•t / Array(’l, •\¦Œ`®) / Array("=®", "FORMULA")
'   r0,c0: •\‚Ì¶ã‚ÌˆÊ’u@other: 2–‡–Ú‚ÌƒV[ƒg‚Ì“à—eiEmpty ‚È‚çì‚ç‚È‚¢j
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

' İ’èƒtƒ@ƒCƒ‹‚ğì‚éi‚Ğ‚ÈŒ`‚ğƒRƒs[‚µ‚Ä’l‚ğ‘‚«Š·‚¦‚éj
Sub WriteConfig(path)
    Dim wb, k, i
    gFso.CopyFile gTemplate, path, True
    Set wb = gXL.Workbooks.Open(path, XL_NO_UPDATE_LINKS, False)
    For Each k In gSet.Keys
        SetKeyValue wb.Worksheets("İ’è"), k, gSet(k)
    Next
    For Each k In gSys.Keys
        SetKeyValue wb.Worksheets("ƒVƒXƒeƒ€İ’è"), k, gSys(k)
    Next
    For i = 0 To UBound(gRemoveKeys)
        RemoveKey wb.Worksheets("İ’è"), gRemoveKeys(i)
    Next
    FillTable wb.Worksheets("ƒ}ƒbƒsƒ“ƒO"), gMap, MAP_COLS
    FillTable wb.Worksheets("ƒR[ƒh•ÏŠ·"), gCodes, SIMPLE_COLS
    FillTable wb.Worksheets("’ŠoğŒ"), gFilters, SIMPLE_COLS
    wb.Save
    wb.Close False
End Sub

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

Sub SetKeyValue(ws, k, v)
    Dim r
    r = FindKeyRow(ws, k)
    If r = 0 Then Err.Raise vbObjectError + 1, "RunTests", "‚Ğ‚ÈŒ`‚Éİ’èƒL[‚ª‚ ‚è‚Ü‚¹‚ñ: " & k
    ws.Cells(r, 2).NumberFormat = "@"
    ws.Cells(r, 2).Value = v
End Sub

Sub RemoveKey(ws, k)
    Dim r
    r = FindKeyRow(ws, k)
    If r > 0 Then
        ws.Cells(r, 1).Value = Empty
        ws.Cells(r, 2).Value = Empty
    End If
End Sub

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

' ƒc[ƒ‹‚ğÀs‚µ‚ÄI—¹ƒR[ƒh‚ğ•Ô‚·iExcel ƒvƒƒZƒX‚Ìc‚è‚àŠm”Fj
Function RunTool()
    Dim cfg, before
    cfg = gFso.BuildPath(gDir, "config.xlsx")
    WriteConfig cfg
    before = ExcelCount()
    RunTool = gSh.Run("cscript //nologo """ & gTool & """ """ & cfg & """", 0, True)
    If Not IsMock() Then
        If WaitExcelCount(before) Then
            Check "Excel‚ÌƒvƒƒZƒX‚ªc‚ç‚È‚¢", True, ""
        Else
            Check "Excel‚ÌƒvƒƒZƒX‚ªc‚ç‚È‚¢", False, "Às‘O " & before & " ŒÂ ¨ ÀsŒã " & ExcelCount() & " ŒÂ"
        End If
    End If
End Function

Function ExcelCount()
    If IsMock() Then
        ExcelCount = 0
    Else
        ExcelCount = GetObject("winmgmts:\\.\root\cimv2").ExecQuery("Select * From Win32_Process Where Name = 'EXCEL.EXE'").Count
    End If
End Function

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
' ŒŸØ
'==============================================================
Sub ExpectExit(rc, expected)
    Check "I—¹ƒR[ƒh " & expected, (rc = expected), "ÀÛ: " & rc
End Sub

Sub ExpectText(rel, cs, expected)
    Dim p, actual
    p = gFso.BuildPath(gDir, rel)
    If Not gFso.FileExists(p) Then
        Check "o—Í“à—e " & rel, False, "ƒtƒ@ƒCƒ‹‚ª‚ ‚è‚Ü‚¹‚ñ"
        Exit Sub
    End If
    actual = ReadTextFile(p, cs)
    If actual = expected Then
        Check "o—Í“à—e " & rel, True, ""
    Else
        Check "o—Í“à—e " & rel, False, DiffText(expected, actual)
    End If
End Sub

Sub ExpectHex(rel, hexPrefix)
    Dim p, h
    If IsMock() Then
        SkipCheck "ƒoƒCƒgŠm”F " & rel & "iWindowsÀ‹@‚Ì‚İj"
        Exit Sub
    End If
    p = gFso.BuildPath(gDir, rel)
    h = ReadHex(p, Len(hexPrefix) / 2)
    Check "æ“ªƒoƒCƒg " & rel, (h = hexPrefix), "Šú‘Ò: " & hexPrefix & " / ÀÛ: " & h
End Sub

Sub ExpectLogHex(hexPrefix)
    Dim p
    If IsMock() Then
        SkipCheck "ƒƒO‚ÌBOMiWindowsÀ‹@‚Ì‚İj"
        Exit Sub
    End If
    p = LogPath()
    If p = "" Then
        Check "ƒƒO‚ÌBOM", False, "ƒƒO‚ª‚ ‚è‚Ü‚¹‚ñ"
    Else
        Check "ƒƒO‚ÌBOM", (ReadHex(p, Len(hexPrefix) / 2) = hexPrefix), ""
    End If
End Sub

Sub ExpectNoFile(rel)
    Check "o—Í‚³‚ê‚È‚¢ " & rel, Not gFso.FileExists(gFso.BuildPath(gDir, rel)), "ƒtƒ@ƒCƒ‹‚ªì‚ç‚ê‚Ä‚¢‚Ü‚·"
End Sub

Sub ExpectXlsx(rel, sheetName, rows, requireText)
    Dim p, wb, ws, r, c, v, s, e, ok, detail, lastR, lastC, nCols
    p = gFso.BuildPath(gDir, rel)
    If Not gFso.FileExists(p) Then
        Check "Excelo—Í " & rel, False, "ƒtƒ@ƒCƒ‹‚ª‚ ‚è‚Ü‚¹‚ñ"
        Exit Sub
    End If
    Set wb = gXL.Workbooks.Open(p, XL_NO_UPDATE_LINKS, True)
    Set ws = wb.Worksheets(1)
    Check "ƒV[ƒg–¼ " & sheetName, (ws.Name = sheetName), "ÀÛ: " & ws.Name
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
                If detail = "" Then detail = (r + 1) & "s" & (c + 1) & "—ñ Šú‘Ò[" & e & "] ÀÛ[" & s & "]"
            ElseIf requireText And e <> "" Then
                If VarType(v) <> vbString Or ws.Cells(r + 1, c + 1).NumberFormat <> "@" Then
                    ok = False
                    If detail = "" Then detail = (r + 1) & "s" & (c + 1) & "—ñ ‚ª•¶š—ñƒZƒ‹‚Å‚Í‚ ‚è‚Ü‚¹‚ñ"
                End If
            End If
        Next
    Next
    lastR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    lastC = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If lastR > UBound(rows) + 1 Or lastC > nCols Then
        For r = 1 To lastR
            For c = 1 To lastC
                If r > UBound(rows) + 1 Or c > nCols Then
                    If Not IsEmpty(ws.Cells(r, c).Value) Then
                        ok = False
                        If detail = "" Then detail = "—]•ª‚Èƒf[ƒ^: " & r & "s" & c & "—ñ"
                    End If
                End If
            Next
        Next
    End If
    wb.Close False
    Check "Excelo—Í " & rel, ok, detail
End Sub

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
        Check "ƒƒO " & kind & "u" & text & "v", False, "ƒƒOƒtƒ@ƒCƒ‹‚ª‚ ‚è‚Ü‚¹‚ñ"
        Exit Sub
    End If
    lines = Split(ReadTextFile(p, CS_UTF8), vbCrLf)
    found = False
    For i = 0 To UBound(lines)
        If InStr(lines(i), "," & kind & ",") > 0 And InStr(lines(i), text) > 0 Then found = True
    Next
    Check "ƒƒO " & kind & "u" & text & "v", found, "ŠY“–‚·‚és‚ª‚ ‚è‚Ü‚¹‚ñ"
End Sub

Function DiffText(expected, actual)
    Dim i, n
    n = Len(expected)
    If Len(actual) < n Then n = Len(actual)
    For i = 1 To n
        If Mid(expected, i, 1) <> Mid(actual, i, 1) Then Exit For
    Next
    DiffText = i & "•¶š–Ú‚©‚çˆá‚¢‚Ü‚·@Šú‘Ò[" & Visible(Mid(expected, i, 30)) & "] ÀÛ[" & Visible(Mid(actual, i, 30)) & "]"
End Function

Function Visible(s)
    Visible = Replace(Replace(Replace(s, vbCr, "\r"), vbLf, "\n"), vbTab, "\t")
End Function


'==============================================================
' Œ‹‰Ê‚Ì‹L˜^
'==============================================================
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

Sub SkipCheck(name)
    gSkip = gSkip + 1
    AddResult gCase, gTitle, name, "SKIP", ""
End Sub

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

Sub WriteResults()
    Dim lines, i
    ReDim lines(gResN - 1)
    For i = 0 To gResN - 1
        lines(i) = gResults(i)
    Next
    WriteTextFile gFso.BuildPath(gBase, RESULT_FILE), Join(lines, vbCrLf) & vbCrLf, CS_UTF8BOM
End Sub
