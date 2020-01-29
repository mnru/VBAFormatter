Attribute VB_Name = "MainModule"
Option Explicit
Private tabCount     As Integer
Private cBeginList    As New Collection
Private cEndList     As New Collection
Private cBeginDef    As New Collection
Private clsVBAFormatMenu As VBAFormatMenu
Private cBLList     As New Collection


Sub addButton()
    Dim wVBAFormatterMenu As CommandBarControl
    Dim wOptionMenu    As CommandBarControl
    Dim wFormatExecMenu  As CommandBarControl
    Set wVBAFormatterMenu = Application.VBE.CommandBars("メニュー バー").Controls.Add(Type:=msoControlPopup, ID:=1)
    wVBAFormatterMenu.Caption = "VBAFormatter(&Z)"
    Set wFormatExecMenu = wVBAFormatterMenu.Controls.Add(Type:=msoControlButton)
    Set wOptionMenu = wVBAFormatterMenu.Controls.Add(Type:=msoControlButton)
    wFormatExecMenu.Caption = "フォーマット実行(&F)..."
    wOptionMenu.Caption = "オプション(&O)..."
    Set clsVBAFormatMenu = New VBAFormatMenu
    Call clsVBAFormatMenu.InitializeInstance(wFormatExecMenu, wOptionMenu)
End Sub


Sub dellButton()
    Dim wCtrl As CommandBarControl
    Set clsVBAFormatMenu = Nothing
    For Each wCtrl In Application.VBE.CommandBars("メニュー バー").Controls
        If wCtrl.ID = 1 Then
            wCtrl.Delete
            Exit Sub
        End If
    Next
End Sub


Function indentEdit(prev, this) As String
    Dim space As String
    If (Trim(Left(this, InStr(this, " ")) = "End Select") Or this = "End Select" Or Trim(Left(this, InStr(InStr(this, " ") + 1, this, " "))) = "End Select") Then
        tabCount = tabCount - (2 * cIniKeyList.aTabCnt)
    ElseIf (isMemberOfCollection(cEndList, Trim(Left(this, InStr(this, " ")))) Or isMemberOfCollection(cEndList, this)) And this <> "End" Then
        tabCount = tabCount - cIniKeyList.aTabCnt
    End If
    If (Trim(Left(prev, InStr(prev, " ")) = "Select Case") Or prev = "Select Case" Or Trim(Left(prev, InStr(InStr(prev, " ") + 1, prev, " "))) = "Select Case") Then
        tabCount = tabCount + (2 * cIniKeyList.aTabCnt)
    ElseIf (isMemberOfCollection(cBeginList, Trim(Left(prev, InStr(prev, " ")))) Or isMemberOfCollection(cBeginList, prev) Or isMemberOfCollection(cBeginList, Trim(Left(prev, InStr(InStr(prev, " ") + 1, prev, " "))))) And (isOneLineCode(prev) = False) Then
        tabCount = tabCount + cIniKeyList.aTabCnt
    End If
    If tabCount < 0 Then
        tabCount = 0
    End If
    While Len(space) < tabCount
        space = space & " "
    Wend
    indentEdit = space & Trim(this)
End Function


Sub readPrintTxt(aCodeModule As CodeModule)
    Dim prevBuf As Variant
    Dim thisBuf As Variant
    Dim i    As Long
    Dim wNewThis As String
    prevBuf = ""
    thisBuf = ""
    For i = 1 To aCodeModule.CountOfLines
        If cIniKeyList.aIsCommentExec = True Or Left(Trim(aCodeModule.Lines(i, 1)), 1) <> "'" Then
            prevBuf = thisBuf
            thisBuf = aCodeModule.Lines(i, 1)
            wNewThis = indentEdit(Trim(prevBuf), Trim(thisBuf))
            aCodeModule.ReplaceLine i, wNewThis
            If Trim(prevBuf) = "" Then
                cBLList.Add New Dictionary
                cBLList.item(cBLList.Count).Add i, wNewThis
            Else
                cBLList.item(cBLList.Count).Add i, wNewThis
            End If
        End If
    Next i
End Sub


Sub init()
    Set cBeginList = New Collection
    Set cEndList = New Collection
    Set cBLList = New Collection
    Set cBeginDef = New Collection
    tabCount = 0
    cBeginList.Add "If"
    cBeginList.Add "Else"
    cBeginList.Add "ElseIf"
    cBeginList.Add "Enum"
    cBeginList.Add "Sub"
    cBeginList.Add "With"
    cBeginList.Add "While"
    cBeginList.Add "For"
    cBeginList.Add "Do"
    cBeginList.Add "Function"
    cBeginList.Add "Public Function"
    cBeginList.Add "Private Function"
    cBeginList.Add "Public Sub"
    cBeginList.Add "Private Sub"
    cBeginList.Add "Property"
    cBeginList.Add "Type"
    cBeginList.Add "Private Type"
    cBeginList.Add "Public Type"
    cBeginList.Add "Public Property"
    '  cBeginList.Add "Then"
    cBeginList.Add "Public Enum"
    cBeginList.Add "Case"
    cEndList.Add "End"
    cEndList.Add "Next"
    cEndList.Add "Loop"
    cEndList.Add "Else"
    cEndList.Add "Wend"
    cEndList.Add "ElseIf"
    cEndList.Add "Case"
    cBeginDef.Add "Function"
    cBeginDef.Add "Public Function"
    cBeginDef.Add "Private Function"
    cBeginDef.Add "Property"
    cBeginDef.Add "Type"
    cBeginDef.Add "Private Type"
    cBeginDef.Add "Public Type"
    cBeginDef.Add "Public Property"
    cBeginDef.Add "Sub"
    cBeginDef.Add "Public Sub"
    cBeginDef.Add "Private Sub"
    cBeginDef.Add "Enum"
    cBeginDef.Add "Public Enum"
    cBeginDef.Add "Private Enum"
    '
End Sub


Function isMemberOfCollection(col As Collection, query As Variant) As Boolean
    Dim item
    For Each item In col
        If item = query Then
            isMemberOfCollection = True
            Exit Function
        End If
    Next
    isMemberOfCollection = False
End Function


Function isOneLineCode(str As Variant) As Boolean
    Dim buff As Variant
    isOneLineCode = False
    If InStr(str, "'") = 0 Then
        If (InStr(str, "Then") <> 0 And InStr(str, "Then") + 3 < Len(str)) Or InStr(str, "End Function") <> 0 Or InStr(str, "End Sub") Or InStr(str, "End Property") <> 0 Or InStr(str, "End If") <> 0 Then
        isOneLineCode = True
    End If
Else
    buff = Trim(StrConv(LeftB(StrConv(str, vbFromUnicode), Instr2(1, str, "'") - 1), vbUnicode))
    If (InStr(buff, "Then") <> 0 And InStr(buff, "Then") + 3 < Len(buff)) Or InStr(buff, "End Function") <> 0 Or InStr(buff, "End Sub") Or InStr(str, "End Property") <> 0 Or InStr(buff, "End If") <> 0 Then
    isOneLineCode = True
End If
End If
End Function


Sub FormatExecMain()
    Dim mVBComp As VBComponent
    Dim cmps As VBComponents
    If Not IsExistsIni Then
        Call CreateIniFile
    End If
    Call IniRead
    Call init
    If cIniKeyList.aIsAllModuleExec Then
        'For Each mVBComp In ActiveWorkbook.VBProject.VBComponents
        Set cmps = Application.VBE.ActiveVBProject.VBComponents
        For Each mVBComp In cmps
            Call Exec(mVBComp.CodeModule)
            Set cBLList = New Collection
            tabCount = 0
        Next mVBComp
    Else
        Call Exec(ActiveWorkbook.Application.VBE.SelectedVBComponent.CodeModule)
    End If
End Sub


Sub Exec(aCodeModule As CodeModule)
    Call readPrintTxt(aCodeModule)
    If cIniKeyList.aIsAsFormat Then
        Call FixAs(aCodeModule)
    End If
    If cIniKeyList.aIsCommentFormat Then
        Call FixCom(aCodeModule)
    End If
    If cIniKeyList.aIsDeleteNewLine Then
        Call deleteNewLine(aCodeModule)
    End If
    If cIniKeyList.aIsInsertNewLine Then
        Call insertNewLineBeforeDef(aCodeModule)
    End If
End Sub


Sub OptionMain()
    FOption.Show
End Sub

Sub FixAs(aCodeModule As CodeModule)
    Dim i, j  As Integer
    Dim wDic  As Dictionary
    Dim wKeys
    Dim wKey  As Variant
    Dim wStr  As Variant
    Dim wMax  As Integer
    For i = 1 To cBLList.Count
        wMax = 0
        Set wDic = cBLList(i)
        wKeys = wDic.Keys

        For Each wKey In wKeys
            wStr = wDic(wKey)
            If (InStrRev(wStr, """", InStrRev(wStr, " As ") + 1) = 0) And Instr2(1, wStr, " As ") > wMax And (Left(Trim(wStr), 4) = "Dim " Or Left(Trim(wStr), 6) = "Const " Or (aCodeModule.CountOfDeclarationLines >= wKey)) Then
                wMax = Instr2(1, wStr, " As ")
            End If
        Next

        For Each wKey In wKeys
            wStr = wDic(wKey)
            If (InStrRev(wStr, """", InStrRev(wStr, " As ") + 1) = 0) And InStr(wStr, " As ") > 0 And (Left(Trim(wStr), 4) = "Dim " Or Left(Trim(wStr), 6) = "Const " Or (aCodeModule.CountOfDeclarationLines >= wKey)) Then
                'wStr = StrConv(LeftB(StrConv(wStr, vbFromUnicode), Instr2(1, wStr, " As ")), vbUnicode) & WorksheetFunction.Rept(" ", wMax - Instr2(1, wStr, " As ")) & StrConv(RightB(StrConv(wStr, vbFromUnicode), LenB(StrConv(wStr, vbFromUnicode)) - Instr2(1, wStr, " As ")), vbUnicode)
                Dim spaceNum As Long
                spaceNum = wMax - Instr2(1, wStr, " As ")
                If spaceNum < 0 Then spaceNum = 0
                wStr = StrConv(LeftB(StrConv(wStr, vbFromUnicode), Instr2(1, wStr, " As ")), vbUnicode) & space(spaceNum) & StrConv(RightB(StrConv(wStr, vbFromUnicode), LenB(StrConv(wStr, vbFromUnicode)) - Instr2(1, wStr, " As ")), vbUnicode)
                aCodeModule.ReplaceLine wKey, wStr
            End If
            wDic(wKey) = wStr
        Next
    Next i
End Sub

Sub FixCom(aCodeModule As CodeModule)
    Dim i, j    As Integer
    Dim wDic    As Dictionary
    Dim wKeys
    Dim wKey    As Variant
    Dim wStr    As Variant
    Dim wMax    As Integer
    Dim tempStr As Variant
    For i = 1 To cBLList.Count
        wMax = 0
        Set wDic = cBLList(i)
        wKeys = wDic.Keys

        For Each wKey In wKeys
            wStr = wDic(wKey)
            tempStr = wStr
            While InStr(tempStr, """") > 1
                tempStr = Replace(tempStr, Mid(tempStr, InStr(tempStr, """"), InStr(InStr(tempStr, """"), tempStr, """") + 1 - InStr(tempStr, """")), "")
            Wend
            If (Left(Trim(wStr), 1) <> "'") And (InStr(tempStr, "'") > 0) And (Instr2(1, wStr, "'") > wMax) Then
                wMax = Instr2(1, wStr, "'")
            End If
        Next

        For Each wKey In wKeys
            wStr = wDic(wKey)
            tempStr = wStr
            While InStr(tempStr, """") > 1
                tempStr = Replace(tempStr, Mid(tempStr, InStr(tempStr, """"), InStr(InStr(tempStr, """"), tempStr, """") + 1 - InStr(tempStr, """")), "")
            Wend
            If (Left(Trim(wStr), 1) <> "'") And (InStr(tempStr, "'") > 0) Then
                'wStr = StrConv(LeftB(StrConv(wStr, vbFromUnicode), Instr2(1, wStr, "'") - 1), vbUnicode) & WorksheetFunction.Rept(" ", wMax - Instr2(1, wStr, "'")) & StrConv(RightB(StrConv(wStr, vbFromUnicode), LenB(StrConv(wStr, vbFromUnicode)) - Instr2(1, wStr, "'") + 1), vbUnicode)
                Dim spaceNum As Long
                spaceNum = wMax - Instr2(1, wStr, "           '")
                If spaceNum < 0 Then spaceNum = 0
                wStr = StrConv(LeftB(StrConv(wStr, vbFromUnicode), Instr2(1, wStr, "'") - 1), vbUnicode) & space(spaceNum) & StrConv(RightB(StrConv(wStr, vbFromUnicode), LenB(StrConv(wStr, vbFromUnicode)) - Instr2(1, wStr, "'") + 1), vbUnicode)
                aCodeModule.ReplaceLine wKey, wStr
            End If
        Next
    Next i
End Sub

'Sub FixCom(aCodeModule As CodeModule)
'    Dim i, j  As Integer
'    Dim wDic  As Dictionary
'    Dim wKeys
'    Dim wKey  As Variant
'    Dim wStr  As String
'    Dim wMax  As Long
'    Dim tempStr As Variant
'    Dim pos As Long
'    Dim width As Long
'    Dim str1 As String
'    Dim str2 As String
'    Dim spacenum As String
'
'    For i = 1 To cBLList.Count
'        wMax = 0
'        Set wDic = cBLList(i)
'        Debug.Print wDic(1)
'        wKeys = wDic.Keys
'
'        For Each wKey In wKeys
'            wStr = wDic(wKey)
''            str1 = getOutsideStr(wStr, "'")
''            width = getStringWidth(str1)
''            If width > wMax Then wMax = width
'
'            tempStr = wStr
'            While InStr(tempStr, """") > 1
'                tempStr = Replace(tempStr, Mid(tempStr, InStr(tempStr, """"), InStr(InStr(tempStr, """"), tempStr, """") + 1 - InStr(tempStr, """")), "")
'            Wend
'            If (Left(Trim(wStr), 1) <> "'") And (InStr(tempStr, "'") > 0) And (Instr2(1, wStr, "'") > wMax) Then
'                wMax = Instr2(1, wStr, "'")
'            End If
'        Next
'
'        For Each wKey In wKeys
'            wStr = wDic(wKey)
'            tempStr = wStr
'            While InStr(tempStr, """") > 1
'                tempStr = Replace(tempStr, Mid(tempStr, InStr(tempStr, """"), InStr(InStr(tempStr, """"), tempStr, """") + 1 - InStr(tempStr, """")), "")
'            Wend
'            If (Left(Trim(wStr), 1) <> "                      '") And (InStr(tempStr, "'") > 0) Then
''            str1 = getOutsideStr(wStr, "'")
''            str2 = Right(wStr, Len(wStr) - pos)
''            wStr = str1 & space(wMax - getStringWidth(str1)) & str2
'            wStr = StrConv(LeftB(StrConv(wStr, vbFromUnicode), Instr2(1, wStr, "'") - 1), vbUnicode) & Application.WorksheetFunction.Rept(" ", wMax - Instr2(1, wStr, "'")) & StrConv(RightB(StrConv(wStr, vbFromUnicode), LenB(StrConv(wStr, vbFromUnicode)) - Instr2(1, wStr, "'") + 1), vbUnicode)
'                wStr = StrConv(LeftB(StrConv(wStr, vbFromUnicode), Instr2(1, wStr, "'") - 1), vbUnicode) & space(wMax - Instr2(1, wStr, "'")) & StrConv(RightB(StrConv(wStr, vbFromUnicode), LenB(StrConv(wStr, vbFromUnicode)) - Instr2(1, wStr, "'") + 1), vbUnicode)
'
'                aCodeModule.ReplaceLine wKey, wStr
'            End If
'        Next
'    Next i
'End Sub


Function Instr2(aStart As Integer, aString1 As Variant, aString2 As String) As Long
    Instr2 = InStrB(aStart, StrConv(aString1, vbFromUnicode), StrConv(aString2, vbFromUnicode))
End Function


Sub deleteNewLine(aCodeModule As CodeModule)
    Dim i As Long
    Dim x As String
    For i = aCodeModule.CountOfLines To 1 Step -1
        x = aCodeModule.Lines(i, 1)
        x = Trim(x)
        If x = "" Then
            aCodeModule.DeleteLines i, 1
        End If
    Next i
End Sub


Sub insertNewLineBeforeDef(aCodeModule As CodeModule)
    Dim n As Long
    n = aCodeModule.CountOfLines
    Dim this As String
    Dim pre As String
    this = ""
    pre = ""
    Dim bStart As Boolean
    Dim i As Long
    For i = n To 2 Step -1
        this = pre
        pre = Trim(aCodeModule.Lines(i - 1, 1))
        bStart = startDef(this)
        If pre <> "" And startDef(this) Then
            aCodeModule.InsertLines i, vbCrLf
        End If
    Next i
End Sub


Function startDef(str0 As String) As Boolean
    Dim ret As Boolean
    ret = False
    Dim elm As Variant
    For Each elm In cBeginDef
        If startWith(str0, CStr(elm)) Then
            ret = True
            Exit For
        End If
    Next
    startDef = ret
End Function


Function startWith(str0 As String, str1 As String) As Boolean
    Dim ret As Boolean
    Dim n0 As Long
    Dim n1 As Long
    ret = False
    n0 = Len(str0): n1 = Len(str1)
    If n0 = n1 Then
        If str0 = str1 Then ret = True
    ElseIf n0 > n1 Then
        If Left(str0, n1 + 1) = str1 & " " Then ret = True
    End If
    startWith = ret
End Function


Private Function countStr(ByVal str1 As String, ByVal dlm As String) As Long
    Dim ret As Long
    ret = (Len(str1) - Len(Replace(str1, dlm, ""))) / Len(dlm)
    countStr = ret
End Function


Function getOutsidePos(ByVal str1 As String, ByVal dlm As String) As Long
    Dim ret As Long
    Dim tmp As String
    Dim pos0 As Long
    Dim pos1 As Long
    Dim cnt As Long
    pos0 = 1
    pos1 = 0
    cnt = 0
    ret = 0
    Do
        pos1 = InStr(pos0, str1, dlm)
        If pos1 = 0 Then
            ret = 0
            Exit Do
        End If
        tmp = Left(str1, pos1 - 1)
        cnt = countStr(tmp, """")
        If cnt Mod 2 = 0 Then
            ret = pos1 - 1
            Exit Do
        End If
        pos0 = pos1
    Loop
    getOutsidePos = ret
End Function


Function getOutsideStr(ByVal str1 As String, ByVal dlm As String) As String
    Dim ret As String
    Dim tmp As String
    Dim pos0 As Long
    Dim pos1 As Long
    Dim cnt As Long
    pos0 = 1
    pos1 = 0
    cnt = 0
    ret = ""
    tmp = ""
    Do
        pos1 = InStr(pos0, str1, dlm)
        If pos1 = 0 Then
            ret = ""
            Exit Do
        End If
        tmp = Left(str1, pos1 - 1)
        cnt = countStr(tmp, """")
        If cnt Mod 2 = 0 Then
            ret = tmp
            Exit Do
        End If
        pos0 = pos1
    Loop
    getOutsideStr = ret
End Function


Function getStringWidth(str As String) As Long
    Dim ret As Long
    ret = LenB(StrConv(str, vbFromUnicode))
    getStringWidth = ret
End Function
