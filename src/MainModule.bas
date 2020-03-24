Attribute VB_Name = "MainModule"
Option Explicit
Option Base 0
Private tabCount   As Integer
Private cBeginList  As New Collection
Private cEndList   As New Collection
Private cBeginDef  As New Collection
Private clsVBAFormatMenu As VBAFormatMenu
Private cBLList   As New Collection

Sub addButton()
    Dim wVBAFormatterMenu As CommandBarControl
    Dim wOptionMenu  As CommandBarControl
    Dim wFormatExecMenu As CommandBarControl
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

Function indentEdit(prev As String, this As String) As String
    this = Trim(this)
    prev = Trim(prev)
    If likeKeyword("End Select", this) Then
        tabCount = tabCount - (2 * cIniKeyList.aTabCnt)
    ElseIf likeMemberOfCollection(cEndList, this) And this <> "End" Then
        tabCount = tabCount - cIniKeyList.aTabCnt
    End If
    If likeKeyword("Select Case", prev) Then
        tabCount = tabCount + (2 * cIniKeyList.aTabCnt)
    ElseIf likeMemberOfCollection(cBeginList, prev) And Not isOneLineCode(prev) Then
        tabCount = tabCount + cIniKeyList.aTabCnt
    End If
    If tabCount < 0 Then
        tabCount = 0
    End If
    indentEdit = space(tabCount) & Trim(this)
End Function

Sub readPrintTxt(aCodeModule As CodeModule)
    Dim prevBuf As Variant
    Dim thisBuf As Variant
    Dim i  As Long
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
    cBeginList.Add "Private Property"
    ' cBeginList.Add "Then"
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
    cBeginDef.Add "Private Property"
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

Function likeKeyword(key, x) As Boolean
    Dim ret As Boolean
    ret = (x = key Or x Like (key & " *"))
    likeKeyword = ret
End Function

Function likeMemberOfCollection(col As Collection, query As Variant) As Boolean
    Dim item
    Dim ret As Boolean
    ret = False
    For Each item In col
        If likeKeyword(item, query) Then
            ret = True
            Exit For
        End If
    Next
    likeMemberOfCollection = ret
End Function

Function isOneLineCode(str As String) As Boolean
    Dim buff As String
    Dim ret As Boolean
    Dim word, endWords
    ret = False
    buff = Trim(delComment(str))
    endWords = Array("End Sub", "End Function", "End Property", "End If")
    For Each word In endWords
        If buff Like "*" & word Then ret = True
    Next
    If getPosOverQuote(buff, "Then") > 0 And Not buff Like "* Then" Then ret = True
    isOneLineCode = ret
End Function

Sub FormatExecMain()
    Dim mVBComp As VBComponent
    Dim cmps As VBComponents
    Dim cmp As VBComponent
    If Not IsExistsIni Then
        Call CreateIniFile
    End If
    Call IniRead
    Call init
    Set cmps = Application.VBE.ActiveVBProject.VBComponents
    If cIniKeyList.aIsAllModuleExec Then
        'For Each mVBComp In ActiveWorkbook.VBProject.VBComponents
        For Each mVBComp In cmps
            Call Exec(mVBComp.CodeModule)
            Set cBLList = New Collection
            tabCount = 0
        Next mVBComp
    Else
        ' Call Exec(ActiveWorkbook.Application.VBE.SelectedVBComponent.CodeModule)
        Set cmp = cmps.VBE.SelectedVBComponent
        Call Exec(cmp.CodeModule)
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
    Dim i, j As Integer
    Dim wDic As Dictionary
    Dim wKeys
    Dim wKey As Variant
    Dim wStr As Variant
    Dim wMax As Integer
    Dim spaceNum As Long
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
    Dim i, j As Integer
    Dim wDic As Dictionary
    Dim dicHit As Dictionary
    Dim wKeys
    Dim wKey As Variant
    Dim wStr As String
    Dim wMax As Long
    Dim str1 As String
    Dim str2 As String
    Dim width As Long
    Dim pos As Long
    For i = 1 To cBLList.Count
        wMax = 0
        Set wDic = cBLList(i)
        Set dicHit = New Dictionary
        'Debug.Print wDic(1)
        wKeys = wDic.Keys
        For Each wKey In wKeys
            wStr = wDic(wKey)
            pos = getPosOverQuote(wStr, "'")
            If pos > 0 Then
                width = getStringWidth(Left(wStr, pos - 1))
                If width > wMax Then
                    wMax = width
                End If
                dicHit.Add wKey, pos
            End If
        Next
        For Each wKey In dicHit.Keys
            wStr = wDic(wKey)
            pos = dicHit(wKey)
            str1 = Left(wStr, pos - 1)
            str2 = Right(wStr, Len(wStr) - Len(str1))
            width = getStringWidth(str1)
            wStr = str1 & space(wMax - width) & str2
            aCodeModule.ReplaceLine wKey, wStr
            wDic(wKey) = wStr
        Next
    Next i
End Sub

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
            aCodeModule.InsertLines i, ""
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

Function delComment(sLine As String) As String
    Dim ret As String
    Dim n As Long
    ret = sLine
    n = getPosOverQuote(sLine, "'")
    If n > 0 Then ret = Left(sLine, n - 1)
    delComment = ret
End Function

Private Function countStr(ByVal str1 As String, ByVal dlm As String) As Long
    Dim ret As Long
    ret = (Len(str1) - Len(Replace(str1, dlm, ""))) / Len(dlm)
    countStr = ret
End Function

Function getPosOverQuote(ByVal str1 As String, ByVal dlm As String) As Long
    Dim ret As Long
    Dim tmp As String
    Dim pos0 As Long
    Dim pos1 As Long
    Dim cnt As Long
    pos0 = 1
    pos1 = 0
    cnt = 0
    ret = 0
    tmp = ""
    Do
        pos1 = InStr(pos0, str1, dlm)
        If pos1 = 0 Then
            ret = 0
            Exit Do
        End If
        tmp = Left(str1, pos1 - 1)
        cnt = countStr(tmp, """")
        If cnt Mod 2 = 0 Then
            ret = pos1
            Exit Do
        End If
        pos0 = pos1 + 1
        If pos0 > Len(str1) Then
            ret = 0
            Exit Do
        End If
    Loop
    getPosOverQuote = ret
End Function

Function getStringWidth(str As String) As Long
    Dim ret As Long
    ret = LenB(StrConv(str, vbFromUnicode))
    getStringWidth = ret
End Function

Sub testqqq()
    Const x    As Long = 1
    Const y    As String = "abc"
    Const z    As String = "qqqq"
    Dim w     As String
    Dim s     As String
    Dim t     As String
    Dim u     As String
    Dim v     As String
    w = "'abc"                      'test
    s = """'abc"                    'test 'test
    t = "'""abc"                   'test 'test
    u = "'""abc"""                'test
    v = "'""abc"                'test
End Sub
