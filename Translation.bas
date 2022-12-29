Attribute VB_Name = "modTranslation"
Option Explicit

Private Const LANGCOUNT = 234
Private Const LANG_TAG = "#BT_LANG#"

Private FL(1 To LANGCOUNT) As String
Private needTranslate As Boolean

Public Function LangFilesFillList(CObj As Object) As Integer
    Dim tmp As String, j As Long
    Dim tmp2 As String
    
    LangFilesFillList = 0
    CObj.AddItem "English"
    
    tmp = Dir(CPath & LANGDIR & "*.lng")
    j = FreeFile
    Do While Len(tmp) > 0
        Open CPath & LANGDIR & tmp For Input As #j
        Line Input #j, tmp2
        If tmp2 = LANG_TAG Then
            tmp = GetFileNameNoExt(tmp)
            CObj.AddItem tmp
            If LCase(tmp) = LCase(prgSettings.LangSelected) Then _
                LangFilesFillList = CObj.ListCount - 1
        End If
        Close j
        tmp = Dir
    Loop
End Function

Public Sub LoadCStr()
    Dim i As Long, j As Long, indx As Long
    Dim langFile As String
    Dim tmp As String
    
    needTranslate = False
    If prgSettings.LangSelected = "english" Then Exit Sub
    
    langFile = CPath & LANGDIR & prgSettings.LangSelected & ".lng"
    If Not FileExist(langFile) Then Exit Sub
    
    j = FreeFile
    Open langFile For Input As #j
    Line Input #j, tmp
    If tmp <> LANG_TAG Then
        Close j
        Exit Sub
    End If
    Line Input #j, tmp 'Version
    Line Input #j, tmp 'Author
    While Not EOF(j)
        Line Input #j, tmp
        i = InStr(1, tmp, "=")
        If i > 0 Then
            indx = val(Trim(Left(tmp, i - 1)))
            If indx > 0 And indx <= LANGCOUNT Then
                tmp = Mid(tmp, i + 1)
                str_replace "\n", vbCrLf, tmp
                If Len(tmp) > 0 Then FL(indx) = tmp
            End If
        End If
    Wend
    Close j
    needTranslate = True
End Sub

Public Function getCStr(ind As Long, def_str As String) As String
    If needTranslate And Len(FL(ind)) > 0 Then getCStr = FL(ind) _
       Else getCStr = def_str
End Function

Public Sub ReTranslateNow()
    Dim curFrm As Object
    
    LoadCStr
    For Each curFrm In Forms
        TranslateForm curFrm
    Next
End Sub

Public Sub TranslateForm(CurForm As Form)
    Dim tObj As Object, i As Long
    
    If Not needTranslate Then Exit Sub
    
    For Each tObj In CurForm.Controls
        If TypeOf tObj Is CommandButton Then
            TransStandart tObj
        ElseIf TypeOf tObj Is CheckBox Then
            TransStandart tObj
        ElseIf TypeOf tObj Is OptionButton Then
            TransStandart tObj
        ElseIf TypeOf tObj Is Frame Then
            TransStandart tObj
        ElseIf TypeOf tObj Is ListView Then
            TransListView tObj
        ElseIf TypeOf tObj Is Label Then
            TransStandart tObj
        ElseIf TypeOf tObj Is Toolbar Then
            TransToolbar tObj
        ElseIf TypeOf tObj Is Menu Then
            TransStandart tObj
        ElseIf TypeOf tObj Is TabStrip Then
            TransTabStrip tObj
        End If
    Next
    TransStandart CurForm
End Sub

Private Function checkTrans(someObj As Object) As Long
    Dim num As Long
    
    checkTrans = 0
    If Left(someObj.Tag, 1) = "~" Then
        num = val(Mid(someObj.Tag, 2))
        If num > 0 And num <= LANGCOUNT Then
            If Len(FL(num)) > 0 Then checkTrans = num
        End If
    End If
End Function

Private Sub TransStandart(someObj As Object)
    Dim i As Long
    i = checkTrans(someObj)
    If i > 0 Then someObj.Caption = FL(i)
End Sub

Private Sub TransListView(someObj As Object)
    Dim i As Long, j As Long
    For i = 1 To someObj.ColumnHeaders.Count
        j = checkTrans(someObj.ColumnHeaders(i))
        If j > 0 Then someObj.ColumnHeaders(i).Text = FL(j)
    Next
End Sub

Private Sub TransToolbar(someObj As Object)
    Dim i As Long, j As Long
    Dim i2 As Long, j2 As Long
    For i = 1 To someObj.Buttons.Count
        j = checkTrans(someObj.Buttons(i))
        If j > 0 Then someObj.Buttons(i).Caption = FL(j)
        For i2 = 1 To someObj.Buttons(i).ButtonMenus.Count
            j2 = checkTrans(someObj.Buttons(i).ButtonMenus(i2))
            If j2 > 0 Then someObj.Buttons(i).ButtonMenus(i2).Text = FL(j2)
        Next
    Next
End Sub

Private Sub TransTabStrip(someObj As Object)
    Dim i As Long, j As Long
    For i = 1 To someObj.Tabs.Count
        j = checkTrans(someObj.Tabs(i))
        If j > 0 Then someObj.Tabs(i).Caption = FL(j)
    Next
End Sub
