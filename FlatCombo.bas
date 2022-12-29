Attribute VB_Name = "FlatCombo"
Option Explicit
'********************************************************************
'*                      Team HomeWork                               *
'*                      e-mail: sne_pro@mail.ru                     *
'********************************************************************
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal y As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function PtInRect Lib "user32.dll" (lpRect As RECT, ByVal X As Long, ByVal y As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long

Private Declare Function SetTimer Lib "user32.dll" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32.dll" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Private Type gbHWCtrlData
    BaseProc        As Long
    EditProc        As Long
    hwnd            As Long
    hWndEdit        As Long
    m_BackClr       As Long
    m_ShadowClr     As Long
    m_HighlightClr  As Long
    m_bLBtnDown     As Boolean
    m_bPainted      As Boolean
    m_bHasFocus     As Boolean
End Type
Private Type RECT
    Left            As Long
    Top             As Long
    Right           As Long
    Bottom          As Long
End Type
Private Type POINTAPI
    X As Long
    y As Long
End Type

Private Const CBS_DROPDOWN As Long = &H2&

Private ctlData() As gbHWCtrlData

Public Sub MakeThinAll(CurForm As Form)
    Dim tObj As Object
    
    For Each tObj In CurForm.Controls
        If TypeOf tObj Is TextBox Then
            TextMakeThin tObj.hwnd
        ElseIf TypeOf tObj Is ListBox Then
            TextMakeThin tObj.hwnd
        End If
    Next
End Sub

Private Sub TextMakeThin(ByVal hwnd As Long)
    Dim j As Long
    j = GetWindowLong(hwnd, GWL_STYLE)
    j = j Xor WS_BORDER
    SetWindowLong hwnd, GWL_STYLE, j
End Sub

Public Sub SetComboFlat(ByVal lClHandle As Long, Optional ByVal m_BackClr As Long = vbButtonFace, Optional ByVal m_ShadowClr As Long = vbButtonShadow, Optional ByVal m_HighlightClr As Long = vb3DHighlight)
    Dim lStyle  As Long

    If Not FindCtlByKey(lClHandle) = &HFFFF Then Exit Sub
                                                         
    lStyle = GetWindow(lClHandle, &H5)
    If Not lStyle = &H0 Then
        Call AddCtl(SetWindowLong(lClHandle, &HFFFC, AddressOf flatComboProc), SetWindowLong(lStyle, &HFFFC, AddressOf flatComboProc), lClHandle, lStyle, m_BackClr, m_ShadowClr, m_HighlightClr)
    Else
        Call AddCtl(SetWindowLong(lClHandle, &HFFFC, AddressOf flatComboProc), &H0, lClHandle, &H0, m_BackClr, m_ShadowClr, m_HighlightClr)
    End If
End Sub

Public Sub DelComboFlat(ByVal lClHandle As Long)
    Dim Index As Integer

    Index = FindCtlByKey(lClHandle)
    If Index = &HFFFF Then Exit Sub
                                                
    If ctlData(Index).EditProc Then Call SetWindowLong(ctlData(Index).hWndEdit, &HFFFC, ctlData(Index).EditProc)
    If ctlData(Index).BaseProc Then Call SetWindowLong(lClHandle, &HFFFC, ctlData(Index).BaseProc)

    Call RemoveCtl(Index)
End Sub
                                                
Public Function flatComboProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim Index As Integer

    On Error Resume Next

    Index = FindCtlByKey(hwnd)
    With ctlData(Index)
        If hwnd = ctlData(Index).hwnd Then
            flatComboProc = CallWindowProc(.BaseProc, hwnd, msg, wParam, lParam)

        ElseIf hwnd = ctlData(Index).hWndEdit Then
            flatComboProc = CallWindowProc(.EditProc, hwnd, msg, wParam, lParam)
            hwnd = ctlData(Index).hwnd
        End If

        Select Case msg
            Case Is = &H200                                     ' WM_MOUSEMOVE
                Call SetTimer(hwnd, &H1, &HA, AddressOf OnTimer)
                Call OnTimer(hwnd, &H0, &H0, &H0)

            Case Is = &H201                                     ' OnLButtonDown
                .m_bLBtnDown = True

            Case Is = &H202                                     ' OnLButtonUp
                .m_bLBtnDown = False

            Case Is = &HF                                       ' WM_PAINT
                If .m_bHasFocus Then _
                    Call DrawCombo(hwnd, &H1, .m_ShadowClr, .m_HighlightClr, .m_BackClr, .m_ShadowClr, .m_HighlightClr) _
                Else _
                    Call DrawCombo(hwnd, &H0, .m_BackClr, .m_BackClr, .m_BackClr, .m_ShadowClr, .m_HighlightClr)

            Case Is = &H7                                       ' WM_SETFOCUS
                ctlData(Index).m_bHasFocus = True
                Call DrawCombo(hwnd, &H1, .m_ShadowClr, .m_HighlightClr, .m_BackClr, .m_ShadowClr, .m_HighlightClr)

            Case Is = &H8                                       ' WM_KILLFOCUS
                ctlData(Index).m_bHasFocus = False
                Call DrawCombo(hwnd, &H0, .m_BackClr, .m_BackClr, .m_BackClr, .m_ShadowClr, .m_HighlightClr)
        End Select
    End With
End Function

Private Function FindCtlByKey(ByVal hwnd As Long) As Integer
    On Error GoTo er
    Dim fk As Long: FindCtlByKey = &HFFFF
    For fk = 0 To UBoundCtl
        If hwnd = ctlData(fk).hwnd Or hwnd = ctlData(fk).hWndEdit Then FindCtlByKey = fk: Exit For
    Next
er:
End Function

Private Function AddCtl(ByVal BaseProc As Long, ByVal EditProc As Long, ByVal hwnd As Long, ByVal hWndEdit As Long, ByVal m_BackClr As Long, ByVal m_ShadowClr As Long, ByVal m_HighlightClr As Long) As Long
    On Error GoTo er
    Dim Index As Integer: Index = UBoundCtl + &H1
    
    ReDim Preserve ctlData(Index)
    With ctlData(Index)
        .BaseProc = BaseProc
        .EditProc = EditProc
        .hwnd = hwnd
        .hWndEdit = hWndEdit
        .m_BackClr = m_BackClr
        .m_ShadowClr = m_ShadowClr
        .m_HighlightClr = m_HighlightClr
    End With
    
    AddCtl = Index
er:
End Function

Private Function RemoveCtl(ByVal Index As Integer) As Integer
    On Error GoTo er
    Dim ub As Long: ub = UBoundCtl
    
    If Not ub = 0 And Not ub = &HFFFF Then
       Call CopyMemory(ctlData(Index), ctlData(Index + 1), Len(ctlData(Index)) * (ub - Index))
    Else
        Erase ctlData
    End If
er:
End Function

Private Function UBoundCtl() As Long
    On Error GoTo er
    UBoundCtl = UBound(ctlData)
    Exit Function
er:     UBoundCtl = &HFFFF
End Function

Private Sub DrawCombo(ByVal hwnd As Long, ByVal eState As Integer, ByVal clrTopLeft As Long, ByVal clrBottomRight As Long, ByVal clrBack As Long, ByVal clrShadow As Long, ByVal clrHighlight As Long)

    Dim rcItem As RECT, pDC As Long, m_nOffset As Integer

    Call GetClientRect(hwnd, rcItem)
    pDC = GetDC(hwnd)
    m_nOffset = GetSystemMetrics(&HA)

    Call Draw3DRect(pDC, rcItem, clrTopLeft, clrBottomRight)
    Call InflateRect(rcItem, &HFFFF, &HFFFF)

    If IsWindowEnabled(hwnd) Then _
        Call Draw3DRect(pDC, rcItem, clrBack, clrBack) _
    Else _
        Call Draw3DRect(pDC, rcItem, clrHighlight, clrHighlight)

    Call InflateRect(rcItem, &HFFFF, &HFFFF)
    rcItem.Left = rcItem.Right - m_nOffset
    Call Draw3DRect(pDC, rcItem, clrBack, clrBack)

    Call InflateRect(rcItem, &HFFFF, &HFFFF)
    Call Draw3DRect(pDC, rcItem, clrBack, clrBack)

    If IsWindowEnabled(hwnd) = &H0 Then Exit Sub

    Select Case eState
        Case Is = &H0
            rcItem.Top = rcItem.Top - &H1
            rcItem.Bottom = rcItem.Bottom + &H1
            Call Draw3DRect(pDC, rcItem, clrHighlight, clrHighlight)
            rcItem.Left = rcItem.Left - &H1
            Call Draw3DRect(pDC, rcItem, clrHighlight, clrHighlight)

        Case Is = &H1
            rcItem.Top = rcItem.Top - &H1
            rcItem.Bottom = rcItem.Bottom + &H1
            Call Draw3DRect(pDC, rcItem, clrHighlight, clrShadow)

        Case Is = &H2
            rcItem.Top = rcItem.Top - &H1
            rcItem.Bottom = rcItem.Bottom + &H1
            Call Draw3DRect(pDC, rcItem, clrShadow, clrHighlight)
    End Select

    Call ReleaseDC(hwnd, pDC):  Call DeleteDC(pDC)
End Sub

Private Sub Draw3DRect(ByVal hdc As Long, ByRef rcItem As RECT, ByVal oTopLeftColor As OLE_COLOR, ByVal oBottomRightColor As OLE_COLOR, Optional ByVal bMod As Boolean = False)
    Dim hPen As Long, hPenOld As Long, tp As POINTAPI

    hPen = CreatePen(&H0, &H1, TrClr(oTopLeftColor))
    hPenOld = SelectObject(hdc, hPen)

    Call MoveToEx(hdc, rcItem.Left, rcItem.Bottom - &H1, tp)
    Call LineTo(hdc, rcItem.Left, rcItem.Top)
    Call LineTo(hdc, rcItem.Right - &H1, rcItem.Top)

    Call SelectObject(hdc, hPenOld)
    Call DeleteObject(hPen)

    If Not rcItem.Left = rcItem.Right Then
        hPen = CreatePen(&H0, &H1, TrClr(oBottomRightColor))
        hPenOld = SelectObject(hdc, hPen)

        Call LineTo(hdc, rcItem.Right - &H1, rcItem.Bottom - &H1)
        Call LineTo(hdc, rcItem.Left - (IIf(bMod, &H1, &H0)), rcItem.Bottom - &H1)

        Call SelectObject(hdc, hPenOld)
        Call DeleteObject(hPen)
    End If
End Sub

Private Function TrClr(ByVal clr As Long, Optional ByVal hPal As Long = &H0) As Long
    If OleTranslateColor(clr, hPal, TrClr) Then TrClr = &HFFFF
End Function

Private Function OnTimer(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long) As Long
    Dim pt As POINTAPI, rcItem As RECT, Index As Integer
    Call GetCursorPos(pt)
    Call GetWindowRect(hwnd, rcItem)

    Index = FindCtlByKey(hwnd)
    With ctlData(Index)
        If .m_bLBtnDown Then
            Call KillTimer(hwnd, &H1)
            If .m_bPainted Then
                Call DrawCombo(hwnd, &H2, .m_ShadowClr, .m_HighlightClr, .m_BackClr, .m_ShadowClr, .m_HighlightClr)
                ctlData(Index).m_bPainted = False
            End If
            Exit Function
        End If
    
        If PtInRect(rcItem, pt.X, pt.y) = &H0 And Not .m_bHasFocus Then
            Call KillTimer(hwnd, &H1)
            If .m_bPainted Then
                Call DrawCombo(hwnd, &H0, .m_BackClr, .m_BackClr, .m_BackClr, .m_ShadowClr, .m_HighlightClr)
                .m_bPainted = False
            End If
            Exit Function
        End If
    
        If .m_bPainted Then
            Exit Function
        Else
            .m_bPainted = True
            Call DrawCombo(hwnd, &H1, .m_ShadowClr, .m_HighlightClr, .m_BackClr, .m_ShadowClr, .m_HighlightClr)
        End If
    End With
End Function
