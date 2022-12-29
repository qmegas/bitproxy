Attribute VB_Name = "MinButton"
Option Explicit

Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Declare Function GetTitleBarInfo Lib "user32" (ByVal hWnd As Long, pti As TitleBarInfo) As Boolean
Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal Y As Long) As Long

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type TitleBarInfo
    cbSize As Long
    rcTitleBar As RECT
    rgState(5) As Long
End Type

Type TOOLINFO
    cbSize As Long
    uFlags As Long
    hWnd As Long
    uid As Long
    RECT As RECT
    hinst As Long
    lpszText As String
    lParam As Long
End Type

Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
        uid As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        sztip As String * 64
End Type

Public Type POINTAPI
   x As Long
   Y As Long
End Type
            
Const DFC_BUTTON = 4
Const DFCS_BUTTONPUSH = &H10
Const DFCS_PUSHED = &H200

Const SM_CXFRAME = 32
Const COLOR_BTNTEXT = 18
Dim lDC As Long
Public R As RECT

Const WS_EX_TOPMOST = &H8&
Const TTS_ALWAYSTIP = &H1
Const HWND_TOPMOST = -1

Const SWP_NOACTIVATE = &H10
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1

Const WM_USER = &H400
Const TTM_ADDTOOLA = (WM_USER + 4)
Const TTF_SUBCLASS = &H10
Public Const TRAY_BACK = (WM_USER + 200)

Const NIM_ADD = &H0
Const NIM_DELETE = &H2
Const NIM_MODIFY = &H1
Const NIF_MESSAGE = &H1
Const NIF_ICON = &H2
Const NIF_TIP = &H4

Const MF_GRAYED = &H1&
Const MF_STRING = &H0&
Const MF_SEPARATOR = &H800&
Const TPM_NONOTIFY = &H80&
Const TPM_RETURNCMD = &H100&

Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, ByVal lprc As Any) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long

Public hWndTT As Long
Public bTraySet As Boolean
Dim lMenu As Long

Public Sub ButtonDraw(frm As Form, bState As Boolean)
    Dim TBButtons As Integer
    Dim TBarHeight As Integer
    Dim TBButtonHeight As Integer
    Dim TBButtonWidth As Integer
    Dim DrawWidth As Integer
    Dim TBI As TitleBarInfo
    Dim TBIRect As RECT
    Dim bRslt As Boolean
    Dim WinBorder As Integer
    With frm
        If .BorderStyle = 0 Then Exit Sub
        '----How Many Buttons in TitleBar------------------------------------------
        If Not .ControlBox Then TBButtons = 0
        If .ControlBox Then TBButtons = 1
        If .ControlBox And .WhatsThisButton Then
            If .BorderStyle < 4 Then
                TBButtons = 2
            Else
                TBButtons = 1
            End If
        End If
        If .ControlBox And .MinButton And .BorderStyle = 2 Then TBButtons = 3
        If .ControlBox And .MinButton And .BorderStyle = 5 Then TBButtons = 1
        If .ControlBox And .MaxButton And .BorderStyle = 2 Then TBButtons = 3
        If .ControlBox And .MaxButton And .BorderStyle = 5 Then TBButtons = 1
        '----Get height of Titlebar----------------------------------------------
        'Using this method gets the height of the titlebar regardless of the window
        'style. It does, however, restrict its use to Win98/2000. So if you want to
        'use this code in Win95, then call GetSystemMetrics to find the windowstyle
        'and titlebar size.
        TBI.cbSize = Len(TBI)
        bRslt = GetTitleBarInfo(.hWnd, TBI)
        TBIRect = TBI.rcTitleBar
        TBarHeight = TBIRect.Bottom - TBIRect.Top - 1
        '----Get WindowBorder Size----------------------------------------------
        If .BorderStyle = 2 Or .BorderStyle = 5 Then
            R.Top = GetSystemMetrics(32) + 2
            WinBorder = R.Top - 6
        Else
            R.Top = 5
            WinBorder = -1
        End If
    End With
    '----Use Titlebar Height to determin button size----------------------------
    TBButtonHeight = TBarHeight - 4
    TBButtonWidth = TBButtonHeight + 2
    'and the size and space of the dot on the button
    DrawWidth = TBarHeight / 8
    '---------------------------------------------------------------------------
    '----Determin the position of our button------------------------------------
    R.Bottom = R.Top + TBButtonHeight
    Select Case TBButtons
        Case 1
            R.Right = frm.ScaleWidth - (TBButtonWidth) + WinBorder
        Case 2
            R.Right = frm.ScaleWidth - ((TBButtonWidth * 2) + 2) + WinBorder
        Case 3
            R.Right = frm.ScaleWidth - ((TBButtonWidth * 3) + 2) + WinBorder
        Case Else
            R.Right = frm.ScaleWidth
    End Select
    R.Left = R.Right - TBButtonWidth
    '----Get the Widow DC so that we may draw in the title bar-----------------
    lDC = GetWindowDC(frm.hWnd)
    '----Устанавливаем позицию точки--------------------------------------
    Dim StartXY As Integer, EndXY As Integer
    Select Case TBarHeight
        Case Is < 20
            StartXY = DrawWidth + 1
            EndXY = DrawWidth - 1
        Case Else
            StartXY = (DrawWidth * 2)
            EndXY = DrawWidth
    End Select
    '----У нас есть вся информация, которая нам нужна чтобы рисовать кнопку----------------
    Dim rDot As RECT
    If bState Then
        DrawFrameControl lDC, R, DFC_BUTTON, DFCS_BUTTONPUSH Or DFCS_PUSHED
        rDot.Left = R.Right - (1 + StartXY)
        rDot.Top = R.Bottom - (1 + StartXY)
        rDot.Right = R.Right - (1 + EndXY)
        rDot.Bottom = R.Bottom - (1 + EndXY)
    Else
        DrawFrameControl lDC, R, DFC_BUTTON, DFCS_BUTTONPUSH
        rDot.Left = R.Right - (2 + StartXY)
        rDot.Top = R.Bottom - (2 + StartXY)
        rDot.Right = R.Right - (2 + EndXY)
        rDot.Bottom = R.Bottom - (2 + EndXY)
    End If
    FillRect lDC, rDot, GetSysColorBrush(COLOR_BTNTEXT)
    '----Set Tooltip------------------------------------------------------------
    Dim TTRect As RECT
    TTRect.Bottom = R.Bottom + (TBarHeight - ((TBarHeight * 2) + WinBorder + 5))
    TTRect.Left = R.Left - (4 - WinBorder)
    TTRect.Right = R.Right - (4 - WinBorder)
    TTRect.Top = R.Top + (TBarHeight - ((TBarHeight * 2) + WinBorder + 5))
    KillTip
    CreateTip appForm.hWnd, "Убрать в ситемный трей", TTRect
End Sub

Public Sub CreateTip(hwndForm As Long, szText As String, rct As RECT)
    hWndTT = CreateWindowEx(WS_EX_TOPMOST, "tooltips_class32", "", TTS_ALWAYSTIP, 0, 0, 0, 0, hwndForm, 0&, App.hInstance, 0&)
    SetWindowPos hWndTT, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
    Dim TI As TOOLINFO
    With TI
        .cbSize = Len(TI)
        .uFlags = TTF_SUBCLASS
        .hWnd = hwndForm
        .hinst = App.hInstance
        .uid = 1&
        .lpszText = szText & vbNullChar
        .RECT = rct
    End With
    SendMessage hWndTT, TTM_ADDTOOLA, 0, TI
End Sub

Public Sub KillTip()
    DestroyWindow hWndTT
End Sub

Public Sub TraySet(frm As Form)
    frm.Hide
    bTraySet = True
End Sub

Public Sub TrayRestore(frm As Form)
    frm.Show
    frm.SetFocus
    bTraySet = False
End Sub

Public Sub ShowTray(frm As Form, sztip As String, hIcon As Long)
    Dim NID As NOTIFYICONDATA
    With NID
        .cbSize = Len(NID)
        .hIcon = hIcon
        .hWnd = frm.hWnd
        .sztip = sztip & vbNullChar
        .uCallbackMessage = TRAY_BACK
        .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
        .uid = 1&
    End With
    Shell_NotifyIcon NIM_ADD, NID
End Sub

Public Sub ChangeTrayTip(frm As Form, sztip As String)
    Dim NID As NOTIFYICONDATA
    With NID
        .cbSize = Len(NID)
        .hWnd = frm.hWnd
        .sztip = sztip & vbNullChar
        .uFlags = NIF_TIP
        .uid = 1&
    End With
    Shell_NotifyIcon NIM_MODIFY, NID
End Sub

Public Sub KillTray(frm As Form)
    Dim NID As NOTIFYICONDATA
    With NID
        .cbSize = Len(NID)
        .hWnd = frm.hWnd
        .uid = 1&
    End With
    Shell_NotifyIcon NIM_DELETE, NID
End Sub
