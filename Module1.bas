Attribute VB_Name = "modCommon"
Option Explicit

Public Sub Main()
    Dim rez As Long
    
    prgSettings.LangSelected = 0
    ifCheckedUp = False
    CPath = App.Path
    If Right(CPath, 1) <> "\" Then CPath = CPath & "\"
    TempPath = Space(255)
    rez = GetTempPath(255, TempPath)
    TempPath = Left(TempPath, rez)
    AppArg = Trim(Command())
    
    Global_TaskbarMsg = RegisterWindowMessage("TaskbarCreated")
    
    Select Case AppArg
        Case "-h" 'Install
            MakeAssociation True
        Case "-u" 'Uninstall
            MakeAssociation False
        Case Else
            rez = FindPrevWindow()
            FilePoolFromCommandLine
            If rez > 0 Then
                If Global_FilePool.Count > 0 Then
                    WriteKey HKEY_LOCAL, MYPATH, "ComSend", Global_FilePool.Item(1)
                    SendMessage rez, MAKE_NEW, 0, Null
                End If
                End
            End If
            
            LoadPSettings
            LoadCStr
            
            MainForm.Show
    End Select
End Sub

Public Function WinProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim tmp As String, skip_it As Boolean
    
    skip_it = False
    Select Case msg
        Case TRAY_BACK 'Tray icon
            If lParam = WM_LBUTTONDBLCLK Then
                MainForm.Tray_Click 1000
            End If
            If lParam = WM_RBUTTONDOWN Then
                'MainForm.SetFocus
                MenuTrack
            End If
        Case MAKE_NEW
            If MainForm.Visible = False Then MainForm.Visible = True
            MainForm.SetFocus
            tmp = GetKeyValue(HKEY_LOCAL, MYPATH, "ComSend", "-")
            If tmp <> "-" Then
                FilePoolClear
                Global_FilePool.Add tmp
                MainForm.MakeAsk
            End If
        Case WM_DROPFILES
            FilePoolFromMainDrop wParam
            MainForm.MakeAsk
        Case Global_TaskbarMsg
            MainForm.AddSysTray
    End Select
    
    If Not skip_it Then _
        WinProc = CallWindowProc(prev, hwnd, msg, wParam, lParam)
End Function

Public Sub PopupMenuCreate()
    Dim mii As MENUITEMINFO
    
    hPopMenu = CreatePopupMenu()
    
    AppendMenu hPopMenu, MF_DEFAULT, 1000, getCStr(161, "Open BitTorrent Proxy window")
    AppendMenu hPopMenu, MF_SEPARATOR, 0&, 0&
    AppendMenu hPopMenu, 0&, 1100, getCStr(162, "Add downloading task...")
    AppendMenu hPopMenu, 0&, 1200, getCStr(163, "Add emulation task...")
    AppendMenu hPopMenu, 0&, 1300, getCStr(164, "Settings...")
    AppendMenu hPopMenu, MF_SEPARATOR, 0&, 0&
    AppendMenu hPopMenu, 0&, 1400, getCStr(165, "Check for new version")
    AppendMenu hPopMenu, MF_SEPARATOR, 0&, 0&
    AppendMenu hPopMenu, 0&, 1500, getCStr(166, "Instruction/Help")
    AppendMenu hPopMenu, 0&, 1600, getCStr(167, "BitTorrent Proxy home page")
    AppendMenu hPopMenu, 0&, 1700, getCStr(168, "BitTorrent Proxy forum")
    AppendMenu hPopMenu, MF_SEPARATOR, 0&, 0&
    AppendMenu hPopMenu, 0&, 1800, getCStr(169, "About...")
    AppendMenu hPopMenu, MF_SEPARATOR, 0&, 0&
    AppendMenu hPopMenu, 0&, 1900, getCStr(170, "Exit")
    
    mii.cbSize = Len(mii)
    mii.fMask = MIIM_STATE
    mii.fState = MFS_DEFAULT
    SetMenuItemInfo hPopMenu, 0, API_TRUE, mii
End Sub

Public Sub MenuTrack()
    Dim sMenu As Long, mp As POINTAPI
    
    GetCursorPos mp
    SetForegroundWindow MainForm.hwnd
    sMenu = TrackPopupMenu(hPopMenu, TPM_RETURNCMD, mp.X, mp.y, 0, MainForm.hwnd, 0&)
    MainForm.Tray_Click sMenu
End Sub

Public Function MakeNDigit(inx As String, toNum As Integer, peer_id As String) As String
    Dim tmp As String, i As Integer, b As String * 1, j As String
    
    tmp = peer_id
    j = 1
    For i = 1 To toNum
        b = Mid(tmp, j, 1)
        If b = "%" Then j = j + 3 Else j = j + 1
    Next
    MakeNDigit = inx & Mid(tmp, j)
End Function

Public Sub LoadPSettings()
    Dim i As Integer, tmp As String
    On Error Resume Next
    
    With prgSettings
        .Sett = GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_SETT, "non")
        .LangSelected = GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_LANG, "english")
        .txtServer = GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_HOST, "torrents.ru")
        .txtPort = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_PORT, "6666"))
        '=====
        i = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_MODE, "0"))
        If i < 1 Or i > 2 Then i = 1
        .optMode = i
        .txtUpload = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_UPLOAD, "5"))
        .txtM2from = GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_M2FROM, "1.5")
        .txtM2to = GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_M2TO, "2")
        .chkDown = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_DWNUSE, "0"))
        .txtDown = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_DWNVAL, "2"))
        .chkDnotsend = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_DWNNOTSEND, "0"))
        '=====
        .chkVersion = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_USEVER, "0"))
        .cmbVersion = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_VERTYPE, "0"))
        '=====
        .chkMinimize = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_MINIMIZE, "0"))
        .chkAutoUpdate = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_AUTOCHECK, "1"))
        .chkUseProxy = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_USEPROXY, "0"))
        .txtProxyIp = GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_PROXYIP, "0.0.0.0")
        .txtProxyPort = GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_PROXYPORT, "80")
        .chkRetracker = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_RETRACKER, "0"))
        '=====
        .chkSmart = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_SMARTUSE, "0"))
        .txtSmartA = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_SMARTA, "30"))
        .txtSmartP = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_SMARTP, "30"))
        '=====
        .defAction = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_DEFACTION, "0"))
        '=====
        .emul_client = GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_EMULCLIENT, vbNullString)
        .emul_dw1 = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_EMULDW1, "30"))
        .emul_dw2 = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_EMULDW2, "40"))
        .emul_up1 = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_EMULUP1, "300"))
        .emul_up2 = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_EMULUP2, "400"))
        .emul_port = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_EMULPORT, CStr(generatePort)))
        .chkIgnorT = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_USEIGNOR, "0"))
        .txtIgnorT = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_IGNORTIME, "10"))
        .SaveList = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_SAVELIST, "1"))
        .chkUseScrape = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_USESCRAPE, "0"))
        .chkIgnorServerError = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_IGNORSERVERR, "0"))
        .chkIgnorSocketError = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_IGNORSOCKETERR, "0"))
        .txtConnectTries = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_CONNTRIES, "20"))
        .chkFroze = (val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_FROZE, "1")) = 1)
        .chkStepModeD = (val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_STEPMODED, "0")) = 1)
        .chkStepModeU = (val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_STEPMODEU, "0")) = 1)
        .txtSMDown = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_STEPMODEDVAL, "0"))
        .txtSMUp = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_STEPMODEUVAL, "0"))
        .txtHave = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_EMULHAVE, "0"))
        .chkSameHash = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_SAMEHASH, "0"))
        '=====
        .chkRemote = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_REMOTEUSE, "0"))
        .txtRCPort = val(GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_REMOTEPORT, "80"))
        .txtRCPass = GetKeyValue(HKEY_LOCAL, MYPATH, REG_SETTINGS_REMOTEPASS, vbNullString)
    End With
End Sub

Public Function SaveWindowPos(somefrm As Form, regname As String)
    If somefrm.WindowState = 0 Then
        WriteKey HKEY_LOCAL, MYPATH, regname & "_x", CStr(somefrm.Width)
        WriteKey HKEY_LOCAL, MYPATH, regname & "_y", CStr(somefrm.Height)
    End If
End Function

Public Function LoadWindowPos(somefrm As Form, regname As String, defx As Long, defy As Long)
    Dim XX As Long, yy As Long
    
    XX = val(GetKeyValue(HKEY_LOCAL, MYPATH, regname & "_x", CStr(defx)))
    yy = val(GetKeyValue(HKEY_LOCAL, MYPATH, regname & "_y", CStr(defy)))
    somefrm.Width = XX
    somefrm.Height = yy
End Function

Public Function FileExist(ByVal Fname As String) As Boolean
    Dim lRetVal As Long
    Dim OfSt As OFSTRUCT

    lRetVal = OpenFile(Fname, OfSt, OF_EXIST)
    FileExist = (lRetVal <> HFILE_ERROR)
End Function

Public Function getRandomN(minVal As Single, maxVal As Single) As Single
    Randomize Timer
    getRandomN = (maxVal - minVal) * Rnd(Timer) + minVal
End Function

Public Sub MakeAssociation(isit As Boolean)
    Dim tmp As String
    
    tmp = GetKeyValue(HKEY_CLASSES_ROOT, ".torrent", vbNullString, vbNullString)
    If Len(tmp) > 0 Then
        If isit Then
            WriteKey HKEY_CLASSES_ROOT, tmp & "\shell\bitproxy", vbNullString, getCStr(1, "Open with BitTorrent Proxy")
            WriteKey HKEY_CLASSES_ROOT, tmp & "\shell\bitproxy\command", vbNullString, Chr(34) & CPath & App.EXEName & ".exe"" %1"
        Else
            DeleteKey HKEY_CLASSES_ROOT, tmp & "\shell\bitproxy", "command"
            DeleteKey HKEY_CLASSES_ROOT, tmp & "\shell", "bitproxy"
        End If
    Else
        If isit Then _
            MsgBox getCStr(41, "'torren' extension not associated with any program!"), vbCritical, App.Title
    End If
End Sub

Public Sub TimerProc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
    KillTimer MainForm.hwnd, 0
    CheckUpdate False
End Sub

Public Sub AutoCloserProc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
    KillTimer hwnd, 123 'Debug
    'MainForm.Socket2_Close
End Sub

Function isInList(strSearchString As String, lHwndListbox As Long) As Boolean
    Const CB_FINDSTRING = &H14C
    isInList = (SendMessage(lHwndListbox, CB_FINDSTRING, -1, ByVal strSearchString) >= 0)
End Function

Function GetComboIndex(strSearchString As String, lHwndListbox As Long) As Long
    Const CB_FINDSTRING = &H14C
    GetComboIndex = SendMessage(lHwndListbox, CB_FINDSTRING, -1, ByVal strSearchString)
End Function

Function getHost(data As String) As String
    Dim i As Long, i2 As Long
    
    i = InStr(data, "://") + 3
    i2 = InStr(i, data, "/") - i
    getHost = Mid(data, i, i2)
End Function

Function MakeErrorLog(dt As String)
    On Error Resume Next
    
    If FileExist(CPath & ERRORLOG) Then Open CPath & ERRORLOG For Append As #8 _
        Else Open CPath & ERRORLOG For Output As #8
    Print #8, Now
    Print #8, vbNullString
    Print #8, dt
    Close 8
End Function

Public Sub ShowTray(frm As Form, szTip As String, hIcon As Long)
    Dim NID As NOTIFYICONDATA
    With NID
        .cbSize = Len(NID)
        .hIcon = hIcon
        .hwnd = frm.hwnd
        .szTip = szTip & vbNullChar
        .uCallbackMessage = TRAY_BACK
        .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
        .uID = 1&
    End With
    Shell_NotifyIcon NIM_ADD, NID
End Sub

Public Sub ChangeTrayTip(frm As Form, szTip As String)
    Dim NID As NOTIFYICONDATA
    With NID
        .cbSize = Len(NID)
        .hwnd = frm.hwnd
        .szTip = szTip & vbNullChar
        .uFlags = NIF_TIP
        .uID = 1&
    End With
    Shell_NotifyIcon NIM_MODIFY, NID
End Sub

Public Sub KillTray(frm As Form)
    Dim NID As NOTIFYICONDATA
    With NID
        .cbSize = Len(NID)
        .hwnd = frm.hwnd
        .uID = 1&
    End With
    Shell_NotifyIcon NIM_DELETE, NID
End Sub

Public Function PopupMsgTray(frm As Form, Message As String, Title As String, msgType As Long)
    Dim NID As NOTIFYICONDATA
    With NID
        .cbSize = Len(NID)
        .hwnd = frm.hwnd
        .uID = 1&
        .uFlags = NIF_INFO
        .szTip = Title & vbNullChar
        .dwState = 0
        .dwStateMask = 0
        .szInfo = Message & Chr(0)
        .szInfoTitle = Title & Chr(0)
        .dwInfoFlags = msgType
    End With
    Shell_NotifyIcon NIM_MODIFY, NID
End Function

Public Function FindPrevWindow() As Long
    Dim rez As Long, tmp As String, i As Long
    Const MAXLEN = 255
    Const MCLS = "ThunderRT5MDIForm"
    Const WNAME = "BitTorrent Proxy"
    
    FindPrevWindow = 0
    
    rez = FindWindow(MCLS, vbNullString)
    Do While rez <> 0
        tmp = Space(MAXLEN)
        i = GetWindowText(rez, tmp, MAXLEN)
        tmp = Left(tmp, i)
        If Left(tmp, Len(WNAME)) = WNAME Then _
            FindPrevWindow = rez
        rez = FindWindowEx(0, rez, MCLS, vbNullString)
    Loop
End Function

Public Function GetFileNameNoExt(nFile As String) As String
    Dim intPos As String
    Dim intPosSave As String
    
    If InStr(nFile, ".") = 0 Then
        GetFileNameNoExt = nFile
        Exit Function
    End If
    
    intPos = 1
    Do
        intPos = InStr(intPos, nFile, ".")
        If intPos = 0 Then
            Exit Do
        Else
            intPos = intPos + 1
            intPosSave = intPos - 2
        End If
    Loop

    GetFileNameNoExt = Left(nFile, intPosSave)
End Function

Public Function GetByteFormat(num As Currency) As String
    Dim rez As Single
    
    If num < 0 Then
        GetByteFormat = "N/A"
        Exit Function
    End If
    If (num >= 0) And (num < CKILO) Then
        GetByteFormat = CStr(num) & getCStr(119, " byte")
        Exit Function
    End If
    If (num >= CKILO) And (num < CMEGA) Then
        rez = num / CKILO
        GetByteFormat = Format(rez, "#0.0") & " Kb"
        Exit Function
    End If
    If (num >= CMEGA) And (num < CGIGA) Then
        rez = num / CMEGA
        GetByteFormat = Format(rez, "#0.0") & " Mb"
        Exit Function
    End If
    If num >= CGIGA Then
        rez = num / CGIGA
        GetByteFormat = Format(rez, "#0.0") & " Gb"
        Exit Function
    End If
End Function

Public Function GetTimeFormat(numSec As Long) As String
    Dim ss As String
    
    ss = CStr(numSec Mod 60)
    If Len(ss) = 1 Then ss = "0" & ss
    GetTimeFormat = CStr(numSec \ 60) & ":" & ss
End Function

Public Sub str_replace(ByVal Find As String, ByVal Replace As String, ByRef Expression As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare)
    On Error Resume Next
    Dim l As Long
    Dim lenR As Long
    Dim p1 As Long
    Dim p2 As Long
    Dim p21 As Long
    Dim s As String
    
    l = Len(Find)
    If (l = 0) Then Exit Sub
    
    lenR = Len(Replace)
    If (lenR = l) Then
        p1 = 1
        p2 = InStr(p1, Expression, Find, Compare)
        Do While (p2)
            Mid$(Expression, p1) = Mid$(Expression, p1, p2 - p1)
            Mid$(Expression, p2) = Replace
            p1 = p2 + l
            p2 = InStr(p1, Expression, Find, Compare)
        Loop
        Exit Sub
        
    ElseIf (lenR > l) Then
        s = Space$(Len(Expression) + (Len(Expression) \ l) * (lenR - l))
        
    Else
        s = Space$(Len(Expression))
    End If
    
    p21 = 1
    p1 = 1
    p2 = InStr(p1, Expression, Find, Compare)
    Do While (p2)
        Mid$(s, p21) = Mid(Expression, p1, p2 - p1)
        p21 = p21 + p2 - p1
        Mid$(s, p21) = Replace
        p21 = p21 + lenR
        p1 = p2 + l
        p2 = InStr(p1, Expression, Find, Compare)
    Loop
    Mid$(s, p21) = Mid$(Expression, p1)
    p21 = p21 + Len(Mid$(Expression, p1))
    s = Left$(s, p21 - 1)
    Expression = s
End Sub

Public Sub AddToFileLog(desc As String, sData As String)
    #If debugver = 1 Then
        Print #1, desc & "(" & Date$ & " " & time$ & "):"
        Print #1, sData & vbCrLf
    #End If
End Sub

Public Function IsFormLoaded(ByVal form_name As String) As Boolean
    Dim frm As Form

    IsFormLoaded = False
    For Each frm In Forms
        If frm.name = form_name Then
            IsFormLoaded = True
            Exit For
        End If
    Next frm
End Function

Public Function getINIString(iSec As String, iKey As String, FileN As String) As String
    Dim rez As Long
    Dim tmp As String
    Const STRMAX = 500
    
    tmp = Space(STRMAX)
    rez = GetPrivateProfileString(iSec, iKey, vbNullString, tmp, STRMAX, FileN)
    If rez = 0 Then
        getINIString = vbNullString
    Else
        getINIString = Left(tmp, rez)
    End If
End Function

Public Function MakeHash(tHash As String) As String
    Dim tmp As String, i As Integer
    
    tmp = vbNullString
    tHash = Left(tHash, 40)
    For i = 1 To Len(tHash) Step 2
        tmp = tmp & "%" & Mid(tHash, i, 2)
    Next
    MakeHash = tmp
End Function

Public Function generatePort() As Long
    Randomize Timer
    generatePort = CLng(Rnd * 10000) + 10000
End Function

Public Function getAutoRun() As Integer
    Dim tmp As String
    tmp = GetKeyValue(HKEY_CURRENT_USER, AUTORUN, "BitProxy", vbNullString)
    getAutoRun = IIf((Len(tmp) = 0), 0, 1)
End Function

Public Sub OpenHomepage()
    ShellExecute MainForm.hwnd, vbNullString, HTTP & SITEURL, vbNullString, App.Path, vbNormalFocus
End Sub

Public Sub OpenForum()
    ShellExecute MainForm.hwnd, vbNullString, HTTP & "forum." & SITEURL & FORUMURL, vbNullString, App.Path, vbNormalFocus
End Sub

Public Function MakeShortName(longName As String) As String
    Dim sShortFile As String * 250
    Dim lResult As Long
    
    lResult = GetShortPathName(longName, sShortFile, Len(sShortFile))
    MakeShortName = StrConv(Left(sShortFile, lResult * 2), vbFromUnicode)
End Function

Public Function StringToDouble(str As String) As Double
    str_replace ",", ".", str
    StringToDouble = val(str)
End Function

Public Sub FilePoolClear()
    Set Global_FilePool = Nothing
    Set Global_FilePool = New Collection
End Sub

Public Sub FilePoolFromCommandLine()
    Dim buffer() As Byte
    Dim StrLen As Long
    Dim Cmd As Long
    Dim tmp As String, i As Integer
    Dim sShortFile As String * 250
    Dim lResult As Long
    
    FilePoolClear
    Cmd = GetCommandLine
    If Cmd Then
        StrLen = lstrlenW(Cmd) * 2
        If StrLen Then
            ReDim buffer(0 To (StrLen - 1)) As Byte
            CopyMemory buffer(0), ByVal Cmd, StrLen
            tmp = buffer
        
            i = InStr(1, tmp, ChrW(34)) 'Is program path quoted?
            If i = 1 Then
                i = InStr(2, tmp, ChrW(34))
                If i > 0 Then tmp = Mid(tmp, i + 2)
            Else
                i = InStr(1, tmp, ChrW(32))
                If i > 0 Then tmp = Mid(tmp, i + 1)
            End If
            tmp = Trim(tmp)
            If Left(tmp, 1) = """" And Right(tmp, 1) = """" Then _
                tmp = Mid(tmp, 2, Len(tmp) - 2)
            tmp = StrConv(tmp, vbUnicode)
            
            If (Len(tmp) > 0) Then _
                Global_FilePool.Add MakeShortName(tmp)
        End If
    End If
End Sub

Public Sub FilePoolFromMainDrop(hDrop As Long)
    Const MAX_PATH = 256
    Dim rez As Long
    Dim num As Long, i As Long
    Dim Filename As String * MAX_PATH
    Dim tmp As String
    
    FilePoolClear
    
    num = DragQueryFile(hDrop, True, Filename, MAX_PATH)
    For i = 1 To num
        rez = DragQueryFile(hDrop, i - 1, Filename, MAX_PATH)
        tmp = Left(Filename, rez * 2)
        If StrConv(Right(tmp, Len(TRR) * 2), vbFromUnicode) = TRR Then _
            Global_FilePool.Add MakeShortName(tmp)
   Next
   DragFinish (hDrop)
End Sub

Public Sub FilePoolFromEmulationDrop(data As MSComctlLib.DataObject)
    Dim i As Long, tmp As String
    
    FilePoolClear

    If data.GetFormat(vbCFFiles) Then
        For i = 1 To data.Files.Count
            If Right(data.Files(i), Len(TRR)) = TRR Then _
                Global_FilePool.Add data.Files(i)
        Next
    End If
End Sub

Public Function VB6_Split(ByVal sIn As String, Optional sDelim As String, Optional nLimit As Long = -1, Optional bCompare As VbCompareMethod = vbBinaryCompare) As Variant
    Dim sRead As String, sOut() As String, nC As Integer
    
    If Len(sDelim) = 0 Then
        VB6_Split = sIn
        Exit Function
    End If
    
    sRead = ReadUntil(sIn, sDelim, bCompare)
    Do
        ReDim Preserve sOut(nC)
        sOut(nC) = sRead
        nC = nC + 1
        If nLimit <> -1 And nC >= nLimit Then _
            Exit Do
        sRead = ReadUntil(sIn, sDelim)
    Loop While sRead <> ""
    ReDim Preserve sOut(nC)
    sOut(nC) = sIn
    VB6_Split = sOut
End Function

Public Function ReadUntil(ByRef sIn As String, sDelim As String, Optional bCompare As VbCompareMethod = vbBinaryCompare) As String
    Dim nPos As String
    nPos = InStr(1, sIn, sDelim, bCompare)
    If nPos > 0 Then
        ReadUntil = Left(sIn, nPos - 1)
        sIn = Mid(sIn, nPos + Len(sDelim))
    End If
End Function

Public Function fix16K(num As Currency) As Currency
    Const MY16K = 16384#
    fix16K = Int(num / MY16K) * MY16K
End Function

