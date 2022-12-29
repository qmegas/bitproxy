VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDownloadManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private curJob As clsDownloadItem
Private mylist As New Collection
Private curIndex As Integer

Public Property Let m1_dwn_notsend(ByVal vData As Boolean)
    curJob.m1_dwnnotsend = vData
End Property

Public Property Get m1_dwn_notsend() As Boolean
    m1_dwn_notsend = curJob.m1_dwnnotsend
End Property

Public Property Let smart_p(ByVal vData As Long)
    curJob.smart_p = vData
End Property

Public Property Get smart_p() As Long
    smart_p = curJob.smart_p
End Property

Public Property Let smart_a(ByVal vData As Long)
    curJob.smart_a = vData
End Property

Public Property Get smart_a() As Long
    smart_a = curJob.smart_a
End Property

Public Property Let smart_use(ByVal vData As Boolean)
    curJob.smart_use = vData
End Property

Public Property Get smart_use() As Boolean
    smart_use = curJob.smart_use
End Property

Public Property Let last_cur_speed(ByVal vData As Long)
    curJob.last_cur_speed = vData
End Property

Public Property Get last_cur_speed() As Long
    smart_p = curJob.last_cur_speed
End Property

Public Property Let last_cur_time(ByVal vData As Long)
    curJob.last_cur_time = vData
End Property

Public Property Get last_cur_time() As Long
    last_cur_time = curJob.last_cur_time
End Property

Public Property Let last_key(ByVal vData As String)
    curJob.last_key = vData
End Property

Public Property Get last_key() As String
    last_key = curJob.last_key
End Property

Public Property Let last_no_peer_id(ByVal vData As Long)
    curJob.last_no_peer_id = vData
End Property

Public Property Get last_no_peer_id() As Long
    last_no_peer_id = curJob.last_no_peer_id
End Property

Public Property Let last_compact(ByVal vData As Long)
    curJob.last_compact = vData
End Property

Public Property Get last_compact() As Long
    last_compact = curJob.last_compact
End Property

Public Property Let last_numwant(ByVal vData As Long)
    curJob.last_numwant = vData
End Property

Public Property Get last_numwant() As Long
    last_numwant = curJob.last_numwant
End Property

Public Property Let last_left(ByVal vData As Currency)
    If (Not curJob.m1_dwnnotsend) Or (curJob.last_event = TORRENT_START) Then _
        curJob.last_left = vData
End Property

Public Property Get last_left() As Currency
    last_left = curJob.last_left
End Property

Public Property Let last_port(ByVal vData As Long)
    curJob.last_port = vData
End Property

Public Property Get last_port() As Long
    last_port = curJob.last_port
End Property

Public Property Let last_real_upload(ByVal vData As Currency)
    curJob.last_real_upload = vData
End Property

Public Property Get last_real_upload() As Currency
    last_real_upload = curJob.last_real_upload
End Property

Public Property Let last_event(ByVal vData As String)
    curJob.last_event = vData
End Property

Public Property Get last_event() As String
    last_event = curJob.last_event
End Property

Public Property Let last_downloaded(ByVal vData As Currency)
    If (Not curJob.m1_dwnnotsend) Or (curJob.last_event = TORRENT_START) Then _
        curJob.last_downloaded = vData
End Property

Public Property Get last_downloaded() As Currency
    last_downloaded = curJob.last_downloaded
End Property

Public Property Let last_uploaded(ByVal vData As Currency)
    curJob.last_uploaded = vData
End Property

Public Property Get last_uploaded() As Currency
    last_uploaded = curJob.last_uploaded
End Property

Public Property Let last_peer_id(ByVal vData As String)
    curJob.last_peer_id = vData
End Property

Public Property Get last_peer_id() As String
    last_peer_id = curJob.last_peer_id
End Property

Public Property Let last_info_hash(ByVal vData As String)
    curJob.last_info_hash = vData
End Property

Public Property Get last_info_hash() As String
    last_info_hash = curJob.last_info_hash
End Property

Public Property Let StepMode(ByVal vData As Integer)
    curJob.StepMode = vData
End Property

Public Property Get StepMode() As Integer
    StepMode = curJob.StepMode
End Property

Public Property Let TryingToConnect(ByVal vData As Boolean)
    curJob.TryingToConnect = vData
End Property

Public Property Get TryingToConnect() As Boolean
    TryingToConnect = curJob.TryingToConnect
End Property

Public Property Let conErr(ByVal vData As Boolean)
    curJob.conErr = vData
End Property

Public Property Get conErr() As Boolean
    conErr = curJob.conErr
End Property

Public Property Let SocketOutBuffer(ByVal vData As String)
    curJob.SocketOutBuffer = vData
End Property

Public Property Get SocketOutBuffer() As String
    SocketOutBuffer = curJob.SocketOutBuffer
End Property

Public Property Let Status(ByVal vData As Integer)
    curJob.Status = vData
End Property

Public Property Get Status() As Integer
    Status = curJob.Status
End Property

Public Property Let ClientID(ByVal vData As Long)
    curJob.ClientID = vData
End Property

Public Property Get ClientID() As Long
    ClientID = curJob.ClientID
End Property

Public Property Let useClient(ByVal vData As Boolean)
    curJob.useClient = vData
End Property

Public Property Get useClient() As Boolean
    useClient = curJob.useClient
End Property

Public Property Let m2_to(ByVal vData As Single)
    curJob.m2_to = vData
End Property

Public Property Get m2_to() As Single
    m2_to = curJob.m2_to
End Property

Public Property Let m2_from(ByVal vData As Single)
    curJob.m2_from = vData
End Property

Public Property Get m2_from() As Single
    m2_from = curJob.m2_from
End Property

Public Property Let m1_downcoef(ByVal vData As Double)
    curJob.m1_downcoef = vData
End Property

Public Property Get m1_downcoef() As Double
    m1_downcoef = curJob.m1_downcoef
End Property

Public Property Let m1_isdown(ByVal vData As Boolean)
    curJob.m1_isdown = vData
End Property

Public Property Get m1_isdown() As Boolean
    m1_isdown = curJob.m1_isdown
End Property

Public Property Let m1_coef(ByVal vData As Double)
    curJob.m1_coef = vData
End Property

Public Property Get m1_coef() As Double
    m1_coef = curJob.m1_coef
End Property

Public Property Let Mode(ByVal vData As Integer)
    curJob.Mode = vData
End Property

Public Property Get Mode() As Integer
    Mode = curJob.Mode
End Property

Public Property Let port(ByVal vData As Long)
    curJob.port = vData
End Property

Public Property Get port() As Long
    port = curJob.port
End Property

Public Property Let URL(ByVal vData As String)
    curJob.URL = vData
End Property

Public Property Get URL() As String
    URL = curJob.URL
End Property

Public Property Let name(ByVal vData As String)
    If Len(vData) = 0 Then vData = "Unnamed"
    curJob.name = vData
End Property

Public Property Get name() As String
    name = curJob.name
End Property

Public Sub DrawFullList(lv As ListView)
    Dim tmp As ListItem
    Dim t_icon As Integer
    
    lv.ListItems.Clear
    For Each curJob In mylist
        Set tmp = lv.ListItems.Add(, "k" & CStr(curJob.Id), curJob.name, , IIf(curJob.Status = STATUS_ON, 1, 2))
        tmp.SubItems(1) = curJob.URL
        tmp.SubItems(2) = CStr(curJob.port)
        tmp.SubItems(3) = IIf(curJob.Mode = 1, getCStr(200, "Static"), getCStr(201, "Dynamic"))
    Next
End Sub

'Redraw current
Public Sub DrawSingle(lv As ListView)
    Dim i As Long
    
    For i = 1 To lv.ListItems.Count
        If lv.ListItems(i).key = "k" & CStr(curJob.Id) Then
            lv.ListItems(i).SmallIcon = IIf(curJob.Status = STATUS_ON, 1, 2)
            lv.ListItems(i).Text = curJob.name
            lv.ListItems(i).SubItems(1) = curJob.URL
            lv.ListItems(i).SubItems(2) = CStr(curJob.port)
            lv.ListItems(i).SubItems(3) = IIf(curJob.Mode = 1, getCStr(200, "Static"), getCStr(201, "Dynamic"))
        End If
    Next
End Sub

Public Function SelectJob(jID As Integer) As Boolean
    SelectJob = False
    If Not (curJob Is Nothing) Then
        If curJob.Id = jID Then
            SelectJob = True
            Exit Function
        End If
    End If
    
    For Each curJob In mylist
        If curJob.Id = jID Then
            SelectJob = True
            Exit For
        End If
    Next
End Function

Private Function getFreePort()
    Dim max_port As Integer, tmp As clsDownloadItem
    
    If curIndex = 0 Then
        getFreePort = prgSettings.txtPort
    Else
        max_port = prgSettings.txtPort - 1
        For Each tmp In mylist
            If tmp.port > max_port Then _
                max_port = tmp.port
        Next
        getFreePort = max_port + 1
    End If
End Function

Public Function AddJob() As Integer
    Dim i As New clsDownloadItem
    
    curIndex = curIndex + 1
    With prgSettings
        i.Id = curIndex
        i.port = getFreePort
        i.Mode = IIf(.optMode = 1, 1, 2)
        i.m1_coef = .txtUpload
        i.m1_isdown = (.chkDown = 1)
        i.m1_downcoef = .txtDown
        i.m1_dwnnotsend = (.chkDnotsend = 1)
        i.m2_from = val(.txtM2from)
        i.m2_to = val(.txtM2to)
        i.smart_use = (.chkSmart = 1)
        i.smart_a = .txtSmartA
        i.smart_p = .txtSmartP
        i.useClient = (.chkVersion = 1)
        i.ClientID = .cmbVersion
        i.Status = STATUS_OFF
        i.TryingToConnect = False
        i.conErr = False
    End With
    
    Load frmMain.Socket1(curIndex)
    Load frmMain.Socket2(curIndex)
    Load frmMain.Socket3(curIndex)
    Load frmMain.Socket4(curIndex)
    Load frmMain.lstLog(curIndex)
    
    mylist.Add i
    
    AddJob = curIndex
End Function

Public Function DeleteJob(jID As Integer) As Boolean
    Dim i As Integer, tmp As clsDownloadItem
    
    DeleteJob = False
    For i = 1 To mylist.Count
        Set tmp = mylist.Item(i)
        If tmp.Id = jID Then
            DeleteJob = True
            mylist.Remove i
            Unload frmMain.Socket1(jID)
            Unload frmMain.Socket2(jID)
            Unload frmMain.Socket3(jID)
            Unload frmMain.Socket4(jID)
            Unload frmMain.lstLog(jID)
            Exit For
        End If
    Next
    
    Set tmp = Nothing
End Function

Public Sub SaveState()
    Dim i As Long, cnt As Long
    On Error Resume Next
    
    If prgSettings.SaveList = 0 Then Exit Sub
    
    cnt = 0
    If FileExist(CPath & SAVEDOWNF) Then
        Kill CPath & SAVEDOWNF
        err.Clear
    End If
    
    i = FreeFile
    Open CPath & SAVEDOWNF For Output As #i
    
    Print #i, "[General]"
    Print #i, "ID=" & BP10
    Print #i, "Total=" & CStr(mylist.Count)
    
    For Each curJob In mylist
        cnt = cnt + 1
        Print #i, "[" & INIITEM & CStr(cnt) & "]"
        'General
        Print #i, "Name=" & curJob.name
        Print #i, "Tracker=" & curJob.URL
        Print #i, "Port=" & CStr(curJob.port)
        'Mode
        Print #i, "Mode=" & CStr(curJob.Mode)
        If curJob.Mode = 1 Then
            Print #i, "Upload=" & Trim(str(curJob.m1_coef))
            Print #i, "UseDownload=" & CStr(IIf(curJob.m1_isdown, 1, 0))
            Print #i, "Download=" & Trim(str(curJob.m1_downcoef))
            Print #i, "DownNotSend=" & CStr(IIf(curJob.m1_dwnnotsend, 1, 0))
        Else
            Print #i, "From=" & Trim(str(curJob.m2_from))
            Print #i, "To=" & Trim(str(curJob.m2_to))
        End If
        'Smart system
        Print #1, "SmartUse=" & CStr(IIf(curJob.smart_use, 1, 0))
        Print #1, "SmartA=" & CStr(curJob.smart_a)
        Print #1, "SmartP=" & CStr(curJob.smart_p)
        'Client
        Print #i, "CVersion=" & CStr(IIf(curJob.useClient, 1, 0))
        Print #i, "Client=" & CStr(curJob.ClientID)
    Next
    
    Close i
End Sub

Public Sub LoadState()
    Dim i As Long, j As Long, k As Long
    Dim curID As Integer
    
    i = val(getINIString("General", "Total", CPath & SAVEDOWNF))
    If i = 0 Then Exit Sub
    
    For j = 1 To i
        curID = Me.AddJob
        Me.SelectJob curID
        
        'Caption
        Me.name = getINIString(INIITEM & CStr(j), "Name", CPath & SAVEDOWNF)
        Me.URL = getINIString(INIITEM & CStr(j), "Tracker", CPath & SAVEDOWNF)
        Me.port = val(getINIString(INIITEM & CStr(j), "Port", CPath & SAVEDOWNF))
        'Mode
        k = val(getINIString(INIITEM & CStr(j), "Mode", CPath & SAVEDOWNF))
        If k <> 1 Then k = 2
        Me.Mode = k
        If k = 1 Then
            Me.m1_coef = val(getINIString(INIITEM & CStr(j), "Upload", CPath & SAVEDOWNF))
            Me.m1_downcoef = val(getINIString(INIITEM & CStr(j), "Download", CPath & SAVEDOWNF))
            Me.m1_isdown = (val(getINIString(INIITEM & CStr(j), "UseDownload", CPath & SAVEDOWNF)) = 1)
            Me.m1_dwn_notsend = (val(getINIString(INIITEM & CStr(j), "DownNotSend", CPath & SAVEDOWNF)) = 1)
        Else
            Me.m2_from = val(getINIString(INIITEM & CStr(j), "From", CPath & SAVEDOWNF))
            Me.m2_to = val(getINIString(INIITEM & CStr(j), "To", CPath & SAVEDOWNF))
        End If
        'Smart system
        Me.smart_use = (getINIString(INIITEM & CStr(j), "SmartUse", CPath & SAVEDOWNF) = "1")
        Me.smart_a = val(getINIString(INIITEM & CStr(j), "SmartA", CPath & SAVEDOWNF))
        Me.smart_p = val(getINIString(INIITEM & CStr(j), "SmartP", CPath & SAVEDOWNF))
        'Client
        Me.useClient = (val(getINIString(INIITEM & CStr(j), "CVersion", CPath & SAVEDOWNF)) = 1)
        k = val(getINIString(INIITEM & CStr(j), "Client", CPath & SAVEDOWNF))
        If (k >= 0) And (clnt.getCount >= k) Then
            Me.ClientID = k
        Else
            Me.ClientID = 0
            Me.useClient = False
        End If
    Next
    Me.DrawFullList frmMain.lv1
    frmMain.FillForms
End Sub

Public Sub AddHistoryData()
    If curJob.last_event = TORRENT_START Then curJob.ClearData
    curJob.AddData ToUnixTime(Now), curJob.last_uploaded, curJob.last_downloaded
    If frmMain.ts.SelectedItem.Index = GRAPH_TAB Then _
        Me.DrawGraph
End Sub

Public Sub DrawGraph()
    If Not (curJob Is Nothing) Then _
        curJob.DrawData frmMain.GraphikObj, (frmMain.chkGUpload.Value = 1), (frmMain.chkGDownload = 1)
End Sub

Private Sub Class_Initialize()
    curIndex = 0
End Sub

Private Sub Class_Terminate()
    Set mylist = Nothing
End Sub
