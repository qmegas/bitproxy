VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDownloadItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private s_info As New Collection

Private mvarName As String
Private mvarURL As String
Private mvarPort As Long
Private mvarMode As Integer
Private mvarm1_coef As Double
Private mvarm1_isdown As Boolean
Private mvarm1_downcoef As Double
Private mvarm2_from As Single
Private mvarm2_to As Single
Private mvaruseClient As Boolean
Private mvarClientID As Long
Private mvarStatus As Integer
Private mvarupdatesList As Collection
Private mvarId As Long
Private mvarSocketOutBuffer As String
Private mvarconErr As Boolean
Private mvarTryingToConnect As Boolean
Private mvarWeQuit As Boolean
Private mvarStepMode As Integer
Private mvarlast_info_hash As String
Private mvarlast_peer_id As String
Private mvarlast_uploaded As Currency
Private mvarlast_downloaded As Currency
Private mvarlast_event As String
Private mvarlast_real_upload As Currency
Private mvarlast_port As Long
Private mvarlast_left As Currency
Private mvarlast_numwant As Long
Private mvarlast_compact As Long
Private mvarlast_no_peer_id As Long
Private mvarlast_key As String
Private mvarlast_cur_time As Long
Private mvarsmart_use As Boolean
Private mvarsmart_a As Long
Private mvarsmart_p As Long
Private mvarlast_cur_speed As Long
Private mvarm1_dwnnotsend As Boolean

Public Property Let m1_dwnnotsend(ByVal vData As Boolean)
    mvarm1_dwnnotsend = vData
End Property

Public Property Get m1_dwnnotsend() As Boolean
    m1_dwnnotsend = mvarm1_dwnnotsend
End Property

Public Property Let smart_p(ByVal vData As Long)
    mvarsmart_p = vData
End Property

Public Property Get smart_p() As Long
    smart_p = mvarsmart_p
End Property

Public Property Let smart_a(ByVal vData As Long)
    mvarsmart_a = vData
End Property

Public Property Get smart_a() As Long
    smart_a = mvarsmart_a
End Property

Public Property Let smart_use(ByVal vData As Boolean)
    mvarsmart_use = vData
End Property

Public Property Get smart_use() As Boolean
    smart_use = mvarsmart_use
End Property

Public Property Let last_cur_speed(ByVal vData As Long)
    mvarlast_cur_speed = vData
End Property

Public Property Get last_cur_speed() As Long
    smart_p = mvarlast_cur_speed
End Property

Public Property Let last_cur_time(ByVal vData As Long)
    mvarlast_cur_time = vData
End Property

Public Property Get last_cur_time() As Long
    last_cur_time = mvarlast_cur_time
End Property

Public Property Let last_key(ByVal vData As String)
    mvarlast_key = vData
End Property

Public Property Get last_key() As String
    last_key = mvarlast_key
End Property

Public Property Let last_no_peer_id(ByVal vData As Long)
    mvarlast_no_peer_id = vData
End Property

Public Property Get last_no_peer_id() As Long
    last_no_peer_id = mvarlast_no_peer_id
End Property

Public Property Let last_compact(ByVal vData As Long)
    mvarlast_compact = vData
End Property

Public Property Get last_compact() As Long
    last_compact = mvarlast_compact
End Property

Public Property Let last_numwant(ByVal vData As Long)
    mvarlast_numwant = vData
End Property

Public Property Get last_numwant() As Long
    last_numwant = mvarlast_numwant
End Property

Public Property Let last_left(ByVal vData As Currency)
    mvarlast_left = vData
End Property

Public Property Get last_left() As Currency
    last_left = mvarlast_left
End Property

Public Property Let last_port(ByVal vData As Long)
    mvarlast_port = vData
End Property

Public Property Get last_port() As Long
    last_port = mvarlast_port
End Property

Public Property Let last_real_upload(ByVal vData As Currency)
    mvarlast_real_upload = vData
End Property

Public Property Get last_real_upload() As Currency
    last_real_upload = mvarlast_real_upload
End Property

Public Property Let last_event(ByVal vData As String)
    mvarlast_event = vData
End Property

Public Property Get last_event() As String
    last_event = mvarlast_event
End Property

Public Property Let last_downloaded(ByVal vData As Currency)
    mvarlast_downloaded = vData
End Property

Public Property Get last_downloaded() As Currency
    last_downloaded = mvarlast_downloaded
End Property

Public Property Let last_uploaded(ByVal vData As Currency)
    mvarlast_uploaded = vData
End Property

Public Property Get last_uploaded() As Currency
    last_uploaded = mvarlast_uploaded
End Property

Public Property Let last_peer_id(ByVal vData As String)
    mvarlast_peer_id = vData
End Property

Public Property Get last_peer_id() As String
    last_peer_id = mvarlast_peer_id
End Property

Public Property Let last_info_hash(ByVal vData As String)
    mvarlast_info_hash = vData
End Property

Public Property Get last_info_hash() As String
    last_info_hash = mvarlast_info_hash
End Property

Public Property Let StepMode(ByVal vData As Integer)
    mvarStepMode = vData
End Property

Public Property Get StepMode() As Integer
    StepMode = mvarStepMode
End Property

Public Property Let TryingToConnect(ByVal vData As Boolean)
    mvarTryingToConnect = vData
End Property

Public Property Get TryingToConnect() As Boolean

    TryingToConnect = mvarTryingToConnect
End Property

Public Property Let conErr(ByVal vData As Boolean)
    mvarconErr = vData
End Property

Public Property Get conErr() As Boolean
    conErr = mvarconErr
End Property

Public Property Let SocketOutBuffer(ByVal vData As String)
    mvarSocketOutBuffer = vData
End Property

Public Property Get SocketOutBuffer() As String
    SocketOutBuffer = mvarSocketOutBuffer
End Property

Public Property Let Id(ByVal vData As Integer)
    mvarId = vData
End Property

Public Property Get Id() As Integer
    Id = mvarId
End Property

Public Property Let Status(ByVal vData As Integer)
    mvarStatus = vData
End Property

Public Property Get Status() As Integer
    Status = mvarStatus
End Property

Public Property Let ClientID(ByVal vData As Long)
    mvarClientID = vData
End Property

Public Property Get ClientID() As Long
    ClientID = mvarClientID
End Property

Public Property Let useClient(ByVal vData As Boolean)
    mvaruseClient = vData
End Property

Public Property Get useClient() As Boolean
    useClient = mvaruseClient
End Property

Public Property Let m2_to(ByVal vData As Single)
    mvarm2_to = vData
End Property

Public Property Get m2_to() As Single
    m2_to = mvarm2_to
End Property

Public Property Let m2_from(ByVal vData As Single)
    mvarm2_from = vData
End Property

Public Property Get m2_from() As Single
    m2_from = mvarm2_from
End Property

Public Property Let m1_downcoef(ByVal vData As Double)
    mvarm1_downcoef = vData
End Property

Public Property Get m1_downcoef() As Double
    m1_downcoef = mvarm1_downcoef
End Property

Public Property Let m1_isdown(ByVal vData As Boolean)
    mvarm1_isdown = vData
End Property

Public Property Get m1_isdown() As Boolean
    m1_isdown = mvarm1_isdown
End Property

Public Property Let m1_coef(ByVal vData As Double)
    mvarm1_coef = vData
End Property

Public Property Get m1_coef() As Double
    m1_coef = mvarm1_coef
End Property

Public Property Let Mode(ByVal vData As Integer)
    mvarMode = vData
End Property

Public Property Get Mode() As Integer
    Mode = mvarMode
End Property

Public Property Let port(ByVal vData As Long)
    mvarPort = vData
End Property

Public Property Get port() As Long
    port = mvarPort
End Property

Public Property Let URL(ByVal vData As String)
    mvarURL = vData
End Property

Public Property Get URL() As String
    URL = mvarURL
End Property

Public Property Let name(ByVal vData As String)
    mvarName = vData
End Property

Public Property Get name() As String
    name = mvarName
End Property

Public Sub ClearData()
    Set s_info = Nothing
    Set s_info = New Collection
End Sub

Public Sub AddData(time As Long, upload As Currency, download As Currency)
    s_info.Add CStr(time) & "|" & CStr(upload) & "|" & CStr(download)
End Sub

Public Sub DrawData(gObj As clsGrafik, upl As Boolean, down As Boolean)
    Dim i As Integer, main_time As Long
    Dim upLine As Long, dwnLine As Long
    Dim tm1 As Long, upl1 As Currency, dwn1 As Currency
    Dim tm2 As Long, upl2 As Currency, dwn2 As Currency
    
    'Dim tmp As String 'debug
    
    gObj.ResetData
    If s_info.Count < 2 Then
        gObj.DrawGraphik
        Exit Sub
    End If
    
    'tmp = s_info(1) 'debug
    ParseHistoryData s_info(1), tm1, upl1, dwn1
    main_time = tm1
    If upl Then upLine = gObj.AddLine(vbRed, True)
    If down Then dwnLine = gObj.AddLine(vbBlue, True)
    
    For i = 2 To s_info.Count
        ParseHistoryData s_info(i), tm2, upl2, dwn2
        'tmp = tmp & s_info(i) 'debug
        If upl Then _
            gObj.AddPoint upLine, tm2 - main_time, val(((upl2 - upl1) / (tm2 - tm1)) / 1024)
        If down Then _
            gObj.AddPoint dwnLine, tm2 - main_time, val(((dwn2 - dwn1) / (tm2 - tm1)) / 1024)
            
        tm1 = tm2
        upl1 = upl2
        dwn1 = dwn2
    Next
    'Clipboard.SetText tmp 'debug
    gObj.DrawGraphik
End Sub

Private Sub ParseHistoryData(data As String, time As Long, upl As Currency, dwn As Currency)
    Dim i As Integer, j As Integer
    
    i = InStr(1, data, "|")
    time = val(Left(data, i - 1))
    j = InStr(i + 1, data, "|")
    upl = CCur(Mid(data, i + 1, j - i - 1))
    dwn = CCur(Mid(data, j + 1))
End Sub

