VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEmulationManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private TotalIDs As Long
Private mycoll As New Collection
Private curJob As clsEmulationItem

Public Property Let tAddedTime(ByVal vData As String)
    curJob.tAddedTime = vData
End Property

Public Property Get tAddedTime() As String
    tAddedTime = curJob.tAddedTime
End Property

Public Property Let iEvent(ByVal vData As String)
    curJob.iEvent = vData
End Property

Public Property Get iEvent() As String
    iEvent = curJob.iEvent
End Property

Public Property Let tFrozen(ByVal vData As Boolean)
    curJob.tFrozen = vData
End Property

Public Property Get tFrozen() As Boolean
    tFrozen = curJob.tFrozen
End Property

Public Property Let tFrozenSkip(ByVal vData As Boolean)
    curJob.tFrozenSkip = vData
End Property

Public Property Get tFrozenSkip() As Boolean
    tFrozenSkip = curJob.tFrozenSkip
End Property

Public Property Let tmpDataRecieved(ByVal vData As String)
    curJob.tmpDataRecieved = vData
End Property

Public Property Get tmpDataRecieved() As String
    tmpDataRecieved = curJob.tmpDataRecieved
End Property

Public Property Let tConnectTries(ByVal vData As Long)
    curJob.tConnectTries = vData
End Property

Public Property Get tConnectTries() As Long
    tConnectTries = curJob.tConnectTries
End Property

Public Property Let tStepUploadTime(ByVal vData As Long)
    curJob.tStepUploadTime = vData
End Property

Public Property Get tStepUploadTime() As Long
    tStepUploadTime = curJob.tStepUploadTime
End Property

Public Property Let tStepDownloadTime(ByVal vData As Long)
    curJob.tStepDownloadTime = vData
End Property

Public Property Get tStepDownloadTime() As Long
    tStepDownloadTime = curJob.tStepDownloadTime
End Property

Public Property Let tUseStepUpload(ByVal vData As Boolean)
    curJob.tUseStepUpload = vData
End Property

Public Property Get tUseStepUpload() As Boolean
    tUseStepUpload = curJob.tUseStepUpload
End Property

Public Property Let tUseStepDownload(ByVal vData As Boolean)
    curJob.tUseStepDownload = vData
End Property

Public Property Get tUseStepDownload() As Boolean
    tUseStepDownload = curJob.tUseStepDownload
End Property

Public Property Let tScrapeStatus(ByVal vData As Integer)
    curJob.tScrapeStatus = vData
End Property

Public Property Get tScrapeStatus() As Integer
    tScrapeStatus = curJob.tScrapeStatus
End Property

Public Property Let tScrapeTime(ByVal vData As Long)
    curJob.tScrapeTime = vData
End Property

Public Property Get tScrapeTime() As Long
    tScrapeTime = curJob.tScrapeTime
End Property

Public Property Let tUseScrape(ByVal vData As Boolean)
    curJob.tUseScrape = vData
End Property

Public Property Get tUseScrape() As Boolean
    tUseScrape = curJob.tUseScrape
End Property

Public Property Let tStopValue(ByVal vData As Currency)
    curJob.tStopValue = vData
End Property

Public Property Get tStopValue() As Currency
    tStopValue = curJob.tStopValue
End Property

Public Property Let tStopMode(ByVal vData As Integer)
    curJob.tStopMode = vData
End Property

Public Property Get tStopMode() As Integer
    tStopMode = curJob.tStopMode
End Property

Public Property Let tIgnorTime(ByVal vData As Integer)
    curJob.tIgnorTime = vData
End Property

Public Property Get tIgnorTime() As Integer
    tIgnorTime = curJob.tIgnorTime
End Property

Public Property Let tUseIgnorTime(ByVal vData As Boolean)
    curJob.tUseIgnorTime = vData
End Property

Public Property Get tUseIgnorTime() As Boolean
    tUseIgnorTime = curJob.tUseIgnorTime
End Property

Public Property Let tLeft(ByVal vData As Currency)
    curJob.tLeft = vData
End Property

Public Property Get tLeft() As Currency
    tLeft = curJob.tLeft
End Property

Public Property Let tConnectionRiz(ByVal vData As Integer)
    curJob.tConnectionRiz = vData
End Property

Public Property Get tConnectionRiz() As Integer
    tConnectionRiz = curJob.tConnectionRiz
End Property

Public Property Let tClient(ByVal vData As String)
    curJob.tClient = vData
End Property

Public Property Get tClient() As String
    tClient = curJob.tClient
End Property

Public Property Let tUTime(ByVal vData As Long)
    curJob.tUTime = vData
End Property

Public Property Get tUTime() As Long
    tUTime = curJob.tUTime
End Property

Public Property Get cID() As Long
    cID = curJob.tID
End Property

Public Property Get sUp2() As Long
    sUp2 = curJob.sUp2
End Property

Public Property Get sUp1() As Long
    sUp1 = curJob.sUp1
End Property

Public Property Get sDw2() As Long
    sDw2 = curJob.sDw2
End Property

Public Property Get sDw1() As Long
    sDw1 = curJob.sDw1
End Property

Public Property Let tName(ByVal vData As String)
    curJob.tName = vData
End Property

Public Property Get tName() As String
    tName = curJob.tName
End Property

Public Property Let tTracker(ByVal vData As String)
    curJob.tTracker = vData
    curJob.tHost = getHost(vData)
End Property

Public Property Get tTracker() As String
    tTracker = curJob.tTracker
End Property

Public Property Get tHost() As String
    tHost = curJob.tHost
End Property

Public Property Let tEvent(ByVal vData As Integer)
    curJob.tStatus = vData
End Property

Public Property Get tEvent() As Integer
    tEvent = curJob.tStatus
End Property

Public Property Let tKey(ByVal vData As String)
     curJob.tKey = vData
End Property

Public Property Get tKey() As String
    tKey = curJob.tKey
End Property

Public Property Let tFullSize(ByVal vData As Currency)
    curJob.tFullSize = vData
End Property

Public Property Get tFullSize() As Currency
    tFullSize = curJob.tFullSize
End Property

Public Property Let tDownloaded(ByVal vData As Currency)
    curJob.tDownloaded = vData
End Property

Public Property Get tDownloaded() As Currency
    tDownloaded = curJob.tDownloaded
End Property

Public Property Let tUploaded(ByVal vData As Currency)
    curJob.tUploaded = vData
End Property

Public Property Get tUploaded() As Currency
    tUploaded = curJob.tUploaded
End Property

Public Property Let tPort(ByVal vData As Long)
    curJob.tPort = vData
End Property

Public Property Get tPort() As Long
    tPort = curJob.tPort
End Property

Public Property Let tPeerID(ByVal vData As String)
    curJob.tPeerID = vData
End Property

Public Property Get tPeerID() As String
    tPeerID = curJob.tPeerID
End Property

Public Property Let tHash(ByVal vData As String)
    curJob.tHash = vData
End Property

Public Property Get tHash() As String
    tHash = curJob.tHash
End Property

Public Sub GenerateScrapeURL(sTracker As String)
    Dim disScrape As Boolean
    Dim j As Long, i As Long
    
    Const TMP1 = "announce"
    Const tmp2 = ".php"
    Const TMP3 = "scrape"
    
    disScrape = False
    
    If curJob.tUseScrape Then
        If Len(sTracker) > 0 Then
            'Find last / symbol
            sTracker = getPath(sTracker)
            j = 0
            Do
                i = InStr(j + 1, sTracker, "/")
                If i > 0 Then j = i
            Loop Until i = 0
            
            If j > 0 Then
                If LCase(Mid(sTracker, j + 1, Len(TMP1))) = TMP1 Then
                    curJob.tScrapeURL = Left(sTracker, j) & TMP3
                    If LCase(Mid(sTracker, j + Len(TMP1) + 1, Len(tmp2))) = tmp2 Then _
                        curJob.tScrapeURL = curJob.tScrapeURL & tmp2
                    curJob.tScrapeURL = curJob.tScrapeURL & "?info_hash=" & _
                        MakeHash(curJob.tHash)
                Else
                    disScrape = True
                End If
            Else
                disScrape = True
            End If
        Else
            disScrape = True
        End If
        
        If disScrape Then
            MainForm.stb.Panels(1).Text = emul.tName & ": " & getCStr(141, "This tracker does not support scrape function")
            emul.tUseScrape = False
        End If
    End If
End Sub

Public Function MakeStep(TimerStep As Long) As Boolean
    Dim i As Currency
    Dim sec As Long
    Dim completedM As Boolean
    Dim stopFlag As Boolean
    Dim step_coef As Double
    Dim stepper As Single
    
    On Error Resume Next
    
    Randomize Timer
    sec = TimerStep / 1000
    completedM = False
    stopFlag = False
    stepper = 60 / sec
    
    'Scrape
    If (curJob.tUseScrape) And (curJob.tScrapeStatus = 0) Then
        curJob.tScrapeTime = curJob.tScrapeTime - sec
        If curJob.tScrapeTime <= 0 Then
            curJob.tScrapeTime = 0
            curJob.tScrapeStatus = 1
            MakeScrape
        End If
    End If
    
    If curJob.tStatus <> EMUL_TIMEDOUT Then
        curJob.tInnerCounter = curJob.tInnerCounter + 1
        
        'Change download
        If curJob.tLeft > 0 Then
            'Step mode?
            If curJob.tUseStepDownload And (curJob.tInnerCounter / stepper < curJob.tStepDownloadTime) Then _
                step_coef = curJob.tInnerCounter / (curJob.tStepDownloadTime * stepper) _
            Else _
                step_coef = 1
            'if frozen?
            If curJob.tFrozen Then step_coef = 0
            
            i = (Rnd * (curJob.sDw2 - curJob.sDw1) + curJob.sDw1) * 1024 * sec
            i = fix16K(CCur(i * step_coef))
            curJob.tDownloaded = curJob.tDownloaded + i
            curJob.tLeft = curJob.tLeft - i
            If curJob.tLeft <= 0 Then
                curJob.tLeft = 0
                curJob.tDownloaded = curJob.tFullSize
                completedM = True
            End If
        End If
        
        'Change upload
        'Step mode?
        If curJob.tUseStepUpload And (curJob.tInnerCounter / stepper < curJob.tStepUploadTime) Then _
            step_coef = curJob.tInnerCounter / (curJob.tStepUploadTime * stepper) _
        Else _
            step_coef = 1
        'if frozen?
        If curJob.tFrozen Then step_coef = 0
        
        i = (Rnd * (curJob.sUp2 - curJob.sUp1) + curJob.sUp1) * 1024 * sec
        i = fix16K(CCur(i * step_coef))
        curJob.tUploaded = curJob.tUploaded + i
    End If
    
    'Change time
    curJob.tUTime = curJob.tUTime - sec
    If (curJob.tUTime <= 0) Or completedM Then
        curJob.tUTime = 0
        If curJob.tStatus = EMUL_TIMEDOUT Then
            curJob.tStatus = EMUL_UPDATE
            frmEmulate.ws(curJob.tID).Connect
            If err.Number <> 0 Then
                MainForm.stb.Panels(1).Text = "TimeOut Error: " & err.Description
                err.Clear
                frmEmulate.ws(curJob.tID).Close
            End If
        Else
            curJob.tStatus = EMUL_UPDATE
            curJob.tConnectionRiz = CONNECT_UPDATE
            If completedM Then UpdateData "completed" _
                Else UpdateData vbNullString
        End If
    End If
    
    If (curJob.tStatus = EMUL_WORK) And (curJob.tStopMode <> 0) Then
        Select Case curJob.tStopMode
            Case 1
                stopFlag = (curJob.tUploaded > curJob.tStopValue)
            Case 2
                If curJob.tDownloaded > 0 Then
                    i = curJob.tUploaded / curJob.tDownloaded
                    stopFlag = (i > curJob.tStopValue)
                End If
            Case 3
                stopFlag = (ToUnixTime(Now) > curJob.tStopValue)
        End Select
        If stopFlag Then
            curJob.tStopMode = 0
            curJob.tStatus = EMUL_UPDATE
            curJob.tConnectionRiz = CONNECT_STOP
            UpdateData "stopped"
        End If
    End If
End Function

Public Function SelectJob(jID As Long) As Boolean
    For Each curJob In mycoll
        If curJob.tID = jID Then Exit For
    Next
End Function

Public Function AddJob() As Long
    Dim i As New clsEmulationItem
    
    TotalIDs = TotalIDs + 1
    i.tID = TotalIDs
    i.tDownloaded = 0
    i.tKey = vbNullString
    i.tName = vbNullString
    i.tPeerID = vbNullString
    i.tStatus = EMUL_ADD
    i.tTracker = vbNullString
    i.tUploaded = 0
    i.tUTime = 0
    i.tStopMode = 0
    i.tScrapeTime = 5 'Scrape after 5 sec from starting
    i.tScrapeStatus = 0
    i.tUseStepDownload = False
    i.tUseStepUpload = False
    i.tStepDownloadTime = 0
    i.tStepUploadTime = 0
    i.tFrozen = False
    i.tFrozenSkip = False
    
    mycoll.Add i
    SelectJob TotalIDs
    Load frmEmulate.ws(TotalIDs)
    Load frmEmulate.sc(TotalIDs)
    
    AddJob = TotalIDs
End Function

Public Function DeleteJob(jID As Long) As Boolean
    Dim i As Integer, tmp As clsEmulationItem
    Dim ret As Boolean
    
    ret = False
    For i = 1 To mycoll.Count
        Set tmp = mycoll.Item(i)
        If tmp.tID = jID Then
            ret = True
            mycoll.Remove i
            Unload frmEmulate.ws(jID)
            Unload frmEmulate.sc(jID)
            Exit For
        End If
    Next
    
    Set tmp = Nothing
    DeleteJob = ret
End Function

Public Sub SetTimedOut()
    Dim i As Long, tmp As String
    
    curJob.tStatus = EMUL_TIMEDOUT
    curJob.tConnectTries = curJob.tConnectTries + 1
    
    If (prgSettings.txtConnectTries <> 0) And (curJob.tConnectTries > prgSettings.txtConnectTries) Then
        curJob.tStatus = EMUL_ERROR
        tmp = getCStr(180, "Failed to connect for ") & CStr(curJob.tConnectTries - 1) & " " & getCStr(181, "times")
        MainForm.stb.Panels(1).Text = curJob.tName & ": " & tmp
        PopupMsgTray MainForm, tmp, curJob.tName, NIIF_ERROR
        Exit Sub
    End If
    
    i = curJob.tConnectTries * 10
    If i > 90 Then i = 90
    curJob.tUTime = i
End Sub

Public Sub SetSpeed(up1 As Long, up2 As Long, dw1 As Long, dw2 As Long)
    curJob.sDw1 = dw1
    curJob.sDw2 = dw2
    curJob.sUp1 = up1
    curJob.sUp2 = up2
End Sub

Public Sub SetHave(have As Integer)
    curJob.tLeft = CCur((1 - (have / 100)) * curJob.tFullSize)
End Sub

Public Function GetHave() As String
    GetHave = Format((1 - (curJob.tLeft / curJob.tFullSize)) * 100, "0")
End Function

Public Sub StartDownload()
    curJob.tScrapeTime = 5
    curJob.tInnerCounter = 0
    curJob.tStatus = EMUL_UPDATE
    curJob.tConnectionRiz = CONNECT_START
    curJob.tUploaded = 0
    curJob.tDownloaded = 0
    UpdateData TORRENT_START
End Sub

Public Sub StopDownload()
    curJob.tStatus = EMUL_UPDATE
    curJob.tConnectionRiz = CONNECT_STOP
    UpdateData "stopped"
End Sub

Public Sub SaveList()
    Dim i As Long, j As Long
    On Error Resume Next
    
    If prgSettings.SaveList = 0 Then Exit Sub
    
    If FileExist(CPath & SAVEDOWNF) Then
        Kill CPath & SAVEDOWNF
        err.Clear
    End If
    
    i = FreeFile
    Open CPath & SAVELISTF For Output As #i
    
    Print #i, "[General]"
    Print #i, "ID=" & BP10
    Print #i, "Total=" & CStr(mycoll.Count)
    For j = 1 To mycoll.Count
        Print #i, vbNullString
        Print #i, "[" & INIITEM & CStr(j) & "]"
        Print #i, "Client=" & mycoll(j).tClient
        Print #i, "IgnorTime=" & CStr(mycoll(j).tIgnorTime)
        Print #i, "UseIgnor=" & CStr(mycoll(j).tUseIgnorTime)
        Print #i, "Tracker=" & mycoll(j).tTracker
        Print #i, "Name=" & mycoll(j).tName
        Print #i, "Hash=" & mycoll(j).tHash
        Print #i, "Size=" & CStr(mycoll(j).tFullSize)
        Print #i, "PeerID=" & mycoll(j).tPeerID
        Print #i, "Key=" & mycoll(j).tKey
        Print #i, "Left=" & CStr(mycoll(j).tLeft)
        Print #i, "Dw1=" & CStr(mycoll(j).sDw1)
        Print #i, "Dw2=" & CStr(mycoll(j).sDw2)
        Print #i, "Up1=" & CStr(mycoll(j).sUp1)
        Print #i, "Up2=" & CStr(mycoll(j).sUp2)
        Print #i, "Scrape=" & CStr(mycoll(j).tUseScrape)
        Print #i, "UseStepD=" & CStr(mycoll(j).tUseStepDownload)
        Print #i, "UseStepU=" & CStr(mycoll(j).tUseStepUpload)
        Print #i, "StepDVal=" & CStr(mycoll(j).tStepDownloadTime)
        Print #i, "StepUVal=" & CStr(mycoll(j).tStepUploadTime)
        Print #i, "AddTime=" & mycoll(j).tAddedTime
    Next
    Close #i
End Sub

Public Sub UpdateData(iEvent As String, Optional cHost As String = vbNullString)
    Dim msg As String
    Dim mUrl As String
    Dim cur_host As String
    Dim i As Long, i2 As Long
    
    On Error Resume Next
    
    If Len(cHost) = 0 Then cur_host = curJob.tHost _
        Else cur_host = cHost
    
    If prgSettings.chkUseProxy = 1 Then
        mUrl = curJob.tTracker
    Else
        i = InStr(1, curJob.tTracker, "://")
        i2 = InStr(i + 3, curJob.tTracker, "/")
        If i2 = 0 Then mUrl = "/" _
            Else mUrl = Mid(curJob.tTracker, i2)
    End If
    i = InStr(1, mUrl, "?")
    If i = 0 Then mUrl = mUrl & "?" Else mUrl = mUrl & "&"
    
    Call clnt.LoadClient(curJob.tClient)
    msg = clnt.BuildHTTP(mUrl, curJob.tHash, curJob.tPeerID, CStr(curJob.tPort), _
        Format(curJob.tUploaded, "0"), Format(curJob.tDownloaded, "0"), _
        Format(curJob.tLeft, "0"), curJob.tKey, iEvent, cur_host)
    
    If prgSettings.chkUseProxy = 1 Then
        frmEmulate.ws(curJob.tID).RemoteHost = prgSettings.txtProxyIp
        frmEmulate.ws(curJob.tID).RemotePort = prgSettings.txtProxyPort
    Else
        i = InStr(1, cur_host, ":")
        If i > 0 Then
            frmEmulate.ws(curJob.tID).RemoteHost = Left(cur_host, i - 1)
            frmEmulate.ws(curJob.tID).RemotePort = val(Mid(cur_host, i + 1))
        Else
            frmEmulate.ws(curJob.tID).RemoteHost = curJob.tHost
            frmEmulate.ws(curJob.tID).RemotePort = 80
        End If
    End If
    curJob.iEvent = iEvent
    curJob.tmpDataRecieved = vbNullString
    frmEmulate.ws(curJob.tID).Tag = msg
    frmEmulate.ws(curJob.tID).Connect
    
    If err.Number <> 0 Then
        curJob.tStatus = EMUL_ERROR
        MainForm.stb.Panels(1).Text = "UpdateData: " & err.Description
    End If
End Sub

Public Function checkExistence(hash As String) As Boolean
    checkExistence = False
    For Each curJob In mycoll
        If curJob.tHash = hash Then
            checkExistence = True
            Exit For
        End If
    Next
End Function

Private Sub MakeScrape()
    Dim msg As String
    Dim i As Long
    
    On Error Resume Next
    
    msg = curJob.tScrapeURL
    If prgSettings.chkUseProxy = 1 Then
        i = InStr(1, curJob.tTracker, "://")
        If i > 0 Then
            i = InStr(i + 3, curJob.tTracker, "/")
            If i = 0 Then msg = curJob.tTracker & curJob.tScrapeURL _
                Else msg = Left(curJob.tTracker, i - 1) & curJob.tScrapeURL
        End If
    End If
    
    Call clnt.LoadClient(curJob.tClient)
    msg = clnt.BuildScrapeHTTP(msg, curJob.tHost)
    
    If prgSettings.chkUseProxy = 1 Then
        frmEmulate.sc(curJob.tID).RemoteHost = prgSettings.txtProxyIp
        frmEmulate.sc(curJob.tID).RemotePort = prgSettings.txtProxyPort
    Else
        i = InStr(1, curJob.tHost, ":")
        If i > 0 Then
            frmEmulate.sc(curJob.tID).RemoteHost = Left(curJob.tHost, i - 1)
            frmEmulate.sc(curJob.tID).RemotePort = val(Mid(curJob.tHost, i + 1))
        Else
            frmEmulate.sc(curJob.tID).RemoteHost = curJob.tHost
            frmEmulate.sc(curJob.tID).RemotePort = 80
        End If
    End If
    frmEmulate.sc(curJob.tID).Tag = msg
    frmEmulate.sc(curJob.tID).Connect
    
    If err.Number <> 0 Then
        MainForm.stb.Panels(1).Text = "Scrape: " & err.Description
    End If
End Sub

Private Sub Class_Initialize()
    TotalIDs = 0
End Sub

Private Function getHost(sUrl As String) As String
    Dim i As Integer, i2 As Integer
    i = InStr(1, sUrl, "://")
    If i > 0 Then
        i2 = InStr(i + 3, sUrl, "/")
        If i2 = 0 Then getHost = Mid(sUrl, i + 3) _
            Else getHost = Mid(sUrl, i + 3, i2 - i - 3)
    End If
End Function

Private Function getPath(sUrl As String) As String
    Dim i As Integer
    getPath = vbNullString
    i = InStr(1, sUrl, "://")
    If i > 0 Then
        i = InStr(i + 3, sUrl, "/")
        If i > 0 Then getPath = Mid(sUrl, i) _
            Else getPath = "/"
    End If
End Function
