VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl32.ocx"
Begin VB.Form frmEmulate 
   Caption         =   "Emulation"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10335
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5430
   ScaleWidth      =   10335
   Tag             =   "~83"
   Begin MSComctlLib.Toolbar tb1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   1217
      ButtonWidth     =   1482
      ButtonHeight    =   1164
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add"
            Object.Tag             =   "~84"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Object.Tag             =   "~85"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Start"
            Object.Tag             =   "~86"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
            Object.Tag             =   "~87"
            ImageIndex      =   4
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "~178"
                  Text            =   "Normal Stop"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "~179"
                  Text            =   "Stop without updating data"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Properties"
            Object.Tag             =   "~88"
            ImageIndex      =   5
         EndProperty
      EndProperty
      Begin MSWinsockLib.Winsock sc 
         Index           =   0
         Left            =   6360
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Timer Timer1 
         Interval        =   2000
         Left            =   6840
         Top             =   240
      End
      Begin MSWinsockLib.Winsock ws 
         Index           =   0
         Left            =   7320
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   7800
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmulate.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmulate.frx":0354
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmulate.frx":06A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmulate.frx":09FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmulate.frx":0D50
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmulate.frx":10A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmulate.frx":13F8
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmulate.frx":174C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8400
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   24
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmulate.frx":1AA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmulate.frx":21B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmulate.frx":28C8
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmulate.frx":2FDC
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmulate.frx":36F0
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView plist 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8281
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "~97"
         Text            =   "Name"
         Object.Width           =   4789
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "~98"
         Text            =   "Tracker"
         Object.Width           =   3120
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "~99"
         Text            =   "Status"
         Object.Width           =   1983
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "~100"
         Text            =   "Updating in"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "~101"
         Text            =   "Uploaded"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "~102"
         Text            =   "Downloaded"
         Object.Width           =   1429
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "~103"
         Text            =   "Completed"
         Object.Width           =   1508
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "~104"
         Text            =   "Rating"
         Object.Width           =   1164
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "~140"
         Text            =   "Seed/Peer"
         Object.Width           =   1376
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Tag             =   "~226"
         Text            =   "Added at"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu cmm 
      Caption         =   "cmm"
      Visible         =   0   'False
      Begin VB.Menu cmt 
         Caption         =   "Properties"
         Index           =   1
      End
      Begin VB.Menu cmt 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu cmt 
         Caption         =   "Start"
         Index           =   3
      End
      Begin VB.Menu cmt 
         Caption         =   "Stop"
         Index           =   4
      End
      Begin VB.Menu cmt 
         Caption         =   "Delete"
         Index           =   5
      End
      Begin VB.Menu cmt 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu cmt 
         Caption         =   "Update statistics on tracker"
         Index           =   7
      End
      Begin VB.Menu cmt 
         Caption         =   "Autostop"
         Index           =   8
      End
   End
   Begin VB.Menu cmCuston 
      Caption         =   "cmCuston"
      Visible         =   0   'False
      Begin VB.Menu cmCustItem 
         Caption         =   "cmCustItem"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmEmulate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private redirect_host As String

Dim sortColumn As Integer

Private Sub cmt_Click(Index As Integer)
    Select Case Index
        Case 1: EditTorrent
        Case 3: StartTorrent
        Case 4: StopTorrent False
        Case 5: DeleteTorrent
        Case 7
            Call emul.SelectJob(GetListItemID(plist.SelectedItem.Index))
            ws(emul.cID).Close
            If emul.tEvent = EMUL_WORK Then
                emul.tUTime = 0
            End If
        Case 8
            frmStopSett.Show vbModal
    End Select
End Sub

Private Sub Form_Load()
    DrawForm
    Show
    LoadSavedList
    LoadWindowPos Me, "emulf", 10455, 5835
    Global_MenuCustomizer.LoadState plist
    sortColumn = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim i As Long, jID As Long
    
    emul.SaveList
    
    For i = 1 To plist.ListItems.Count
        jID = GetListItemID(i)
        emul.DeleteJob jID
    Next
    
    SaveWindowPos Me, "emulf"
    Global_MenuCustomizer.SaveState plist
End Sub

Private Sub Form_Resize()
    Static cc As Long
    Dim vOffset As Long
    Dim hOffset As Long
    
    On Error Resume Next
    
    vOffset = (GetSystemMetrics(SM_CYCAPTION) + (GetSystemMetrics(SM_CYFRAME)) * 2) * 15
    hOffset = (GetSystemMetrics(SM_CXFRAME) * 2) * 15
    
    cc = cc + 1
    If Me.WindowState = vbNormal Then
        plist.Width = Me.Width - hOffset
        plist.Height = Me.Height - tb1.Height - vOffset - 30
        If cc < 3 Then _
            plist.Height = plist.Height - (GetSystemMetrics(SM_CYFRAME) * 2) * 15
    End If
    err.Clear
End Sub

Private Sub plist_AfterLabelEdit(Cancel As Integer, NewString As String)
    If Len(Trim(NewString)) = 0 Then
        Cancel = 1
    Else
        emul.SelectJob GetListItemID(plist.SelectedItem.Index)
        emul.tName = NewString
    End If
End Sub

Private Sub plist_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If plist.SortKey = ColumnHeader.Index - 1 Then
        If plist.SortOrder = lvwAscending Then plist.SortOrder = lvwDescending _
            Else plist.SortOrder = lvwAscending
    Else
        plist.SortOrder = lvwAscending
    End If
    plist.SortKey = ColumnHeader.Index - 1
    plist.Sorted = True
    plist.Sorted = False
End Sub

Private Sub plist_DblClick()
    tb1_ButtonClick tb1.Buttons(5)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyA) And (GetAsyncKeyState(vbKeyControl) <> 0) Then _
        SelectAllItems
    If KeyCode = vbKeyDelete Then DeleteTorrent
End Sub

Private Sub SelectAllItems()
    Dim i As Long
    
    For i = 1 To plist.ListItems.Count
        plist.ListItems(i).Selected = True
    Next
End Sub

Private Sub plist_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim ind As Long
    On Error Resume Next
    
    If Button = vbRightButton Then
        ind = plist.SelectedItem.Index
        If err.Number = 0 Then
            Me.PopupMenu cmm, , , , cmt(1)
        Else
            err.Clear
        End If
    End If
End Sub

Private Sub plist_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
    FilePoolFromEmulationDrop data
    ProcceedFilePool
End Sub

Private Sub sc_Connect(Index As Integer)
    AddToFileLog "Scrape send [" & CStr(Index) & "]", sc(Index).Tag
    sc(Index).SendData sc(Index).Tag
End Sub

Private Sub sc_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strDataIn As String
    
    sc(Index).GetData strDataIn, vbString
    AddToFileLog "Scrape recieved [" & CStr(Index) & "]", strDataIn
    
    sc(Index).Tag = vbNullString
    Call emul.SelectJob(CLng(Index))
    Call GetPeersSeeds(strDataIn)
    
    sc(Index).Close
End Sub

Private Sub sc_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call emul.SelectJob(CLng(Index))
    If Number = 10060 Then
        emul.tScrapeTime = 30
        emul.tScrapeStatus = 0
    Else
        MainForm.stb.Panels(1).Text = emul.tName & ": Scrape Error - " & Description
        emul.tUseScrape = False
    End If
    sc(Index).Close
End Sub

Private Sub tb1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Add
            AddFile
        Case 2 'Delete
            DeleteTorrent
        Case 3 'Start
            StartTorrent
        Case 4 'Stop
            StopTorrent False
        Case 5 'Settings
            EditTorrent
    End Select
End Sub

Private Sub DrawForm()
    Me.Icon = MainForm.Icon
    
    cmt(1).Tag = "~88"
    cmt(3).Tag = "~86"
    cmt(4).Tag = "~87"
    cmt(5).Tag = "~85"
    cmt(7).Tag = "~107"
    cmt(8).Tag = "~132"
    
    TranslateForm Me
End Sub

Public Sub ProcceedFilePool()
    Dim i As Long, view_form As Boolean
    
    view_form = True
    
    For i = 1 To Global_FilePool.Count
        RunAddTorrent Global_FilePool.Item(i), view_form
        If (Global_FilePool.Count > 1) And (i = 1) Then
            If MsgBox(getCStr(224, "TUse same settings for other torrents?"), vbYesNo + vbQuestion, App.Title) = vbYes Then _
                view_form = False
        End If
    Next
End Sub

Public Sub AddFile()
    Dim i As New clsCommonDialog
    Dim tmp As String
    
    tmp = vbNullString
    If i.OpenFileName(tmp, , True, False, False, False, "Torrent files|*.torrent") Then
        AddToFileLog "Open file", tmp
        
        tmp = MakeShortName(tmp)
        RunAddTorrent tmp
    End If
    
    Set i = Nothing
End Sub

Public Sub RunAddTorrent(Fname As String, Optional view_form As Boolean = True)
    Dim tw As New clsTorrentObj
    Dim i As Long, errorMsg As String
    
    If tw.LoadTorrent(Fname) Then
        If prgSettings.chkSameHash = 0 And emul.checkExistence(tw.tHash) Then
            errorMsg = getCStr(225, "This torrent already exist in the list")
            MainForm.stb.Panels(1).Text = tw.tName & ": " & errorMsg
            PopupMsgTray MainForm, errorMsg, tw.tName, NIIF_ERROR
        Else
            i = emul.AddJob
            emul.tAddedTime = Format(Now, "yyyy.mm.dd hh:mm")
            emul.tFullSize = tw.tLength
            emul.tHash = tw.tHash
            emul.tName = tw.tName
            emul.tTracker = tw.tTrackerURL
            emul.tPort = prgSettings.emul_port
            emul.tUseIgnorTime = (prgSettings.chkIgnorT = 1)
            emul.tIgnorTime = prgSettings.txtIgnorT
            emul.tUseScrape = (prgSettings.chkUseScrape = 1)
            emul.SetSpeed prgSettings.emul_up1, prgSettings.emul_up2, prgSettings.emul_dw1, prgSettings.emul_dw2
            emul.tUseStepDownload = prgSettings.chkStepModeD
            emul.tUseStepUpload = prgSettings.chkStepModeU
            emul.tStepDownloadTime = prgSettings.txtSMDown
            emul.tStepUploadTime = prgSettings.txtSMUp
            emul.SetHave prgSettings.txtHave
            
            plist.ListItems.Add , "k" & CStr(i), emul.tName, , 3
            
            Load frmAddEmul
            If view_form Then
                frmAddEmul.Show vbModal
            Else
                frmAddEmul.Command1_Click
            End If
        End If
    End If
    
    Set tw = Nothing
End Sub

Private Sub StartTorrent()
    Dim i As Long
    
    For i = plist.ListItems.Count To 1 Step -1
        If plist.ListItems(i).Selected Then
            Call emul.SelectJob(GetListItemID(i))
            If (emul.tEvent = EMUL_STOP) Or (emul.tEvent = EMUL_ERROR) Then
                emul.StartDownload
                RedrawCurrentJob i, True
            End If
        End If
    Next
End Sub

Private Sub StopTorrent(CrudeStop As Boolean)
    Dim i As Long
    
    For i = plist.ListItems.Count To 1 Step -1
        If plist.ListItems(i).Selected Then
            Call emul.SelectJob(GetListItemID(i))
            If CrudeStop And (emul.tEvent = EMUL_TIMEDOUT Or emul.tEvent = EMUL_WORK) Then
                emul.tEvent = EMUL_STOP
                RedrawCurrentJob i, True
            End If
            If emul.tEvent = EMUL_WORK Then
                emul.StopDownload
                RedrawCurrentJob i, True
            End If
        End If
    Next
End Sub

Private Sub DeleteTorrent()
    Dim curJob As Long, ask As Long
    Dim i As Long, asked As Boolean
    
    asked = False
    
    For i = plist.ListItems.Count To 1 Step -1
        If plist.ListItems(i).Selected Then
            If Not asked Then
                ask = MsgBox(getCStr(77, "Are yor sure you want to delete this task?"), vbQuestion + vbYesNo, App.Title)
                If ask <> vbYes Then Exit For
                asked = True
            End If
            curJob = GetListItemID(i)
            emul.DeleteJob curJob
            plist.ListItems.Remove i
        End If
    Next
    emul.SaveList
End Sub

Private Sub EditTorrent()
    Dim ind As Long, curJob As Long
    
    ind = getSelectionCount
    If ind = 0 Then Exit Sub
    
    If ind > 1 Then
        frmAddEmulMulti.Show vbModal
    Else
        ind = plist.SelectedItem.Index
        curJob = GetListItemID(ind)
        Call emul.SelectJob(curJob)
        frmAddEmul.Show vbModal
    End If
End Sub

Private Function getSelectionCount() As Integer
    Dim i As Integer
    
    getSelectionCount = 0
    For i = 1 To plist.ListItems.Count
        If plist.ListItems(i).Selected Then _
            getSelectionCount = getSelectionCount + 1
    Next
End Function

Public Sub DeleteFromList(jID As Long)
    Dim i As Integer
    
    For i = 1 To plist.ListItems.Count
        If plist.ListItems(i).key = "k" & CStr(jID) Then
            plist.ListItems.Remove i
            Exit For
        End If
    Next
End Sub

Public Sub RedrawCurrentJob(Optional li As Long = 0, Optional small As Boolean = False)
    Dim curIndx As Long, i As Integer, color_need As Long
    
    If li = 0 Then
        For i = 1 To plist.ListItems.Count
            If plist.ListItems(i).key = "k" & CStr(emul.cID) Then
                curIndx = i
                Exit For
            End If
        Next
    Else
        curIndx = li
    End If
    
    color_need = vbBlack
    
    Select Case emul.tEvent
        Case EMUL_ADD, EMUL_STOP
            plist.ListItems(curIndx).SmallIcon = 3
            plist.ListItems(curIndx).SubItems(2) = getCStr(89, "Stopped")
        Case EMUL_WORK
            If emul.tLeft > 0 Then
                plist.ListItems(curIndx).SmallIcon = IIf(emul.tFrozen, 7, 1)
                plist.ListItems(curIndx).SubItems(2) = getCStr(90, "Downloading")
            Else
                plist.ListItems(curIndx).SmallIcon = IIf(emul.tFrozen, 8, 4)
                plist.ListItems(curIndx).SubItems(2) = getCStr(91, "Seeding")
            End If
            If emul.tStopMode > 0 Then _
                color_need = vbBlue
        Case EMUL_UPDATE
            plist.ListItems(curIndx).SmallIcon = 2
            plist.ListItems(curIndx).SubItems(2) = getCStr(92, "Connecting")
        Case EMUL_TIMEDOUT
            plist.ListItems(curIndx).SmallIcon = 5
            plist.ListItems(curIndx).SubItems(2) = getCStr(93, "Waiting")
        Case EMUL_ERROR
            plist.ListItems(curIndx).SmallIcon = 6
            plist.ListItems(curIndx).SubItems(2) = getCStr(94, "Error")
    End Select
    
    If plist.ListItems(curIndx).ForeColor <> color_need Then _
        plist.ListItems(curIndx).ForeColor = color_need
    
    If Not small Then
        plist.ListItems(curIndx).Text = emul.tName
        plist.ListItems(curIndx).SubItems(1) = emul.tTracker
        plist.ListItems(curIndx).SubItems(9) = emul.tAddedTime
    End If
    plist.ListItems(curIndx).SubItems(3) = GetTimeFormat(emul.tUTime)
    plist.ListItems(curIndx).SubItems(4) = GetByteFormat(emul.tUploaded)
    plist.ListItems(curIndx).SubItems(5) = GetByteFormat(emul.tDownloaded)
    plist.ListItems(curIndx).SubItems(6) = emul.GetHave & "%"
    If emul.tDownloaded = 0 Then _
        plist.ListItems(curIndx).SubItems(7) = "0.0" _
        Else _
        plist.ListItems(curIndx).SubItems(7) = Format((emul.tUploaded / emul.tDownloaded), "#0.0")
    
End Sub

Private Sub tb1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Index
        Case 1 'Normal stop
            StopTorrent False
        Case 2
            StopTorrent True
    End Select
End Sub

Private Sub Timer1_Timer()
    Dim i As Long
    Dim curID As Long
    
    For i = 1 To plist.ListItems.Count
        curID = GetListItemID(i)
        emul.SelectJob curID
        If (emul.tEvent = EMUL_WORK) Or (emul.tEvent = EMUL_TIMEDOUT) Then
            emul.MakeStep Timer1.Interval
            RedrawCurrentJob i, True
            plist.Refresh
        End If
    Next
End Sub

Public Function GetListItemID(lIndx As Long) As Long
    GetListItemID = val(Mid(plist.ListItems(lIndx).key, 2))
End Function

Private Sub ws_Close(Index As Integer)
    Dim i As Long
    
    ws(Index).Tag = vbNullString
    Call emul.SelectJob(CLng(Index))
    emul.tConnectTries = 0
    
    If isRedirect(emul.tmpDataRecieved) Then
        ws(Index).Close
        DoEvents
        emul.tEvent = EMUL_UPDATE
        emul.UpdateData emul.iEvent, redirect_host
        Exit Sub
    End If
    
    Select Case emul.tConnectionRiz
        Case CONNECT_START, CONNECT_UPDATE
            emul.tEvent = EMUL_WORK
            If emul.tUseIgnorTime Then
                emul.tUTime = emul.tIgnorTime * 60
            Else
                i = GetTrackerTime(emul.tmpDataRecieved)
                If i = 0 Then
                    emul.tEvent = EMUL_ERROR
                Else
                    emul.tUTime = i
                End If
            End If
        Case CONNECT_STOP
            emul.tEvent = EMUL_STOP
            emul.tFrozenSkip = False
    End Select
    
    RedrawCurrentJob
    
    ws(Index).Close
End Sub

Private Sub ws_Connect(Index As Integer)
    AddToFileLog "Emulation send [" & CStr(Index) & "]", ws(Index).Tag
    ws(Index).SendData ws(Index).Tag
End Sub

Private Sub ws_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strDataIn As String
    
    ws(Index).GetData strDataIn, vbString
    AddToFileLog "Emulation recieved [" & CStr(Index) & "]", strDataIn
    
    Call emul.SelectJob(CLng(Index))
    emul.tmpDataRecieved = emul.tmpDataRecieved & strDataIn
End Sub

Private Sub ws_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Dim tmp As String
    
    Call emul.SelectJob(CLng(Index))
    If Number = 10060 Or Number = 10053 Or prgSettings.chkIgnorSocketError = 1 Then
        emul.SetTimedOut
    Else
        tmp = getCStr(123, "Socket error") & " ¹" & CStr(Number)
        MainForm.stb.Panels(1).Text = tmp & ": " & Description
        PopupMsgTray MainForm, Description, tmp, NIIF_ERROR
        emul.tEvent = EMUL_ERROR
    End If
    RedrawCurrentJob
    ws(Index).Close
End Sub

Private Function isRedirect(data As String) As Boolean
    Dim i As Long, j As Long
    Dim tmp As String
    
    Dim cur_host As String
    Dim cur_path As String
    Dim cur_port As Long
    
    Const HTTP_LOCATION = "Location: "
    Const HTTP_PREFIX = "://"
    Const HTTP_301 = "301 Moved Permanently"
    
    isRedirect = False
    
    i = InStr(1, data, HTTP_301)
    If i = 0 Then Exit Function
    
    i = InStr(1, data, HTTP_LOCATION)
    j = InStr(i, data, vbCrLf)
    If i = 0 Or j = 0 Then Exit Function
    
    tmp = Trim(Mid(data, i + Len(HTTP_LOCATION), j - i - Len(HTTP_LOCATION)))
    If Len(tmp) = 0 Then Exit Function
    
    'Get Host and path
    i = InStr(1, tmp, HTTP_PREFIX)
    j = InStr(i + Len(HTTP_PREFIX), tmp, "/")
    If i = 0 Then Exit Function
    If j = 0 Then
        redirect_host = Mid(tmp, i + Len(HTTP_PREFIX))
    Else
        redirect_host = Mid(tmp, i + Len(HTTP_PREFIX), j - i - Len(HTTP_PREFIX))
    End If
    
    'Reget data
    AddToFileLog "Redirection found", "Host: " & redirect_host
    
    isRedirect = True
End Function

Private Sub GetPeersSeeds(sData As String)
    Dim i As Long, tmp As String
    Dim seeds As Long, peers As Long
    Dim curIndx As Long
    
    Const HTTP_200 = "200 OK"
    Const HTTP_301 = "301 Moved Permanently"
    Const TMP1 = ":complete"
    Const tmp2 = "incomplete"
        
    i = InStr(1, sData, vbCrLf)
    tmp = Left(sData, i)
    i = InStr(1, tmp, HTTP_200)
    If i = 0 Then 'Result is not ok
        MainForm.stb.Panels(1).Text = emul.tName & ": " & getCStr(141, "This tracker does not support scrape function")
        'PopupMsgTray MainForm, emul.tName & ": " & FL(140), App.Title, NIIF_ERROR
        emul.tUseScrape = False
        Exit Sub
    End If
    
    tmp = DecodeHTTP(sData)
    
    'seeds
    i = InStr(1, tmp, TMP1)
    If i > 0 Then seeds = GetTorLongBuffer(tmp, i + Len(TMP1))
    i = InStr(1, tmp, tmp2)
    If i > 0 Then peers = GetTorLongBuffer(tmp, i + Len(tmp2))
    
    'Draw it
    curIndx = 0
    For i = 1 To plist.ListItems.Count
        If plist.ListItems(i).key = "k" & CStr(emul.cID) Then
            curIndx = i
            Exit For
        End If
    Next
    If curIndx > 0 Then
        plist.ListItems(curIndx).SubItems(8) = CStr(seeds) & "/" & CStr(peers)
    End If
    
    emul.tScrapeTime = SCRAPE_STEP
    emul.tScrapeStatus = 0
End Sub

Private Function GetTrackerTime(sData As String) As Long
    Dim i As Long, tmp As String
    
    Const TMP1 = "interval"
    Const tmp2 = "min interval"
    Const TMP3 = "failure reason"
    Const TMP4 = "5:peers6:"
    Const HTTP_200 = "200 OK"
    
    i = InStr(1, sData, vbCrLf)
    tmp = Left(sData, i)
    i = InStr(1, tmp, HTTP_200)
    If i = 0 Then 'File not found or something like that
        If prgSettings.chkIgnorServerError = 1 Then
            emul.SetTimedOut
            MainForm.stb.Panels(1).Text = "Debug 500"
        Else
            MainForm.stb.Panels(1).Text = emul.tName & ": " & getCStr(95, "Server returned error")
            PopupMsgTray MainForm, getCStr(95, "Server returned error"), emul.tName, NIIF_ERROR
            GetTrackerTime = 0
        End If
        Exit Function
    End If
    
    tmp = DecodeHTTP(sData)
    
    i = InStr(1, tmp, TMP3)
    If i > 0 Then 'Tracker return error
        tmp = GetTorStringBuffer(tmp, i + Len(TMP3))
        MainForm.stb.Panels(1).Text = emul.tName & ": " & getCStr(96, "Tracker returned error") & " - " & tmp
        PopupMsgTray MainForm, getCStr(96, "Tracker returned error") & " - " & tmp, emul.tName, NIIF_ERROR
        GetTrackerTime = 0
        Exit Function
    End If
    
    'Is frozen?
    emul.tFrozen = False
    If prgSettings.chkFroze Then
        i = InStr(1, tmp, TMP4)
        If i > 0 Then
            If Not emul.tFrozenSkip Then
                emul.tFrozen = True
                MainForm.stb.Panels(1).Text = emul.tName & ": " & getCStr(222, "Tracker returned 1 peer")
            End If
        Else
            emul.tFrozenSkip = True
        End If
    End If
    
    i = InStr(1, tmp, TMP1)
    If i > 0 Then
        GetTrackerTime = GetTorLongBuffer(tmp, i + Len(TMP1))
    Else
        i = InStr(1, tmp, tmp2)
        If i > 0 Then
            GetTrackerTime = GetTorLongBuffer(tmp, i + Len(tmp2))
        Else
            GetTrackerTime = 600
        End If
    End If
End Function

Private Function DecodeHTTP(sData As String) As String
    Dim i As Long, tmp As String, i2 As Long, rez As Long
    Dim encType As String
    
    i = InStr(1, sData, "Content-Encoding:")
    If i > 0 Then
        Dim zlib As New clsZLIB
        
        i = InStr(i, sData, " ")
        i2 = InStr(i, sData, vbCrLf)
        encType = LCase(Mid(sData, i + 1, i2 - i - 1))
        i = InStr(1, sData, vbCrLf & vbCrLf)
        tmp = Mid(sData, i + 4)
        Select Case encType
            Case "gzip": rez = zlib.UncompressString(tmp, Z_AUTO)
            Case "deflate": rez = zlib.UncompressString(tmp, Z_DEFLATE)
            Case Else: MsgBox "Unknown compression " & encType, vbCritical, App.Title 'Debug
        End Select
        Set zlib = Nothing
        'AddToFileLog "Work data (" & CStr(rez) & ")", "'" & tmp & "'"
    Else
        i = InStr(1, sData, vbCrLf & vbCrLf)
        tmp = Mid(sData, i + 4)
    End If
    DecodeHTTP = tmp
End Function

Private Sub LoadSavedList()
    Dim tmp As String, i As Long, j As Long
    On Error Resume Next
    
    If (prgSettings.SaveList = 1) And FileExist(CPath & SAVELISTF) Then
        tmp = getINIString("General", "ID", CPath & SAVELISTF)
        If tmp <> BP10 Then Exit Sub
        i = getINIString("General", "Total", CPath & SAVELISTF)
        For j = 1 To i
            Call emul.AddJob
            emul.tPort = prgSettings.emul_port
            emul.tClient = getINIString(INIITEM & CStr(j), "Client", CPath & SAVELISTF)
            emul.tIgnorTime = val(getINIString(INIITEM & CStr(j), "IgnorTime", CPath & SAVELISTF))
            emul.tUseIgnorTime = CBool(getINIString(INIITEM & CStr(j), "UseIgnor", CPath & SAVELISTF))
            emul.tTracker = getINIString(INIITEM & CStr(j), "Tracker", CPath & SAVELISTF)
            emul.tName = getINIString(INIITEM & CStr(j), "Name", CPath & SAVELISTF)
            emul.tHash = getINIString(INIITEM & CStr(j), "Hash", CPath & SAVELISTF)
            emul.tFullSize = CCur(getINIString(INIITEM & CStr(j), "Size", CPath & SAVELISTF))
            emul.tKey = getINIString(INIITEM & CStr(j), "Key", CPath & SAVELISTF)
            emul.tPeerID = getINIString(INIITEM & CStr(j), "PeerID", CPath & SAVELISTF)
            emul.tLeft = CCur(getINIString(INIITEM & CStr(j), "Left", CPath & SAVELISTF))
            emul.SetSpeed CCur(getINIString(INIITEM & CStr(j), "Up1", CPath & SAVELISTF)), _
                CCur(getINIString(INIITEM & CStr(j), "Up2", CPath & SAVELISTF)), _
                CCur(getINIString(INIITEM & CStr(j), "Dw1", CPath & SAVELISTF)), _
                CCur(getINIString(INIITEM & CStr(j), "Dw2", CPath & SAVELISTF))
            emul.tUseScrape = CBool(getINIString(INIITEM & CStr(j), "Scrape", CPath & SAVELISTF))
            emul.tUseStepDownload = CBool(getINIString(INIITEM & CStr(j), "UseStepD", CPath & SAVELISTF))
            emul.tUseStepUpload = CBool(getINIString(INIITEM & CStr(j), "UseStepU", CPath & SAVELISTF))
            emul.tStepDownloadTime = CLng(getINIString(INIITEM & CStr(j), "StepDVal", CPath & SAVELISTF))
            emul.tStepUploadTime = CLng(getINIString(INIITEM & CStr(j), "StepUVal", CPath & SAVELISTF))
            emul.tEvent = EMUL_STOP
            emul.tAddedTime = getINIString(INIITEM & CStr(j), "AddTime", CPath & SAVELISTF)
            emul.GenerateScrapeURL emul.tTracker
            
            plist.ListItems.Add , "k" & CStr(emul.cID), emul.tName, , 3
            RedrawCurrentJob
        Next
    End If
End Sub
