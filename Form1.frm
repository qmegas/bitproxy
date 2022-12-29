VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Downloading"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4515
   ScaleWidth      =   8415
   Tag             =   "~106"
   Begin VB.Frame Frame 
      Height          =   3375
      Index           =   0
      Left            =   0
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Label lblNosel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Downloading job not selected"
         Height          =   195
         Left            =   990
         TabIndex        =   10
         Tag             =   "~199"
         Top             =   1560
         Width           =   2115
      End
   End
   Begin VB.Frame Frame 
      Height          =   3375
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   8415
      Begin VB.ListBox lstLog 
         Height          =   3060
         Index           =   0
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   8175
      End
   End
   Begin VB.Frame Frame 
      Height          =   3375
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   8415
      Begin VB.PictureBox picGrf 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2655
         Left            =   120
         MousePointer    =   99  'Custom
         ScaleHeight     =   2625
         ScaleWidth      =   8145
         TabIndex        =   8
         Top             =   600
         Width           =   8175
      End
      Begin VB.CheckBox chkGUpload 
         Caption         =   "Draw uploading speed"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Tag             =   "~198"
         Top             =   240
         Width           =   2895
      End
      Begin VB.CheckBox chkGDownload 
         Caption         =   "Draw downloading speed"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Tag             =   "~197"
         Top             =   240
         Width           =   2775
      End
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   3375
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5953
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "~188"
         Text            =   "Task name"
         Object.Width           =   5345
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "~189"
         Text            =   "Tracker"
         Object.Width           =   3281
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "~190"
         Text            =   "Port"
         Object.Width           =   1270
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "~191"
         Text            =   "Mode"
         Object.Width           =   2170
      EndProperty
   End
   Begin MSWinsockLib.Winsock Socket4 
      Index           =   0
      Left            =   9240
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Socket3 
      Index           =   0
      Left            =   8760
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Socket2 
      Index           =   0
      Left            =   8280
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Socket1 
      Index           =   0
      Left            =   7800
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   1217
      ButtonWidth     =   1217
      ButtonHeight    =   1164
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add"
            Object.Tag             =   "~184"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "~186"
                  Text            =   "Open torrent file"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "~187"
                  Text            =   "Add new task with default settings"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Object.Tag             =   "~185"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Turn On"
            Object.Tag             =   "~21"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Turn off"
            Object.Tag             =   "~22"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Settings"
            Object.Tag             =   "~193"
            ImageIndex      =   5
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   6600
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":365A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":39AE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   7200
         Top             =   0
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
               Picture         =   "Form1.frx":3D02
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":4416
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":4B2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":523E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":5952
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.TabStrip ts 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   4080
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   661
      MultiRow        =   -1  'True
      Placement       =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Task list"
            Object.Tag             =   "~192"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Log"
            Object.Tag             =   "~194"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Graph"
            Object.Tag             =   "~195"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu cMenu 
      Caption         =   "xmenu"
      Visible         =   0   'False
      Begin VB.Menu cmlist 
         Caption         =   "Clear log"
         Index           =   1
      End
      Begin VB.Menu cmlist 
         Caption         =   "Copy to clipboard"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WeQuit As Boolean
Public GraphikObj As New clsGrafik

Private Sub chkGDownload_Click()
    Jobs.SelectJob getSelectedJobKey
    Jobs.DrawGraph
End Sub

Private Sub chkGUpload_Click()
    Jobs.SelectJob getSelectedJobKey
    Jobs.DrawGraph
End Sub

Private Sub cmlist_Click(Index As Integer)
    Dim jID As Long
    Dim i As Long, tmp As String
    
    jID = getSelectedJobKey
    If jID = 0 Then Exit Sub
    
    Select Case Index
        Case 1
            lstLog(jID).Clear
        Case 2
            For i = 1 To lstLog(jID).ListCount
                tmp = tmp & lstLog(jID).List(i - 1) & vbCrLf
            Next
            Clipboard.Clear
            Clipboard.SetText tmp
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyA) And (GetAsyncKeyState(vbKeyControl) <> 0) Then _
        SelectAllItems
    If KeyCode = vbKeyDelete Then DeleteTorrentJob
End Sub

Private Sub Form_Load()
    WeQuit = False
    
    'Design
    DrawForm
    LoadWindowPos Me, "downf", 7050, 4920
    Global_MenuCustomizer.LoadState lv1
    GraphikObj.Init picGrf
    GraphikObj.DrawGraphik
    
    If prgSettings.SaveList Then Jobs.LoadState
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    WeQuit = True
    
    Jobs.SaveState
    Global_MenuCustomizer.SaveState lv1
    
    Set GraphikObj = Nothing
    Set Jobs = Nothing
End Sub

Private Sub Form_Resize()
    Dim i As Integer, tmp As ListBox
    Dim vOffset As Long
    
    Dim hOffset As Long
    
    vOffset = (GetSystemMetrics(SM_CYCAPTION) + (GetSystemMetrics(SM_CYFRAME)) * 2) * 15
    hOffset = (GetSystemMetrics(SM_CXFRAME) * 2) * 15
    
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Height < 3000 Then Exit Sub
    
    ts.Width = Me.Width - hOffset
    ts.Top = Me.Height - ts.Height - vOffset - 30
    
    lv1.Width = Me.Width - hOffset
    lv1.Height = ts.Top - lv1.Top + 15
    
    For i = Frame.LBound To Frame.UBound
        Frame(i).Width = Me.Width - hOffset
        Frame(i).Height = ts.Top - Frame(i).Top + 15
    Next
    
    picGrf.Width = Frame(2).Width - 240
    picGrf.Height = Frame(2).Height - picGrf.Top - 120
    
    For Each tmp In lstLog
        tmp.Width = Frame(1).Width - 240
        tmp.Height = Frame(1).Height - tmp.Top - 75
    Next
    
    lblNosel.Left = (Frame(0).Width - lblNosel.Width) / 2
    lblNosel.Top = (Frame(0).Height - lblNosel.Height) / 2
    If ts.SelectedItem.Index = GRAPH_TAB Then
        Jobs.DrawGraph
    End If
End Sub

Private Sub Frame_DblClick(Index As Integer)
    MsgBox Me.Width
    MsgBox Frame(Index).Width
End Sub

Private Sub lstLog_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then Me.PopupMenu cMenu
End Sub

Private Sub lv1_AfterLabelEdit(Cancel As Integer, NewString As String)
    Dim ind As Long
    
    If Len(Trim(NewString)) = 0 Then
        Cancel = 1
    Else
        Jobs.SelectJob getSelectedJobKey
        Jobs.name = NewString
    End If
End Sub

Private Sub lv1_Click()
    FillForms
End Sub

Private Sub optMode_Click(Index As Integer)
    Jobs.SelectJob getSelectedJobKey
    Jobs.Mode = Index
    Jobs.DrawSingle lv1
End Sub

Private Sub lv1_DblClick()
    If getSelectedJobKey > 0 Then _
        frmDownloadSett.Show vbModal
End Sub

Private Sub lv1_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
    FilePoolFromEmulationDrop data
    ProcceedFilePool
    Jobs.SaveState
End Sub

Private Sub Socket1_Close(Index As Integer)
    AddToLog Index, getCStr(11, "Data was NOT updated")
    Socket1(Index).Close
    Socket2(Index).Close
    Jobs.SelectJob Index
    Jobs.StepMode = 0
End Sub

Private Sub Socket1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Jobs.SelectJob Index
    If Number = 10060 Then Jobs.conErr = True _
        Else AddToLog Index, "Error: (" & CStr(Number) & ") " & Description
    Socket1(Index).Close
    Socket2(Index).Close
End Sub

Private Sub Socket1_SendComplete(Index As Integer)
    Socket1(Index).Close
    Jobs.SelectJob Index
    Jobs.SocketOutBuffer = vbNullString
End Sub

Private Sub Socket2_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Jobs.SelectJob Index
    If Number = 10060 Then Jobs.conErr = True _
        Else AddToLog Index, "Error: (" & CStr(Number) & ") " & Description
    Socket2(Index).Close
End Sub

Private Sub Socket3_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim conTime As Long
    Dim i As Long, remPort As Long, remServer As String
    On Error Resume Next
    
    Jobs.SelectJob Index
    
    If Jobs.TryingToConnect Then
        If Socket4(Index).State <> sckClosed Then Socket4(Index).Close
        Socket4(Index).Accept requestID
        Socket4(Index).Close
        AddToLog Index, getCStr(12, "Connection request") & " [" & requestID & "] " & getCStr(13, "rejected")
        Exit Sub
    End If
    
    AddToLog Index, getCStr(12, "Connection request") & " [" & requestID & "]"
    If (Socket1(Index).State <> sckClosed) Or (Socket2(Index).State <> sckClosed) Then
        Socket1(Index).Close
        Socket2(Index).Close
    End If
    
    If prgSettings.chkUseProxy = 1 Then
        Socket2(Index).RemoteHost = Trim(prgSettings.txtProxyIp)
        Socket2(Index).RemotePort = prgSettings.txtProxyPort
    Else
        remPort = 80  'HTTP
        remServer = Jobs.URL
        'Check if URL looks like url.com:port
        i = InStr(1, remServer, ":")
        If i > 0 Then
            remPort = val(Mid(remServer, i + 1))
            If remPort = 0 Then remPort = 80
            remServer = Left(remServer, i - 1)
        End If
        Socket2(Index).RemotePort = remPort
        Socket2(Index).RemoteHost = remServer
    End If
    Socket2(Index).Connect
        
    conTime = Timer
    Jobs.conErr = False
    Jobs.TryingToConnect = True
        
    Do Until (Socket2(Index).State = sckConnected) Or Jobs.conErr 'Trying to conect server first
        Jobs.conErr = ((conTime + 20 < Timer) Or WeQuit) 'Time out 20 seconds
        DoEvents
        Jobs.SelectJob Index '!
    Loop
        
    If Not Jobs.conErr Then
        Jobs.SocketOutBuffer = vbNullString
        Socket1(Index).Accept (requestID)
            
        If err.Number <> 0 Then
            AddToLog Index, getCStr(14, "Error on accept: ") & err.Description
            err.Clear
        End If
    Else
        AddToLog Index, getCStr(15, "Time out")
    End If
        
    Jobs.TryingToConnect = False
End Sub

Private Sub Socket1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strDataOut As String, i As Long, i2 As Long
    Dim remPort As Long, remServer As String
    
    Jobs.SelectJob Index
    Socket1(Index).GetData strDataOut, vbString
    
    'If data alredy passed
    If Jobs.StepMode <> 1 Then
        AddToLog Index, "DEBUG: double sending"
        Exit Sub
    End If
        
    MainProccess Index, strDataOut

    'Host
    i = InStr(1, strDataOut, HOST)
    If i > 0 Then i2 = InStr(i + 1, strDataOut, Chr(&HD), vbBinaryCompare)

    If i > 0 And i2 > 0 Then
        strDataOut = Left(strDataOut, i - 1) & HOST & Trim(Jobs.URL) & Mid(strDataOut, i2)
    End If
    
    If prgSettings.chkUseProxy = 1 Then
        'Check URL for proxy
        If Left(strDataOut, 4) = "GET " Then _
            strDataOut = "GET http://" & Jobs.URL & Mid(strDataOut, 5)
    End If
    
    Socket2(Index).SendData strDataOut
    Jobs.StepMode = 2 'Data sent
    
    'log
    AddToFileLog "Sending", strDataOut
End Sub

Public Sub Socket2_Close(Index As Integer)
    AddToLog Index, getCStr(17, "Data updated successfully")
    Socket2(Index).Close
    Jobs.SelectJob Index
    Socket1(Index).SendData Jobs.SocketOutBuffer
    Jobs.SocketOutBuffer = vbNullString
    Jobs.StepMode = 3 'Final step
End Sub

Private Sub Socket2_Connect(Index As Integer)
    AddToLog Index, getCStr(171, "Connected to: ") & Socket2(Index).RemoteHost
    Jobs.SelectJob Index
    Jobs.StepMode = 1 'Connected to tracker
End Sub

Private Sub Socket2_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strDataIn As String
    On Error Resume Next
    
    Jobs.SelectJob Index
    Socket2(Index).GetData strDataIn, vbString
    
    Jobs.SocketOutBuffer = Jobs.SocketOutBuffer & strDataIn
    DoEvents
    If err.Number <> 0 Then
        AddToLog Index, getCStr(18, "Error: ") & CStr(err.Number)
        err.Clear
    End If
    
    AddToFileLog "Recieving", strDataIn
End Sub

Public Sub AddToLog(jID As Integer, inx As String)
    lstLog(jID).AddItem Format(time, "hh:mm") & ": " & inx, 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Add
            Add_OpenTorrent
        Case 2 'Delete
            DeleteTorrentJob
        Case 3 'Turn on
            TurnOnAllSelected
        Case 4 'Turn off
            TurnOffAllSelected
        Case 5 'Settings
            lv1_DblClick
    End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Index
        Case 1
            Add_OpenTorrent
        Case 2
            RunWorkDefault
    End Select
End Sub

Private Sub ts_Click()
    Dim i As Integer, cV As Integer
    
    If ts.SelectedItem.Index = 1 Then
        cV = -1
    Else
        cV = IIf(getSelectedJobKey = 0, 0, ts.SelectedItem.Index - 1)
    End If
    
    lv1.Visible = (cV = -1)
    For i = Frame.LBound To Frame.UBound
        Frame(i).Visible = (cV = i)
    Next
    If cV = GRAPH_TAB - 1 Then
        Jobs.SelectJob getSelectedJobKey
        Jobs.DrawGraph
    End If
End Sub

Private Sub MainProccess(jID As Integer, ByRef strDataOut As String)
    Dim i As Long, i2 As Long
    Dim XX As Integer
    Dim oldBytes As Currency, newBytes As Currency, oldEvent As String
    Dim sVal As Currency
    Dim m2min As Double
    'Dim tmpUpSpeed As Long
    Dim tmp As String
    Dim curSpeed As Long
    Dim curTime As Long
    
    AddToFileLog "Before sending", strDataOut
    'For mode 2
    oldBytes = Jobs.last_uploaded
    oldEvent = Jobs.last_event
    
    FillMsgData jID, strDataOut
    curTime = ToUnixTime(Now)
    '=====Modify head
    
    i = InStr(1, strDataOut, UPKEY)
    If i > 0 Then
        i2 = InStr(i + 1, strDataOut, "&")
        If i2 = 0 Then i2 = InStr(i + 1, strDataOut, " ")
    End If
    If (i > 0) And (i2 > 0) Then
        
        'Mode 1
        If Jobs.Mode = 1 Then
            oldBytes = CCur(Mid(strDataOut, i + Len(UPKEY), i2 - i - Len(UPKEY)))
            XX = val(Jobs.m1_coef)
            newBytes = oldBytes * XX
        End If
        
        'Mode 2
        If Jobs.Mode = 2 Then
            If Jobs.m2_from <= 0 Then Jobs.m2_from = 1
            If Jobs.m2_to < Jobs.m2_from Then Jobs.m2_to = 1.5
            m2min = getRandomN(Jobs.m2_from, Jobs.m2_to)
            newBytes = Format(Jobs.last_downloaded * m2min, "#0")
            If (oldEvent = "stopped") And (Jobs.last_event = TORRENT_START) Then oldBytes = 0
            If newBytes < oldBytes Then newBytes = oldBytes
        End If
        
        
        If curTime = Jobs.last_cur_time Then
            curSpeed = 0
        Else
            curSpeed = ((newBytes - Jobs.last_uploaded) / (curTime - Jobs.last_cur_time)) / CKILO
        End If
        
        'Smart system
        If Jobs.smart_use And (Jobs.last_cur_time > 0) And (Jobs.last_event <> TORRENT_START) Then
            'AddToLog jID, "Debug: Activate smart system" 'Debug
            'AddToLog jID, "Debug: current speed " & CStr(curSpeed) 'Debug
            'AddToLog jID, "Debug: last speed " & CStr(Jobs.last_cur_speed) 'Debug
            If Abs(Jobs.last_cur_speed - curSpeed) > Jobs.smart_a Then
                'AddToLog jID, "Debug: Activate changing" 'Debug
                XX = Sgn(curSpeed - Jobs.last_cur_speed)
                If XX = 0 Then XX = 1
                curSpeed = (1 + (Jobs.smart_p / 100) * XX) * Jobs.last_cur_speed
                If Abs(curSpeed - Jobs.last_cur_speed) < Jobs.smart_a Then
                    'AddToLog jID, "Debug: to small change, add smart_a" 'Debug
                    curSpeed = Jobs.last_cur_speed + (Jobs.smart_a * XX)
                    If curSpeed < 0 Then curSpeed = 0
                End If
                'AddToLog jID, "Debug: changed to " & CStr(curSpeed) 'Debug
                newBytes = Jobs.last_uploaded + curSpeed * (curTime - Jobs.last_cur_time)
            End If
        End If
        
        newBytes = fix16K(newBytes)
        If (Jobs.last_uploaded > newBytes) Then newBytes = Jobs.last_uploaded
        
        'Log
        If Jobs.Mode = 1 Then
            AddToLog jID, getCStr(16, "Uploaded: ") & CStr(oldBytes) & " => " & CStr(newBytes)
        Else
            AddToLog jID, getCStr(16, "Uploaded: ") & CStr(newBytes) & " (k=" & Format(m2min, "#0.00") & ")"
        End If
        
        Jobs.last_cur_speed = curSpeed
        Jobs.last_uploaded = newBytes
        Jobs.last_real_upload = oldBytes
        Jobs.last_cur_time = curTime
        strDataOut = Left(strDataOut, i - 1) & UPKEY & CStr(newBytes) & Mid(strDataOut, i2)
    End If
    
    'Change Download
    If Jobs.Mode = 1 Then
        i = InStr(1, strDataOut, DOWNKEY)
        If i > 0 Then
            i2 = InStr(i + 1, strDataOut, "&")
            If i2 = 0 Then i2 = InStr(i + 1, strDataOut, " ")
        End If
        If (i > 0) And (i2 > 0) Then
            'Do not send download data
            If Jobs.m1_dwn_notsend Then
                AddToLog jID, getCStr(172, "Downloaded: ") & CStr(Jobs.last_downloaded)
                strDataOut = Left(strDataOut, i - 1) & DOWNKEY & CStr(Jobs.last_downloaded) & Mid(strDataOut, i2)
                
                'Change left
                i = InStr(1, strDataOut, LEFTKEY)
                If i > 0 Then
                    i2 = InStr(i + 1, strDataOut, "&")
                    If i2 = 0 Then i2 = InStr(i + 1, strDataOut, " ")
                End If
                If (i > 0) And (i2 > 0) Then
                    strDataOut = Left(strDataOut, i - 1) & LEFTKEY & CStr(Jobs.last_left) & Mid(strDataOut, i2)
                End If
            End If
            
            'Decrease download data
            If Jobs.m1_isdown Then
                sVal = val(Jobs.m1_downcoef)
                If sVal < 1 Then sVal = 1
                newBytes = fix16K(Jobs.last_downloaded / sVal)
                AddToLog jID, getCStr(172, "Downloaded: ") & CStr(Jobs.last_downloaded) & " => " & Format(newBytes, "#0")
                Jobs.last_downloaded = Format(newBytes, "#0")
                strDataOut = Left(strDataOut, i - 1) & DOWNKEY & CStr(newBytes) & Mid(strDataOut, i2)
            End If
        End If
    End If
    
    'Redraw graph
    Jobs.AddHistoryData
    If Me.ts.SelectedItem.Index = GRAPH_TAB Then _
        Jobs.DrawGraph
    
    'Client version
    If Jobs.useClient And (Jobs.ClientID > 0) And (clnt.getCount >= Jobs.ClientID) Then
        clnt.LoadClientById Jobs.ClientID
        newClient.peer_id = MakeNDigit(clnt.Prefix, clnt.PrefixSize, Jobs.last_peer_id)
        newClient.Cleint = clnt.UserAgent
        'Change peer_id
        i = InStr(1, strDataOut, PEERID)
        If i > 0 Then
            i2 = InStr(i + 1, strDataOut, "&")
            If i2 = 0 Then i2 = InStr(i + 1, strDataOut, " ")
        End If
        If (i > 0) And (i2 > 0) Then
            strDataOut = Left(strDataOut, i - 1) & PEERID & newClient.peer_id & Mid(strDataOut, i2)
        End If
        'Change user-Agent
        If newClient.Cleint <> vbNullString Then
            i = InStr(1, strDataOut, UAGENT)
            If i > 0 Then i2 = InStr(i + 1, strDataOut, Chr(&HD), vbBinaryCompare)
            If i > 0 And i2 > 0 Then
                strDataOut = Left(strDataOut, i - 1) & UAGENT & newClient.Cleint & Mid(strDataOut, i2)
            End If
        End If
    End If
End Sub

Private Sub FillMsgData(jID As Integer, dt As String)
    Dim i As Long, i2 As Long, tmp As String, tmp2 As String
    Dim uKey As String, uVal As String
    Const MSTART = "GET"
    Const MEND = "HTTP"
    
    i = InStr(1, dt, MSTART)
    If i = 0 Then
        AddToLog jID, "URL Parsing error 1"
        MakeErrorLog dt
        Exit Sub 'Some kind of error
    End If
    i = i + Len(MSTART)
    i2 = InStr(i, dt, MEND)
    If i2 = 0 Then
        AddToLog jID, "URL Parsing error 2"
        MakeErrorLog dt
        Exit Sub 'Some kind of error
    End If
    
    tmp = Trim(Mid(dt, i, i2 - i)) ' Get url
    i = InStr(1, tmp, "?")
    If i > 0 Then tmp = Mid(tmp, i + 1)
    clearLastData
    Do
        i = InStr(1, tmp, "&")
        If i > 0 Then
            tmp2 = Mid(tmp, 1, i - 1)
            tmp = Mid(tmp, i + 1)
        Else
            tmp2 = tmp
            tmp = vbNullString
        End If
        i2 = InStr(1, tmp2, "=")
        If i2 = 0 Then
            uKey = tmp2
            uVal = vbNullString
        Else
            uKey = Left(tmp2, i2 - 1)
            uVal = Mid(tmp2, i2 + 1)
        End If
        Select Case LCase(uKey)
            Case "info_hash": Jobs.last_info_hash = uVal
            Case "peer_id": Jobs.last_peer_id = uVal
            Case "uploaded": Jobs.last_uploaded = uVal
            Case "downloaded": Jobs.last_downloaded = uVal
            Case "event": Jobs.last_event = uVal
            Case "port": Jobs.last_port = val(uVal)
            Case "left": Jobs.last_left = uVal
            Case "numwant": Jobs.last_numwant = val(uVal)
            Case "no_peer_id": Jobs.last_no_peer_id = val(uVal)
            Case "compact": Jobs.last_compact = val(uVal)
            Case "key": Jobs.last_key = uVal
        End Select
    Loop Until i = 0
End Sub

Private Sub clearLastData()
    Jobs.last_event = "?"
End Sub

Private Sub DrawForm()
    cmlist(1).Caption = getCStr(36, "Clear log")
    cmlist(2).Caption = getCStr(37, "Copy to clipboard")
    
    TranslateForm Me
End Sub

Public Sub ProcceedFilePool()
    Dim i As Long
    
    For i = 1 To Global_FilePool.Count
        RunWorkTorrent Global_FilePool.Item(i)
    Next
End Sub

Public Sub Add_OpenTorrent()
    Dim i As New clsCommonDialog
    Dim tmp As String
    
    tmp = vbNullString
    If i.OpenFileName(tmp, , True, False, False, False, "Torrent files|*.torrent") Then
        tmp = MakeShortName(tmp)
        RunWorkTorrent tmp
        Jobs.SaveState
    End If
    
    Set i = Nothing
End Sub

Public Sub RunWorkTorrent(FileN As String)
    Dim dt As D_OUTDATA
    Dim jID As Integer
    
    jID = Jobs.AddJob
    Jobs.SelectJob jID
    
    If RunTorrentWork(FileN, Jobs.port, dt) Then
        Jobs.name = dt.torName
        Jobs.URL = dt.realURL
        
        ShellExecute Me.hwnd, "open", dt.tmpFile, vbNullString, vbNullString, vbNormalFocus
        
        Jobs.DrawFullList lv1
        FillForms
        TurnJobOn jID
    Else
        Jobs.DeleteJob jID
    End If
End Sub

Public Sub RunWorkDefault()
    Dim jID As Integer
    
    jID = Jobs.AddJob
    Jobs.SelectJob jID
    Jobs.name = vbNullString
    Jobs.URL = prgSettings.txtServer
    Jobs.DrawFullList lv1
    FillForms
End Sub

Public Function getSelectedJobKey() As Long
    Dim tmp As String
    On Error Resume Next
    
    tmp = lv1.SelectedItem.key
    If err.Number <> 0 Then
        getSelectedJobKey = 0
        Exit Function
    End If
    
    getSelectedJobKey = val(Mid(tmp, 2))
End Function

Public Sub FillForms()
    Dim i As Integer, tmp As ListBox
    
    i = getSelectedJobKey
    If i = 0 Then Exit Sub
    
    If Jobs.SelectJob(i) Then
        For Each tmp In lstLog
            tmp.Visible = (tmp.Index = i)
        Next
    End If
End Sub

Private Sub DeleteTorrentJob()
    Dim i As Integer, j As Integer
    
    i = getSelectedJobKey
    If i = 0 Then Exit Sub
    
    If MsgBox(getCStr(183, "Are you sure you want remove selected task(s)?"), vbYesNo + vbExclamation, App.Title) = vbYes Then
        For i = lv1.ListItems.Count To 1 Step -1
            If lv1.ListItems(i).Selected Then
                j = val(Mid(lv1.ListItems(i).key, 2))
                Jobs.DeleteJob j
                lv1.ListItems.Remove i
            End If
        Next
        ts_Click
        Jobs.SaveState
    End If
End Sub

Private Sub TurnOnAllSelected()
    Dim i As Integer
    
    For i = 1 To Me.lv1.ListItems.Count
        If Me.lv1.ListItems(i).Selected Then _
            TurnJobOn CInt(val(Mid(Me.lv1.ListItems(i).key, 2)))
    Next
End Sub

Private Sub TurnOffAllSelected()
    Dim i As Integer
    
    For i = 1 To Me.lv1.ListItems.Count
        If Me.lv1.ListItems(i).Selected Then _
            TurnJobOff CInt(val(Mid(Me.lv1.ListItems(i).key, 2)))
    Next
End Sub

Public Function TurnJobOn(jID As Integer) As Boolean
    On Error Resume Next

    TurnJobOn = False
    Jobs.SelectJob jID
    
    If Jobs.Status = STATUS_ON Then
        TurnJobOn = True
        Exit Function
    End If
    
    If (val(Jobs.port) > 65000) Or (val(Jobs.port) < 1) Then
        MsgBox getCStr(6, "Wrong port"), vbCritical, App.Title
        Exit Function
    End If
    If Trim(Jobs.URL) = vbNullString Then
        MsgBox getCStr(7, "Wrong tracker URL"), vbCritical, App.Title
        Exit Function
    End If
    With Socket3(jID)
        .LocalPort = val(Jobs.port)
        .Listen
    End With
    If err.Number <> 0 Then
        AddToLog jID, getCStr(8, "Error on listen: ") & err.Description
        err.Clear
    Else
        Jobs.Status = STATUS_ON
        AddToLog jID, getCStr(9, "Awaiting BitTorrent client connection")
        Jobs.DrawSingle Me.lv1
        TurnJobOn = True
    End If
End Function

Public Function TurnJobOff(jID As Integer) As Boolean
    Jobs.SelectJob jID
    
    If Jobs.Status = STATUS_OFF Then
        TurnJobOff = True
        Exit Function
    End If
    
    Socket3(jID).Close
    Jobs.Status = STATUS_OFF
    AddToLog jID, getCStr(10, "Turning off")
    Jobs.DrawSingle Me.lv1
    
    TurnJobOff = True
End Function

Private Sub SelectAllItems()
    Dim i As Long
    
    For i = 1 To lv1.ListItems.Count
        lv1.ListItems(i).Selected = True
    Next
End Sub
