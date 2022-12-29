VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl32.ocx"
Begin VB.MDIForm MainForm 
   BackColor       =   &H8000000C&
   Caption         =   "BitTorrent Proxy v"
   ClientHeight    =   6300
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10710
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock rc 
      Left            =   10080
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stb 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6045
      Width           =   10710
      _ExtentX        =   18891
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17833
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   529
            MinWidth        =   529
            Text            =   "R"
            TextSave        =   "R"
            Object.ToolTipText     =   "Remote Control"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10080
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":365A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":3D6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":4482
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":4B96
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":52AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":59BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":60D2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10710
      _ExtentX        =   18891
      _ExtentY        =   1217
      ButtonWidth     =   1799
      ButtonHeight    =   1164
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Downloading"
            Object.Tag             =   "~111"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Emulation"
            Object.Tag             =   "~112"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Settings"
            Object.Tag             =   "~113"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Object.Tag             =   "~142"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            Object.Tag             =   "~114"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Object.Tag             =   "~115"
            ImageIndex      =   5
         EndProperty
      EndProperty
      Begin VB.PictureBox picDonate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   9195
         MousePointer    =   10  'Up Arrow
         Picture         =   "MainForm.frx":67E6
         ScaleHeight     =   600
         ScaleWidth      =   1515
         TabIndex        =   2
         Top             =   30
         Width           =   1515
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RemoteControl As New clsRemote

Public Sub Tray_Click(Index As Long)
    On Error Resume Next
    
    Select Case Index
        Case 1000 'Show form
            Me.Show
        Case 1100 'Add download window
            Me.Show
            AppArg = vbNullString
            frmMain.Show
            frmMain.Add_OpenTorrent
        Case 1200 'Add emulation
            Me.Show
            frmEmulate.Show
            frmEmulate.AddFile
        Case 1300 'Settings
            Me.Show
            frmSettings.Show , Me
        Case 1400 'Check update
            CheckUpdate True
        Case 1500 ' Help
            OpenInstruction
        Case 1600 'Home
            OpenHomepage
        Case 1700 'Forum
            OpenForum
        Case 1800 'About
            Me.Show
            frmAbout.Show , Me
        Case 1900 'Exit
            Unload Me
    End Select
    
    err.Clear
End Sub

Private Sub MDIForm_Load()
    Randomize Timer
    
    PopupMenuCreate
    
    prev = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WinProc)

    'Update check
    #If Not debugver Then
        If (Rnd() * 5) > 4 Then _
            SetTimer Me.hwnd, 0, 30000, AddressOf TimerProc
    #End If
    
    'Logs
    #If debugver Then
        If FileExist(CPath & FLOG) Then _
            Open CPath & FLOG For Append As #1 _
        Else _
            Open CPath & FLOG For Output As #1
    #End If
    
    'Tray
    AddSysTray
    
    DrawForm
    LoadWindowPos Me, "mainf", 10830, 6705
    Show
    DragAcceptFiles Me.hwnd, True
    LoadClients
    
    'Remote control
    RemoteControl.set_socket rc
    If prgSettings.chkRemote = 1 Then _
        RemoteControl.StartRC
    
    If Global_FilePool.Count > 0 Then _
        MakeAsk
End Sub

Private Sub DrawForm()
    Me.Caption = Me.Caption & CStr(App.Major) & "." & CStr(App.Minor)
    TranslateForm Me
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If prgSettings.chkMinimize And UnloadMode = 0 Then
        Cancel = 1
        Me.WindowState = vbMinimized
    Else
        #If debugver Then
            Close 1
        #End If
        KillTray Me
        SaveWindowPos Me, "mainf"
        DestroyMenu hPopMenu
    End If
End Sub

Private Sub MDIForm_Resize()
    On Error Resume Next
    
    If Me.WindowState = vbMinimized Then
        Me.Visible = False
        Me.WindowState = vbNormal
    Else
        picDonate.Left = Me.Width - picDonate.Width - 120
    End If
    err.Clear
End Sub

Private Sub picDonate_Click()
    ShellExecute MainForm.hwnd, vbNullString, HTTP & SITEURL & DONATIONURL, vbNullString, App.Path, vbNormalFocus
End Sub

Private Sub rc_ConnectionRequest(ByVal requestID As Long)
    If rc.State <> sckClosed Then rc.Close
    rc.Accept requestID
End Sub

Private Sub rc_DataArrival(ByVal bytesTotal As Long)
    Dim tmp As String
    
    rc.GetData tmp, vbString
    RemoteControl.got_data tmp
End Sub

Private Sub rc_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RemoteControl.ShowSocketError Description
End Sub

Private Sub rc_SendComplete()
    rc.Close
    rc.Listen
End Sub

Private Sub stb_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
    Me.stb.Panels(1).Text = vbNullString
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            frmMain.Show
        Case 2
            frmEmulate.Show
        Case 3
            frmSettings.Show vbModal
            RemoteControl.StopRC
            If prgSettings.chkRemote = 1 Then _
                RemoteControl.StartRC
        Case 5
            OpenInstruction
        Case 6
            frmAbout.Show , Me
        Case 8
            Unload Me
    End Select
End Sub

Public Sub MakeAsk()
    Select Case prgSettings.defAction
        Case 0
            frmMain.Show
            frmMain.ProcceedFilePool
        Case 1
            frmEmulate.Show
            frmEmulate.ProcceedFilePool
        Case 2
            frmActionSelect.Show vbModal, MainForm
    End Select
End Sub

Private Sub LoadClients()
    Dim tmp As String, ret As Boolean
    
    ret = False
    tmp = Dir(CPath & CLIENT_DIR & "*.ini")
    Do While Len(tmp) > 0
        stb.Panels(1).Text = getCStr(173, "Loading clients: ") & tmp
        stb.Refresh
        ret = ret Or clnt.AddClient(CPath & CLIENT_DIR & tmp)
        tmp = Dir
    Loop
    stb.Panels(1).Text = vbNullString
    If Not ret Then
        MsgBox getCStr(174, "Can not find any client file!"), vbCritical, App.Title
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(2).Enabled = False
    End If
End Sub

Private Sub OpenInstruction()
    ShellExecute Me.hwnd, vbNullString, HTTP & SITEURL & HELPURL, vbNullString, App.Path, vbNormalFocus
End Sub

Public Sub AddSysTray()
    ShowTray Me, App.Title, Me.Icon.Handle
End Sub
