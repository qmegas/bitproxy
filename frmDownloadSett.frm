VERSION 5.00
Begin VB.Form frmDownloadSett 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Download task settings"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   5895
   StartUpPosition =   1  'CenterOwner
   Tag             =   "~202"
   Begin VB.Frame Frame2 
      Caption         =   "Smooth speed change system"
      Height          =   1455
      Left            =   0
      TabIndex        =   31
      Tag             =   "~207"
      Top             =   4320
      Width           =   5895
      Begin VB.TextBox txtSmartP 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4920
         TabIndex        =   13
         Top             =   1050
         Width           =   375
      End
      Begin VB.TextBox txtSmartA 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4920
         TabIndex        =   12
         Top             =   675
         Width           =   375
      End
      Begin VB.CheckBox chkSmart 
         Caption         =   "Use smooth speed change system"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Tag             =   "~209"
         Top             =   360
         Width           =   5535
      End
      Begin VB.Label Label12 
         Caption         =   "%"
         Height          =   255
         Left            =   5400
         TabIndex        =   35
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Not let to speed change more than on:"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Tag             =   "~211"
         Top             =   1080
         Width           =   4695
      End
      Begin VB.Label Label10 
         Caption         =   "Kb/s"
         Height          =   255
         Left            =   5400
         TabIndex        =   33
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Activate when speed changes more than on:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Tag             =   "~210"
         Top             =   720
         Width           =   4695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Client settings"
      Height          =   735
      Left            =   0
      TabIndex        =   30
      Tag             =   "~208"
      Top             =   5880
      Width           =   5895
      Begin VB.ComboBox cmbVersion 
         Height          =   315
         ItemData        =   "frmDownloadSett.frx":0000
         Left            =   3240
         List            =   "frmDownloadSett.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   330
         Width           =   2535
      End
      Begin VB.CheckBox chkVersion 
         Caption         =   "Change BitTorrent client version to:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Tag             =   "~34"
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save && Close"
      Height          =   375
      Left            =   2640
      TabIndex        =   18
      Tag             =   "~220"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4440
      TabIndex        =   17
      Tag             =   "~219"
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Turn On"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Tag             =   "~21"
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Frame Frame 
      Caption         =   "General settings"
      Height          =   1455
      Index           =   1
      Left            =   0
      TabIndex        =   26
      Tag             =   "~205"
      Top             =   0
      Width           =   5895
      Begin VB.ComboBox txtServer 
         Height          =   315
         ItemData        =   "frmDownloadSett.frx":0004
         Left            =   1440
         List            =   "frmDownloadSett.frx":0006
         Sorted          =   -1  'True
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   300
         Width           =   3495
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   1
         Text            =   "6666"
         Top             =   690
         Width           =   855
      End
      Begin VB.TextBox txtJobName 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   1050
         Width           =   3495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tracker URL:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Tag             =   "~20"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Port:"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Tag             =   "~26"
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Job name:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Tag             =   "~196"
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Mode settings"
      Height          =   2655
      Index           =   2
      Left            =   0
      TabIndex        =   19
      Tag             =   "~206"
      Top             =   1560
      Width           =   5895
      Begin VB.CheckBox chkDnotsend 
         Caption         =   "Do not send downloading information"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Tag             =   "~218"
         Top             =   1320
         Width           =   5295
      End
      Begin VB.TextBox txtUpload 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         MaxLength       =   7
         TabIndex        =   4
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton optMode 
         Caption         =   "Static change mode"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Tag             =   "~27"
         Top             =   240
         Value           =   -1  'True
         Width           =   3735
      End
      Begin VB.OptionButton optMode 
         Caption         =   "Dynamic change mode"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Tag             =   "~30"
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox txtM2from 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3960
         TabIndex        =   9
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox txtM2to 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5160
         TabIndex        =   10
         Top             =   2160
         Width           =   495
      End
      Begin VB.CheckBox chkDown 
         Caption         =   "Decrease downloading in:"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Tag             =   "~105"
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtDown 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2640
         MaxLength       =   7
         TabIndex        =   6
         Top             =   960
         Width           =   495
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   120
         X2              =   5760
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label5 
         Caption         =   "Keep rating in limits:"
         Height          =   255
         Left            =   480
         TabIndex        =   25
         Tag             =   "~31"
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "from:"
         Height          =   255
         Left            =   3360
         TabIndex        =   24
         Tag             =   "~32"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "to:"
         Height          =   255
         Left            =   4680
         TabIndex        =   23
         Tag             =   "~33"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Increase upload in:"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Tag             =   "~28"
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label7 
         Height          =   255
         Left            =   2760
         TabIndex        =   21
         Tag             =   "~29"
         Top             =   600
         Width           =   375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   120
         X2              =   5760
         Y1              =   1665
         Y2              =   1665
      End
      Begin VB.Label Label9 
         Height          =   255
         Left            =   3240
         TabIndex        =   20
         Tag             =   "~29"
         Top             =   960
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmDownloadSett"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
        
Private curID As Integer
Private canSave As Boolean
        
Private Sub Check1_Click()
    If canSave Then cmdSave.Visible = True
End Sub

Private Sub chkDnotsend_Click()
    If canSave Then cmdSave.Visible = True
    
    If chkDnotsend.Value = 1 And chkDown.Value = 1 Then _
        chkDown.Value = 0
End Sub

Private Sub chkDown_Click()
    If canSave Then cmdSave.Visible = True
    
    If chkDown.Value = 1 And chkDnotsend.Value = 1 Then _
        chkDnotsend.Value = 0
End Sub

Private Sub chkSmart_Click()
    If canSave Then cmdSave.Visible = True
End Sub

Private Sub chkVersion_Click()
    If canSave Then cmdSave.Visible = True
End Sub

Private Sub cmbVersion_Click()
    If canSave Then cmdSave.Visible = True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Jobs.SelectJob curID
    
    If val(txtSmartP.Text) < 1 Or val(txtSmartP.Text) > 100 Then _
        txtSmartP.Text = 30
    
    If Jobs.Status = STATUS_OFF Then
        Jobs.URL = Trim(txtServer.Text)
        Jobs.port = val(txtPort.Text)
    End If
    Jobs.name = IIf(Len(txtJobName.Text) = 0, "Unnamed", txtJobName.Text)
    Jobs.Mode = IIf(optMode(1).Value, 1, 2)
    Jobs.m1_coef = StringToDouble(txtUpload.Text)
    Jobs.m1_isdown = (Me.chkDown.Value = 1)
    Jobs.m1_downcoef = val(txtDown.Text)
    Jobs.m1_dwn_notsend = (Me.chkDnotsend.Value = 1)
    Jobs.m2_from = CSng(StringToDouble(txtM2from.Text))
    Jobs.m2_to = CSng(StringToDouble(txtM2to.Text))
    Jobs.smart_use = (chkSmart.Value = 1)
    Jobs.smart_a = val(txtSmartA.Text)
    Jobs.smart_p = val(txtSmartP.Text)
    Jobs.useClient = (chkVersion.Value = 1)
    Jobs.ClientID = clnt.getIdByName(cmbVersion.Text)
    Jobs.DrawSingle frmMain.lv1
    
    Jobs.SaveState
        
    SavePSettings
    Unload Me
End Sub

Private Sub Form_Initialize()
    canSave = False
End Sub

Private Sub optMode_Click(Index As Integer)
    If canSave Then cmdSave.Visible = True
End Sub

Private Sub txtDown_Change()
    If canSave Then cmdSave.Visible = True
End Sub

Private Sub txtDown_GotFocus()
    txtDown.SelStart = 0
    txtDown.SelLength = Len(txtDown)
End Sub

Private Sub txtJobName_Change()
    If canSave Then cmdSave.Visible = True
End Sub

Private Sub txtJobName_GotFocus()
    txtJobName.SelStart = 0
    txtJobName.SelLength = Len(txtJobName)
End Sub

Private Sub txtM2from_Change()
    If canSave Then cmdSave.Visible = True
End Sub

Private Sub txtM2from_GotFocus()
    txtM2from.SelStart = 0
    txtM2from.SelLength = Len(txtM2from)
End Sub

Private Sub txtM2to_Change()
    If canSave Then cmdSave.Visible = True
End Sub

Private Sub txtM2to_GotFocus()
    txtM2to.SelStart = 0
    txtM2to.SelLength = Len(txtM2to)
End Sub

Private Sub txtPort_Change()
    If canSave Then cmdSave.Visible = True
End Sub

Private Sub txtPort_GotFocus()
    txtPort.SelStart = 0
    txtPort.SelLength = Len(txtPort)
End Sub

Private Sub txtServer_Change()
    If canSave Then cmdSave.Visible = True
End Sub

Private Sub txtServer_Click()
    If canSave Then cmdSave.Visible = True
End Sub

Private Sub txtSmartA_Change()
    If canSave Then cmdSave.Visible = True
End Sub

Private Sub txtSmartA_GotFocus()
    txtSmartA.SelStart = 0
    txtSmartA.SelLength = Len(txtSmartA)
End Sub

Private Sub txtSmartP_Change()
    If canSave Then cmdSave.Visible = True
End Sub

Private Sub txtSmartP_GotFocus()
    txtSmartP.SelStart = 0
    txtSmartP.SelLength = Len(txtSmartP)
End Sub

Private Sub txtUpload_Change()
    If canSave Then cmdSave.Visible = True
End Sub

Private Sub txtUpload_GotFocus()
    txtUpload.SelStart = 0
    txtUpload.SelLength = Len(txtUpload)
End Sub

Private Sub Form_Load()
    curID = frmMain.getSelectedJobKey
    
    DrawForm
    DrawSettings
    canSave = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call DelComboFlat(txtServer.hwnd)
    Call DelComboFlat(cmbVersion.hwnd)
End Sub

Private Sub DrawForm()
    Me.Icon = MainForm.Icon
    MakeThinAll Me
    Call SetComboFlat(txtServer.hwnd)
    Call SetComboFlat(cmbVersion.hwnd)
    
    TranslateForm Me
End Sub
    
Public Sub Command1_Click()
    Jobs.SelectJob curID
    If Jobs.Status = STATUS_OFF Then
        If frmMain.TurnJobOn(curID) Then
            Command1.Caption = getCStr(22, "Turn off")
            txtPort.Enabled = False
            txtServer.Enabled = False
        End If
    Else
        If frmMain.TurnJobOff(curID) Then
            Command1.Caption = getCStr(21, "Turn on")
            txtPort.Enabled = True
            txtServer.Enabled = True
        End If
    End If
End Sub

Private Sub FillTracklist()
    Dim i As Long, tmp As String
    For i = 1 To MAX_LIST
        tmp = GetKeyValue(HKEY_LOCAL, MYPATH, "tl" & CStr(i), vbNullString)
        If tmp <> vbNullString Then _
            Me.txtServer.AddItem tmp
    Next
End Sub

Private Sub DrawSettings()
    Dim tmp As String
    clnt.FillList Me.cmbVersion
    FillTracklist
    
    Jobs.SelectJob curID
    
    If Jobs.Status = STATUS_OFF Then Command1.Caption = getCStr(21, "Turn On") _
        Else Command1.Caption = getCStr(22, "Turn Off")
    Me.txtJobName.Text = Jobs.name
    Me.txtPort.Text = Jobs.port
    Me.txtPort.Enabled = (Jobs.Status = STATUS_OFF)
    Me.txtServer.Text = Jobs.URL
    Me.txtServer.Enabled = (Jobs.Status = STATUS_OFF)
        
    Me.optMode(Jobs.Mode).Value = True
    Me.txtUpload = Trim(str(Jobs.m1_coef))
    Me.txtDown = Trim(str(Jobs.m1_downcoef))
    Me.chkDown.Value = IIf(Jobs.m1_isdown, 1, 0)
    Me.chkDnotsend.Value = IIf(Jobs.m1_dwn_notsend, 1, 0)
        
    Me.txtM2from = Trim(str(Jobs.m2_from))
    Me.txtM2to = Trim(str(Jobs.m2_to))
    
    Me.chkSmart = IIf(Jobs.smart_use, 1, 0)
    Me.txtSmartA = Jobs.smart_a
    Me.txtSmartP = Jobs.smart_p
        
    Me.chkVersion.Value = IIf(Jobs.useClient, 1, 0)
    tmp = clnt.getClientNameById(Jobs.ClientID)
    If Len(tmp) = 0 Then Me.cmbVersion.ListIndex = 0 _
        Else Me.cmbVersion.Text = tmp
    cmdSave.Visible = False
End Sub

Private Sub SavePSettings()
    Dim i As Integer
    With Me
        WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_HOST, .txtServer
        prgSettings.txtServer = .txtServer
        '=====
        If .optMode(1) Then prgSettings.optMode = 1 Else prgSettings.optMode = 2
        WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_MODE, CStr(prgSettings.optMode)
        WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_UPLOAD, .txtUpload
        prgSettings.txtUpload = val(.txtUpload)
        WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_M2FROM, .txtM2from
        prgSettings.txtM2from = .txtM2from
        WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_M2TO, .txtM2to
        prgSettings.txtM2to = .txtM2to
        WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_DWNUSE, CStr(.chkDown.Value)
        prgSettings.chkDown = .chkDown.Value
        WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_DWNVAL, CStr(.txtDown)
        prgSettings.txtDown = .txtDown
        WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_SMARTUSE, CStr(.chkSmart.Value)
        prgSettings.chkSmart = .chkSmart.Value
        WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_SMARTA, CStr(.txtSmartA)
        prgSettings.txtSmartA = val(.txtSmartA)
        WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_SMARTP, CStr(.txtSmartP)
        prgSettings.txtSmartP = val(.txtSmartP)
        WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_DWNNOTSEND, CStr(.chkDnotsend)
        prgSettings.chkDnotsend = .chkDnotsend.Value
        '=====
        WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_USEVER, CStr(.chkVersion.Value)
        prgSettings.chkVersion = .chkVersion.Value
        WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_VERTYPE, CStr(.cmbVersion.ListIndex)
        prgSettings.cmbVersion = .cmbVersion.ListIndex
        
        'Saving URL list
        If Not isInList(txtServer.Text, txtServer.hwnd) Then _
            txtServer.AddItem txtServer.Text, 0
        For i = 1 To .txtServer.ListCount
            WriteKey HKEY_LOCAL, MYPATH, "tl" & CStr(i), .txtServer.List(i - 1)
        Next
    End With
End Sub
