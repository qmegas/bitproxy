VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl32.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "~110"
   Begin VB.Frame Frame 
      Height          =   3975
      Index           =   3
      Left            =   240
      TabIndex        =   34
      Top             =   360
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CheckBox chkSameHash 
         Caption         =   "Allow to add torrents with same hash"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Tag             =   "~234"
         Top             =   2160
         Width           =   4215
      End
      Begin VB.CheckBox chkFroze 
         Caption         =   "Set speeds to 0 if tracker returns one peer"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Tag             =   "~221"
         Top             =   1800
         Width           =   4335
      End
      Begin VB.TextBox txtConnectTries 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   600
         MaxLength       =   3
         TabIndex        =   18
         Text            =   "0"
         Top             =   1410
         Width           =   495
      End
      Begin VB.CheckBox chkIgnorSocketErr 
         Caption         =   "Ignore socket errors"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Tag             =   "~175"
         Top             =   720
         Width           =   4335
      End
      Begin VB.CheckBox chkIgnorServerError 
         Caption         =   "Ignore server errors"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Tag             =   "~160"
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label5 
         Caption         =   "times (0 = unlimited times)"
         Height          =   255
         Left            =   1200
         TabIndex        =   36
         Tag             =   "~177"
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "Try connect to server for:"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Tag             =   "~176"
         Top             =   1080
         Width           =   3495
      End
   End
   Begin VB.Frame Frame 
      Height          =   3975
      Index           =   2
      Left            =   240
      TabIndex        =   25
      Top             =   360
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CheckBox chkRetracker 
         Caption         =   "Do not remove retracker from announce list"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Tag             =   "~223"
         Top             =   600
         Width           =   4215
      End
      Begin VB.TextBox txtDefPort 
         Height          =   285
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtProxyPort 
         Height          =   285
         Left            =   3480
         MaxLength       =   5
         TabIndex        =   13
         Top             =   2610
         Width           =   855
      End
      Begin VB.TextBox txtProxyIp 
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   2610
         Width           =   1335
      End
      Begin VB.CheckBox chkUseProxy 
         Caption         =   "Use proxy"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Tag             =   "~44"
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   120
         X2              =   4440
         Y1              =   2175
         Y2              =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   120
         X2              =   4440
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label lbPWarn 
         Alignment       =   2  'Center
         Caption         =   $"frmSettings.frx":0000
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   120
         TabIndex        =   33
         Tag             =   "~151"
         Top             =   3000
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "Default port:"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Tag             =   "~109"
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Port:"
         Height          =   255
         Left            =   2880
         TabIndex        =   27
         Tag             =   "~46"
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Proxy IP:"
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Tag             =   "~45"
         Top             =   2640
         Width           =   855
      End
   End
   Begin VB.Frame Frame 
      Height          =   3975
      Index           =   5
      Left            =   240
      TabIndex        =   37
      Top             =   360
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CheckBox chkRemote 
         Caption         =   "Activate remote control functions"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Tag             =   "~212"
         Top             =   240
         Width           =   4215
      End
      Begin VB.TextBox txtRCPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   22
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtRCPort 
         Height          =   285
         Left            =   1440
         TabIndex        =   21
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "It is insistently recommended to use in password only Latin symbols and numbers. Password is case sensitive."
         Height          =   615
         Left            =   240
         TabIndex        =   40
         Tag             =   "~216"
         Top             =   1440
         Width           =   4215
      End
      Begin VB.Label Label7 
         Caption         =   "Password:"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Tag             =   "~129"
         Top             =   990
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Port:"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Tag             =   "~128"
         Top             =   630
         Width           =   1095
      End
   End
   Begin VB.Frame Frame 
      Height          =   3975
      Index           =   1
      Left            =   240
      TabIndex        =   28
      Top             =   360
      Width           =   4575
      Begin VB.CheckBox chkAutoR 
         Caption         =   "Run program on Windows startup"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Tag             =   "~150"
         Top             =   1560
         Width           =   4335
      End
      Begin VB.CheckBox chkMinimize 
         Caption         =   "Minimize to tray on window close"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Tag             =   "~147"
         Top             =   1920
         Width           =   4335
      End
      Begin VB.CommandButton cmdAssoc 
         Caption         =   "Add special menu for torrent files"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Tag             =   "~49"
         Top             =   3600
         Width           =   4335
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Check manualy"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Tag             =   "~40"
         Top             =   1080
         Width           =   4335
      End
      Begin VB.CheckBox chkAutoUpdate 
         Caption         =   "Automatically check new versions"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Tag             =   "~39"
         Top             =   720
         Width           =   4095
      End
      Begin VB.ComboBox cmdLang 
         Height          =   315
         ItemData        =   "frmSettings.frx":008D
         Left            =   1800
         List            =   "frmSettings.frx":008F
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   315
         Width           =   2415
      End
      Begin VB.ComboBox cmdAutoJ 
         Height          =   315
         ItemData        =   "frmSettings.frx":0091
         Left            =   2160
         List            =   "frmSettings.frx":009E
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3120
         Width           =   2055
      End
      Begin VB.CheckBox chkSaveEmul 
         Caption         =   "Save emulation and download job list"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Tag             =   "~121"
         Top             =   2280
         Width           =   4215
      End
      Begin VB.Label Label9 
         Caption         =   "Language:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Tag             =   "~38"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Default action:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Tag             =   "~108"
         Top             =   3150
         Width           =   2055
      End
   End
   Begin VB.Frame Frame 
      Height          =   3975
      Index           =   4
      Left            =   240
      TabIndex        =   31
      Top             =   360
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton cmdClients 
         Caption         =   "Download client's files"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Tag             =   "~124"
         Top             =   3480
         Width           =   4335
      End
      Begin VB.ListBox lstClient 
         Height          =   2790
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label3 
         Caption         =   "List of installed clients:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Tag             =   "~149"
         Top             =   240
         Width           =   4335
      End
   End
   Begin MSComctlLib.TabStrip ts1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   7858
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Object.Tag             =   "~144"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Downloading"
            Object.Tag             =   "~145"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Emulation"
            Object.Tag             =   "~153"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Clients"
            Object.Tag             =   "~146"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Remote Control"
            Object.Tag             =   "~43"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   24
      Tag             =   "~80"
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1560
      TabIndex        =   23
      Tag             =   "~79"
      Top             =   4560
      Width           =   1575
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim onLoadLand As Integer

Private Sub cmdAssoc_Click()
    MakeAssociation True
End Sub

Private Sub cmdClients_Click()
    Const FCLIENT = "/custom/bit_clients2.htm"
    ShellExecute Me.hwnd, vbNullString, HTTP & SITEURL & FCLIENT, vbNullString, App.Path, vbNormalFocus
End Sub

Private Sub cmdUpdate_Click()
    Me.MousePointer = 11
    CheckUpdate True
    Me.MousePointer = 0
End Sub

Private Sub Command1_Click()
    If Len(txtRCPass.Text) = 0 And chkRemote.Value = 1 Then
        MsgBox getCStr(213, "Password for remote control can not be empty!"), vbExclamation, App.Title
        Exit Sub
    End If
    
    prgSettings.LangSelected = cmdLang.Text
    prgSettings.chkAutoUpdate = chkAutoUpdate.Value
    prgSettings.defAction = cmdAutoJ.ListIndex
    prgSettings.txtProxyIp = Trim(txtProxyIp)
    prgSettings.txtProxyPort = Trim(txtProxyPort)
    prgSettings.chkUseProxy = chkUseProxy.Value
    prgSettings.txtPort = val(txtDefPort.Text)
    prgSettings.SaveList = chkSaveEmul.Value
    prgSettings.chkMinimize = chkMinimize.Value
    prgSettings.chkIgnorServerError = chkIgnorServerError.Value
    prgSettings.chkIgnorSocketError = chkIgnorSocketErr.Value
    prgSettings.txtConnectTries = val(txtConnectTries.Text)
    If prgSettings.txtConnectTries < 0 Then prgSettings.txtConnectTries = 0
    prgSettings.chkFroze = (chkFroze.Value = 1)
    prgSettings.chkRemote = chkRemote.Value
    prgSettings.txtRCPort = val(txtRCPort.Text)
    prgSettings.txtRCPass = txtRCPass.Text
    prgSettings.chkRetracker = chkRetracker.Value
    prgSettings.chkSameHash = chkSameHash.Value
    
    
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_AUTOCHECK, CStr(prgSettings.chkAutoUpdate)
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_LANG, prgSettings.LangSelected
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_DEFACTION, CStr(prgSettings.defAction)
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_USEPROXY, CStr(prgSettings.chkUseProxy)
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_PROXYIP, prgSettings.txtProxyIp
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_PROXYPORT, CStr(prgSettings.txtProxyPort)
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_PORT, CStr(prgSettings.txtPort)
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_RETRACKER, CStr(prgSettings.chkRetracker)
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_SAVELIST, CStr(prgSettings.SaveList)
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_MINIMIZE, CStr(prgSettings.chkMinimize)
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_IGNORSERVERR, CStr(prgSettings.chkIgnorServerError)
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_IGNORSOCKETERR, CStr(prgSettings.chkIgnorSocketError)
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_CONNTRIES, CStr(prgSettings.txtConnectTries)
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_FROZE, CStr(chkFroze.Value)
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_SAMEHASH, CStr(chkSameHash.Value)
    '5
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_REMOTEUSE, CStr(prgSettings.chkRemote)
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_REMOTEPASS, prgSettings.txtRCPass
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_REMOTEPORT, CStr(prgSettings.txtRCPort)
    
    SaveAutorun
    
    If cmdLang.ListIndex <> onLoadLand Then
        If cmdLang.ListIndex = 0 Then _
            MsgBox getCStr(122, "To see changes in program laguage you need to restart program") _
        Else _
            ReTranslateNow
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    DrawForm
End Sub

Private Sub DrawForm()
    Me.Icon = MainForm.Icon
    
    MakeThinAll Me
    Call SetComboFlat(cmdLang.hwnd)
    Call SetComboFlat(cmdAutoJ.hwnd)
      
    '1
    Me.chkAutoUpdate.Value = prgSettings.chkAutoUpdate
    onLoadLand = LangFilesFillList(Me.cmdLang)
    Me.cmdLang.ListIndex = onLoadLand
    Me.cmdAutoJ.ListIndex = prgSettings.defAction
    Me.txtDefPort = prgSettings.txtPort
    Me.chkSaveEmul.Value = prgSettings.SaveList
    Me.chkMinimize = prgSettings.chkMinimize
    Me.chkAutoR = getAutoRun
    '2
    Me.chkUseProxy.Value = prgSettings.chkUseProxy
    Me.txtProxyIp = prgSettings.txtProxyIp
    Me.txtProxyPort = prgSettings.txtProxyPort
    Me.chkRetracker = prgSettings.chkRetracker
    '3
    Me.chkIgnorServerError.Value = prgSettings.chkIgnorServerError
    Me.chkIgnorSocketErr.Value = prgSettings.chkIgnorSocketError
    Me.txtConnectTries.Text = prgSettings.txtConnectTries
    Me.chkFroze.Value = IIf(prgSettings.chkFroze, vbChecked, vbUnchecked)
    Me.chkSameHash.Value = IIf(prgSettings.chkSameHash, vbChecked, vbUnchecked)
    '5
    Me.chkRemote.Value = prgSettings.chkRemote
    Me.txtRCPass = prgSettings.txtRCPass
    Me.txtRCPort = prgSettings.txtRCPort
    
    TranslateForm Me
    cmdAutoJ.List(0) = getCStr(116, "Downloading")
    cmdAutoJ.List(1) = getCStr(117, "Emulation")
    cmdAutoJ.List(2) = getCStr(118, "Ask me")
    
    clnt.FillList lstClient
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call DelComboFlat(cmdLang.hwnd)
    Call DelComboFlat(cmdAutoJ.hwnd)
End Sub

Private Sub ts1_Click()
    Dim i As Integer
    
    For i = 1 To Me.Frame.UBound
        Me.Frame(i).Visible = False
    Next
    Frame(ts1.SelectedItem.Index).Visible = True
End Sub

Private Sub txtConnectTries_GotFocus()
    txtConnectTries.SelStart = 0
    txtConnectTries.SelLength = Len(txtConnectTries)
End Sub

Private Sub txtDefPort_GotFocus()
    txtDefPort.SelStart = 0
    txtDefPort.SelLength = Len(txtDefPort)
End Sub

Private Sub txtProxyIp_GotFocus()
    txtProxyIp.SelStart = 0
    txtProxyIp.SelLength = Len(txtProxyIp)
End Sub

Private Sub txtProxyPort_GotFocus()
    txtProxyPort.SelStart = 0
    txtProxyPort.SelLength = Len(txtProxyPort)
End Sub

Private Sub txtProxyPort_LostFocus()
    txtProxyPort.Text = val(txtProxyPort.Text)
End Sub

Private Sub SaveAutorun()
    If chkAutoR.Value = 1 Then
        WriteKey HKEY_CURRENT_USER, AUTORUN, "BitProxy", CPath & App.EXEName & ".exe"
    Else
        DeleteValue HKEY_CURRENT_USER, AUTORUN, "BitProxy"
    End If
End Sub
