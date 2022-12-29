VERSION 5.00
Begin VB.Form frmStopSett 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Autostop settings"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "~131"
   Begin VB.Frame Framez 
      Height          =   1215
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   4455
      Begin VB.CommandButton cmdM3Change 
         Caption         =   "Change It"
         Height          =   255
         Left            =   2880
         TabIndex        =   19
         Tag             =   "~233"
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtM3Val 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   16
         Text            =   "0"
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox cmbM3Type 
         Height          =   315
         ItemData        =   "frmStopSett.frx":0000
         Left            =   3240
         List            =   "frmStopSett.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblM3Time 
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   18
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblM3ETitle 
         Alignment       =   1  'Right Justify
         Caption         =   "Emulation will be stopped at:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Tag             =   "~232"
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lblM3 
         Alignment       =   1  'Right Justify
         Caption         =   "Stop emulation in:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Tag             =   "~228"
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Framez 
      Height          =   1215
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   4455
      Begin VB.TextBox txtM2Val 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3240
         TabIndex        =   12
         Text            =   "2.0"
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Stop seeding when reaching ratio:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Tag             =   "~138"
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Tag             =   "~80"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Tag             =   "~79"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.ComboBox cmbMode 
      Height          =   315
      ItemData        =   "frmStopSett.frx":0027
      Left            =   120
      List            =   "frmStopSett.frx":0037
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Frame Framez 
      Height          =   1215
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   4455
      Begin VB.TextBox txtM1FVal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0"
         Top             =   840
         Width           =   2535
      End
      Begin VB.ComboBox cmbM1Type 
         Height          =   315
         ItemData        =   "frmStopSett.frx":00DE
         Left            =   3600
         List            =   "frmStopSett.frx":00EE
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtM1Val 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2400
         MaxLength       =   6
         TabIndex        =   7
         Text            =   "0"
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "This number in bytes:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Tag             =   "~137"
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Stop emulation when uploaded more than"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Tag             =   "~136"
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Framez 
      Height          =   1215
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
End
Attribute VB_Name = "frmStopSett"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim curJob As Long

Private Sub cmbM1Type_Click()
    RecalculateMode1
End Sub

Private Sub cmbMode_Click()
    Dim i As Integer
    For i = Framez.LBound To Framez.UBound
        Framez(i).Visible = (cmbMode.ListIndex = i)
    Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdM3Change_Click()
    txtM3Val.Enabled = True
    cmbM3Type.Enabled = True
    lblM3ETitle.Visible = False
    lblM3Time.Visible = False
    cmdM3Change.Visible = False
End Sub

Private Sub cmdOK_Click()
    Dim tmp As String, i As Long
    On Error Resume Next
    
    If curJob <> 0 Then
        emul.SelectJob curJob
        emul.tStopMode = cmbMode.ListIndex
        Select Case cmbMode.ListIndex
            Case 1
                emul.tStopValue = CCur(txtM1FVal.Text)
            Case 2
                tmp = txtM2Val.Text
                Call str_replace(".", ",", tmp)
                emul.tStopValue = CCur(tmp)
            Case 3
                i = val(txtM3Val.Text)
                If i < 0 Then i = 0
                Select Case cmbM3Type.ListIndex
                    Case 0: i = i * 60
                    Case 1: i = i * 3600
                    Case 2: i = i * 86400
                End Select
                emul.tStopValue = i + ToUnixTime(Now)
        End Select
        err.Clear
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    curJob = 0
    
    DrawForm
    LoadCurJob
End Sub

Private Sub DrawForm()
    cmbM1Type.ListIndex = 0
    
    MakeThinAll Me
    Call SetComboFlat(cmbMode.hwnd)
    Call SetComboFlat(cmbM1Type.hwnd)
    Call SetComboFlat(cmbM3Type.hwnd)
    
    TranslateForm Me
    cmbMode.List(0) = getCStr(133, "Disable autostop")
    cmbMode.List(1) = getCStr(134, "Stop emulation on specific upload rate")
    cmbMode.List(2) = getCStr(135, "Stop emulation on specific ratio")
    cmbMode.List(3) = getCStr(227, "Stop emulation after specific time")
    
    cmbM3Type.List(0) = getCStr(229, "Minutes")
    cmbM3Type.List(1) = getCStr(230, "Hours")
    cmbM3Type.List(2) = getCStr(231, "Days")
End Sub

Private Sub LoadCurJob()
    Dim i As Long
    On Error Resume Next
    
    i = frmEmulate.plist.SelectedItem.Index
    If err.Number <> 0 Then
        cmbMode.ListIndex = 0
        Exit Sub
    End If
    curJob = val(Mid(frmEmulate.plist.ListItems(i).key, 2))
    
    cmbM3Type.ListIndex = 0
    
    emul.SelectJob curJob
    cmbMode.ListIndex = emul.tStopMode
    Select Case emul.tStopMode
        Case 1
            txtM1FVal.Text = CStr(emul.tStopValue)
        Case 2
            txtM2Val.Text = CStr(emul.tStopValue)
        Case 3
            txtM3Val.Enabled = False
            cmbM3Type.Enabled = False
            lblM3ETitle.Visible = True
            lblM3Time.Visible = True
            cmdM3Change.Visible = True
            lblM3Time.Caption = Format(FromUnixTime(emul.tStopValue), "yyyy.mm.dd hh:mm")
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call DelComboFlat(cmbMode.hwnd)
    Call DelComboFlat(cmbM1Type.hwnd)
End Sub

Private Sub RecalculateMode1()
    Dim i As Double, coef As Currency, tmp As String
    On Error Resume Next
    
    tmp = txtM1Val.Text
    str_replace ".", ",", tmp
    i = CDbl(tmp)
    Select Case cmbM1Type.ListIndex
        Case 0: coef = CKILO
        Case 1: coef = CMEGA
        Case 2: coef = CGIGA
        Case 3: coef = CTERA
    End Select
    coef = CCur(i * coef)
    txtM1FVal.Text = Format(coef, "#0")
    err.Clear
End Sub

Private Sub txtM1Val_Change()
    RecalculateMode1
End Sub
