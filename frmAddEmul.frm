VERSION 5.00
Begin VB.Form frmAddEmul 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add file to emulation"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11160
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "~57"
   Begin VB.Frame Frame4 
      Caption         =   "Data change"
      Enabled         =   0   'False
      Height          =   1575
      Left            =   5640
      TabIndex        =   46
      Tag             =   "~74"
      Top             =   2760
      Width           =   5415
      Begin VB.ComboBox cmbAddK2 
         Height          =   315
         ItemData        =   "frmAddEmul.frx":0000
         Left            =   3720
         List            =   "frmAddEmul.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtDownAdd 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         TabIndex        =   23
         Text            =   "0"
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cmbAddK 
         Height          =   315
         ItemData        =   "frmAddEmul.frx":0024
         Left            =   3720
         List            =   "frmAddEmul.frx":0034
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   480
         Width           =   855
      End
      Begin VB.CheckBox chkUpNow 
         Caption         =   "Update tracker statistic now"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Tag             =   "~78"
         Top             =   1200
         Width           =   3855
      End
      Begin VB.TextBox txtUpAdd 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         MaxLength       =   9
         TabIndex        =   21
         Text            =   "0"
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Add to download:"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Tag             =   "~143"
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "This option is enabled only while emulation is active"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Tag             =   "~75"
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Add to upload:"
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Tag             =   "~76"
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   9120
      TabIndex        =   27
      Tag             =   "~80"
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   6840
      TabIndex        =   26
      Tag             =   "~79"
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Caption         =   "Speed"
      Height          =   2655
      Left            =   5640
      TabIndex        =   36
      Tag             =   "~72"
      Top             =   0
      Width           =   5415
      Begin VB.TextBox txtSMUp 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3360
         MaxLength       =   4
         TabIndex        =   20
         Top             =   2250
         Width           =   855
      End
      Begin VB.CheckBox chkStepModeU 
         Caption         =   "Increase upload speed gradually"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Tag             =   "~158"
         Top             =   1920
         Width           =   5055
      End
      Begin VB.TextBox txtSMDown 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3360
         MaxLength       =   4
         TabIndex        =   18
         Top             =   1530
         Width           =   855
      End
      Begin VB.CheckBox chkStepModeD 
         Caption         =   "Increase download speed gradually"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Tag             =   "~155"
         Top             =   1200
         Width           =   5055
      End
      Begin VB.TextBox txtUp2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3960
         MaxLength       =   9
         TabIndex        =   16
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtUp1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         MaxLength       =   9
         TabIndex        =   15
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtDw2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3960
         MaxLength       =   9
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtDw1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         MaxLength       =   9
         TabIndex        =   13
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label24 
         Caption         =   "minutes"
         Height          =   255
         Left            =   4320
         TabIndex        =   56
         Tag             =   "~157"
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "during:"
         Height          =   255
         Left            =   1920
         TabIndex        =   55
         Tag             =   "~156"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label22 
         Caption         =   "minutes"
         Height          =   255
         Left            =   4320
         TabIndex        =   54
         Tag             =   "~157"
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "during:"
         Height          =   255
         Left            =   1920
         TabIndex        =   53
         Tag             =   "~156"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Kb/s"
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   44
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Kb/s"
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   43
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "to:"
         Height          =   255
         Left            =   3480
         TabIndex        =   42
         Tag             =   "~73"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "Kb/s"
         Height          =   255
         Left            =   3000
         TabIndex        =   41
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Uploading from:"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Tag             =   "~72"
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "to:"
         Height          =   255
         Left            =   3480
         TabIndex        =   39
         Tag             =   "~73"
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Kb/s"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   38
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Downloading from:"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Tag             =   "~71"
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Connection"
      Height          =   2535
      Left            =   120
      TabIndex        =   31
      Tag             =   "~63"
      Top             =   1800
      Width           =   5415
      Begin VB.CheckBox chkScrape 
         Caption         =   "Use Scrape to recieve seed and peers number (some trackers do not support this function)"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Tag             =   "~139"
         Top             =   1920
         Width           =   5055
      End
      Begin VB.TextBox txtHave 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   11
         Text            =   "100"
         Top             =   1530
         Width           =   375
      End
      Begin VB.TextBox txtIgnorT 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3600
         MaxLength       =   3
         TabIndex        =   10
         Text            =   "10"
         Top             =   1110
         Width           =   615
      End
      Begin VB.CheckBox chkIgnorT 
         Caption         =   "Ignor tracker update time and update every:"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Tag             =   "~68"
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox txtKey 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   4440
         TabIndex        =   8
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtPeerID 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   720
         Width           =   2535
      End
      Begin VB.ComboBox cmbClient 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   330
         Width           =   2535
      End
      Begin VB.Label Label20 
         Caption         =   "%"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   51
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Already have part of file:"
         Height          =   255
         Left            =   360
         TabIndex        =   50
         Tag             =   "~120"
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label18 
         Caption         =   "minutes"
         Height          =   255
         Left            =   4320
         TabIndex        =   49
         Tag             =   "~69"
         Top             =   1150
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Key:"
         Height          =   255
         Left            =   3720
         TabIndex        =   35
         Tag             =   "~66"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Port:"
         Height          =   255
         Left            =   3720
         TabIndex        =   34
         Tag             =   "~67"
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "PeerID:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Tag             =   "~65"
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Client:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Tag             =   "~64"
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "General information"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Tag             =   "~58"
      Top             =   0
      Width           =   5415
      Begin VB.TextBox txtFullsize 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtHash 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox txtTracker 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Size:"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Tag             =   "~62"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Hash:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Tag             =   "~61"
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Tag             =   "~60"
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tracker:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Tag             =   "~59"
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmAddEmul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim curItemId As Long

Private Sub chkStepModeD_Click()
    If (chkStepModeD.Value = 1) And (Len(txtSMDown.Text) = 0 Or txtSMDown.Text = "0") Then _
        txtSMDown.Text = 15
End Sub

Private Sub chkStepModeU_Click()
    If (chkStepModeU.Value = 1) And (Len(txtSMUp.Text) = 0 Or txtSMUp.Text = "0") Then _
        txtSMUp.Text = 30
End Sub

Private Sub cmbClient_Click()
    clnt.LoadClient cmbClient.Text
    txtPeerID = clnt.GeneratePeerID
    txtKey = clnt.GenerateKey
End Sub

Public Sub Command1_Click()
    Dim txt_have As Integer
    Dim coef As Currency
    On Error Resume Next
    
    If val(txtDw2) < val(txtDw1) Then
        MsgBox getCStr(81, "Wrong download speed!"), vbCritical, App.Title
        Exit Sub
    End If
    If val(txtUp2) < val(txtUp1) Then
        MsgBox getCStr(82, "Wrong upload speed!"), vbCritical, App.Title
        Exit Sub
    End If
    If (chkStepModeD.Value = 1) And val(txtSMDown.Text) < 1 Then
        MsgBox getCStr(159, "Increase time can not be less than 1 minute"), vbCritical, App.Title
        txtSMDown.SetFocus
        Exit Sub
    End If
    If (chkStepModeU.Value = 1) And val(txtSMUp.Text) < 1 Then
        MsgBox getCStr(159, "Increase time can not be less than 1 minute"), vbCritical, App.Title
        txtSMUp.SetFocus
        Exit Sub
    End If
    
    txt_have = val(txtHave.Text)
    If txt_have < 0 Then txtHave.Text = 0
    If txt_have > 100 Then txtHave.Text = 100
     
    Call emul.SelectJob(curItemId)
    
    prgSettings.emul_client = cmbClient.Text
    prgSettings.emul_dw1 = val(txtDw1)
    prgSettings.emul_dw2 = val(txtDw2)
    prgSettings.emul_up1 = val(txtUp1)
    prgSettings.emul_up2 = val(txtUp2)
    prgSettings.emul_port = val(txtPort)
    prgSettings.txtIgnorT = val(txtIgnorT)
    prgSettings.chkIgnorT = chkIgnorT.Value
    prgSettings.chkUseScrape = chkScrape.Value
    prgSettings.chkStepModeD = (chkStepModeD.Value = vbChecked)
    prgSettings.chkStepModeU = (chkStepModeU.Value = vbChecked)
    prgSettings.txtSMDown = val(txtSMDown.Text)
    prgSettings.txtSMUp = val(txtSMUp.Text)
    
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_EMULCLIENT, prgSettings.emul_client
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_EMULDW1, CStr(prgSettings.emul_dw1)
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_EMULDW2, CStr(prgSettings.emul_dw2)
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_EMULUP1, CStr(prgSettings.emul_up1)
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_EMULUP2, CStr(prgSettings.emul_up2)
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_EMULPORT, CStr(prgSettings.emul_port)
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_IGNORTIME, CStr(prgSettings.txtIgnorT)
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_USEIGNOR, CStr(prgSettings.chkIgnorT)
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_USESCRAPE, CStr(prgSettings.chkUseScrape)
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_STEPMODED, CStr(IIf(prgSettings.chkStepModeD, 1, 0))
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_STEPMODEU, CStr(IIf(prgSettings.chkStepModeU, 1, 0))
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_STEPMODEDVAL, CStr(prgSettings.txtSMDown)
    WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_STEPMODEUVAL, CStr(prgSettings.txtSMUp)
    
    emul.SetSpeed prgSettings.emul_up1, prgSettings.emul_up2, prgSettings.emul_dw1, prgSettings.emul_dw2
    emul.tUseIgnorTime = (prgSettings.chkIgnorT = 1)
    emul.tIgnorTime = prgSettings.txtIgnorT
    emul.tUseScrape = (prgSettings.chkUseScrape = 1)
    emul.tScrapeTime = 5
    emul.tName = Trim(txtName.Text)
    
    If (emul.tEvent = EMUL_ADD) Or (emul.tEvent = EMUL_STOP) Or (emul.tEvent = EMUL_ERROR) Then
        emul.SetHave txt_have
        emul.tPeerID = txtPeerID.Text
        emul.tClient = cmbClient.Text
        If emul.tEvent = EMUL_ADD Then _
            emul.tEvent = EMUL_STOP
        emul.tTracker = txtTracker.Text
        emul.tPort = val(txtPort.Text)
        emul.tUseStepDownload = (chkStepModeD.Value = 1)
        emul.tUseStepUpload = (chkStepModeU.Value = 1)
        emul.tStepDownloadTime = val(txtSMDown.Text)
        emul.tStepUploadTime = val(txtSMUp.Text)
        emul.tKey = txtKey.Text
        
        prgSettings.txtHave = txt_have
        WriteKey HKEY_LOCAL, MYPATH, REG_SETTINGS_EMULHAVE, CStr(prgSettings.txtHave)
    End If
    
    If (val(txtUpAdd) > 0) Or (val(txtDownAdd) > 0) Then
        If val(txtUpAdd) > 0 Then
            Select Case cmbAddK.ListIndex
                Case 0: coef = CKILO
                Case 1: coef = CMEGA
                Case 2: coef = CGIGA
                Case 3: coef = CTERA
            End Select
            emul.tUploaded = emul.tUploaded + (val(txtUpAdd) * coef)
        End If
        If val(txtDownAdd) > 0 Then
            Select Case cmbAddK2.ListIndex
                Case 0: coef = CKILO
                Case 1: coef = CMEGA
                Case 2: coef = CGIGA
                Case 3: coef = CTERA
            End Select
            emul.tDownloaded = emul.tDownloaded + (val(txtDownAdd) * coef)
        End If
        err.Clear
        If chkUpNow.Value = 1 Then emul.tUTime = 0
    End If
    
    emul.GenerateScrapeURL emul.tTracker
    emul.SaveList
    frmEmulate.RedrawCurrentJob
    
    Unload Me
End Sub

Private Sub Command2_Click()
    If emul.tEvent = EMUL_ADD Then
        Call emul.SelectJob(curItemId)
        frmEmulate.DeleteFromList emul.cID
        emul.DeleteJob emul.cID
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    curItemId = emul.cID
    DrawForm
    FillData
End Sub

Private Sub DrawForm()
    MakeThinAll Me
    Call SetComboFlat(cmbClient.hwnd)
    Call SetComboFlat(cmbAddK.hwnd)
    Call SetComboFlat(cmbAddK2.hwnd)
    TranslateForm Me
    
    cmbAddK.ListIndex = 0
    cmbAddK2.ListIndex = 0
End Sub

Private Sub FillData()
    Dim i As Long
    
    txtTracker = emul.tTracker
    txtName = emul.tName
    txtHash = emul.tHash
    txtFullsize = GetByteFormat(emul.tFullSize)
    txtKey = emul.tKey
    txtPort = emul.tPort
    txtDw1 = emul.sDw1
    txtDw2 = emul.sDw2
    txtUp1 = emul.sUp1
    txtUp2 = emul.sUp2
    txtIgnorT = emul.tIgnorTime
    txtHave = emul.GetHave
    txtSMDown = emul.tStepDownloadTime
    txtSMUp = emul.tStepUploadTime
    If emul.tUseIgnorTime Then chkIgnorT.Value = 1
    If emul.tUseScrape Then chkScrape.Value = 1
    If emul.tUseStepDownload Then chkStepModeD.Value = 1
    If emul.tUseStepUpload Then chkStepModeU.Value = 1
    
    If (emul.tEvent = EMUL_ADD) Or (emul.tEvent = EMUL_STOP) Or (emul.tEvent = EMUL_ERROR) Then
        clnt.FillList cmbClient
        If emul.tEvent = EMUL_ADD Then
            i = GetComboIndex(prgSettings.emul_client, cmbClient.hwnd)
            If i = -1 Then i = 0
            cmbClient.ListIndex = i
        Else
            cmbClient.ListIndex = GetComboIndex(emul.tClient, cmbClient.hwnd)
        End If
    Else
        cmbClient.Enabled = False
        txtHave.Locked = True
        txtHave.BackColor = vbButtonFace
        txtTracker.Locked = True
        txtTracker.BackColor = vbButtonFace
        txtPort.Locked = True
        txtPort.BackColor = vbButtonFace
        txtPeerID = emul.tPeerID
        chkStepModeD.Enabled = False
        txtSMDown.Enabled = False
        chkStepModeU.Enabled = False
        txtSMUp.Enabled = False
    End If
    
    If (emul.tEvent = EMUL_WORK) Or (emul.tEvent = EMUL_UPDATE) Then
        Label17.Visible = False
        Frame4.Enabled = True
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call DelComboFlat(cmbClient.hwnd)
    Call DelComboFlat(cmbAddK.hwnd)
    Call DelComboFlat(cmbAddK2.hwnd)
End Sub

Private Sub txtDownAdd_GotFocus()
    txtDownAdd.SelStart = 0
    txtDownAdd.SelLength = Len(txtDownAdd.Text)
End Sub

Private Sub txtDw1_GotFocus()
    txtDw1.SelStart = 0
    txtDw1.SelLength = Len(txtDw1.Text)
End Sub

Private Sub txtDw2_GotFocus()
    txtDw2.SelStart = 0
    txtDw2.SelLength = Len(txtDw2.Text)
End Sub

Private Sub txtHave_GotFocus()
    txtHave.SelStart = 0
    txtHave.SelLength = Len(txtHave.Text)
End Sub

Private Sub txtIgnorT_GotFocus()
    txtIgnorT.SelStart = 0
    txtIgnorT.SelLength = Len(txtIgnorT.Text)
End Sub

Private Sub txtIgnorT_LostFocus()
    Dim i As Long
    i = val(txtIgnorT)
    If i < 1 Then i = 1
    txtIgnorT = CStr(i)
End Sub

Private Sub txtPort_GotFocus()
    txtPort.SelStart = 0
    txtPort.SelLength = Len(txtPort.Text)
End Sub

Private Sub txtSMDown_GotFocus()
    txtSMDown.SelStart = 0
    txtSMDown.SelLength = Len(txtSMDown.Text)
End Sub

Private Sub txtSMUp_GotFocus()
    txtSMUp.SelStart = 0
    txtSMUp.SelLength = Len(txtSMUp.Text)
End Sub

Private Sub txtUp1_GotFocus()
    txtUp1.SelStart = 0
    txtUp1.SelLength = Len(txtUp1.Text)
End Sub

Private Sub txtUp2_GotFocus()
    txtUp2.SelStart = 0
    txtUp2.SelLength = Len(txtUp2.Text)
End Sub

Private Sub txtUpAdd_GotFocus()
    txtUpAdd.SelStart = 0
    txtUpAdd.SelLength = Len(txtUpAdd.Text)
End Sub
