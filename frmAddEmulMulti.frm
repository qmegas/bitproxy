VERSION 5.00
Begin VB.Form frmAddEmulMulti 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Multiple change"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "~217"
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Tag             =   "~79"
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Tag             =   "~80"
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Caption         =   "Speed"
      Height          =   1215
      Left            =   0
      TabIndex        =   15
      Tag             =   "~72"
      Top             =   0
      Width           =   5415
      Begin VB.TextBox txtDw1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         MaxLength       =   9
         TabIndex        =   1
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtDw2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3960
         MaxLength       =   9
         TabIndex        =   2
         Text            =   "0"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtUp1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         MaxLength       =   9
         TabIndex        =   3
         Text            =   "0"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtUp2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3960
         MaxLength       =   9
         TabIndex        =   4
         Text            =   "0"
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Downloading from:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Tag             =   "~71"
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Kb/s"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   22
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "to:"
         Height          =   255
         Left            =   3480
         TabIndex        =   21
         Tag             =   "~73"
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Uploading from:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Tag             =   "~72"
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "Kb/s"
         Height          =   255
         Left            =   3000
         TabIndex        =   19
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "to:"
         Height          =   255
         Left            =   3480
         TabIndex        =   18
         Tag             =   "~73"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Kb/s"
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   17
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Kb/s"
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   16
         Top             =   720
         Width           =   375
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Data change"
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Tag             =   "~74"
      Top             =   1320
      Width           =   5415
      Begin VB.TextBox txtUpAdd 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         MaxLength       =   9
         TabIndex        =   5
         Text            =   "0"
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox chkUpNow 
         Caption         =   "Update tracker statistic now"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Tag             =   "~78"
         Top             =   1200
         Width           =   3855
      End
      Begin VB.ComboBox cmbAddK 
         Height          =   315
         ItemData        =   "frmAddEmulMulti.frx":0000
         Left            =   3720
         List            =   "frmAddEmulMulti.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtDownAdd 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         TabIndex        =   7
         Text            =   "0"
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cmbAddK2 
         Height          =   315
         ItemData        =   "frmAddEmulMulti.frx":0024
         Left            =   3720
         List            =   "frmAddEmulMulti.frx":0034
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Add to upload:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Tag             =   "~76"
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "This option is enabled only while emulation is active"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Tag             =   "~75"
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Add to download:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Tag             =   "~143"
         Top             =   840
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmAddEmulMulti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim i As Long, iid As Long
    Dim coef As Currency
    
    If val(txtDw2) < val(txtDw1) Then
        MsgBox getCStr(81, "Wrong download speed!"), vbCritical, App.Title
        Exit Sub
    End If
    If val(txtUp2) < val(txtUp1) Then
        MsgBox getCStr(82, "Wrong upload speed!"), vbCritical, App.Title
        Exit Sub
    End If
    
    With frmEmulate.plist.ListItems
        For i = 1 To .Count
            If .Item(i).Selected Then
                iid = frmEmulate.GetListItemID(.Item(i).Index)
                emul.SelectJob iid
                
                emul.SetSpeed val(txtUp1), val(txtUp2), val(txtDw1), val(txtDw2)
                
                If (emul.tEvent = EMUL_WORK) Or (emul.tEvent = EMUL_UPDATE) Then
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
                    If chkUpNow.value = 1 Then emul.tUTime = 0
                End If
            End If
        Next
    End With
    
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    DrawForm
    FillData
End Sub

Private Sub DrawForm()
    MakeThinAll Me
    Call SetComboFlat(cmbAddK.hwnd)
    Call SetComboFlat(cmbAddK2.hwnd)
    TranslateForm Me
End Sub

Private Sub FillData()
    Dim com_dl1 As Long
    Dim com_dl2 As Long
    Dim com_upl1 As Long
    Dim com_upl2 As Long
    Dim i As Long, iid As Long
    
    With frmEmulate.plist.ListItems
        For i = 1 To .Count
            If .Item(i).Selected Then
                iid = frmEmulate.GetListItemID(.Item(i).Index)
                emul.SelectJob iid
                If com_dl1 <> -1 Then
                    If com_dl1 = 0 Then
                        com_dl1 = emul.sDw1
                    Else
                        If com_dl1 <> emul.sDw1 Then com_dl1 = -1
                    End If
                End If
                If com_dl2 <> -1 Then
                    If com_dl2 = 0 Then
                        com_dl2 = emul.sDw2
                    Else
                        If com_dl2 <> emul.sDw2 Then com_dl2 = -1
                    End If
                End If
                If com_upl1 <> -1 Then
                    If com_upl1 = 0 Then
                        com_upl1 = emul.sUp1
                    Else
                        If com_upl1 <> emul.sUp1 Then com_upl1 = -1
                    End If
                End If
                If com_upl2 <> -1 Then
                    If com_upl2 = 0 Then
                        com_upl2 = emul.sUp2
                    Else
                        If com_upl2 <> emul.sUp2 Then com_upl2 = -1
                    End If
                End If
            End If
        Next
    End With
    
    If com_dl1 <> -1 Then txtDw1.Text = com_dl1
    If com_dl2 <> -1 Then txtDw2.Text = com_dl2
    If com_upl1 <> -1 Then txtUp1.Text = com_upl1
    If com_upl2 <> -1 Then txtUp2.Text = com_upl2
    
    cmbAddK.ListIndex = 0
    cmbAddK2.ListIndex = 0
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

