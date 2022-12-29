VERSION 5.00
Begin VB.Form frmSerial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BitTorrent Proxy"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4140
   Icon            =   "frmSerial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   4140
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSer 
      Height          =   285
      Left            =   120
      MaxLength       =   16
      TabIndex        =   2
      Top             =   2280
      Width           =   2655
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   280
      Left            =   2880
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Please, enter your serial key, and then press OK."
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Tag             =   "~203"
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   $"frmSerial.frx":365A
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Tag             =   "~50"
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmSerial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    This_SID = Trim(txtSer.Text)
    If CheckValidate Then Unload Me
End Sub

Private Sub Form_Load()
    TranslateForm Me
    Me.Show
    txtSer.SetFocus
End Sub
