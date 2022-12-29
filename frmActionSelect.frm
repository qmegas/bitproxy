VERSION 5.00
Begin VB.Form frmActionSelect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select action..."
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "~54"
   Begin VB.CommandButton Command2 
      Caption         =   "Start Emulation"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Tag             =   "~56"
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Downloading"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Tag             =   "~55"
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmActionSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Me.Visible = False
    frmMain.Show
    frmMain.ProcceedFilePool
    Unload Me
End Sub

Private Sub Command2_Click()
    Me.Visible = False
    frmEmulate.Show
    frmEmulate.ProcceedFilePool
    Unload Me
End Sub

Private Sub Form_Load()
    TranslateForm Me
End Sub
