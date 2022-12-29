VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "~52"
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Tag             =   "~152"
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label18 
      Caption         =   "http://qmegas.info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   960
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Copyright by Megas © 2006-2011"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Version"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Tag             =   "~53"
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "BitTorrent Proxy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    DrawForm
End Sub

Private Sub Label18_Click()
    OpenHomepage
End Sub

Private Sub DrawForm()
    TranslateForm Me
    Me.Label2.Caption = Me.Label2.Caption & " " & CStr(App.Major) & "." & CStr(App.Minor)
End Sub
