VERSION 5.00
Begin VB.Form frmDebug1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Debug Form"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   2400
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   5895
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   2640
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "frmDebug1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public debForm As Form

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    debForm.onDebug = False
End Sub

Private Sub Timer1_Timer()
    Dim i As Integer
    
    On Error Resume Next
    
    List1.Clear
    i = debForm.Socket1.State
    If Err.Number <> 0 Then
        List1.AddItem "Debug form is unloaded!"
        Timer1.Enabled = False
    End If
    
    List1.AddItem "Socket1: " & getStateStr(i) & " (Client)"
    List1.AddItem "Socket2: " & getStateStr(debForm.Socket2.State) & " (Tracker)"
    List1.AddItem "Socket3: " & getStateStr(debForm.Socket3.State) & " (Listener)"
    List1.AddItem "Socket4: " & getStateStr(debForm.Socket4.State) & " (Rejecter)"
End Sub

Private Function getStateStr(st As Integer) As String
    Select Case st
        Case sckClosed: getStateStr = "Closed"
        Case sckClosing: getStateStr = "Closing"
        Case sckConnected: getStateStr = "Connected"
        Case sckConnecting: getStateStr = "Connecting"
        Case sckConnectionPending: getStateStr = "Connection Pending"
        Case sckError: getStateStr = "Error"
        Case sckHostResolved: getStateStr = "Host Resolved"
        Case sckListening: getStateStr = "Listening"
        Case sckOpen: getStateStr = "Open"
        Case sckResolvingHost: getStateStr = "Resolving Host"
        Case Else: getStateStr = "Unknown"
    End Select
End Function
