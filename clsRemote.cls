VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRemote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const IS_HTTP = "HTTP"
Private Const IS_POST = "POST "
Private Const SCRIPT_FILE = "remote.htm"
Private Const BTP_KEY = "btp/"

Private Const ERR_PROTOCOL = 1
Private Const ERR_PASSWORD = 2
Private Const ERR_UNKNOWN = 3
Private Const ERR_CUSTOM = 999

Private start_time As Date
Private cur_ws As Winsock

Public Sub ShowSocketError(err As String)
    Dim tmp As String
    tmp = getCStr(214, "Remote Control")
    MainForm.stb.Panels(1).Text = tmp & ": " & err
    PopupMsgTray MainForm, err, tmp, NIIF_ERROR
End Sub

Public Sub StartRC()
    On Error Resume Next
    
    cur_ws.LocalPort = prgSettings.txtRCPort
    cur_ws.Listen
    If err.Number <> 0 Then
        ShowSocketError err.Description
        cur_ws.Close
        err.Clear
    Else
        MainForm.stb.Panels(2).Enabled = True
    End If
End Sub

Public Sub StopRC()
    cur_ws.Close
    MainForm.stb.Panels(2).Enabled = False
End Sub

Public Sub set_socket(ws As Winsock)
    Set cur_ws = ws
End Sub

Public Sub got_data(data As String)
    Dim i As Long, j As Long
    Dim tmp As String
    Dim cur_pass As String
    Dim cur_act As String
    
    i = InStr(1, data, vbCrLf)
    If i = 0 Then
        InstantClose
        Exit Sub
    End If
    
    data = Left(data, i - 1)
    
    If Left(data, 5) = IS_POST Then
        j = InStr(1, data, IS_HTTP)
        If j = 0 Then j = Len(data)
        data = Mid(data, 7, j - 8)
        
        If Left(data, Len(BTP_KEY)) <> BTP_KEY Then
            ReturnError ERR_PROTOCOL, vbNullString
            Exit Sub
        End If
        data = Mid(data, Len(BTP_KEY) + 1)
        
        i = InStr(1, data, "/")
        If i = 0 Then
            ReturnError ERR_PROTOCOL, vbNullString
            Exit Sub
        End If
        
        cur_pass = Left(data, i - 1)
        cur_act = Mid(data, i + 1)
        
        If cur_pass <> prgSettings.txtRCPass Then
            ReturnError ERR_PASSWORD, vbNullString
            Exit Sub
        End If
        
        ReturnMainAction cur_act
    Else
        ReturnFullScript
    End If
End Sub

Private Sub ReturnMainAction(act As String)
    Dim dOut As String
    
    Select Case act
        Case "login"
            dOut = "callback_logged(); " & retInfo
        Case "info"
            dOut = retInfo
        Case "emul"
            dOut = "reset_emulation(); "
            dOut = dOut & retEmulationList & " callback_emul_flash();"
        Case "down"
            dOut = "reset_download(); "
            dOut = dOut & retDownloadList & " callback_down_flash();"
        Case Else
            ReturnError ERR_UNKNOWN, vbNullString
    End Select
    cur_ws.SendData GetHeader & dOut
End Sub

Private Function retDownloadList() As String
    Dim i As Integer
    Dim tmp As String
    
    retDownloadList = vbNullString
    If GetDownloadCount = 0 Then Exit Function
    
    With frmMain.lv1.ListItems
        For i = 1 To .Count
            tmp = "callback_down_add(0,'"
            tmp = tmp & .Item(i).Text & "','"
            tmp = tmp & .Item(i).SubItems(1) & "','"
            tmp = tmp & .Item(i).SubItems(2) & "','"
            tmp = tmp & .Item(i).SubItems(3) & "','"
            tmp = tmp & .Item(i).SmallIcon & "'); "
            retDownloadList = retDownloadList & tmp
        Next
    End With
End Function

Private Function retEmulationList() As String
    Dim i As Integer
    Dim tmp As String
    
    retEmulationList = vbNullString
    If GetEmulationCount = 0 Then Exit Function
    
    With frmEmulate.plist.ListItems
        For i = 1 To .Count
            tmp = "callback_emul_add(0,'"
            tmp = tmp & .Item(i).Text & "','"
            tmp = tmp & .Item(i).SubItems(1) & "','"
            tmp = tmp & .Item(i).SubItems(2) & "','"
            tmp = tmp & .Item(i).SubItems(4) & "','"
            tmp = tmp & .Item(i).SubItems(5) & "','"
            tmp = tmp & .Item(i).SubItems(6) & "','"
            tmp = tmp & .Item(i).SubItems(7) & "','"
            tmp = tmp & .Item(i).SubItems(8) & "'); "
            retEmulationList = retEmulationList & tmp
        Next
    End With
End Function

Private Function retInfo() As String
    retInfo = "callback_info('"
    retInfo = retInfo & CStr(App.Major) & "." & CStr(App.Minor) & "','"
    retInfo = retInfo & CDate(Now - start_time) & "'," & GetDownloadCount & ","
    retInfo = retInfo & GetEmulationCount & ");"
End Function

Private Sub ReturnFullScript()
    Dim tmp As String, tline As String
    Dim j As Long
    
    If Not FileExist(CPath & SCRIPT_FILE) Then
        tmp = getCStr(130, "Remote control script not found")
    Else
        j = FreeFile
        Open CPath & SCRIPT_FILE For Input As #j
        While Not EOF(j)
            Line Input #j, tline
            tmp = tmp & tline & vbCrLf
        Wend
        Close j
    End If
    
    tmp = GetHeader & tmp
    cur_ws.SendData tmp
End Sub

Private Sub ReturnError(eer_code As Integer, info As String)
    Dim tmp As String
    tmp = GetHeader & "callback_error(" & CStr(eer_code) & ",'" & info & "');"
    cur_ws.SendData tmp
End Sub

Private Sub InstantClose()
    cur_ws.Close
    cur_ws.Listen
End Sub

Private Function GetHeader() As String
    Dim tmp As String, tmp2 As String
    
    tmp = "HTTP/1.1 200 OK" & vbCrLf
    tmp = tmp & "Date: " & Format(Now, "ddd, d mmm yyyy hh:mm:ss") & " GMT" & vbCrLf
    tmp = tmp & "Server: BTP Server 1.1" & vbCrLf
    tmp = tmp & "Connection: close" & vbCrLf
    tmp = tmp & "Content-Type: text/html"
    tmp2 = getCStr(215, vbNullString)
    If Len(tmp2) > 0 Then tmp = tmp & "; " & tmp2
    tmp = tmp & vbCrLf & vbCrLf
    
    GetHeader = tmp
End Function

Private Function GetEmulationCount() As Long
    If Not IsFormLoaded("frmEmulate") Then
        GetEmulationCount = 0
    Else
        GetEmulationCount = frmEmulate.plist.ListItems.Count
    End If
End Function

Private Function GetDownloadCount() As Long
    If Not IsFormLoaded("frmMain") Then
        GetDownloadCount = 0
    Else
        GetDownloadCount = frmMain.lv1.ListItems.Count
    End If
End Function

Private Sub Class_Initialize()
    start_time = Now
End Sub
