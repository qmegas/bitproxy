Attribute VB_Name = "modTorrentWork"
Option Explicit

Public Type D_OUTDATA
    tmpFile As String
    torName As String
    realURL As String
End Type

Private tTempName As String
Private tNewAnons As String
Private tFullNewAnons As String
Private prevStr As String
Private lvlCnt As Integer
Private realURL As String
Private torName As String
Private torEncode As String

Public Function RunTorrentWork(FileN As String, port As Long, dt As D_OUTDATA) As Boolean
    Dim tmp As String
    On Error Resume Next
    
    RunTorrentWork = False
    
    If Len(Dir(FileN)) > 0 Then
        realURL = vbNullString
        torName = vbNullString
        tFullNewAnons = vbNullString
        dt.tmpFile = TempPath & "~BP" & CStr(Timer) & ".torrent"
        tNewAnons = "127.0.0.1:" & CStr(port)
        
        Open FileN For Binary As 10
        Open dt.tmpFile For Binary As 11
        If err.Number > 0 Then
            MsgBox err.Description, vbCritical, App.Title
            Exit Function
        End If
        
        TorrentList
        
        Close 10
        Close 11
        
        If torEncode = "utf-8" Then
            tmp = torName
            torName = Left(torName, MultiByteToWideChar(CP_UTF8, 0, tmp, -1, StrPtr(torName), LenB(torName)))
        End If
        
        dt.realURL = realURL
        dt.torName = torName
        RunTorrentWork = True
    End If
End Function

Private Sub TorrentList()
    Dim b As Byte

    lvlCnt = 0
    Do
        Get 10, , b
        Select Case b
            Case 100 'd
                'Dictionary
                Put 11, , b
            Case 101 'e
                'End List
                lvlCnt = lvlCnt - 1
                If lvlCnt = 0 Then prevStr = vbNullString
                Put 11, , b
            Case 105 'i
                GetTorInteger
            Case 108 'L
                'List
                lvlCnt = lvlCnt + 1
                Put 11, , b
            Case Else
                GetTorString
        End Select
    Loop Until Seek(10) > LOF(10)
End Sub

Private Sub GetTorString()
    Dim b As Byte, strSize As Long, i As Long, i2 As Long
    Dim startPoz As Long
    Dim tmp As String
    Dim cur_domain As String
    Dim Buff() As Byte
    
    Const ANONS1 = "announce-list"
    Const ANONS2 = "announce"
    Const RETRACKER = "retracker"
    
    startPoz = Seek(10) - 1
    Seek 10, startPoz
    tmp = vbNullString
    Do
        Get 10, , b
        If b <> 58 Then tmp = tmp & Chr(b)
        If EOF(10) Then Exit Sub
    Loop Until b = 58
    strSize = val(tmp)
    tmp = vbNullString
    If strSize > 0 Then
        ReDim Buff(1 To strSize) As Byte
        Get 10, , Buff
        tmp = Space(strSize)
        CopyMemory ByVal tmp, Buff(1), strSize
    End If
    
    If prevStr = "name" Then torName = tmp
    
    If (prevStr = ANONS2) Or (prevStr = ANONS1) Then
        i = 0: i2 = 0
        
        i = InStr(1, tmp, "//")
        If i > 0 Then i2 = InStr(i + 2, tmp, "/")
        
        If (i > 0) And (i2 > 0) Then
            cur_domain = Mid(tmp, i + 2, i2 - i - 2)
            If prevStr = ANONS2 Then realURL = cur_domain
            
            If (prgSettings.chkRetracker = 1) And (InStr(1, LCase(cur_domain), RETRACKER) > 0) Then _
                'Is retracker. Do nothing
            Else
                tmp = Left(tmp, i + 1) & tNewAnons & Mid(tmp, i2)
                If prevStr = ANONS1 Then
                    If Len(realURL) = 0 Then realURL = cur_domain 'Torrent file doesn't have announce
                    If Len(tFullNewAnons) > 0 Then tmp = tFullNewAnons _
                        Else tFullNewAnons = tmp
                End If
                strSize = Len(tmp)
                ReDim Buff(1 To strSize) As Byte
                For i = 1 To strSize
                    Buff(i) = Asc(Mid(tmp, i, 1))
                Next
            End If
        Else
            'Some URL error
        End If
    End If
    
    If prevStr = "encoding" Then
        torEncode = LCase(tmp)
    End If
    
    If prevStr <> ANONS1 Then prevStr = tmp
    
    Put 11, , CStr(strSize) & ":"
    Put 11, , Buff
End Sub

Private Sub GetTorInteger()
    Dim b As Byte, tmp As String
    
    Do
        Get 10, , b
        tmp = tmp & Chr(b)
    Loop Until b = 101
    Put 11, , "i" & tmp
End Sub

Public Function GetTorStringBuffer(Buff As String, poz As Long) As String
    Dim i As Long, num As Long
    
    i = InStr(poz, Buff, ":")
    If i = 0 Then
        GetTorStringBuffer = vbNullString
        Exit Function
    End If
    num = val(Mid(Buff, poz, i - poz))
    If num = 0 Then
        GetTorStringBuffer = vbNullString
        Exit Function
    End If
    GetTorStringBuffer = Mid(Buff, i + 1, num)
End Function

Public Function GetTorLongBuffer(Buff As String, poz As Long) As Long
    Dim i As Long
    
    If Mid(Buff, poz, 1) = "i" Then
        i = InStr(poz + 1, Buff, "e")
        If i > 0 Then GetTorLongBuffer = val(Mid(Buff, poz + 1, i - poz - 1)) _
            Else GetTorLongBuffer = 0
    Else
        GetTorLongBuffer = 0
    End If
End Function
