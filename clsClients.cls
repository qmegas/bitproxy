VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsClientManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mycoll As New Collection
Private curClient As clsClientItem
Private LoadErrors As Long

Private Const CLIENT_HEADER = "BP14"
Private Const PART_GENERAL = "General"
Private Const PART_HEADERS = "Headers"
Private Const PART_HASH = "Peerid"
Private Const PART_KEY = "Key"

Public Function getClientNameById(cID As Long) As String
    If (cID > mycoll.Count) Or (cID < 1) Then
        getClientNameById = vbNullString
    Else
        getClientNameById = mycoll(cID).Title
    End If
End Function

Public Function getCount() As Long
    getCount = mycoll.Count
End Function

Public Sub FillList(cmbHwnd As Object)
    Dim i As clsClientItem
    
    For Each i In mycoll
        cmbHwnd.AddItem i.Title
    Next
    
    Set i = Nothing
End Sub

Public Function getIdByName(cName As String) As Long
    Dim i As Long, cl As clsClientItem
    
    getIdByName = 0
    For i = 1 To mycoll.Count
        Set cl = mycoll(i)
        If cl.Title = cName Then
            getIdByName = i
            Exit For
        End If
    Next
    
    Set cl = Nothing
End Function

Public Function BuildHTTP(tTracker As String, tHash As String, tPeerID As String, _
    tPort As String, tUploaded As String, tDownload As String, tLeft As String, _
    tKey As String, tEvent As String, tHost As String) As String
    Dim tmp As String, tmp2 As String
    
    tmp = "GET " & tTracker
    tmp2 = curClient.URL
    str_replace "{thash}", MakeHash(Trim(tHash)), tmp2
    str_replace "{peerid}", tPeerID, tmp2
    str_replace "{port}", tPort, tmp2
    str_replace "{uploaded}", tUploaded, tmp2
    str_replace "{downloaded}", tDownload, tmp2
    str_replace "{left}", tLeft, tmp2
    str_replace "{key}", tKey, tmp2
    If Len(tEvent) = 0 Then str_replace "{event}", vbNullString, tmp2 _
        Else str_replace "{event}", "&event=" & tEvent, tmp2
    tmp = tmp & tmp2 & " " & curClient.Protocol & vbCrLf
    
    tmp2 = curClient.Headers
    str_replace "{host}", tHost, tmp2
    str_replace "{uagent}", curClient.UserAgent, tmp2
    str_replace "{\n}", vbCrLf, tmp2
    tmp = tmp & tmp2 & vbCrLf & vbCrLf
    
    BuildHTTP = tmp
End Function

Public Function BuildScrapeHTTP(mUrl As String, tHost As String) As String
    Dim tmp As String, tmp2 As String
    
    tmp = "GET " & mUrl
    tmp = tmp & " " & curClient.Protocol & vbCrLf
    
    tmp2 = curClient.Headers
    str_replace "{host}", tHost, tmp2
    str_replace "{\n}", vbCrLf, tmp2
    tmp = tmp & tmp2 & vbCrLf & vbCrLf
    
    BuildScrapeHTTP = tmp
End Function

Public Function GenerateKey() As String
    Dim i As Integer, tmp As String, b As String * 1
    Randomize Timer
    
    tmp = vbNullString
    For i = 1 To curClient.KeySize
        Select Case curClient.KeyType
            Case "Num"
                b = CStr(Int(Rnd() * 10))
            Case Else
                b = LCase(Hex(Int(Rnd * 16)))
        End Select
        
        If curClient.KeyUpper Then b = UCase(b)
        
        tmp = tmp & b
    Next
    GenerateKey = tmp
End Function

Public Function GeneratePeerID() As String
    Dim i As Long, k As Integer, b As Byte, sym As String, j As Integer
    Const PEERIDSIZE = 20
    
    Randomize Timer
    
    GeneratePeerID = curClient.Prefix
    i = PEERIDSIZE - curClient.PrifixSize
    For k = i To 1 Step -1
        j = Int(Rnd * 3)
        Select Case j
            Case 0 'numbers
                 b = Int(Rnd * 10) + 48
            Case 1 'Big sym
                b = Int(Rnd * 26) + 65
            Case 2 'Small sym
                b = Int(Rnd * 26) + 97
        End Select
        
        If curClient.HashEncoding Then
            sym = "%" & LCase(Hex(b))
        Else
            sym = Chr(b)
        End If
        
        If curClient.HashUpper Then sym = UCase(sym)
        
        GeneratePeerID = GeneratePeerID & sym
    Next
End Function

Public Function LoadClientById(cID As Long) As Boolean
    If cID > 0 And cID <= mycoll.Count Then
        Set curClient = mycoll(cID)
        LoadClientById = True
    Else
        LoadClientById = False
    End If
End Function

Public Function LoadClient(cTitle As String) As Boolean
    For Each curClient In mycoll
        If curClient.Title = cTitle Then
            LoadClient = True
            Exit Function
        End If
    Next
    LoadClient = False
End Function

Public Function AddClient(FileN As String) As Boolean
    Dim i As New clsClientItem
    
    If Not FileExist(FileN) Then
        AddClient = False
        Exit Function
    End If
    
    LoadErrors = 0
    If getINIString("general", "ID", FileN) <> CLIENT_HEADER Then
        AddClient = False
        Exit Function
    End If
    
    'General
    i.Author = TraceIt(getINIString(PART_GENERAL, "Author", FileN))
    i.Title = TraceIt(getINIString(PART_GENERAL, "Title", FileN))
    'Headers
    i.URL = TraceIt(getINIString(PART_HEADERS, "Urladd", FileN))
    i.Headers = TraceIt(getINIString(PART_HEADERS, "Headers", FileN))
    i.Protocol = TraceIt(getINIString(PART_HEADERS, "Protocol", FileN))
    i.UserAgent = TraceIt(getINIString(PART_HEADERS, "UserAgent", FileN))
    'Hash
    i.Prefix = TraceIt(getINIString(PART_HASH, "Prefix", FileN))
    i.PrifixSize = val(TraceIt(getINIString(PART_HASH, "PrefixSize", FileN)))
    i.HashEncoding = (val(getINIString(PART_HASH, "Encoding", FileN)) <> 0)
    i.HashUpper = (val(getINIString(PART_HASH, "PrefixSize", FileN)) <> 0)
    'Key
    i.KeySize = val(TraceIt(getINIString(PART_KEY, "Size", FileN)))
    i.KeyType = TraceIt(getINIString(PART_KEY, "Type", FileN))
    i.KeyUpper = (val(getINIString(PART_KEY, "Uppercase", FileN)) <> 0)
    
    If LoadErrors = 0 Then
        mycoll.Add i
        Set i = Nothing
        AddClient = True
    Else
        AddClient = False
    End If
End Function

Private Function TraceIt(mStr As String) As String
    If Len(mStr) = 0 Then LoadErrors = LoadErrors + 1
    TraceIt = mStr
End Function

Public Property Get Title() As String
    Title = curClient.Title
End Property

Public Property Get Author() As String
    Author = curClient.Author
End Property

Public Property Get Prefix() As String
    Prefix = curClient.Prefix
End Property

Public Property Get PrefixSize() As Integer
    PrefixSize = curClient.PrifixSize
End Property

Public Property Get UserAgent() As String
    UserAgent = curClient.UserAgent
End Property
