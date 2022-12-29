VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTorrentObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'SHA1
Private Declare Function CreateAlgorithm Lib "sha1.dll" () As Long
Private Declare Function DestroyAlgorithm Lib "sha1.dll" (ByVal hAlgorithm As Long) As Long
Private Declare Function InitAlgorithm Lib "sha1.dll" (ByVal hAlgorithm As Long) As Long
Private Declare Function InputData Lib "sha1.dll" (ByVal hAlgorithm As Long, ByVal data As String, ByVal leng As Long) As Long
Private Declare Function GetHashLength Lib "sha1.dll" (ByVal hAlgorithm As Long) As Long
Private Declare Function GetHash Lib "sha1.dll" (ByVal hAlgorithm As Long, ByVal buffer As String, ByVal bufsize As Long) As Long
Private Declare Function GetHashTextLength Lib "sha1.dll" (ByVal hAlgorithm As Long) As Long
Private Declare Function GetHashText Lib "sha1.dll" (ByVal hAlgorithm As Long, ByVal buffer As String, ByVal bufsize As Long) As Long

Private mvartTrackerURL As String
Private mvartHash As String * 40
Private mvartName As String
Private mvartLength As Currency
Private mvartEncode As String


Private tabC As Integer
Private infoStart As Long
Private infoEnd As Long
Private infoLevel As Long
Private infoLocked As Boolean

Public Property Get tLength() As Currency
    tLength = mvartLength
End Property

Public Property Get tHash() As String
    tHash = mvartHash
End Property

Public Property Get tTrackerURL() As String
    tTrackerURL = mvartTrackerURL
End Property

Public Property Get tName() As String
    tName = mvartName
End Property

Public Function LoadTorrent(FileN As String) As Boolean
    Dim ff As Integer, Buff As String
    
    If FileExist(FileN) Then
        ff = FreeFile
        Open FileN For Binary As #ff
        tabC = 0
        mvartLength = 0
        LoadTorrent = StartTorrentWork(ff)
        If infoLocked Then
            Seek ff, infoStart
            Buff = Input(infoEnd - infoStart, ff)
            mvartHash = SHA1f(Buff)
        Else
            MsgBox "INFO box not found!", vbCritical, App.Title
            LoadTorrent = False
        End If
        
        Close #ff
    Else
        LoadTorrent = False
    End If
End Function

Private Function SHA1f(intxt As String) As String
    Dim hAlgorithm As Long, hashsize As Long, buf As String
    
    hAlgorithm = CreateAlgorithm()
    If hAlgorithm > 0 Then
        If InitAlgorithm(hAlgorithm) = 0 Then
            If InputData(hAlgorithm, intxt, Len(intxt)) = 0 Then
                hashsize = GetHashTextLength(hAlgorithm)
                If hashsize > 0 Then
                    buf = Space(hashsize + 1)
                    If GetHashText(hAlgorithm, buf, Len(buf)) = 0 Then _
                        SHA1f = buf
                End If
            End If
        End If
        Call DestroyAlgorithm(hAlgorithm)
    Else
        MsgBox getCStr(51, "Error on initialisation SHA1 library"), vbCritical, App.Title
    End If
End Function

Private Function StartTorrentWork(hFile As Integer) As Boolean
    Dim b As Byte, isValue As Boolean
    Dim prevStr As String, curStr As String
    
    infoLocked = False
    isValue = False
    infoStart = 0
    Do
        Get hFile, , b
        Select Case b
            Case 100 'd
                If isValue And (prevStr = "info") Then
                    infoStart = Seek(hFile) - 1
                    infoLevel = tabC
                End If
                tabC = tabC + 1
                isValue = False
            Case 101 'e
                tabC = tabC - 1
                If Not infoLocked Then
                    If (infoStart > 0) And (infoLevel = tabC) Then
                        infoEnd = Seek(hFile)
                        infoLocked = True
                    End If
                End If
                isValue = False
            Case 105 'i
                If isValue And (prevStr = "length") Then
                    mvartLength = mvartLength + GetInteger(hFile)
                Else
                    SkipInteger hFile
                End If
                isValue = False
            Case 108 'L
                tabC = tabC + 1
                If prevStr = "announce-list" Then isValue = True _
                    Else isValue = False
            Case Else
                curStr = GetString(hFile)
                If isValue Then
                    Select Case prevStr
                        Case "name"
                            mvartName = curStr
                        Case "announce"
                            mvartTrackerURL = curStr
                        Case "announce-list"
                            If Len(mvartTrackerURL) = 0 Then _
                                mvartTrackerURL = curStr
                        Case "encoding"
                            mvartEncode = LCase(curStr)
                    End Select
                End If
                isValue = Not isValue
                prevStr = curStr
        End Select
    Loop Until (Seek(hFile) > LOF(hFile))
    
    If mvartEncode = "utf-8" Then
        curStr = mvartName
        mvartName = Left(mvartName, MultiByteToWideChar(CP_UTF8, 0, curStr, -1, StrPtr(mvartName), LenB(mvartName)))
    End If
    
    StartTorrentWork = True
End Function

Private Sub SkipInteger(hFile As Integer)
    Dim b As Byte
    Do
        Get hFile, , b
    Loop Until (b = 101) Or (Seek(hFile) > LOF(hFile))
End Sub

Private Function GetInteger(hFile As Integer) As Currency
    Dim b As Byte, tmp As String
    
    Do
        Get hFile, , b
        tmp = tmp & Chr(b)
    Loop Until b = 101 'e
    GetInteger = CCur(Left(tmp, Len(tmp) - 1))
End Function

Private Function GetString(hFile As Integer) As String
    Dim b As Byte, strSize As Long
    Dim tmp As String
    Dim Buff() As Byte
    
    Seek hFile, Seek(hFile) - 1
    tmp = vbNullString
    Do
        Get hFile, , b
        If b <> 58 Then tmp = tmp & Chr(b) ':
    Loop Until b = 58
    strSize = val(tmp)
    tmp = vbNullString
    If strSize > 0 Then
        ReDim Buff(1 To strSize) As Byte
        Get hFile, , Buff
        tmp = Space(strSize)
        CopyMemory ByVal tmp, Buff(1), strSize
    End If
    GetString = tmp
End Function
