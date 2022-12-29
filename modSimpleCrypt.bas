Attribute VB_Name = "modSimpleCrypt"
Option Explicit

Private Const SIMPLE_KEY = "j$s2P1ASgh6-r"
Private Const LIMIT_MIN = 33
Private Const LIMIT_MAX = 126

Public Function SimpleCrypt(data As String) As String
    Dim i As Byte, tmp As String
    
    If Len(data) = 0 Then
        SimpleCrypt = vbNullString
        Exit Function
    End If
    
    Randomize Timer
    i = Rnd(Day(Now)) * 93 + LIMIT_MIN
    tmp = Shuffle(data, i)
    SimpleCrypt = PutKey(Chr(i) & tmp, True)
End Function

Public Function SimpleDecrypt(data As String) As String
    If Len(data) <> 17 Then
        SimpleDecrypt = vbNullString
        Exit Function
    End If
    
    data = PutKey(data, False)
    SimpleDecrypt = DeShuffle(Mid(data, 2), Asc(Mid(data, 1, 1)))
End Function

Private Function Shuffle(data As String, key As Byte) As String
    Dim i As Integer
    Dim getByte As Integer
    Dim swap_temp As Byte
    Dim tmp() As Byte
    ReDim tmp(1 To Len(data)) As Byte
    
    For i = 1 To UBound(tmp)
        tmp(i) = Asc(Mid(data, i, 1))
    Next
    
    For i = 1 To Len(data) - 1
        getByte = i Mod 7
        If ((2 ^ getByte) And key) Then
            swap_temp = tmp(i)
            tmp(i) = tmp(UBound(tmp))
            tmp(UBound(tmp)) = swap_temp
        End If
    Next
    
    Shuffle = vbNullString
    For i = 1 To UBound(tmp)
        Shuffle = Shuffle & Chr(tmp(i))
    Next
End Function

Private Function DeShuffle(data As String, key As Byte) As String
    Dim i As Integer
    Dim getByte As Integer
    Dim swap_temp As Byte
    Dim tmp() As Byte
    ReDim tmp(1 To Len(data)) As Byte
    
    For i = 1 To UBound(tmp)
        tmp(i) = Asc(Mid(data, i, 1))
    Next
    
    For i = Len(data) - 1 To 1 Step -1
        getByte = i Mod 7
        If ((2 ^ getByte) And key) Then
            swap_temp = tmp(i)
            tmp(i) = tmp(UBound(tmp))
            tmp(UBound(tmp)) = swap_temp
        End If
    Next
    
    DeShuffle = vbNullString
    For i = 1 To UBound(tmp)
        DeShuffle = DeShuffle & Chr(tmp(i))
    Next
End Function

Private Function PutKey(data As String, is_crypt As Boolean) As String
    Dim i As Integer
    Dim cur_asc As Integer
    Dim cur_key As Integer
    
    PutKey = vbNullString
    For i = 1 To Len(data)
        cur_asc = Asc(Mid(data, i, 1))
        cur_key = Asc(Mid(SIMPLE_KEY, (i Mod Len(SIMPLE_KEY)) + 1, 1))
        If is_crypt Then cur_asc = cur_asc + cur_key _
            Else cur_asc = cur_asc - cur_key
        If cur_asc < LIMIT_MIN Then cur_asc = cur_asc - LIMIT_MIN + LIMIT_MAX + 1
        If cur_asc > LIMIT_MAX Then cur_asc = cur_asc - LIMIT_MAX + LIMIT_MIN - 1
        PutKey = PutKey & Chr(cur_asc)
    Next
End Function
