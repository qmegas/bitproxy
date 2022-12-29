Attribute VB_Name = "Module1"
Option Explicit

Declare Function CreateAlgorithm Lib "sha1.dll" () As Long
Declare Function DestroyAlgorithm Lib "sha1.dll" (ByVal hAlgorithm As Long) As Long
Declare Function InitAlgorithm Lib "sha1.dll" (ByVal hAlgorithm As Long) As Long
Declare Function InputData Lib "sha1.dll" (ByVal hAlgorithm As Long, ByVal data As String, ByVal leng As Long) As Long
Declare Function GetHashLength Lib "sha1.dll" (ByVal hAlgorithm As Long) As Long
Declare Function GetHash Lib "sha1.dll" (ByVal hAlgorithm As Long, ByVal buffer As String, ByVal bufsize As Long) As Long
Declare Function GetHashTextLength Lib "sha1.dll" (ByVal hAlgorithm As Long) As Long
Declare Function GetHashText Lib "sha1.dll" (ByVal hAlgorithm As Long, ByVal buffer As String, ByVal bufsize As Long) As Long


Public Function SHA1f(intxt As String) As String
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
                Else
                    MsgBox "4"
                End If
            Else
                MsgBox "3"
            End If
        Else
            MsgBox "2"
        End If
        Call DestroyAlgorithm(hAlgorithm)
    Else
        MsgBox "1"
    End If
End Function
