VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCRC32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' This code is taken from the VB.NET CRC32 algorithm
' provided by Paul (wpsjr1@succeed.net) - Excellent work!

Private crc32Table() As Long
Private Const BUFFER_SIZE As Long = 8192

Public Function GetByteArrayCrc32(ByRef buffer() As Byte) As Long
   
   Dim crc32Result As Long
   crc32Result = &HFFFFFFFF
      
   Dim i As Long
   Dim iLookup As Integer
   
   For i = LBound(buffer) To UBound(buffer)
      iLookup = (crc32Result And &HFF) Xor buffer(i)
      crc32Result = ((crc32Result And &HFFFFFF00) \ &H100) And 16777215
      crc32Result = crc32Result Xor crc32Table(iLookup)
   Next i
   
   GetByteArrayCrc32 = Not (crc32Result)

End Function

Private Sub Class_Initialize()

    ' This is the official polynomial used by CRC32 in PKZip.
    ' Often the polynomial is shown reversed (04C11DB7).
    Dim dwPolynomial As Long
    dwPolynomial = &HEDB88320
    Dim i As Integer, j As Integer

    ReDim crc32Table(256)
    Dim dwCrc As Long

    For i = 0 To 255
        dwCrc = i
        For j = 8 To 1 Step -1
            If (dwCrc And 1) Then
                dwCrc = ((dwCrc And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                dwCrc = dwCrc Xor dwPolynomial
            Else
                dwCrc = ((dwCrc And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
            End If
        Next j
        crc32Table(i) = dwCrc
    Next i

End Sub

