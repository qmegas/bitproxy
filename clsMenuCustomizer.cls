VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMenuCustomizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const REG_PREFIX = "Column_"

Public Sub LoadState(lv As ListView)
    Dim tmp As String, i As Long, curID As Long
    Dim columns As Variant
    Dim vars As Variant
    
    On Error GoTo ErrorHandler
    
    tmp = GetKeyValue(HKEY_LOCAL, MYPATH, REG_PREFIX & lv.name, vbNullString)
    If Len(tmp) = 0 Then Exit Sub
    
    columns = VB6_Split(tmp, ";")
    
    For i = LBound(columns) To UBound(columns)
        vars = VB6_Split(columns(i), "|")
        
        curID = val(vars(0))
        
        If lv.ColumnHeaders.Count >= curID Then
            lv.ColumnHeaders(curID).Width = val(vars(1))
            lv.ColumnHeaders(curID).Position = val(vars(2))
        End If
    Next
    
ErrorHandler:
End Sub

Public Sub SaveState(lv As ListView)
    Dim tmp As String
    Dim i As Long
    
    For i = 1 To lv.ColumnHeaders.Count
        If Len(tmp) > 0 Then tmp = tmp & ";"
        tmp = tmp & CStr(i) & "|" & lv.ColumnHeaders(i).Width & "|" & lv.ColumnHeaders(i).Position
    Next
    WriteKey HKEY_LOCAL, MYPATH, REG_PREFIX & lv.name, tmp
End Sub
