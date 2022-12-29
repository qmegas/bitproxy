VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGrafik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type POINTDATA
    X As Long
    y As Long
End Type

Private Type LINEDATA
    color As Long
    cur_index As Long
    points() As POINTDATA
End Type

Private Const SCHEME_COLOR = &H606060
Private Const LITE_SCHEME_COLOR = &HA0A0A0
Private Const BACK_COLOR = &HFFFFFF

Private Const FROM_BORDERS = 4
Private Const FROM_LEFT = 40
Private Const FROM_BOTOM = 20

Private Const MIN_Y_STEP = 25
Private Const MIN_X_STEP = 50

Private lines() As LINEDATA
Private cur_pic As PictureBox
Private cur_index As Long
'Drawinf
Private val_per_pix_y As Double
Private val_per_pix_x As Double
Private max_x As Long
Private max_y As Long

Public Sub Init(pic As PictureBox)
    Set cur_pic = pic
    
    cur_pic.AutoRedraw = True
    cur_pic.BorderStyle = 0
    cur_pic.ScaleMode = vbPixels
End Sub

Public Sub ResetData()
    Erase lines
    cur_index = 0
End Sub

Public Function AddLine(color As Long, must_zero As Boolean) As Long
    Dim new_indx As Long
    
    cur_index = cur_index + 1
    ReDim Preserve lines(1 To cur_index) As LINEDATA
    lines(cur_index).color = color
    lines(cur_index).cur_index = 0
    If must_zero Then AddPoint cur_index, 0, 0
    
    AddLine = cur_index
End Function

Public Sub AddPoint(hLine As Long, X As Long, y As Long)
    Dim new_pi As Long
    
    If y < 0 Then y = 0
    
    new_pi = lines(hLine).cur_index + 1
    lines(hLine).cur_index = new_pi
    
    ReDim Preserve lines(hLine).points(1 To new_pi) As POINTDATA
    lines(hLine).points(new_pi).X = X
    lines(hLine).points(new_pi).y = y
End Sub

Public Sub DrawGraphik()
    cur_pic.Cls
    calculate_max_values
    draw_scheme
    draw_lines
End Sub

Private Sub draw_lines()
    Dim i As Long, j As Long
    Dim prevX As Long, prevY As Long
    Dim newX As Long, newY As Long
    
    For i = 1 To cur_index
        If lines(i).cur_index > 0 Then
            prevX = getXPoz(lines(i).points(1).X)
            prevY = getYPoz(lines(i).points(1).y)
        End If
        For j = 2 To lines(i).cur_index
            newX = getXPoz(lines(i).points(j).X)
            newY = getYPoz(lines(i).points(j).y)
            cur_pic.Line (prevX, prevY)-(newX, newY), lines(i).color
            prevX = newX
            prevY = newY
        Next
    Next
End Sub

Private Sub draw_scheme()
    Dim rounded As Long, steps As Long
    Dim i As Long
    Dim poz As Long
    Dim tmp As String
    
    'main
    cur_pic.Line (0, 0)-(cur_pic.ScaleWidth, cur_pic.ScaleHeight), BACK_COLOR, BF
    cur_pic.Line (FROM_LEFT, FROM_BORDERS)-(FROM_LEFT, cur_pic.ScaleHeight - FROM_BOTOM), SCHEME_COLOR
    cur_pic.Line (FROM_LEFT, cur_pic.ScaleHeight - FROM_BOTOM)-(cur_pic.ScaleWidth - FROM_BORDERS, cur_pic.ScaleHeight - FROM_BOTOM), SCHEME_COLOR
    
    'helper lines
    'Y
    rounded = RoundItUp(CLng(val_per_pix_y * MIN_Y_STEP))
    steps = max_y \ rounded
    For i = 1 To steps
        poz = getYPoz(rounded * i)
        cur_pic.Line (FROM_LEFT + 1, poz)-(cur_pic.ScaleWidth - FROM_BORDERS, poz), LITE_SCHEME_COLOR
        cur_pic.CurrentX = 2
        cur_pic.CurrentY = poz - 7
        cur_pic.Print getYLabel(rounded * i)
    Next
    'X
    rounded = RoundItUp(CLng(val_per_pix_x * MIN_X_STEP))
    steps = max_x \ rounded
    For i = 1 To steps
        poz = getXPoz(rounded * i)
        cur_pic.Line (poz, FROM_BORDERS)-(poz, cur_pic.ScaleHeight - FROM_BOTOM), LITE_SCHEME_COLOR
        cur_pic.CurrentX = poz - 10
        cur_pic.CurrentY = cur_pic.ScaleHeight - (FROM_BOTOM * 0.9)
        cur_pic.Print getXLabel(rounded * i)
    Next
End Sub

Private Sub calculate_max_values()
    Dim i As Long, j As Long
    
    max_x = 0
    max_y = 0
    For i = 1 To cur_index
        For j = 1 To lines(i).cur_index
            If max_x < lines(i).points(j).X Then max_x = lines(i).points(j).X
            If max_y < lines(i).points(j).y Then max_y = lines(i).points(j).y
        Next
    Next
    val_per_pix_y = max_y / (cur_pic.ScaleHeight - FROM_BOTOM - FROM_BORDERS - 1)
    val_per_pix_x = max_x / (cur_pic.ScaleWidth - FROM_LEFT - FROM_BORDERS - 1)
    If val_per_pix_y = 0 Then val_per_pix_y = 1
    If val_per_pix_x = 0 Then val_per_pix_x = 1
End Sub

Private Function RoundItUp(num As Long) As Long
    Dim coef As Long
    If num < 10 Then
        RoundItUp = 10
    Else
        coef = Int(Log(num) / Log(10#))
        num = Int(num / (10 ^ coef)) + 1
        RoundItUp = num * (10 ^ coef)
    End If
End Function

Private Function getYPoz(val As Long) As Long
    getYPoz = CLng((max_y - val) / val_per_pix_y) + FROM_BORDERS
End Function

Private Function getXPoz(val As Long) As Long
    getXPoz = CLng(val / val_per_pix_x) + FROM_LEFT + 1
End Function

Private Function getXLabel(val As Long) As String
    Dim h As Long, m As Long
    
    val = val \ 60
    h = val \ 60
    m = val Mod 60
    getXLabel = CStr(h) & ":" & Format(m, "#00")
End Function

Private Function getYLabel(val As Long) As String
    Const MEGA = 1000
    Const GIGA = 1000000
    Const TERA = 1000000000
    If val >= TERA Then
        getYLabel = CStr(val / TERA) & "T"
    ElseIf val >= GIGA Then
        getYLabel = CStr(val / GIGA) & "G"
    ElseIf val >= MEGA Then
        getYLabel = CStr(val / MEGA) & "M"
    Else
        getYLabel = CStr(val) & "K"
    End If
End Function

Private Sub Class_Initialize()
    cur_index = 0
    max_x = 0
    max_y = 0
End Sub
