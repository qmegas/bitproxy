Attribute VB_Name = "modUpdater"
Option Explicit

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetConnectA Lib "wininet.dll" (ByVal hInternetSession As Long, ByVal lpszServerName As String, ByVal nProxyPort As Integer, ByVal lpszUsername As String, ByVal lpszPassword As String, ByVal dwService As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function HttpAddRequestHeadersA Lib "wininet.dll" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lModifiers As Long) As Integer
Private Declare Function HttpOpenRequestA Lib "wininet.dll" (ByVal hInternetSession As Long, ByVal lpszVerb As String, ByVal lpszObjectName As String, ByVal lpszVersion As String, ByVal lpszReferer As String, ByVal lpszAcceptTypes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function HttpSendRequestA Lib "wininet.dll" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal sOptional As String, ByVal lOptionalLength As Long) As Boolean
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Const UPDATE_OK = 0
Public Const UPDATE_NOCONNECT = 1
Public Const UPDATE_WRONGDATA = 2

Public Sub CheckUpdate(Loud As Boolean)
    On Error Resume Next
    Dim rez As Integer
    Dim cID As Long, i As Long

    Const PURL = "/progs/1.htm"
    
    rez = SendRequest("check", vbNullString)
    If rez <> UPDATE_OK Then
        If Loud Then _
            ShowStandartErrorMsg rez
    Else
        cID = val(WebBuff)
        Select Case cID
            Case 0 'No new version
                If Loud Then _
                    MsgBox getCStr(2, "You have the latest version of program"), vbInformation, App.Title
            Case 1 'Have new version
                If (prgSettings.chkAutoUpdate = 1) Or Loud Then
                    i = MsgBox(getCStr(3, "New version found. Do you want visit program's homepage to download it?"), vbInformation + vbYesNo, App.Title)
                    If i = vbYes Then _
                        ShellExecute MainForm.hwnd, vbNullString, HTTP & SITEURL & PURL, vbNullString, App.Path, vbNormalFocus
                End If
        End Select
    End If
End Sub

Private Function SendRequest(act As String, add_info As String) As Integer
    Dim pVer As String, pData As String
    
    Const BROWZ = "update_checker v1.1"
    Const UURL = "/updater.php"
    
    pVer = App.Major & "." & App.Minor
    hSession = InternetOpen(BROWZ, 1, vbNullString, vbNullString, 0)
    pData = "progID=bitproxy&ver=" & pVer
    If Len(act) > 0 Then pData = pData & "&act=" & act
    If Len(add_info) > 0 Then pData = pData & "&" & add_info
    
    If PostUrlData(HTTP & SITEURL & UURL, pData) Then
        If Left(WebBuff, 4) = "uok|" Then
            WebBuff = Mid(WebBuff, 5)
            SendRequest = UPDATE_OK
        Else
            SendRequest = UPDATE_WRONGDATA
        End If
    Else
        SendRequest = UPDATE_NOCONNECT
    End If
End Function

Private Function PostUrlData(ByVal URL As String, ByRef data As String) As Boolean
    Dim Server As String
    Dim Path As String
    Dim hConnect As Long
    Dim hRequest As Long
    Dim buffer As String * 2048
    Dim Bytes  As Long
    Dim i As Long
  
    WebBuff = vbNullString
    URL = Trim(URL)
    If LCase(Left(URL, 7)) = "http://" Then URL = Mid(URL, 8)
    If LCase(Left(URL, 8)) = "https://" Then URL = Mid(URL, 9)
    i = InStr(URL, "/")
    If i > 0 Then
        Server = Left(URL, i - 1)
        Path = Mid(URL, i)
    Else
        Server = URL
        Path = "/"
    End If
  
    hConnect = InternetConnectA(hSession, Server, 80, vbNullString, vbNullString, INTERNET_SERVICE_HTTP, 0, 0)
    hRequest = HttpOpenRequestA(hConnect, "POST", Path, "HTTP/1.0", vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
  
    HttpAddRequestHeadersA hRequest, INET_ContentType, Len(INET_ContentType), HTTP_ADDREQ_FLAG_REPLACE Or HTTP_ADDREQ_FLAG_ADD
    HttpSendRequestA hRequest, vbNullString, 0, data, Len(data)
  
    Do
        InternetReadFile hRequest, buffer, Len(buffer), Bytes
        If Bytes = 0 Then Exit Do
        WebBuff = WebBuff & Left(buffer, Bytes)
    Loop
  
    PostUrlData = (WebBuff <> vbNullString)
  
    InternetCloseHandle hRequest
    InternetCloseHandle hConnect
End Function

Private Sub ShowStandartErrorMsg(err_code As Integer)
    Select Case err_code
        Case UPDATE_NOCONNECT
            MsgBox getCStr(5, "Can not connect to program's homepage"), vbCritical, App.Title
        Case UPDATE_WRONGDATA
            MsgBox getCStr(4, "Error recieving data from the server"), vbCritical, App.Title
    End Select
End Sub
