Attribute VB_Name = "modRegWork"
Option Explicit

Public Const REG_SETTINGS_SETT = "Sett"
Public Const REG_SETTINGS_LANG = "Lang"
Public Const REG_SETTINGS_HOST = "Host"
Public Const REG_SETTINGS_PORT = "Port"
Public Const REG_SETTINGS_MODE = "UseMode"
Public Const REG_SETTINGS_UPLOAD = "Upl_count"
Public Const REG_SETTINGS_M2FROM = "m2_from"
Public Const REG_SETTINGS_M2TO = "m2_to"
Public Const REG_SETTINGS_DWNUSE = "Dwn_use"
Public Const REG_SETTINGS_DWNVAL = "Dwn_val"
Public Const REG_SETTINGS_DWNNOTSEND = "Dwn_notsend"
Public Const REG_SETTINGS_USEVER = "UseVer"
Public Const REG_SETTINGS_VERTYPE = "Ver_type"
Public Const REG_SETTINGS_MINIMIZE = "Minimize"
Public Const REG_SETTINGS_AUTOCHECK = "AutoCheck"
Public Const REG_SETTINGS_USEPROXY = "UseProxy"
Public Const REG_SETTINGS_PROXYIP = "ProxyIP"
Public Const REG_SETTINGS_PROXYPORT = "ProxyPort"
Public Const REG_SETTINGS_RETRACKER = "Retracker"
Public Const REG_SETTINGS_SMARTUSE = "SmartUse"
Public Const REG_SETTINGS_SMARTA = "SmartA"
Public Const REG_SETTINGS_SMARTP = "SmartP"
Public Const REG_SETTINGS_DEFACTION = "DefAction"
Public Const REG_SETTINGS_EMULCLIENT = "EmulDefClient"
Public Const REG_SETTINGS_EMULDW1 = "EmulDw1"
Public Const REG_SETTINGS_EMULDW2 = "EmulDw2"
Public Const REG_SETTINGS_EMULUP1 = "EmulUp1"
Public Const REG_SETTINGS_EMULUP2 = "EmulUp2"
Public Const REG_SETTINGS_EMULPORT = "EmulPort"
Public Const REG_SETTINGS_USEIGNOR = "UseIgnor"
Public Const REG_SETTINGS_IGNORTIME = "IgnorTime"
Public Const REG_SETTINGS_SAVELIST = "SaveList"
Public Const REG_SETTINGS_USESCRAPE = "UseScrape"
Public Const REG_SETTINGS_IGNORSERVERR = "IgnorServErr"
Public Const REG_SETTINGS_IGNORSOCKETERR = "IgnorSocketErr"
Public Const REG_SETTINGS_CONNTRIES = "ConnectTries"
Public Const REG_SETTINGS_FROZE = "Froze"
Public Const REG_SETTINGS_STEPMODED = "StepModeD"
Public Const REG_SETTINGS_STEPMODEU = "StepModeU"
Public Const REG_SETTINGS_STEPMODEDVAL = "StepModeDVal"
Public Const REG_SETTINGS_STEPMODEUVAL = "StepModeUVal"
Public Const REG_SETTINGS_EMULHAVE = "EmulHave"
Public Const REG_SETTINGS_SAMEHASH = "SameHash"
Public Const REG_SETTINGS_REMOTEUSE = "RemoteUse"
Public Const REG_SETTINGS_REMOTEPORT = "RemotePort"
Public Const REG_SETTINGS_REMOTEPASS = "RemotePass"


Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL = &H80000002
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const REG_OPTION_NON_VOLATILE = 0
Public Const REG_SZ = 1
Public Const REG_DWORD = 4
Public Const REG_BINARY = 3

Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                       
Public Const MYPATH = "Software\Megas\BitProxy"
Public Const AUTORUN = "Software\Microsoft\Windows\CurrentVersion\Run"

Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Public Function WriteKey(KeyRoot As Long, KeyName As String, SubKeyName As String, SubKeyValue As String) As Boolean
    Dim rc As Long                                      ' Return Code
    Dim hKey As Long                                    ' Handle To A Registry Key
    Dim hDepth As Long                                  '
    Dim lpAttr As SECURITY_ATTRIBUTES                   ' Registry Security Type
    lpAttr.nLength = 50                                 ' Set Security Attributes To Defaults...
    lpAttr.lpSecurityDescriptor = 0                     ' ...
    lpAttr.bInheritHandle = True                        ' ...
    rc = RegCreateKeyEx(KeyRoot, KeyName, 0, REG_SZ, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, hKey, hDepth)
    rc = RegSetValueEx(hKey, SubKeyName, 0, REG_SZ, SubKeyValue, Len(SubKeyValue))
    rc = RegCloseKey(hKey)                              ' Close Key
End Function

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, Optional AsDef As String) As String
    Dim rc As Long
    Dim hKey As Long
    Dim KeyValType As Long
    Dim tmpVal As String
    Dim KeyValSize As Long
    
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
    tmpVal = String(1024, 0)
    KeyValSize = 1024
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)
    If KeyValSize = 0 Then
        tmpVal = vbNullString
    Else
        If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then
            tmpVal = Left(tmpVal, KeyValSize - 1)
        Else
            tmpVal = Left(tmpVal, KeyValSize)
        End If
        If Asc(Left(tmpVal, 1)) = 0 Then _
            tmpVal = AsDef
    End If
    GetKeyValue = tmpVal
    rc = RegCloseKey(hKey)
End Function

Public Sub DeleteKey(KeyRoot As Long, KeyName As String, SubKey As String)
    Dim rc As Long
    Dim hKey As Long
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
    rc = RegDeleteKey(hKey, SubKey)
    rc = RegCloseKey(hKey)
End Sub

Public Sub DeleteValue(KeyRoot As Long, KeyName As String, SubKey As String)
    Dim rc As Long
    Dim hKey As Long
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
    rc = RegDeleteValue(hKey, SubKey)
    rc = RegCloseKey(hKey)
End Sub

