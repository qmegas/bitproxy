Attribute VB_Name = "Globals"
Option Explicit

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Public Type POINTAPI
    X As Long
    y As Long
End Type

Public Type CLIENT_ID
    peer_id As String
    Cleint As String
End Type

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

Public Type MSG_DATA
    info_hash As String
    peer_id As String
    uploaded As Currency
    downloaded As Currency
    event As String
    real_upload As Currency
    port As Long
    Left As Currency
    numwant As Long
    compact As Long
    no_peer_id As Long
    key As String
    cur_time As Long
End Type

Public Type PROG_SETT
    Sett As String
    '========
    LangSelected As String
    txtServer As String
    txtPort As Long
    '========
    optMode As Integer
    txtUpload As Double
    txtM2from As String
    txtM2to As String
    chkDown As Integer
    txtDown As Long
    chkDnotsend As Integer
    '========
    chkVersion As Integer
    cmbVersion As Long
    '========
    chkMinimize As Integer
    chkAutoUpdate As Integer
    chkUseProxy As Integer
    txtProxyIp As String
    txtProxyPort As Long
    chkRetracker As Integer
    '========
    chkSmart As Integer
    txtSmartA As Long
    txtSmartP As Integer
    '========
    defAction As Integer
    '========
    emul_up1 As Long
    emul_up2 As Long
    emul_dw1 As Long
    emul_dw2 As Long
    emul_client As String
    emul_port As Long
    chkIgnorT As Integer
    txtIgnorT As Integer
    SaveList As Integer
    chkUseScrape As Integer
    chkIgnorServerError As Integer
    chkIgnorSocketError As Integer
    txtConnectTries As Integer
    chkFroze As Boolean
    chkStepModeD As Boolean
    chkStepModeU As Boolean
    txtSMDown As Integer
    txtSMUp As Integer
    txtHave As Integer
    chkSameHash As Integer
    '========
    chkRemote As Integer
    txtRCPort As Integer
    txtRCPass As String
End Type

Public Type NOTIFYICONDATA
   cbSize As Long
   hwnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * 128
   dwState As Long
   dwStateMask As Long
   szInfo As String * 256
   uTimeout As Long
   szInfoTitle As String * 64
   dwInfoFlags As Long
End Type

Public Const ERROR_ALREADY_EXISTS = 183&

Public Const STATUS_ON = 1
Public Const STATUS_OFF = 0
                       
Public Const TRR = ".torrent"
Public Const FLOG = "log.txt"
Public Const SAVELISTF = "emullist.sav"
Public Const SAVEDOWNF = "savelist.sav"
Public Const ERRORLOG = "error.txt"
Public Const LANGDIR = "lang\"
Public Const CLIENT_DIR = "clients\"
Public Const UPKEY = "uploaded="
Public Const DOWNKEY = "downloaded="
Public Const LEFTKEY = "left="
Public Const PEERID = "peer_id="
Public Const UAGENT = "User-Agent: "
Public Const HOST = "Host: "
Public Const GZIPKEY = "Content-Encoding: gzip"
Public Const BP10 = "BP14"
Public Const INIITEM = "Item"
Public Const HTTP = "http://"
Public Const HTTPS = "https://"
Public Const SITEURL = "qmegas.info"
Public Const HELPURL = "/custom/bit_help.htm"
Public Const DONATIONURL = "/custom/donate.htm"
Public Const FORUMURL = "/index.php?board=27.0"
Public Const MAX_LIST = 10
Public Const CHECKSUM1 = 1921559657
Public Const SCRAPE_STEP = 1200
Public Const GRAPH_TAB = 3
Public Const REDIRECT_COMMAND = -301

Public Const TORRENT_START = "started"

Public Const CKILO = 1024@
Public Const CMEGA = 1048576@
Public Const CGIGA = 1073741824@
Public Const CTERA = 1099511627776@

Public Const GWL_STYLE = (-16)
Public Const WS_BORDER = &H800000
Public Const GWL_WNDPROC = (-4&)

Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_USER = &H400
Public Const TRAY_BACK = (WM_USER + 200)
Public Const MAKE_NEW = (WM_USER + 300)
Public Const WM_DROPFILES = &H233

Public Const NIM_ADD = &H0&
Public Const NIM_MODIFY = &H1&
Public Const NIM_DELETE = &H2&

Public Const NIF_MESSAGE = &H1&
Public Const NIF_ICON = &H2&
Public Const NIF_TIP = &H4&
Public Const NIF_INFO = &H10

Public Const NIIF_ERROR = &H3

Public Const INTERNET_FLAG_RELOAD = &H80000000
Public Const INTERNET_SERVICE_HTTP = 3
Public Const INET_ContentType = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
Public Const HTTP_ADDREQ_FLAG_ADD = &H20000000
Public Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000

Public Const EMUL_ADD = 0
Public Const EMUL_STOP = 1
Public Const EMUL_WORK = 2
Public Const EMUL_UPDATE = 3
Public Const EMUL_TIMEDOUT = 4
Public Const EMUL_ERROR = 5

Public Const CONNECT_START = 1
Public Const CONNECT_STOP = 2
Public Const CONNECT_UPDATE = 3

Public Const OF_EXIST As Long = &H4000
Public Const OFS_MAXPATHNAME As Long = 128
Public Const HFILE_ERROR As Long = -1

Public Const CP_UTF8 = 65001

Public Const TPM_RETURNCMD = &H100&

Public Const MFS_DEFAULT As Long = &H1000
Public Const MF_DEFAULT = &H1000&
Public Const MF_SEPARATOR = &H800&
Public Const MF_OWNERDRAW = &H100&
Public Const MF_DISABLED = &H2&
Public Const MIIM_STATE As Long = &H1
Public Const API_TRUE As Long = 1&
Public Const WM_MEASUREITEM = &H2C
Public Const WM_DRAWITEM = &H2B

Public Const DC_GRADIENT = &H20
Public Const DC_ACTIVE = &H1
'Public Const DC_ICON = &H4
Public Const DC_SMALLCAP = &H2
Public Const DC_TEXT = &H8

Public Const SM_CYCAPTION = 4
Public Const SM_CXFRAME = 32
Public Const SM_CYFRAME = 33

Public Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpSection As String, ByVal lpKey As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal lprc As Any) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineW" () As Long
Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameW" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Declare Sub DragAcceptFiles Lib "shell32" (ByVal hwnd As Long, ByVal fAccept As Long)
Declare Sub DragFinish Lib "shell32" (ByVal hDrop As Long)
Declare Function DragQueryFile Lib "shell32" Alias "DragQueryFileW" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long

Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public CPath As String
Public TempPath As String
Public newClient As CLIENT_ID
Public SocetInBuffer As String
Public AppArg As String
Public hSession As Long
Public WebBuff As String
Public prgSettings As PROG_SETT
Public PortIncr As Long
Public ifCheckedUp As Boolean
Public editedEmulation As Long

Public Global_TaskbarMsg As Long

Public prev As Long 'Main Window handler
Public hPopMenu As Long 'Popup menu handler

Public emul As New clsEmulationManager
Public clnt As New clsClientManager
Public Jobs As New clsDownloadManager
Public Global_FilePool As New Collection
Public Global_MenuCustomizer As New clsMenuCustomizer
