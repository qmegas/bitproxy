VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsClientItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarTitle As String
Private mvarAuthor As String
Private mvarPrefix As String
Private mvarPrifixSize As Integer
Private mvarHeaders As String
Private mvarProtocol As String
Private mvarURL As String
Private mvarHashEncoding As Boolean
Private mvarHashUpper As Boolean
Private mvarKeySize As Integer
Private mvarKeyType As String
Private mvarKeyUpper As Boolean
Private mvarUserAgent As String

Public Property Let UserAgent(ByVal vData As String)
    mvarUserAgent = vData
End Property

Public Property Get UserAgent() As String
    UserAgent = mvarUserAgent
End Property

Public Property Let KeyUpper(ByVal vData As Boolean)
    mvarKeyUpper = vData
End Property

Public Property Get KeyUpper() As Boolean
    KeyUpper = mvarKeyUpper
End Property

Public Property Let KeyType(ByVal vData As String)
    mvarKeyType = vData
End Property

Public Property Get KeyType() As String
    KeyType = mvarKeyType
End Property

Public Property Let KeySize(ByVal vData As Integer)
    mvarKeySize = vData
End Property

Public Property Get KeySize() As Integer
    KeySize = mvarKeySize
End Property

Public Property Let HashUpper(ByVal vData As Boolean)
    mvarHashUpper = vData
End Property

Public Property Get HashUpper() As Boolean
    HashUpper = mvarHashUpper
End Property

Public Property Let HashEncoding(ByVal vData As Boolean)
    mvarHashEncoding = vData
End Property

Public Property Get HashEncoding() As Boolean
    HashEncoding = mvarHashEncoding
End Property

Public Property Let URL(ByVal vData As String)
    mvarURL = vData
End Property

Public Property Get URL() As String
    URL = mvarURL
End Property

Public Property Let Protocol(ByVal vData As String)
    mvarProtocol = vData
End Property

Public Property Get Protocol() As String
    Protocol = mvarProtocol
End Property

Public Property Let Headers(ByVal vData As String)
    mvarHeaders = vData
End Property

Public Property Get Headers() As String
    Headers = mvarHeaders
End Property

Public Property Let PrifixSize(ByVal vData As Integer)
    mvarPrifixSize = vData
End Property

Public Property Get PrifixSize() As Integer
    PrifixSize = mvarPrifixSize
End Property

Public Property Let Prefix(ByVal vData As String)
    mvarPrefix = vData
End Property

Public Property Get Prefix() As String
    Prefix = mvarPrefix
End Property

Public Property Let Author(ByVal vData As String)
    mvarAuthor = vData
End Property

Public Property Get Author() As String
    Author = mvarAuthor
End Property

Public Property Let Title(ByVal vData As String)
    mvarTitle = vData
End Property

Public Property Get Title() As String
    Title = mvarTitle
End Property

