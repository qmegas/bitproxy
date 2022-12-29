VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEmulationItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvartDownloaded As Currency
Private mvartUploaded As Currency
Private mvartFullSize As Currency
Private mvartID As Long
Private mvartStatus As Integer
Private mvartName As String
Private mvartHash As String
Private mvartKey As String
Private mvartHost As String
Private mvartTracker As String
Private mvartPeerId As String
Private mvartPort As Long
Private mvarsUp1 As Long
Private mvarsUp2 As Long
Private mvarsDw1 As Long
Private mvarsDw2 As Long
Private mvartUTime As Long
Private mvartClient As String
Private mvartConnectionRiz As Integer
Private mvartLeft As Currency
Private mvartUseIgnorTime As Boolean
Private mvartIgnorTime As Integer
Private mvartStopMode As Integer
Private mvarsStopValue As Currency
Private mvartUseScrape As Boolean
Private mvartScrapeTime As Long
Private mvartScrapeStatus As Integer
Private mvartScrapeURL As String
Private mvartUseStepDownload As Boolean
Private mvartUseStepUpload As Boolean
Private mvartStepDownloadTime As Long
Private mvartStepUploadTime As Long
Private mvartInnerCounter As Long
Private mvartConnectTries As Long
Private mvartmpDataRecieved As String
Private mvartFrozen As Boolean
Private mvartFrozenSkip As Boolean
Private mvariEvent As String
Private mvartAddedTime As String

Public Property Let tAddedTime(ByVal vData As String)
    mvartAddedTime = vData
End Property

Public Property Get tAddedTime() As String
    tAddedTime = mvartAddedTime
End Property

Public Property Let iEvent(ByVal vData As String)
    mvariEvent = vData
End Property

Public Property Get iEvent() As String
    iEvent = mvariEvent
End Property

Public Property Let tFrozen(ByVal vData As Boolean)
    mvartFrozen = vData
End Property

Public Property Get tFrozen() As Boolean
    tFrozen = mvartFrozen
End Property

Public Property Let tFrozenSkip(ByVal vData As Boolean)
    mvartFrozenSkip = vData
End Property

Public Property Get tFrozenSkip() As Boolean
    tFrozenSkip = mvartFrozenSkip
End Property

Public Property Let tmpDataRecieved(ByVal vData As String)
    mvartmpDataRecieved = vData
End Property

Public Property Get tmpDataRecieved() As String
    tmpDataRecieved = mvartmpDataRecieved
End Property

Public Property Let tConnectTries(ByVal vData As Long)
    mvartConnectTries = vData
End Property

Public Property Get tConnectTries() As Long
    tConnectTries = mvartConnectTries
End Property

Public Property Let tInnerCounter(ByVal vData As Long)
    mvartInnerCounter = vData
End Property

Public Property Get tInnerCounter() As Long
    tInnerCounter = mvartInnerCounter
End Property

Public Property Let tStepUploadTime(ByVal vData As Long)
    mvartStepUploadTime = vData
End Property

Public Property Get tStepUploadTime() As Long
    tStepUploadTime = mvartStepUploadTime
End Property

Public Property Let tStepDownloadTime(ByVal vData As Long)
    mvartStepDownloadTime = vData
End Property

Public Property Get tStepDownloadTime() As Long
    tStepDownloadTime = mvartStepDownloadTime
End Property

Public Property Let tUseStepUpload(ByVal vData As Boolean)
    mvartUseStepUpload = vData
End Property

Public Property Get tUseStepUpload() As Boolean
    tUseStepUpload = mvartUseStepUpload
End Property

Public Property Let tUseStepDownload(ByVal vData As Boolean)
    mvartUseStepDownload = vData
End Property

Public Property Get tUseStepDownload() As Boolean
    tUseStepDownload = mvartUseStepDownload
End Property

Public Property Let tScrapeURL(ByVal vData As String)
    mvartScrapeURL = vData
End Property

Public Property Get tScrapeURL() As String
    tScrapeURL = mvartScrapeURL
End Property

Public Property Let tScrapeStatus(ByVal vData As Integer)
    mvartScrapeStatus = vData
End Property

Public Property Get tScrapeStatus() As Integer
    tScrapeStatus = mvartScrapeStatus
End Property

Public Property Let tScrapeTime(ByVal vData As Long)
    mvartScrapeTime = vData
End Property

Public Property Get tScrapeTime() As Long
    tScrapeTime = mvartScrapeTime
End Property

Public Property Let tUseScrape(ByVal vData As Boolean)
    mvartUseScrape = vData
End Property

Public Property Get tUseScrape() As Boolean
    tUseScrape = mvartUseScrape
End Property

Public Property Let tStopValue(ByVal vData As Currency)
    mvarsStopValue = vData
End Property

Public Property Get tStopValue() As Currency
    tStopValue = mvarsStopValue
End Property

Public Property Let tStopMode(ByVal vData As Integer)
    mvartStopMode = vData
End Property

Public Property Get tStopMode() As Integer
    tStopMode = mvartStopMode
End Property

Public Property Let tIgnorTime(ByVal vData As Integer)
    mvartIgnorTime = vData
End Property

Public Property Get tIgnorTime() As Integer
    tIgnorTime = mvartIgnorTime
End Property

Public Property Let tUseIgnorTime(ByVal vData As Boolean)
    mvartUseIgnorTime = vData
End Property

Public Property Get tUseIgnorTime() As Boolean
    tUseIgnorTime = mvartUseIgnorTime
End Property

Public Property Let tLeft(ByVal vData As Currency)
    mvartLeft = vData
End Property

Public Property Get tLeft() As Currency
    tLeft = mvartLeft
End Property

Public Property Let tConnectionRiz(ByVal vData As Integer)
    mvartConnectionRiz = vData
End Property

Public Property Get tConnectionRiz() As Integer
    tConnectionRiz = mvartConnectionRiz
End Property

Public Property Let tClient(ByVal vData As String)
    mvartClient = vData
End Property

Public Property Get tClient() As String
    tClient = mvartClient
End Property

Public Property Let tUTime(ByVal vData As Long)
    mvartUTime = vData
End Property

Public Property Get tUTime() As Long
    tUTime = mvartUTime
End Property

Public Property Let sDw2(ByVal vData As Long)
    mvarsDw2 = vData
End Property

Public Property Get sDw2() As Long
    sDw2 = mvarsDw2
End Property

Public Property Let sDw1(ByVal vData As Long)
    mvarsDw1 = vData
End Property

Public Property Get sDw1() As Long
    sDw1 = mvarsDw1
End Property

Public Property Let sUp2(ByVal vData As Long)
    mvarsUp2 = vData
End Property

Public Property Get sUp2() As Long
    sUp2 = mvarsUp2
End Property

Public Property Let sUp1(ByVal vData As Long)
    mvarsUp1 = vData
End Property

Public Property Get sUp1() As Long
    sUp1 = mvarsUp1
End Property

Public Property Let tPort(ByVal vData As Long)
    mvartPort = vData
End Property

Public Property Get tPort() As Long
    tPort = mvartPort
End Property

Public Property Let tPeerID(ByVal vData As String)
    mvartPeerId = vData
End Property

Public Property Get tPeerID() As String
    tPeerID = mvartPeerId
End Property

Public Property Let tTracker(ByVal vData As String)
    mvartTracker = vData
End Property

Public Property Get tTracker() As String
    tTracker = mvartTracker
End Property

Public Property Let tHost(ByVal vData As String)
    mvartHost = vData
End Property

Public Property Get tHost() As String
    tHost = mvartHost
End Property

Public Property Let tKey(ByVal vData As String)
    mvartKey = vData
End Property

Public Property Get tKey() As String
    tKey = mvartKey
End Property

Public Property Let tHash(ByVal vData As String)
    mvartHash = vData
End Property

Public Property Get tHash() As String
    tHash = mvartHash
End Property

Public Property Let tName(ByVal vData As String)
    mvartName = vData
End Property

Public Property Get tName() As String
    tName = mvartName
End Property

Public Property Let tStatus(ByVal vData As Integer)
    mvartStatus = vData
End Property

Public Property Get tStatus() As Integer
    tStatus = mvartStatus
End Property

Public Property Let tID(ByVal vData As Long)
    mvartID = vData
End Property

Public Property Get tID() As Long
    tID = mvartID
End Property

Public Property Let tFullSize(ByVal vData As Currency)
    mvartFullSize = vData
End Property

Public Property Get tFullSize() As Currency
    tFullSize = mvartFullSize
End Property

Public Property Let tUploaded(ByVal vData As Currency)
    mvartUploaded = vData
End Property

Public Property Get tUploaded() As Currency
    tUploaded = mvartUploaded
End Property

Public Property Let tDownloaded(ByVal vData As Currency)
    mvartDownloaded = vData
End Property

Public Property Get tDownloaded() As Currency
    tDownloaded = mvartDownloaded
End Property
