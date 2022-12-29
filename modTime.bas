Attribute VB_Name = "modTime"
Option Explicit

Private Type SystemTime
       wYear As Integer
       wMonth As Integer
       wDayOfWeek As Integer
       wDay As Integer
       wHour As Integer
       wMinute As Integer
       wSecond As Integer
       wMilliseconds As Integer
End Type
Private Type TIME_ZONE_INFORMATION
       Bias As Long
       StandardName(32) As Integer
       StandardDate As SystemTime
       StandardBias As Long
       DaylightName(32) As Integer
       DaylightDate As SystemTime
       DaylightBias As Long
End Type
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Function FromUnixTime(ByVal sUnixTime As Long) As Date
    Dim NTime As Date, STime As Date
    Dim TZ As TIME_ZONE_INFORMATION
    STime = #1/1/1970#
    NTime = DateAdd("s", sUnixTime, STime)
    GetTimeZoneInformation TZ
    NTime = DateAdd("n", -TZ.Bias, NTime)
    FromUnixTime = NTime
End Function

Public Function ToUnixTime(ByVal STime As Date) As Long
    Dim NTime As Date, sUnix As Date, sUnixTime As Long
    Dim TZ As TIME_ZONE_INFORMATION
    sUnix = #1/1/1970#
    GetTimeZoneInformation TZ
    NTime = DateAdd("n", TZ.Bias, STime)
    sUnixTime = DateDiff("s", sUnix, NTime)
    ToUnixTime = sUnixTime
End Function
