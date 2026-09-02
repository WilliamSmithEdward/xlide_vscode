Attribute VB_Name = "OS"
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    OSVersionInfoSize As Long
    MajorVersion As Long
    MinorVersion As Long
    BuildNumber As Long
    PlatformId As Long
    szCSDVersion As String * 128
End Type
Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type
Dim NotCompatible As Boolean

Public Function ItIsWin7() As Boolean
Dim OS As OSVERSIONINFO
Dim durum As Boolean
Dim version As String
    ItIsWin7 = False
    OS.OSVersionInfoSize = Len(OS)
    durum = GetVersionEx(OS)
    version = OS.PlatformId & "." & OS.MajorVersion & "." & OS.MinorVersion
    Select Case version
      Case "1.4.0"  'Win 95"
      Case "1.4.10" 'Win 98"
      Case "1.4.98" 'Win ME"
      Case "2.3.51" 'Win NT 3"
      Case "2.4.0"  'Win NT 4"
      Case "2.5.0"  'Win 2000"
      Case "2.5.1"  'Win XP"
      Case "2.6.0"  'Win Vista"
      Case "2.6.1": ItIsWin7 = True
      Case Else: NotCompatible = True
    End Select
End Function

