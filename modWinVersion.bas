Attribute VB_Name = "modWinVersion"
'modWinVersion.bas

Option Explicit

Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer

Public Type OSVERSIONINFO
dwOSVersionInfoSize As Long
dwMajorVersion      As Long
dwMinorVersion      As Long
dwBuildNumber       As Long
dwPlatformId        As Long
szCSDVersion        As String * 128
End Type
    
Public Function GetWindowsVersion() As String
Dim getVersion     As String
Dim OSInfo         As OSVERSIONINFO
Dim Ret            As Integer
   
OSInfo.dwOSVersionInfoSize = 148
OSInfo.szCSDVersion = Space$(128)
Ret = GetVersionExA(OSInfo)

With OSInfo
Select Case .dwPlatformId

'Windows 95/98
Case 1
If .dwMinorVersion = 0 Then
getVersion = "Windows 95"
ElseIf .dwMinorVersion = 10 Then
getVersion = "Windows 98"
End If

'Windows NT 3.51/NT 4.0
Case 2
If .dwMajorVersion = 3 Then
getVersion = "Windows NT 3.51"
ElseIf .dwMajorVersion = 4 Then
getVersion = "Windows NT 4.0"
End If
      
'No version of Windows
Case Else
getVersion = "No versions of Windows detected"
End Select
End With
GetWindowsVersion = getVersion
End Function

'No support for NT 5.0 (Windows 2000) not supported, sorry.



