Attribute VB_Name = "modUserName"
'modUserName.bas

Option Explicit

Private Declare Function GetUserName Lib "advapi32.dll" _
Alias "GetUserNameA" (ByVal lpBuffer As String, _
nSize As Long) As Long

Public Function UserName() As String

Dim llReturn    As Long
Dim lsUserName  As String
Dim lsBuffer    As String
    
lsUserName = ""
lsBuffer = Space$(255)
llReturn = GetUserName(lsBuffer, 255)
    
If llReturn Then
lsUserName = Left$(lsBuffer, InStr(lsBuffer, Chr(0)) - 1)
End If
    
UserName = lsUserName
End Function


