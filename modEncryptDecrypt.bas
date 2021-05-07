Attribute VB_Name = "modEncryptDecrypt"
'modEncryptDecrypt.bas
Function encdec(inputstrinG As String) As String

If Len(inputstrinG) = 0 Then Exit Function

Dim p As String
Dim o As String
Dim k As String
Dim s As String
Dim tempstr As String

For i = 1 To Len(inputstrinG)
p = Mid$(inputstrinG, i, 1)

o = Asc(p)
k = o Xor 2
s = Chr$(k)
tempstr = tempstr & s

Next i

encdec = tempstr
End Function



