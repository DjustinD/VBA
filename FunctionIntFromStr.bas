Attribute VB_Name = "FunctionIntFromStr"
Option Compare Database
Option Explicit
Option Base 0

Sub testIntFromStr()
    Dim SomeStr As String
    SomeStr = "abc */- 987 "
    'SomeStr = "kljlaui asd /*-+ "
    Debug.Print Chr(10) & funIntFromStr(SomeStr) & Chr(10)
End Sub


Function funIntFromStr(SomeStr As String) As Integer
On Error GoTo errorIntFromStr
Dim strTemp As String
Dim x As Integer
SomeStr = Trim(SomeStr)
For x = 1 To Len(SomeStr)
If IsNumeric(Right(Left(SomeStr, x), 1)) Then
strTemp = strTemp & Right(Left(SomeStr, x), 1)
End If
Next x
If Len(strTemp) = 0 Then
    ' simply exit the function and the return value will be zero
    Exit Function
End If
funIntFromStr = CInt(strTemp)
Exit Function
errorIntFromStr:
Stop
End Function
