Attribute VB_Name = "FunctionDoubleFromString"
Option Compare Database
Option Explicit
Option Base 0

Sub testDoubleFromString()
Dim SomeString As String

SomeString = "154.25s"
Debug.Print funDoubleFromString(SomeString)

End Sub

Function funDoubleFromString(SomeString As String) As Double
On Error GoTo DoubleFromStringError
Dim TempChar As String

Dim strTemp As String
Dim x As Integer
SomeString = Trim(SomeString)
For x = 1 To Len(SomeString)
TempChar = Right(Left(SomeString, x), 1)
If InStr("0123456789.-", TempChar) > 0 Then
strTemp = strTemp & TempChar
End If

Next x
funDoubleFromString = CDbl(strTemp)

Exit Function
DoubleFromStringError:
Stop
End Function


