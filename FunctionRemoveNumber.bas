Attribute VB_Name = "FunctionRemoveNumber"
Option Compare Database
Option Explicit

Sub testRemoveNumber()
Dim strVariable As String

strVariable = "a76"
Debug.Print funRemoveNumber(strVariable)

End Sub
Function funRemoveNumber(strVariable As String) As String
Dim strRemove As String
Dim intX As Integer

strRemove = "0123456789"
For intX = 1 To Len(strRemove)

strVariable = Replace(strVariable, Mid(strRemove, intX, 1), "")
Next intX
funRemoveNumber = strVariable

End Function


