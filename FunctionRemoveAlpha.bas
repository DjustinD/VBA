Attribute VB_Name = "FunctionRemoveAlpha"
Option Compare Database
Option Explicit

Sub testRemoveAlpha()
Dim strVariable As String

strVariable = "a76"
Debug.Print funRemoveAlpha(strVariable)

End Sub
Function funRemoveAlpha(strVariable As String) As String
Dim strRemove As String
Dim intX As Integer

strRemove = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
For intX = 1 To Len(strRemove)

strVariable = Replace(strVariable, Mid(strRemove, intX, 1), "")
Next intX
funRemoveAlpha = strVariable

End Function

