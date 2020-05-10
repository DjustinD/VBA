Attribute VB_Name = "FunctionMaxOf"
Option Explicit
Option Compare Database

Sub testMaxOf()
Dim dbl1 As Double
Dim dbl2 As Double
dbl1 = 1.234
dbl2 = 2.432
Debug.Print funMaxOf(dbl1, dbl2)

End Sub


Function funMaxOf(dbl1 As Double, dbl2 As Double) As Double
On Error GoTo errorMaxOf

    If dbl1 > dbl2 Then
        funMaxOf = dbl1
    Else
        funMaxOf = dbl2
    End If

Exit Function
errorMaxOf:
Stop

End Function
