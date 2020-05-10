Attribute VB_Name = "FunctionYearFromInt"
Option Compare Database
Option Explicit
Option Base 0

Sub testYearFromInt()
Dim YearInt As Integer
YearInt = 14
Debug.Print funYearFromInt(YearInt)

End Sub

Function funYearFromInt(YearInt) As Integer
On Error GoTo errorYearFromInt

Dim x As Integer
    x = Len(YearInt)
    If x = 1 Then funYearFromInt = Int(Year(Now()) / 10) * 10 + YearInt
    If x = 2 Then funYearFromInt = Int(Year(Now()) / 100) * 100 + YearInt
    If x = 3 Then funYearFromInt = Int(Year(Now()) / 1000) * 1000 + YearInt
    If x = 4 Then funYearFromInt = YearInt

Exit Function
errorYearFromInt:

End Function
