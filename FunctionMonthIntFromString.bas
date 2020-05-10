Attribute VB_Name = "FunctionMonthIntFromString"
Option Compare Database
Option Explicit
Option Base 0

Sub testMonthIntFromStr()
    Dim MonthStr As String
    MonthStr = "Januar 08, 2014"
    Debug.Print funMonthIntFromStr(MonthStr)
End Sub
Function funMonthIntFromStr(MonthStr As String) As Integer
On Error GoTo errorMonthIntFromStr

If IsDate(MonthStr) Then: funMonthIntFromStr = Month(MonthStr): Exit Function

If InStr(MonthStr, "January") > 0 Then: funMonthIntFromStr = 1: Exit Function
If InStr(MonthStr, "February") > 0 Then: funMonthIntFromStr = 2: Exit Function
If InStr(MonthStr, "March") > 0 Then: funMonthIntFromStr = 3: Exit Function
If InStr(MonthStr, "April") > 0 Then: funMonthIntFromStr = 4: Exit Function
If InStr(MonthStr, "May") > 0 Then: funMonthIntFromStr = 5: Exit Function
If InStr(MonthStr, "June") > 0 Then: funMonthIntFromStr = 6: Exit Function
If InStr(MonthStr, "July") > 0 Then: funMonthIntFromStr = 7: Exit Function
If InStr(MonthStr, "August") > 0 Then: funMonthIntFromStr = 8: Exit Function
If InStr(MonthStr, "September") > 0 Then: funMonthIntFromStr = 9: Exit Function
If InStr(MonthStr, "October") > 0 Then: funMonthIntFromStr = 10: Exit Function
If InStr(MonthStr, "November") > 0 Then: funMonthIntFromStr = 11: Exit Function
If InStr(MonthStr, "December") > 0 Then: funMonthIntFromStr = 12: Exit Function

If InStr(MonthStr, "spot") > 0 Then: funMonthIntFromStr = Month(Now()): Exit Function
If InStr(MonthStr, "mjj") > 0 Then: funMonthIntFromStr = 5: Exit Function
If InStr(MonthStr, "S/O/N") > 0 Then: funMonthIntFromStr = 6: Exit Function
If InStr(MonthStr, "SON") > 0 Then: funMonthIntFromStr = 9: Exit Function
If InStr(MonthStr, "JFM") > 0 Then: funMonthIntFromStr = 1: Exit Function
If InStr(MonthStr, "A/S") > 0 Then: funMonthIntFromStr = 8: Exit Function
If InStr(MonthStr, "fall") > 0 Then: funMonthIntFromStr = 9: Exit Function
If InStr(MonthStr, "j/j") > 0 Then: funMonthIntFromStr = 6: Exit Function
If InStr(MonthStr, "jj") > 0 Then: funMonthIntFromStr = 6: Exit Function
If InStr(MonthStr, "o/n") > 0 Then: funMonthIntFromStr = 10: Exit Function
If InStr(MonthStr, "m/j") > 0 Then: funMonthIntFromStr = 5: Exit Function

If InStr(MonthStr, "Sept") > 0 Then: funMonthIntFromStr = 9: Exit Function

If InStr(MonthStr, "Jan") > 0 Then: funMonthIntFromStr = 1: Exit Function
If InStr(MonthStr, "Feb") > 0 Then: funMonthIntFromStr = 2: Exit Function
If InStr(MonthStr, "Mar") > 0 Then: funMonthIntFromStr = 3: Exit Function
If InStr(MonthStr, "Apr") > 0 Then: funMonthIntFromStr = 4: Exit Function
If InStr(MonthStr, "Jun") > 0 Then: funMonthIntFromStr = 6: Exit Function
If InStr(MonthStr, "Jul") > 0 Then: funMonthIntFromStr = 7: Exit Function
If InStr(MonthStr, "Aug") > 0 Then: funMonthIntFromStr = 8: Exit Function
If InStr(MonthStr, "Sep") > 0 Then: funMonthIntFromStr = 9: Exit Function
If InStr(MonthStr, "Oct") > 0 Then: funMonthIntFromStr = 10: Exit Function
If InStr(MonthStr, "Nov") > 0 Then: funMonthIntFromStr = 11: Exit Function
If InStr(MonthStr, "Dec") > 0 Then: funMonthIntFromStr = 12: Exit Function

If InStr(MonthStr, "Ja") > 0 Then: funMonthIntFromStr = 1: Exit Function
If InStr(MonthStr, "Fe") > 0 Then: funMonthIntFromStr = 2: Exit Function
If InStr(MonthStr, "Mr") > 0 Then: funMonthIntFromStr = 3: Exit Function
If InStr(MonthStr, "Ar") > 0 Then: funMonthIntFromStr = 4: Exit Function
If InStr(MonthStr, "Ap") > 0 Then: funMonthIntFromStr = 4: Exit Function
If InStr(MonthStr, "My") > 0 Then: funMonthIntFromStr = 5: Exit Function
If InStr(MonthStr, "Ju") > 0 Then: funMonthIntFromStr = 6: Exit Function
If InStr(MonthStr, "Jn") > 0 Then: funMonthIntFromStr = 6: Exit Function
If InStr(MonthStr, "Jl") > 0 Then: funMonthIntFromStr = 7: Exit Function
If InStr(MonthStr, "Au") > 0 Then: funMonthIntFromStr = 8: Exit Function
If InStr(MonthStr, "Ag") > 0 Then: funMonthIntFromStr = 8: Exit Function
If InStr(MonthStr, "Sp") > 0 Then: funMonthIntFromStr = 9: Exit Function
If InStr(MonthStr, "Se") > 0 Then: funMonthIntFromStr = 9: Exit Function
If InStr(MonthStr, "Oc") > 0 Then: funMonthIntFromStr = 10: Exit Function
If InStr(MonthStr, "Nv") > 0 Then: funMonthIntFromStr = 11: Exit Function
If InStr(MonthStr, "Dc") > 0 Then: funMonthIntFromStr = 12: Exit Function

Stop
funMonthIntFromStr = Month(Now() + 5)
Exit Function
errorMonthIntFromStr:
Stop
End Function


