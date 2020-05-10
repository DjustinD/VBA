Attribute VB_Name = "FunctionDateFromString"
Option Compare Database
Option Explicit

Sub testDateFromString()
Dim strDate As String
strDate = "january 18, 2008"

Debug.Print funDateFromString(strDate, 1, 1, 1)

End Sub

Function funDateFromString(strDate As String, intGuessDay As Integer, intGuessMonth As Integer, intGuessYear As Integer) As Date
On Error GoTo DateFromStringERROR

Dim intReturnDay As Integer
Dim intReturnMonth As Integer
Dim intReturnYear As Integer

If IsDate(strDate) Then: funDateFromString = strDate: Exit Function

strDate = Replace(strDate, " ", "")
'If IsNumeric(strstrDate) Then: funDateFromString = CDate(strDate): exit function
If InStr(strDate, "spot") = True Then: funDateFromString = Int(Now()): Exit Function




'find day
intReturnDay = intGuessDay
FinishedDay:


'Find Month
If InStr(strDate, "Jan") > 0 Then: intReturnMonth = 1: GoTo FinishedMonth
If InStr(strDate, "January") > 0 Then: intReturnMonth = 1: GoTo FinishedMonth
If InStr(strDate, "Feb") > 0 Then: intReturnMonth = 2: GoTo FinishedMonth
If InStr(strDate, "February") > 0 Then: intReturnMonth = 2: GoTo FinishedMonth
If InStr(strDate, "Mar") > 0 Then: intReturnMonth = 3: GoTo FinishedMonth
If InStr(strDate, "March") > 0 Then: intReturnMonth = 3: GoTo FinishedMonth
If InStr(1, strDate, "Apr") > 0 Then: intReturnMonth = 4: GoTo FinishedMonth
If InStr(1, strDate, "April") > 0 Then: intReturnMonth = 4: GoTo FinishedMonth
If InStr(strDate, "May") > 0 Then: intReturnMonth = 5: GoTo FinishedMonth
If InStr(strDate, "Jun") > 0 Then: intReturnMonth = 6: GoTo FinishedMonth
If InStr(strDate, "June") > 0 Then: intReturnMonth = 6: GoTo FinishedMonth
If InStr(strDate, "Jul") > 0 Then: intReturnMonth = 7: GoTo FinishedMonth
If InStr(strDate, "July") > 0 Then: intReturnMonth = 7: GoTo FinishedMonth
If InStr(strDate, "Aug") > 0 Then: intReturnMonth = 8: GoTo FinishedMonth
If InStr(strDate, "August") > 0 Then: intReturnMonth = 8: GoTo FinishedMonth
If InStr(strDate, "Sep") > 0 Then: intReturnMonth = 9: GoTo FinishedMonth
If InStr(strDate, "Sept") > 0 Then: intReturnMonth = 9: GoTo FinishedMonth
If InStr(strDate, "September") > 0 Then: intReturnMonth = 9: GoTo FinishedMonth
If InStr(strDate, "Oct") > 0 Then: intReturnMonth = 10: GoTo FinishedMonth
If InStr(strDate, "October") > 0 Then: intReturnMonth = 10: GoTo FinishedMonth
If InStr(strDate, "Nov") > 0 Then: intReturnMonth = 11: GoTo FinishedMonth
If InStr(strDate, "November") > 0 Then: intReturnMonth = 11: GoTo FinishedMonth
If InStr(strDate, "Dec") > 0 Then: intReturnMonth = 12: GoTo FinishedMonth
If InStr(strDate, "December") > 0 Then: intReturnMonth = 12: GoTo FinishedMonth

intReturnMonth = intGuessMonth
FinishedMonth:
'Find Year



intReturnYear = intGuessYear
FinishedYear:
funDateFromString = DateSerial(intReturnYear, intReturnMonth, intReturnDay)
Exit Function
DateFromStringERROR:
Stop
End Function

