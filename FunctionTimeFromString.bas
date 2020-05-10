Attribute VB_Name = "FunctionTimeFromString"
Option Compare Database
Option Explicit
Option Base 0

Sub TestTimeFromString()
    Dim TimeString As String
    TimeString = "11:52 AM"
    Debug.Print TimeFromString(TimeString)
End Sub
Function TimeFromString(TimeString As String) As Double
    On Error GoTo TimeFromStringError
    Dim strDate As String
    strDate = "01 Jan 2014"
    Dim strDateTime As String
    strDateTime = strDate & " " & TimeString
    Dim DateTime As Date
    DateTime = CDate(strDateTime)
    TimeFromString = DateTime - Int(DateTime)
    Exit Function
TimeFromStringError:
    TimeFromString = 0
End Function
