Attribute VB_Name = "FunctionFuturesMonthFromFuturesCode"
Option Compare Database
Option Explicit
Option Base 0

Sub TestFuturesMonthFromFuturesCode()
Dim FunctionReturn As Variant
Dim FuturesCodeString As String
FuturesCodeString = "@RSX4"


 FunctionReturn = FuturesMonthFromFuturesCode(FuturesCodeString)
 Debug.Print FunctionReturn(0)
 Debug.Print FunctionReturn(1)
 
End Sub

Function FuturesMonthFromFuturesCode(FuturesCodeString As String) As Variant ' this will return a two dimential array, (Futures Month integer, original string with futurs month removed)
    Dim FuturesMonth As Integer
    Dim ReturnString As String
    Dim ReturnArray(2) As Variant
    
    
    
    If InStr(FuturesCodeString, "F") > 0 Then
        ReturnArray(0) = 1
        ReturnArray(1) = Replace(FuturesCodeString, "F", "")
        GoTo Finished
    End If
    
    If InStr(FuturesCodeString, "G") > 0 Then
        ReturnArray(0) = 2
        ReturnArray(1) = Replace(FuturesCodeString, "G", "")
        GoTo Finished
    End If
    
    If InStr(FuturesCodeString, "H") > 0 Then
        ReturnArray(0) = 3
        ReturnArray(1) = Replace(FuturesCodeString, "H", "")
        GoTo Finished
    End If
    
    If InStr(FuturesCodeString, "J") > 0 Then
        ReturnArray(0) = 4
        ReturnArray(1) = Replace(FuturesCodeString, "J", "")
        GoTo Finished
    End If
    
    If InStr(FuturesCodeString, "K") > 0 Then
        ReturnArray(0) = 5
        ReturnArray(1) = Replace(FuturesCodeString, "K", "")
        GoTo Finished
    End If
    
    If InStr(FuturesCodeString, "M") > 0 Then
        ReturnArray(0) = 6
        ReturnArray(1) = Replace(FuturesCodeString, "M", "")
        GoTo Finished
    End If
    
    If InStr(FuturesCodeString, "N") > 0 Then
        ReturnArray(0) = 7
        ReturnArray(1) = Replace(FuturesCodeString, "N", "")
        GoTo Finished
    End If
    
    If InStr(FuturesCodeString, "Q") > 0 Then
        ReturnArray(0) = 8
        ReturnArray(1) = Replace(FuturesCodeString, "Q", "")
        GoTo Finished
    End If
    
    If InStr(FuturesCodeString, "U") > 0 Then
        ReturnArray(0) = 9
        ReturnArray(1) = Replace(FuturesCodeString, "U", "")
        GoTo Finished
    End If
    
    If InStr(FuturesCodeString, "V") > 0 Then
        ReturnArray(0) = 10
        ReturnArray(1) = Replace(FuturesCodeString, "V", "")
        GoTo Finished
    End If
    
    If InStr(FuturesCodeString, "X") > 0 Then
        ReturnArray(0) = 11
        ReturnArray(1) = Replace(FuturesCodeString, "X", "")
        GoTo Finished
    End If
    
    If InStr(FuturesCodeString, "Z") > 0 Then
        ReturnArray(0) = 12
        ReturnArray(1) = Replace(FuturesCodeString, "Z", "")
        GoTo Finished
    End If
    
Finished:
FuturesMonthFromFuturesCode = ReturnArray

End Function

