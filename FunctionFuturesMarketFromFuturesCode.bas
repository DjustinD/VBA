Attribute VB_Name = "FunctionFuturesMarketFromFuturesCode"
Option Compare Database
Option Explicit
Option Base 0

Sub TestFuturesMarketFromFuturesCode()
    Dim FuturesCodeString As String
    Dim ReturnArray
    
    FuturesCodeString = " 876 */-+ Wheat)"
    ReturnArray = FuturesMarketFromFuturesCode(FuturesCodeString)
    Debug.Print ReturnArray(0)
    Debug.Print ReturnArray(1)
End Sub

Function FuturesMarketFromFuturesCode(FuturesCodeString As String) As Variant
    Dim ReturnArray(2)
    
    If InStr(FuturesCodeString, "KCBT") > 0 Then
        ReturnArray(0) = "KW"
        ReturnArray(1) = Replace(FuturesCodeString, "KCBT", "")
        GoTo Finished
    End If
    
    If InStr(FuturesCodeString, "Spring Wheat") > 0 Then
        ReturnArray(0) = "MW"
        ReturnArray(1) = Replace(FuturesCodeString, "Spring Wheat", "")
        GoTo Finished
    End If
    
    If InStr(FuturesCodeString, "Wheat") > 0 Then
        ReturnArray(0) = "W"
        ReturnArray(1) = Replace(FuturesCodeString, "Wheat", "")
        GoTo Finished
    End If
    
    If InStr(FuturesCodeString, "Corn") > 0 Then
        ReturnArray(0) = "C"
        ReturnArray(1) = Replace(FuturesCodeString, "Corn", "")
        GoTo Finished
    End If
    
    If InStr(FuturesCodeString, "Soybeans") > 0 Then
        ReturnArray(0) = "S"
        ReturnArray(1) = Replace(FuturesCodeString, "Soybeans", "")
        GoTo Finished
    End If
    
    If InStr(FuturesCodeString, "Soy") > 0 Then
        ReturnArray(0) = "S"
        ReturnArray(1) = Replace(FuturesCodeString, "Soy", "")
        GoTo Finished
    End If
    
    If InStr(FuturesCodeString, "@RS") > 0 Then
        ReturnArray(0) = "RS"
        ReturnArray(1) = Replace(FuturesCodeString, "@RS", "")
        GoTo Finished
    End If
    
    If InStr(FuturesCodeString, "RS") > 0 Then
        ReturnArray(0) = "RS"
        ReturnArray(1) = Replace(FuturesCodeString, "RS", "")
        GoTo Finished
    End If
    
    If InStr(FuturesCodeString, "@O") > 0 Then
        ReturnArray(0) = "O"
        ReturnArray(1) = Replace(FuturesCodeString, "@O", "")
        GoTo Finished
    End If
    
    If InStr(FuturesCodeString, "@MW") > 0 Then
        ReturnArray(0) = "MW"
        ReturnArray(1) = Replace(FuturesCodeString, "@MW", "")
        GoTo Finished
    End If
    If InStr(FuturesCodeString, "MW") > 0 Then
        ReturnArray(0) = "MW"
        ReturnArray(1) = Replace(FuturesCodeString, "MW", "")
        GoTo Finished
    End If
    
    If InStr(FuturesCodeString, "KW") > 0 Then
        ReturnArray(0) = "KW"
        ReturnArray(1) = Replace(FuturesCodeString, "KW", "")
        GoTo Finished
    End If
    
    If InStr(FuturesCodeString, "KE") > 0 Then
        ReturnArray(0) = "KW"
        ReturnArray(1) = Replace(FuturesCodeString, "KE", "")
        GoTo Finished
    End If
    '
    If InStr(FuturesCodeString, "@W") > 0 Then
        ReturnArray(0) = "W"
        ReturnArray(1) = Replace(FuturesCodeString, "@W", "")
        GoTo Finished
    End If
    '
    If InStr(FuturesCodeString, "@C") > 0 Then
        ReturnArray(0) = "C"
        ReturnArray(1) = Replace(FuturesCodeString, "@C", "")
        GoTo Finished
    End If



    
Finished:

FuturesMarketFromFuturesCode = ReturnArray

End Function
