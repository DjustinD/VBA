Attribute VB_Name = "FunctionNewCropDelivery"
Option Compare Database
Option Explicit
Option Base 0

Sub testNewCropDelivery()
Dim CommodityString As String
CommodityString = "sdfd8798 wheat asdfa"
Debug.Print funNewCropDelivery(CommodityString)
End Sub

Function funNewCropDelivery(CommodityString As String) As Date


    If InStr(CommodityString, "Wheat") > 0 Then
        If InStr(CommodityString, "Winter") > 0 Then
            funNewCropDelivery = DateSerial(Year(Now()), 7, 1)
            GoTo Finished
        End If
        If InStr(CommodityString, "Spring") > 0 Then
            funNewCropDelivery = DateSerial(Year(Now()), 9, 1)
            GoTo Finished
        End If
        funNewCropDelivery = DateSerial(Year(Now()), 7, 1)
        GoTo Finished
    End If
    '
    If InStr(CommodityString, "Corn") > 0 Then
    funNewCropDelivery = DateSerial(Year(Now()), 9, 1)
        GoTo Finished
    End If
    '
    If InStr(CommodityString, "Soybeans") > 0 Then
    funNewCropDelivery = DateSerial(Year(Now()), 9, 1)
        GoTo Finished
    End If
    '
    If InStr(CommodityString, "Canola") > 0 Then
    funNewCropDelivery = DateSerial(Year(Now()), 9, 1)
        GoTo Finished
    End If
    If InStr(CommodityString, "Barley") > 0 Then
    funNewCropDelivery = DateSerial(Year(Now()), 9, 1)
        GoTo Finished
    End If
    If InStr(CommodityString, "Sorghum") > 0 Then
    funNewCropDelivery = DateSerial(Year(Now()), 9, 1)
        GoTo Finished
    End If
    If InStr(CommodityString, "Flax") > 0 Then
    funNewCropDelivery = DateSerial(Year(Now()), 9, 1)
        GoTo Finished
    End If
    If InStr(CommodityString, "Milo") > 0 Then
    funNewCropDelivery = DateSerial(Year(Now()), 9, 1)
        GoTo Finished
    End If
    If InStr(CommodityString, "Oats") > 0 Then
    funNewCropDelivery = DateSerial(Year(Now()), 9, 1)
        GoTo Finished
    End If

MsgBox ("can't find the new crop date for this comodity")








Finished:

End Function
