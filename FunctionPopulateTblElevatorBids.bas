Attribute VB_Name = "FunctionPopulateTblElevatorBids"
Option Compare Database
Option Explicit
Option Base 0

Sub testPopulateTblElevatorBids()
Dim CommodityString As String
Dim LocationString As String
Dim CompanyString As String
Dim CurrencyString As String
Dim UnitString As String
Dim DeliveryString As String
Dim DeliveryStartDate As Date
Dim DeliveryEndDate As Date
Dim FieldFilter As String
Dim FuturesMarket As String
Dim FuturesMaturity As Date
Dim FuturesPrice As Double
Dim BasisPrice As Double
Dim NetPrice As Double
Dim EffectiveTime As Date
Dim User As String
Dim URLID As Integer
Dim DateTimeStamp As Date
'
CommodityString = "SomeGrain"
LocationString = "SomeWhere"
CompanyString = "SomeCompany"
CurrencyString = "CAD"
UnitString = "$/Tonne"
DeliveryString = "SomeString"
DeliveryStartDate = #1/5/2014#
DeliveryEndDate = #1/6/2014#
FuturesMarket = "F"
FuturesMaturity = #1/7/2014#
FuturesPrice = 3.2545
BasisPrice = -0.52
NetPrice = 252.25
EffectiveTime = #1/8/2014#
User = "jddaniels"
URLID = 125
'
    If funPopulateTblElevatorBids(CommodityString, LocationString, CompanyString, CurrencyString, DeliveryStartDate, NetPrice, EffectiveTime, User, URLID, DeliveryString:=DeliveryString) = True Then
        MsgBox ("Success!")
    Else
        MsgBox ("Failure!")
    End If
'
End Sub

Function funPopulateTblElevatorBids(CommodityString As String, LocationString As String, CompanyString As String, CurrencyString As String, DeliveryStartDate As Date, NetPrice As Double, EffectiveTime As Date, User As String, URLID As Integer, Optional DeliveryEndDate As Variant, Optional FuturesMarket As Variant, Optional FuturesMaturity As Variant, Optional FuturesPrice As Variant, Optional BasisPrice As Variant, Optional DeliveryString As String) As Boolean

On Error GoTo funPopulateTblElevatorBidsError
Dim sqlSELECT As String
Dim sqlUPDATE As String
Dim sqlFROM As String
Dim SQL As String
Dim rstData As Recordset
'Dim Data As Integer
Dim DataID As Integer

'confirm that futures maturity is either missing or is a date ' this need to be done since FuturesMaturity is optional and a variant
    If Not IsMissing(FuturesMaturity) Then
        If Not IsDate(FuturesMaturity) Then
        MsgBox ("FuturesMaturity Variable is not a date datatype")
        End If
    End If
'
'confirm BasisPrice is either missing or is a date ' this is needed since BasisPrice is optional and a variant
    If Not IsMissing(BasisPrice) Then
        If Not IsNumeric(BasisPrice) Then
        MsgBox ("BasisPrice Variable is not a numeric datatype")
        End If
    End If
'
'confirm FuturesPrice is either missing or is a date ' this is needed since FuturesPrice is optional and a variant
    If Not IsMissing(FuturesPrice) Then
        If Not IsNumeric(FuturesPrice) Then
        MsgBox ("FuturesPrice Variable is not a numeric datatype")
        End If
    End If
'
'note: that futures market is a string, therefore it is not necessary to test the datatype
'note: that DeliveryString is a string, therefore it is not necessary to test the datatype

'
'figure out the new DataID
    sqlSELECT = "SELECT Max(tblElevatorBids.DataID) AS MaxOfData"
    sqlFROM = "FROM tblElevatorBids"
    SQL = sqlSELECT & " " & sqlFROM & ";"
'    GoTo SkipTo
    Set rstData = CurrentDb.OpenRecordset(SQL)
    'If rstData.RecordCount = 0 Then
    If IsNull(rstData!MaxofData) Then
        DataID = 1
    Else
        DataID = 1 + rstData!MaxofData
    End If
    Set rstData = Nothing
'
'set the Object "tblElevatorBids" to the database table "tblElevatorBids"
    Set rstData = CurrentDb.OpenRecordset("tblElevatorBids")
    rstData.AddNew
    rstData!DataID = DataID
    rstData!URLID = URLID
    rstData!CommodityString = CommodityString
    rstData!LocationString = LocationString
    rstData!CompanyString = CompanyString
    rstData!CurrencyString = CurrencyString
    rstData!DeliveryStartDate = DeliveryStartDate
    If Not IsMissing(DeliveryEndDate) Then rstData!DeliveryEndDate = DeliveryEndDate ' test this variable since it is optional
    If Not IsMissing(FuturesMarket) Then: rstData!FuturesMarket = FuturesMarket ' test this variable since it is optional
    If Not IsMissing(FuturesMaturity) Then: rstData!FuturesMaturity = FuturesMaturity ' test this variable since it is optional
    If Not IsMissing(FuturesPrice) Then rstData!FuturesPrice = FuturesPrice ' test this variable since it is optional
    If Not IsMissing(BasisPrice) Then rstData!BasisPrice = BasisPrice ' test this variable since it is optional
    If Not IsMissing(DeliveryString) Then: rstData!DeliveryString = DeliveryString ' test this variable since it is optional
    
    rstData!NetPrice = NetPrice
    rstData!EffectiveTime = EffectiveTime
    rstData!User = User
    rstData!DateTimeStamp = Now()
    rstData.Update
    Set rstData = Nothing
SkipTo:
    funPopulateTblElevatorBids = True
Exit Function
funPopulateTblElevatorBidsError:
funPopulateTblElevatorBids = False

End Function


