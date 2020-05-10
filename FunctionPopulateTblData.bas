Attribute VB_Name = "FunctionPopulateTblData"
Option Compare Database
Option Explicit
Option Base 0

Sub testPopulateTblData()
Dim CommodityID As Integer
Dim LocationID As Integer
Dim CompanyID As Integer
Dim CurrencyID As Integer
Dim UnitsID As Integer
Dim DeliveryStartDate As Date
Dim DeliveryEndDate As Date
Dim FieldFilter As String
Dim FuturesMarket As String
Dim FuturesMaturity As Date
Dim FuturesPrice As Double
Dim BasisPrice As Double
Dim NetPrice As Double
Dim EffectiveFrom As Date
Dim User As String
Dim DateTimeStamp As Date
'
CommodityID = 2
LocationID = 2
CompanyID = 2
CurrencyID = 2
UnitsID = 2
DeliveryStartDate = #1/5/2014#
DeliveryEndDate = #1/6/2014#
FuturesMarket = "F"
FuturesMaturity = #1/7/2014#
FuturesPrice = 3.2545
BasisPrice = -0.52
NetPrice = 252.25
EffectiveFrom = #1/8/2014#
User = "jddaniels"
'
    FieldFilter = "CommodityID"
    FieldFilter = FieldFilter & ", LocationID"
    FieldFilter = FieldFilter & ", CompanyID"
    FieldFilter = FieldFilter & ", CurrencyID"
    FieldFilter = FieldFilter & ", UnitsID"
    FieldFilter = FieldFilter & ", DeliveryStartDate"
    FieldFilter = FieldFilter & ", NetPrice"
    FieldFilter = FieldFilter & ", EffectiveFrom"
    FieldFilter = FieldFilter & ", User"
    'FieldFilter = FieldFilter & ", DeliveryEndDate:=DeliveryEndDate, BasisPrice:=BasisPrice"
    'FieldFilter = FieldFilter & ", FuturesMarket:=FuturesMarket"
    'FieldFilter = FieldFilter & ", FuturesMaturity:=FuturesMaturity"
    'FieldFilter = FieldFilter & ", FuturesPrice:=FuturesPrice"
    'FieldFilter = FieldFilter & ", BasisPrice:=BasisPrice"
'
    'If funPopulateTblData(FieldFilter) = True Then
    If funPopulateTblData(CommodityID, LocationID, CompanyID, CurrencyID, UnitsID, DeliveryStartDate, NetPrice, EffectiveFrom, User) = True Then
        MsgBox ("Success!")
    Else
        MsgBox ("Failure!")
    End If
'
End Sub

Function funPopulateTblData(CommodityID As Integer, LocationID As Integer, CompanyID As Integer, CurrencyID As Integer, UnitsID As Integer, DeliveryStartDate As Date, NetPrice As Double, EffectiveFrom As Date, User As String, Optional DeliveryEndDate As Variant, Optional FuturesMarket As Variant, Optional FuturesMaturity As Variant, Optional FuturesPrice As Variant, Optional BasisPrice As Variant) As Boolean

On Error GoTo funPopulateTblDataError
Dim sqlSELECT As String
Dim sqlUPDATE As String
Dim sqlWHERE As String
Dim sqlFROM As String
Dim SQL As String
Dim rstData As Recordset
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
'note: that futures market is a string, therefore it is not nessecary to test the datatype
'
'figure out the new DataID
    sqlSELECT = "SELECT Max(tblData.DataID) AS MaxOfDataID"
    sqlFROM = "FROM tblData"
    SQL = sqlSELECT & " " & sqlFROM & ";"
    Set rstData = CurrentDb.OpenRecordset(SQL)
    If IsNull(rstData!MaxofDataid) Then
        DataID = 1
    Else
        DataID = 1 + rstData!MaxofDataid
    End If
    Set rstData = Nothing
'figure out the previous record with this data, sand update the EffectiveTo field on that record (using the new EffectiveFrom)
    sqlUPDATE = "UPDATE tblData SET tblData.EffectiveTo = #" & EffectiveFrom & "#"
    sqlWHERE = "WHERE (((tblData.CommodityID)=" & CommodityID & ") "
    sqlWHERE = sqlWHERE & "AND ((tblData.LocationID)=" & LocationID & ") "
    sqlWHERE = sqlWHERE & "AND ((tblData.CompanyID)=" & CompanyID & ") "
    sqlWHERE = sqlWHERE & "AND ((tblData.CurrencyID)=" & CurrencyID & ") "
    sqlWHERE = sqlWHERE & "AND ((tblData.UnitsID)=" & UnitsID & ") "
    sqlWHERE = sqlWHERE & "AND ((tblData.DeliveryStartDate)=#" & DeliveryStartDate & "#) "
'
    If IsMissing(DeliveryEndDate) Then
        sqlWHERE = sqlWHERE & "AND ((tblData.DeliveryEndDate) Is Null) "
    Else
       sqlWHERE = sqlWHERE & "AND ((tblData.DeliveryEndDate)=#" & DeliveryEndDate & "#) "
    End If
    '
    If IsMissing(FuturesMarket) Then
        sqlWHERE = sqlWHERE & "AND ((tblData.FuturesMarket) Is Null) "
    Else
       sqlWHERE = sqlWHERE & "AND ((tblData.FuturesMarket)='F') "
    End If
    '
    If IsMissing(FuturesMaturity) Then
        sqlWHERE = sqlWHERE & "AND ((tblData.FuturesMaturity) Is Null) "
    Else
       sqlWHERE = sqlWHERE & "AND ((tblData.FuturesMaturity)=#1/7/2014#) "
    End If
'
    sqlWHERE = sqlWHERE & "AND ((tblData.EffectiveTo) Is Null))"
    SQL = sqlUPDATE & " " & sqlWHERE & ";"
    DoCmd.RunSQL SQL
    DoEvents
'
'set the Object "tblData" to the database table "tblData"
    Set rstData = CurrentDb.OpenRecordset("tblData")
    rstData.AddNew
    rstData!DataID = DataID
    rstData!CommodityID = CommodityID
    rstData!LocationID = LocationID
    rstData!CompanyID = CompanyID
    rstData!CurrencyID = CurrencyID
    rstData!UnitsID = UnitsID
    rstData!DeliveryStartDate = DeliveryStartDate
    If Not IsMissing(DeliveryEndDate) Then rstData!DeliveryEndDate = DeliveryEndDate ' test this variable since it is optional
    If Not IsMissing(FuturesMarket) Then: rstData!FuturesMarket = FuturesMarket ' test this variable since it is optional
    If Not IsMissing(FuturesMaturity) Then: rstData!FuturesMaturity = FuturesMaturity ' test this variable since it is optional
    If Not IsMissing(FuturesPrice) Then rstData!FuturesPrice = FuturesPrice ' test this variable since it is optional
    If Not IsMissing(BasisPrice) Then rstData!BasisPrice = BasisPrice ' test this variable since it is optional
    rstData!NetPrice = NetPrice
    rstData!EffectiveFrom = EffectiveFrom
    rstData!User = User
    rstData!DateTimeStamp = Now()
    rstData.Update
    Set rstData = Nothing
    funPopulateTblData = True
Exit Function
funPopulateTblDataError:
funPopulateTblData = False

End Function





