Attribute VB_Name = "FunctionFuturesMarketFromCommodityName"
Option Compare Database

Sub testFuturesMarketFromCommodityName()
Dim CommodityName As String
CommodityName = "1CWRS13.5"
Debug.Print funFuturesMarketFromCommodityName(CommodityName)


End Sub


Function funFuturesMarketFromCommodityName(CommodityName As String) As String



If InStr(CommodityName, "CWRS") > 0 Then: funFuturesMarketFromCommodityName = "MW": Exit Function
If InStr(CommodityName, "CWRW") > 0 Then: funFuturesMarketFromCommodityName = "KW": Exit Function

If InStr(CommodityName, "Canola") > 0 Then: funFuturesMarketFromCommodityName = "RS": Exit Function
If InStr(CommodityName, "Nexera") > 0 Then: funFuturesMarketFromCommodityName = "RS": Exit Function

If InStr(CommodityName, "Corn") > 0 Then: funFuturesMarketFromCommodityName = "C": Exit Function
If InStr(CommodityName, "Soybean") > 0 Then: funFuturesMarketFromCommodityName = "S": Exit Function
If InStr(CommodityName, "HRS") > 0 Then: funFuturesMarketFromCommodityName = "MW": Exit Function
If InStr(CommodityName, "HRW") > 0 Then: funFuturesMarketFromCommodityName = "KW": Exit Function
If InStr(CommodityName, "Wheat") > 0 Then: funFuturesMarketFromCommodityName = "W": Exit Function

If InStr(CommodityName, "Oats") > 0 Then: funFuturesMarketFromCommodityName = "O": Exit Function

MsgBox ("Can't match FuturesMarket to CommodityName")
Stop

End Function
