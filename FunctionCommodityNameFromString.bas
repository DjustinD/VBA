Attribute VB_Name = "FunctionCommodityNameFromString"
Option Compare Database

Sub testCommodityNameFromString()

Dim CommodityName As String
CommodityName = "CWRS 13.5 delivered russel"
Debug.Print funCommodityNameFromString(CommodityName)

End Sub


Function funCommodityNameFromString(CommodityName As String) As String
On Error GoTo ErrorCommodityNameFromString
Dim ReturnStr As String

ChooseClass:
    If InStr(CommodityName, "CWRS") > 0 Then: ReturnStr = "CWRS": GoTo FindProteinPercent
    If InStr(CommodityName, "CPS") > 0 Then: ReturnStr = "CPSR": GoTo FindProteinPercent
    
    MsgBox ("Unable to find Class!")
    Stop
FindProteinPercent:
    If InStr(CommodityName, "15") > 0 Then: ReturnStr = ReturnStr & "15": GoTo FindProteinDecimal
    If InStr(CommodityName, "14") > 0 Then ReturnStr = ReturnStr & "14": GoTo FindProteinDecimal
    If InStr(CommodityName, "13") > 0 Then ReturnStr = ReturnStr & "13": GoTo FindProteinDecimal
    If InStr(CommodityName, "12") > 0 Then ReturnStr = ReturnStr & "12": GoTo FindProteinDecimal
    If InStr(CommodityName, "11") > 0 Then ReturnStr = ReturnStr & "11": GoTo FindProteinDecimal
    If InStr(CommodityName, "10") > 0 Then ReturnStr = ReturnStr & "10": GoTo FindProteinDecimal
    
FindProteinDecimal:
    If InStr(CommodityName, ".5") > 0 Then: ReturnStr = ReturnStr & ".5"




funCommodityNameFromString = ReturnStr
Exit Function
ErrorCommodityNameFromString:
DoEvents
Stop
DoEvents
End Function
