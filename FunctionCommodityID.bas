Attribute VB_Name = "FunctionCommodityID"
Option Compare Database
Option Explicit
Option Base 0

Sub testCommodityID()
Dim CommodityName As String
'
    CommodityName = "Random"
    MsgBox (funCommodityID(CommodityName))
'
End Sub

Function funCommodityID(CommodityName As String) As Integer
On Error GoTo CommodityIDError
Dim sqlSELECT As String
Dim sqlINSERT As String
Dim sqlFROM As String
Dim sqlWHERE As String
Dim SQL As String
Dim rstCommodityID As Recordset
Dim rstMaxCommodityID As Recordset
'
    DoCmd.SetWarnings False
'
' need to query the tblCommodities to find the Commodity ID for this commodity
    sqlSELECT = "SELECT tblCommodity.CommodityID"
    sqlFROM = "FROM tblCommodity"
    sqlWHERE = "WHERE (((tblCommodity.CommodityName)='" & CommodityName & "') AND ((tblCommodity.EffectiveTo) Is Null))"
    SQL = sqlSELECT & " " & sqlFROM & " " & sqlWHERE & ";"
    'Debug.Print SQL
    Set rstCommodityID = CurrentDb.OpenRecordset(SQL)
    'Debug.Print "rstCommodityID.RecordCount +-*/ " & rstCommodityID.RecordCount
'need to change tempRst to someting like CommodityIDRst
    If rstCommodityID.RecordCount = 0 Then
' there is no records with this commodity name, so we need to make a new record
' need to add a new record to the table with a new commodityID and the new CommodityName
' first figure out the max CommodityID, if there is none then start at 1
        sqlSELECT = "SELECT Max(tblCommodity.CommodityID) AS MaxOfCommodityID"
        sqlFROM = "FROM tblCommodity"
        SQL = sqlSELECT & " " & sqlFROM & ";"
        Debug.Print SQL
        Set rstMaxCommodityID = CurrentDb.OpenRecordset(SQL)
' need to change tblcommodity to something like NewCommodityIDRst
        If IsNull(rstMaxCommodityID!MaxofCommodityID) Then
' if tempRst!MaxofCommodityID is null then there is no data in the recordset so we can create the first record
            funCommodityID = 1
        Else
' if tempRst!MaxofCommodityID is a number, then we will use the next larger (available) number
            funCommodityID = 1 + rstMaxCommodityID!MaxofCommodityID
        End If
        Set rstMaxCommodityID = Nothing
' also need to add the new commodity to the table tblCommodity
        sqlINSERT = "INSERT INTO tblCommodity ( CommodityID, CommodityName, EffectiveFrom, [User], DateTimeStamp )"
        sqlSELECT = "SELECT " & funCommodityID & ", '" & CommodityName & "', #" & Now() & " #, '" & funFSOUserName & "', #" & Now() & "#"
        SQL = sqlINSERT & " " & sqlSELECT & ";"
        Debug.Print SQL
        DoCmd.RunSQL SQL
    Else
' if there is no error, then this commodity name exists in tblCommodity and we can use the CommodityID
        funCommodityID = rstCommodityID!CommodityID
        'Debug.Print "CommodityID +-*/ " & CommodityID
    End If
    Set rstCommodityID = Nothing
Exit Function
CommodityIDError:
Stop
End Function
