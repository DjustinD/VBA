Attribute VB_Name = "mdlYahooDownloader"
Option Compare Database
Option Explicit
' for use with YahooDownloader.accdb
' note: I did note write the important bits of this code...I found it online somewhere...
Sub Test()
    Dim X As Collection
    Dim Y As String
    Set X = FindCookieAndCrumb()
    Debug.Print X!cookie
    Debug.Print X!crumb
    Y = YahooRequest("AAPL", DateValue("31 Dec 2016"), DateValue("30 May 2017"), X)
    Debug.Print Y
    
End Sub


Function FindCookieAndCrumb() As Collection
Dim http    As MSXML2.XMLHTTP60 ' requires reference to library 'Microsoft XML, v 6.0'
Dim cookie  As String
Dim crumb   As String
Dim url     As String
Dim Pos1    As Long
Dim X       As String

    Set FindCookieAndCrumb = New Collection
    Set http = New MSXML2.ServerXMLHTTP60
    url = "https://finance.yahoo.com/quote/MSFT/history"
    http.Open "GET", url, False
    ' http.setProxy 2, "https=127.0.0.1:8888", ""
    ' http.setRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
    ' http.setRequestHeader "Accept-Encoding", "gzip, deflate, sdch, br"
    ' http.setRequestHeader "Accept-Language", "en-ZA,en-GB;q=0.8,en-US;q=0.6,en;q=0.4"
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36"
    http.send
    X = http.ResponseText
    Pos1 = InStr(X, "CrumbStore")
    X = Mid(X, Pos1, 44)
    X = Mid(X, 23, 44)
    Pos1 = InStr(X, """")
    X = Left(X, Pos1 - 1)
    FindCookieAndCrumb.Add X, "Crumb"
    X = http.GetResponseHeader("set-cookie")
    Pos1 = InStr(X, ";")
    X = Left(X, Pos1 - 1)
    FindCookieAndCrumb.Add X, "Cookie"

End Function

Function YahooRequest(ShareCode As String, StartDate As Date, EndDate As Date, CookieAndCrumb As Collection) As String
Dim http            As MSXML2.XMLHTTP60 ' requires reference to library "Microsoft XML, v6.0"
Dim cookie          As String
Dim crumb           As String
Dim url             As String
Dim UnixStartDate   As Long
Dim UnixEndDate     As Long
Dim BaseDate        As Date
    Set http = New MSXML2.ServerXMLHTTP60
    cookie = CookieAndCrumb!cookie
    crumb = CookieAndCrumb!crumb
    BaseDate = DateValue("1 Jan 1970")
    If StartDate = 0 Then StartDate = BaseDate
    UnixStartDate = (StartDate - BaseDate) * 86400
    UnixEndDate = (EndDate - BaseDate) * 86400
    url = "https://query1.finance.yahoo.com/v7/finance/download/" & ShareCode & "?period1=" & UnixStartDate & "&period2=" & UnixEndDate & "&interval=1d&events=history&crumb=" & crumb
    http.Open "GET", url, False
    http.setRequestHeader "Cookie", cookie
    http.send
    YahooRequest = http.ResponseText
End Function
Sub sDownload()
Dim X As Collection
Dim Y As String
Dim Z As Integer
Dim Ya() As String
Dim Yb() As String
Dim strTicker As String
Dim strTemp As String
Dim dtmStart As Date
Dim dtmEnd As Date
Dim dtmTemp As Date
Dim rstTemp As Recordset
Dim strSQL As String
    
    dtmTemp = Now()
    Debug.Print dtmTemp
    DoCmd.SetWarnings False
    Set X = FindCookieAndCrumb()
    Debug.Print X!cookie
    Debug.Print X!crumb
    dtmStart = DateValue("1 Jan 1970")
    Debug.Print dtmStart
    dtmEnd = DateValue("31 Dec 2019")
    Debug.Print dtmEnd
    
    strTemp = "SELECT tblTickers.strTicker, tblTickers.dtmValidFrom, tblTickers.dtmValidTo"
    strTemp = strTemp & " FROM tblTickers"
    strTemp = strTemp & " WHERE (((tblTickers.dtmValidFrom)<=#" & dtmStart & "#) AND ((tblTickers.dtmValidTo)>=#" & dtmEnd & "#)) OR (((tblTickers.dtmValidFrom)<=#" & dtmStart & "#) AND ((tblTickers.dtmValidTo) Is Null)) OR (((tblTickers.dtmValidFrom) Is Null) AND ((tblTickers.dtmValidTo) Is Null));"
    'Debug.Print strTemp
    Set rstTemp = CurrentDb.OpenRecordset(strTemp)
    Debug.Print rstTemp.RecordCount
    rstTemp.MoveFirst
    Do Until rstTemp.EOF
        Y = YahooRequest(rstTemp!strTicker, dtmStart, dtmEnd, X)
        Ya = Split(Y, Chr(10))
        For Z = 1 To UBound(Ya) - 1
            Yb = Split(Ya(Z), Chr(44))
            If Not Yb(1) = "null" Then
                strSQL = "INSERT INTO tblHistoricalPrices (strTicker, dtmDate, curOpen, curHigh, curLow, curClose, curAdj_Close, intVolume) VALUES ('" & rstTemp!strTicker & "', #" & Yb(0) & "#, " & Round(Yb(1), 2) & ", " & Round(Yb(2), 2) & ", " & Round(Yb(3), 2) & ", " & Round(Yb(4), 2) & ", " & Round(Yb(5), 2) & ", " & Yb(6) & ");"
                DoCmd.RunSQL strSQL
            End If
        Next Z
        rstTemp.MoveNext
        'Stop
    Loop
    Set rstTemp = Nothing
Debug.Print "done"
DoCmd.SetWarnings True
Debug.Print Now()
Debug.Print "time elapsed " & Now() - dtmTemp
End Sub
Sub testing()
Dim rstTemp As Recordset
Dim strTemp As String
Dim dtmStart As Date
Dim dtmEnd As Date

    dtmStart = DateValue("1 Jan 1972")
    dtmEnd = DateValue("31 Dec 1982")
    
    strTemp = "SELECT tblTickers.strTicker, tblTickers.dtmValidFrom, tblTickers.dtmValidTo"
    strTemp = strTemp & " FROM tblTickers"
    strTemp = strTemp & " WHERE (((tblTickers.dtmValidFrom)<=#" & dtmStart & "#) AND ((tblTickers.dtmValidTo)>=#" & dtmEnd & "#)) OR (((tblTickers.dtmValidFrom)<=#" & dtmStart & "#) AND ((tblTickers.dtmValidTo) Is Null)) OR (((tblTickers.dtmValidFrom) Is Null) AND ((tblTickers.dtmValidTo) Is Null));"
    Debug.Print strTemp
    Set rstTemp = CurrentDb.OpenRecordset(strTemp)
    Stop
    Debug.Print rstTemp.RecordCount
    rstTemp.MoveFirst
    Do Until rstTemp.EOF
    Debug.Print rstTemp!strTicker
    rstTemp.MoveNext
    Loop
    Set rstTemp = Nothing
End Sub
