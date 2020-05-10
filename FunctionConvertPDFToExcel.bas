Attribute VB_Name = "FunctionConvertPDFToExcel"
Option Compare Database
Option Explicit
Option Base 0

Sub TestConvertPDFToExcel()
Dim PDFFilename As String
PDFFilename = "I:\MktRskRsch\Restricted\MktRsch\RestrictedPrices\TempExcelFiles\222-20140424111240.pdf"
Dim ExcelFilename As String
ExcelFilename = "I:\MktRskRsch\Restricted\MktRsch\RestrictedPrices\TempExcelFiles\222-20140424111240.xlsx"
'
If funConvertPDFToExcel(PDFFilename, ExcelFilename) = True Then
    MsgBox ("Success")
Else
    MsgBox ("Failure")
End If
'
End Sub

Function funConvertPDFToExcel(PDFFilename As String, ExcelFilename As String) As Boolean
On Error GoTo ConvertPDFToExcelError
' NOTE, need to have Adobe Pro installed for this to work.
Dim WaitUntil As Long

Shell "C:\Program Files (x86)\Adobe\Acrobat 9.0\Acrobat\Acrobat.exe", vbMaximizedFocus

WaitUntil = Int(Timer + 5)
Do Until Timer > WaitUntil: DoEvents: Loop
SendKeys "^o", True ' ctrl o to open the open file dialogue
WaitUntil = Int(Timer + 5)
Do Until Timer > WaitUntil: DoEvents: Loop
SendKeys PDFFilename, True ' this will type the fill path and file name of the file to open
' for example I:\MktRskRsch\Restricted\MktRsch\AutoFiles\AutoWITPricesFromWeb\WIT-Weyburn, SK-20130729101928.pdf
WaitUntil = Int(Timer + 5)
Do Until Timer > WaitUntil: DoEvents: Loop
SendKeys "~", True ' ctrl o to click the open button
WaitUntil = Int(Timer + 5)
Do Until Timer > WaitUntil: DoEvents: Loop
SendKeys ("^(+s)") ' save as
WaitUntil = Int(Timer + 5)
Do Until Timer > WaitUntil: DoEvents: Loop
SendKeys (ExcelFilename) ' use the same file name as the pdf
WaitUntil = Int(Timer + 5)
Do Until Timer > WaitUntil: DoEvents: Loop
SendKeys ("{tab}") ' tab down to save as file type
WaitUntil = Int(Timer + 5)
Do Until Timer > WaitUntil: DoEvents: Loop
SendKeys ("t") ' start typing "t" that will select tab deliminated xml file type
WaitUntil = Int(Timer + 5)
Do Until Timer > WaitUntil: DoEvents: Loop
SendKeys ("%s") ' save
WaitUntil = Int(Timer + 15)
Do Until Timer > WaitUntil: DoEvents: Loop
SendKeys ("%{F4}") ' alt F4, this will close the active window
WaitUntil = Int(Timer + 15)
Do Until Timer > WaitUntil: DoEvents: Loop




funConvertPDFToExcel = True
Exit Function

ConvertPDFToExcelError:
funConvertPDFToExcel = False

End Function
