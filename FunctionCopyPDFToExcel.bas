Attribute VB_Name = "FunctionCopyPDFToExcel"
Option Compare Database
Option Explicit
Option Base 0

Sub TestCopyPDFToExcel()
Dim PDFFilename As String
Dim ExcelFilename As String

If funCopyPDFToExcel(PDFFilename, ExcelFilename) = True Then
    MsgBox ("Success!")
Else
    MsgBox ("Failure!")
End If

End Sub

Function funCopyPDFToExcel(PDFFilename As String, ExcelFilename As String) As Boolean
On Error GoTo CopyPDFToExcelError




funCopyPDFToExcel = True

Exit Function
CopyPDFToExcelError:
funCopyPDFToExcel = False

End Function
