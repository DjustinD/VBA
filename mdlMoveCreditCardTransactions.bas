Attribute VB_Name = "mdlMoveCreditCardTransactions"
Option Explicit
' this is the code used to import credit card transactions
' the general idea can also be used to import any series of csv files into a spreadsheet, or even a database.
Private Const strImportFolder = "temp"
Private Const strArchiveFolder = "archive"
Private Const strDataFile = "CreditCardTransactions.xlsx"
Private Const strDataSheet = "CreditCardSpend"
Private Const strCurDir = "C:\Users\djust\OneDrive\Investments\MasterCard"
' the purpose of this code is to move the data from the individual files in the import folder and put the dat into the data file and put the file into the archive folder
Sub sMoveCreditCardTransactions()
Dim strFileName As String
Dim objDataFile As Workbook
Dim objDataSheet As Worksheet
Dim intCurrentRow As Integer
Dim strRowText As String
Dim varRowText As Variant
Dim intFor As Integer
Dim objFileScripting As FileSystemObject ' requires reference to Microsoft Scripting library
Dim intRowCount As Integer
Dim bolNeedHeaders As Boolean

    ' note that application.activeworkbook.path may not work if the folder is on onedrive; so use CurDir instead; that this isn't a stable solution(it doesn't always work)
    Set objFileScripting = CreateObject("Scripting.FileSystemObject")
    Set objDataFile = Application.Workbooks.Open(Filename:=strCurDir & "\" & strDataFile)
    Set objDataSheet = objDataFile.Sheets(strDataSheet)
    intCurrentRow = objDataSheet.UsedRange.Rows.Count
    bolNeedHeaders = True ' set this to true if the spreadhset is empty and the headers need to be added from the first file in the directory
    'Stop ' for development
    strFileName = Dir(strCurDir & "\" & strImportFolder & "\*.csv")
    Do While strFileName <> ""
    '    Debug.Print strFileName ' for debugging
    '    Stop ' for debugging
        Open strCurDir & "\" & strImportFolder & "\" & strFileName For Input As #1
        intRowCount = 1
        Do Until EOF(1)
            Line Input #1, strRowText
            If Not intRowCount = 1 Or bolNeedHeaders = True Then
                strRowText = Replace(strRowText, Chr(34), "") ' remove the double quotes
                varRowText = Split(strRowText, ",") ' split the string based on the deliminator, in this case a comma
                For intFor = LBound(varRowText) To UBound(varRowText)
                    objDataSheet.Cells(intCurrentRow, intFor + 1) = varRowText(intFor)
                Next intFor
                intCurrentRow = intCurrentRow + 1
                If bolNeedHeaders = True Then bolNeedHeaders = False
            End If
            intRowCount = intRowCount + 1
        Loop
        Close #1
        'Stop
        objFileScripting.MoveFile Source:=strCurDir & "\" & strImportFolder & "\" & strFileName, Destination:=strCurDir & "\" & strArchiveFolder & "\" & strFileName
        strFileName = Dir
    Loop
End Sub
