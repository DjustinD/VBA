Attribute VB_Name = "fOSUserName"
Option Compare Database
Option Explicit

'******************** Code Start **************************
' This code was originally written by Dev Ashish.
' It is not to be altered or distributed,
' except as part of an application.
' You are free to use it in any application,
' provided the copyright notice is left unchanged.
'
' Code Courtesy of
' Dev Ashish
'
'*****with some unnotable edits by myself
'
'Private Declare Function apiGetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long ' used in 32 bit systems
Private Declare PtrSafe Function apiGetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long ' updated for use with 64 bit systems, note the variables are still dimentioned as long

Sub testOSUserName()
Debug.Print fOSUserName
End Sub
Function fOSUserName() As String
' Returns the network login name
Dim lngLen As Long
Dim lngTemp As Long
Dim strUserName As String
    
    strUserName = String$(254, 0)
    lngLen = 7
    lngLen = 255
    lngTemp = apiGetUserName(strUserName, lngLen)
    If (lngTemp > 0) Then
        fOSUserName = Left(strUserName, lngLen - 1)
    Else
        fOSUserName = vbNullString
    End If
End Function
'******************** Code End **************************

Sub testUserID()
    Debug.Print fUserID(fOSUserName)
End Sub

Function fUserID(strUser As String) As Integer
' use this function to translate the strUserName to idUserName via an access table
Dim strSELECT As String
Dim strFROM As String
Dim strWHERE As String
Dim rstTemp As Recordset

    strSELECT = "SELECT tblUsers.idUser"
    strFROM = "FROM tblUsers"
    strWHERE = "WHERE (((tblUsers.strUserName)='" & strUser & "') AND ((tblUsers.dtmEffectiveTo) Is Null))"

    Set rstTemp = CurrentDb.OpenRecordset(strSELECT & " " & strFROM & " " & strWHERE & ";")
    fUserID = rstTemp!idUser
    Set rstTemp = Nothing

End Function

