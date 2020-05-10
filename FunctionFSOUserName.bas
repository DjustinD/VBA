Attribute VB_Name = "FunctionFSOUserName"
Option Compare Database
Option Explicit
Option Base 0

' set the variable "User" equal to the FSOUserName as a constant for this project


Private Declare Function apiGetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long



Sub testFSOUserName()
    Debug.Print funFSOUserName()
End Sub


Function funFSOUserName() As String

''******************** Code Start **************************
'' This code was originally written by Dev Ashish.
'' It is not to be altered or distributed,
'' except as part of an application.
'' You are free to use it in any application,
'' provided the copyright notice is left unchanged.
''
'' Code Courtesy of
'' Dev Ashish

' Returns the network login name for the current user
Dim lngLen As Long
Dim lngX As Long
Dim strUserName As String

strUserName = String$(254, 0)
lngLen = 255
lngX = apiGetUserName(strUserName, lngLen)
If (lngX > 0) Then
    funFSOUserName = Left$(strUserName, lngLen - 1)
Else
    funFSOUserName = vbNullString
End If
'
End Function



