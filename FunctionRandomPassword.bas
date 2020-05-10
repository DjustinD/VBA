Attribute VB_Name = "FunctionRandomPassword"
Option Compare Database
Option Explicit
Option Base 0

Sub testRandomPassword()
Dim LengthInt As Integer
LengthInt = Int((Rnd() * 8) + 8)
Debug.Print funRandomPassword(LengthInt)
End Sub


Function funRandomPassword(LengthInt As Integer) As String
Dim x As Integer
For x = 1 To LengthInt
' possible characters
funRandomPassword = funRandomPassword & Left(Right("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()-=<>", Int((Rnd() * 76) + 1)), 1)
Next x
'Debug.Print "password length is " & LengthInt
'Debug.Print "password is " & funRandomPassword
End Function

