Attribute VB_Name = "FunctionFuturesPriceFromEighth"
Option Compare Database
Option Explicit
Option Base 0

Sub testFuturesPriceFromEighth()

Dim FuturesPrice As Integer
FuturesPrice = 8096
Debug.Print funFuturesPriceFromEighth(FuturesPrice)

End Sub


Function funFuturesPriceFromEighth(FuturesPrice As Integer) As Double

funFuturesPriceFromEighth = Int(FuturesPrice / 10) / 100 + Right(FuturesPrice, 1) / 800
End Function

