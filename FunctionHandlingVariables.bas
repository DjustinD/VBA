Attribute VB_Name = "FunctionHandlingVariables"
Option Compare Database
Option Explicit
Option Base 0

Sub testHandlingVariables()

Dim BooleanVariable As Boolean
Dim ByteVariable As Byte
Dim CurrencyVariable As Currency
Dim DateVariable As Date
Dim DoubleVariable As Double
Dim IntegerVariable As Integer
Dim LongVariable As Long
Dim SingleVariable As Single
Dim StringVariable As String

BooleanVariable = False                     ' lower limit
Debug.Print BooleanVariable
BooleanVariable = 0                         ' lower limit
Debug.Print BooleanVariable
BooleanVariable = True                      ' upper limit
Debug.Print BooleanVariable
BooleanVariable = 1                         ' upper limit
Debug.Print BooleanVariable


ByteVariable = 0 ' lower limit
ByteVariable = 255 ' upper limit


CurrencyVariable = -922337203.2514          ' lower limit is -922337203685477.5808 , -9.22337203685477E+14
Debug.Print CurrencyVariable                ' not sure how to get this to accept the full number, it works when using the variable in computations
CurrencyVariable = 922337203.5807           ' upper limit is 922337203685477.5807
Debug.Print CurrencyVariable

DateVariable = Now()
Debug.Print DateVariable
DateVariable = 1
Debug.Print DateVariable
DateVariable = #1/1/100# ' lower limit
Debug.Print DateVariable
DateVariable = #12/31/9999# ' upper limmit
Debug.Print DateVariable


DoubleVariable = -1.79769313486231E+308 ' lower limit of negative number
Debug.Print DoubleVariable
DoubleVariable = -4.94065645841247E-324 ' upper limit of negative number
Debug.Print DoubleVariable
DoubleVariable = 4.94065645841247E-324 ' lower limit of positive number
Debug.Print DoubleVariable
DoubleVariable = 1.79769313486231E+308 ' upper limit of positive number

IntegerVariable = -32768 ' lower limit
Debug.Print IntegerVariable
IntegerVariable = 32767 ' upper limit
Debug.Print IntegerVariable

LongVariable = -2147483648# ' lower limit
Debug.Print LongVariable
LongVariable = 2147483647# ' upper limit
Debug.Print LongVariable

SingleVariable = -3.402823E+38 ' lower limit of negative number
Debug.Print SingleVariable
SingleVariable = -1.401298E-45 ' upper limit of negative number
Debug.Print SingleVariable
SingleVariable = 1.401298E-45  ' lower limit of positive number
Debug.Print SingleVariable
SingleVariable = 3.402823E+38 ' upper limit of positive number
Debug.Print SingleVariable

StringVariable = "Justin" ' the string variable is variable length
Debug.Print StringVariable
'
'Call HandlingVariables(BooleanVariable, ByteVariable, CurrencyVariable, DateVariable, DoubleVariable, IntegerVariable, LongVariable, SingleVariable, StringVariable)
'Call HandlingVariables(BooleanVariable, ByteVariable, CurrencyVariable, DateVariable, DoubleVariable, IntegerVariable, LongVariable, SingleVariable)
'Call HandlingVariables(BooleanVariable, ByteVariable, CurrencyVariable, DateVariable, DoubleVariable, IntegerVariable, LongVariable)
'Call HandlingVariables(BooleanVariable, ByteVariable, CurrencyVariable, DateVariable, DoubleVariable, IntegerVariable)
'Call HandlingVariables(BooleanVariable, ByteVariable, CurrencyVariable, DateVariable, DoubleVariable)
'Call HandlingVariables(BooleanVariable, ByteVariable, CurrencyVariable, DateVariable)
'Call HandlingVariables(BooleanVariable, ByteVariable, CurrencyVariable)
'Call HandlingVariables(BooleanVariable, ByteVariable)
'Call HandlingVariables(BooleanVariable)
'Call HandlingVariables
'Call HandlingVariables(StringVariable:=StringVariable)
Call HandlingVariables(StringVariable:=StringVariable, SingleVariable:=SingleVariable, LongVariable:=LongVariable, IntegerVariable:=IntegerVariable, DoubleVariable:=DoubleVariable)



End Sub

Function HandlingVariables(Optional BooleanVariable As Variant, Optional ByteVariable As Variant, Optional CurrencyVariable As Variant, Optional DateVariable As Variant, Optional DoubleVariable As Variant, Optional IntegerVariable As Variant, Optional LongVariable As Variant, Optional SingleVariable As Variant, Optional StringVariable As Variant) As Boolean
' use variant data type becase we can test if variants are missing
On Error GoTo HandlingVariablesError
'this is an example of how to handle different variable types

'Test DateVariable
If Not IsMissing(DateVariable) Then
    If IsDate(DateVariable) Then
        Debug.Print "DateVariable is " & DateVariable
    Else
    Debug.Print "DateVariable is not a Date"
    End If
Else
    Debug.Print "DateVariable is missing"
End If
'Test BooleanVariable
If Not IsMissing(BooleanVariable) Then
    If IsNumeric(BooleanVariable) Then
        Debug.Print "BooleanVariable is " & BooleanVariable
    Else
    Debug.Print "BooleanVariable is not a number"
    End If
Else
    Debug.Print "BooleanVariable is missing"
End If
'Test ByteVariable
If Not IsMissing(ByteVariable) Then
    If IsNumeric(ByteVariable) Then
        Debug.Print "ByteVariable is " & ByteVariable
    Else
    Debug.Print "ByteVariable is not a number"
    End If
Else
    Debug.Print "ByteVariable is missing"
End If
'Test CurencyVariable
If Not IsMissing(CurrencyVariable) Then
    If IsNumeric(CurrencyVariable) Then
        Debug.Print "CurrencyVariable is " & CurrencyVariable
    Else
    Debug.Print "CurrencyVariable is not a number"
    End If
Else
    Debug.Print "CurrencyVariable is missing"
End If
'Test DoubleVariable
If Not IsMissing(DoubleVariable) Then
    If IsNumeric(DoubleVariable) Then
        Debug.Print "DoubleVariable is " & DoubleVariable
    Else
    Debug.Print "DoubleVariable is not a number"
    End If
Else
    Debug.Print "DoubleVariable is missing"
End If
'Test IntegerVariable
If Not IsMissing(IntegerVariable) Then
    If IsNumeric(IntegerVariable) Then
        Debug.Print "IntegerVariable is " & IntegerVariable
    Else
    Debug.Print "IntegerVariable is not a number"
    End If
Else
    Debug.Print "IntegerVariable is missing"
End If
'Test LongVariable
If Not IsMissing(LongVariable) Then
    If IsNumeric(LongVariable) Then
        Debug.Print "LongVariable is " & LongVariable
    Else
    Debug.Print "LongVariable is not a number"
    End If
Else
    Debug.Print "LongVariable is missing"
End If
'Test SingleVariable
If Not IsMissing(SingleVariable) Then
    If IsNumeric(SingleVariable) Then
        Debug.Print "SingleVariable is " & SingleVariable
    Else
    Debug.Print "SingleVariable is not a number"
    End If
Else
    Debug.Print "SingleVariable is missing"
End If
' notice there is no test to see if a string varialbe is a valid datatype
If IsMissing(StringVariable) Then Debug.Print "StringVariable Is missing" Else Debug.Print "StringVariable is " & StringVariable

HandlingVariables = True
Exit Function
HandlingVariablesError:
Stop
End Function

