Imports System.ComponentModel

Module Convert
    Public Function ConvertDec(quantity As String, desired_unit As String) As Decimal
        Dim QP As New QuantityParser
        Dim Q1 As Quantity
        Dim UP As New UnitParser()

        Q1 = QP.Parse(quantity)
        Return Q1.Convert(UP.Parse(desired_unit)).Value
    End Function

    Public Function ConvertStr(quantity As String, desired_unit As String, Optional format As String = "0.000000") As String
        Dim QP As New QuantityParser
        Dim UP As New UnitParser
        Dim Q1 As Quantity
        Dim U1 As Units

        Q1 = QP.Parse(quantity)
        U1 = UP.Parse(desired_unit)

        Return Q1.Convert(U1).Value.ToString(format) + " " + U1.Symbol
    End Function

    Public Function ConvertDbl(quantity As String, desired_unit As String) As Double
        Return CDbl(ConvertDec(quantity, desired_unit))
    End Function
End Module
