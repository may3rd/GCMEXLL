Public Class Quantity
    Public Property Value As Decimal
    Public Property Unit As Units

    Public Sub New()

    End Sub

    Public Sub New(_Value As Object, _unit As Units)
        If TypeOf _Value Is Double Or TypeOf _Value Is Integer Then
            Value = CDec(_Value)
        ElseIf TypeOf _value Is Decimal Then
            Value = _Value
        Else
            Value = 0D
        End If

        Unit = _unit
    End Sub


    Public Function Convert(desired_unit As Units) As Quantity
        Dim new_quantity As New Quantity
        Dim default_value As Decimal
        Dim desired_value As Decimal

        If Not desired_unit.IsSameDimension(Me.Unit) Then
            Console.WriteLine("Not same dimension")
            new_quantity.Value = 0.0D
            new_quantity.Unit = Me.Unit
        Else
            default_value = Me.Unit.Offset + Me.Value * Me.Unit.Coef
            desired_value = (-desired_unit.Offset + default_value) / desired_unit.Coef
            new_quantity.Value = desired_value
            new_quantity.Unit = desired_unit
        End If

        Convert = new_quantity
    End Function

    Public Overrides Function ToString() As String
        Return $"{Value} {Unit.Symbol}"
    End Function
End Class