Public Class UnitPrefix
    Public Property Symbol As String
    Public Property Name As String
    Public Property Factor As Decimal

    Public Sub New()

    End Sub

    Public Sub New(s As String, n As String, f As Object)
        Symbol = s
        Name = n
        If TypeOf f Is Decimal Then
            Factor = f
        ElseIf TypeOf f Is Double Or TypeOf f Is Integer Then
            Factor = CDec(f)
        End If
    End Sub

    Public Sub New(obj As UnitPrefix)
        Symbol = obj.Symbol
        Name = obj.Name
        Factor = obj.Factor
    End Sub

    Public Overrides Function ToString() As String
        Return $"{Me.Symbol} = {Me.Factor}"
    End Function

    Public Function IsSameFactor(other_prefix As UnitPrefix) As Boolean
        IsSameFactor = Me.factor = other_prefix.factor
    End Function

    Public Overloads Function Equals(other_prefix As UnitPrefix) As Boolean
        Equals = Me.symbol = other_prefix.symbol And Me.name = other_prefix.name And Me.factor = other_prefix.factor
    End Function

    Public Function Multiply(init_unit As Units) As Units
        Dim final_unit As New Units With {
            .Symbol = Me.symbol + init_unit.Symbol,
            .Name = Me.name + init_unit.Name,
            .L = init_unit.L,
            .M = init_unit.M,
            .T = init_unit.T,
            .I = init_unit.I,
            .THETA = init_unit.THETA,
            .N = init_unit.N,
            .J = init_unit.J,
            .Coef = Me.factor * init_unit.Coef,
            .Offset = init_unit.Offset
        }

        Multiply = final_unit
    End Function

    Public Shared Operator *(ByVal obj1 As UnitPrefix, init_unit As Units) As Units
        Dim final_unit As New Units With {
            .Symbol = obj1.symbol + init_unit.Symbol,
            .Name = obj1.name + init_unit.Name,
            .L = init_unit.L,
            .M = init_unit.M,
            .T = init_unit.T,
            .I = init_unit.I,
            .THETA = init_unit.THETA,
            .N = init_unit.N,
            .J = init_unit.J,
            .Coef = obj1.factor * init_unit.Coef,
            .Offset = init_unit.Offset
        }

        Return final_unit
    End Operator

End Class