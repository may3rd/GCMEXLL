Public Class Units
    Public Property Symbol As String
    Public Property Name As String
    Public Property Plural_name As Object = Nothing
    Public Property L As Double = 0  ' Length
    Public Property M As Double = 0  ' Mass
    Public Property T As Double = 0  ' Time
    Public Property I As Double = 0  ' Electric current
    Public Property THETA As Double = 0  ' Themodynamic temperature
    Public Property N As Double = 0  ' Amount of substance
    Public Property J As Double = 0  ' Light intensity
    Public Property Coef As Decimal = 1D
    Public Property Offset As Decimal = 0D

    Public Sub New()

    End Sub

    Public Sub New(_Symbol As String, _Name As String, Optional _L As Double = 0, Optional _M As Double = 0, Optional _T As Double = 0, Optional _I As Double = 0,
                   Optional _THETA As Double = 0, Optional _N As Double = 0, Optional _J As Double = 0, Optional _Coef As Decimal = 1D, Optional _Offset As Decimal = 0D)
        Symbol = _Symbol
        Name = _Name
        L = _L
        M = _M
        T = _T
        I = _I
        THETA = _THETA
        N = _N
        J = _J
        Coef = _Coef
        Offset = _Offset
    End Sub

    Public Sub New(obj As Units)
        Symbol = obj.Symbol
        Name = obj.Name
        L = obj.L
        M = obj.M
        T = obj.T
        I = obj.I
        THETA = obj.THETA
        N = obj.N
        J = obj.J
        Coef = obj.Coef
        Offset = obj.Offset
    End Sub

    Public Overrides Function ToString() As String
        Dim result As String
        Dim dictionary = New Dictionary(Of String, Double)()

        If L <> 0D Then dictionary.Add("m", L)
        If M <> 0D Then dictionary.Add("kg", M)
        If T <> 0D Then dictionary.Add("s", T)
        If I <> 0D Then dictionary.Add("A", I)
        If THETA <> 0D Then dictionary.Add("K", THETA)
        If N <> 0 Then dictionary.Add("mol", N)
        If J <> 0 Then dictionary.Add("cd", J)

        Dim sorted = From pair In dictionary Order By pair.Value Descending
        Dim sortedDictionary As Dictionary(Of String, Double) = sorted.ToDictionary(Function(p) p.Key, Function(p) p.Value)

        result = ""
        For Each pair In sortedDictionary
            result += " " + pair.Key + ":" + CStr(pair.Value)
        Next

        'Return $"m^{L} * kg^{M} * s^{T} * A^{I} * K^{THETA} * mol^{N} * cd^{J}"
        Return $"{Symbol}, Dimension = {result}, Coeficient={Coef}, Offset={Offset}"
    End Function

    Public Function IsSameDimension(other_unit As Units) As Boolean
        IsSameDimension = L = other_unit.L And M = other_unit.M And T = other_unit.T And I = other_unit.I And THETA = other_unit.THETA And N = other_unit.N And J = other_unit.J
    End Function

    Public Shared Operator *(obj1 As Units, obj2 As Units) As Units

        Dim final_unit As New Units With {
            .Symbol = obj1.Symbol + "*" + obj2.Symbol,
            .Name = obj1.Name + "*" + obj2.Name,
            .L = obj1.L + obj2.L,
            .M = obj1.M + obj2.M,
            .T = obj1.T + obj2.T,
            .I = obj1.I + obj2.I,
            .THETA = obj1.THETA + obj2.THETA,
            .N = obj1.N + obj2.N,
            .J = obj1.J + obj2.J,
            .Coef = obj1.Coef * obj2.Coef,
            .Offset = obj1.Offset + obj2.Offset
        }

        If final_unit.Symbol.StartsWith("*"c) Then final_unit.Symbol = final_unit.Symbol.Substring(1)

        If final_unit.Name.StartsWith("*"c) Then final_unit.Name = final_unit.Name.Substring(1)

        ' set offset of multiplied unit to 0.0D
        If Array.IndexOf({"K", "C", "F", "R"}, final_unit.Symbol) >= 0 Or final_unit.Symbol.EndsWith("g"c) Then GoTo 1000

        final_unit.Offset = 0D
1000:
        Return final_unit
    End Operator

    Public Shared Operator ^(obj1 As Units, power As Object) As Units

        Dim new_offset As Decimal

        If TypeOf power Is Decimal Then
            new_offset = CDec(Math.Pow(CDbl(obj1.Offset), CDbl(power)))
        ElseIf TypeOf power Is Double Or TypeOf power Is Integer Then
            If CDbl(obj1.Offset) = 0.0 Then
                new_offset = 0D
            Else
                new_offset = CDec(Math.Pow(CDbl(obj1.Offset), CDbl(power)))
            End If
        Else
            new_offset = obj1.Offset
        End If

        Dim final_unit As New Units With {
            .Symbol = obj1.Symbol + "^" + CStr(power),
            .Name = obj1.Name + "^" + CStr(power),
            .L = obj1.L * CDbl(power),
            .M = obj1.M * CDbl(power),
            .T = obj1.T * CDbl(power),
            .I = obj1.I * CDbl(power),
            .THETA = obj1.THETA * CDbl(power),
            .N = obj1.N * CDbl(power),
            .J = obj1.J * CDbl(power),
            .Coef = CDec(Math.Pow(CDbl(obj1.Coef), CDbl(power))),
            .Offset = new_offset
        }

        Dim Coef As Double

        Coef = Math.Pow(CDbl(obj1.Coef), CDbl(power))

        final_unit.Coef = Coef

        Return final_unit
    End Operator

    Public Shared Operator /(obj1 As Units, obj2 As Units) As Units
        Return obj1 * obj2 ^ -1
    End Operator

End Class
