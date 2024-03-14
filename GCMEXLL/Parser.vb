Imports System.Linq.Expressions
Imports System.Numerics
Imports System.Text.RegularExpressions

Public Class QuantityParser
    Private ReadOnly Quantity_re As New Regex("(?<value>\d+[.,]?\d*)? *(?<unit>.*)")

    Public Function Parse(quantity As String) As Quantity
        Dim match As Match = Quantity_re.Match(quantity)
        Dim Value As Decimal
        Dim Unit As Units
        Dim UParser As New UnitParser

        If match.Groups(1).Value IsNot Nothing Then
            Value = CDec(match.Groups(1).Value)
            Unit = UParser.Parse(match.Groups(2).Value)
        Else
            Unit = New Units
        End If

        Return New Quantity(Value, Unit)
    End Function
End Class

Public Class UnitParser
    Private ReadOnly Unit_re As New Regex("(?<unit>[a-zA-Z°Ωµ\(\)]+)\^?(?<pow>[-+]?[0-9]*\.?[0-9]*)")

    ' constants for unit of measurement conversion
    'Private ReadOnly PI As Decimal = 3.14159265358979

    ' physical constants
    'Private ReadOnly Avogadro As Decimal = 602214076000000000000000D ' [1/mole]
    'Private ReadOnly Boltzmann As Decimal = 0.0000000000000000000000138065D ' [j/K]
    'Private ReadOnly R As Decimal = Avogadro * Boltzmann ' [j/mole.K]
    'Private ReadOnly C As Decimal = 299792458D
    Private ReadOnly G_constant As Decimal = 9.80665D
    Private ReadOnly ATM As Decimal = 101325D

    ' mass in kg
    Private ReadOnly gram As Decimal = 0.001D
    'Private ReadOnly metric_ton As Decimal = 1000D
    Private ReadOnly ton As Decimal = 1000D
    'Private ReadOnly tonne As Decimal = 1000D
    Private ReadOnly grain As Decimal = 0.00006479891D
    Private ReadOnly pound As Decimal = grain * 7000D
    'Private ReadOnly lb As Decimal = pound
    'Private ReadOnly lbm As Decimal = pound
    'Private ReadOnly oz As Decimal = pound / 16D
    'Private ReadOnly stone As Decimal = 14D * pound
    'Private ReadOnly short_ton As Decimal = 2000D * pound
    'Private ReadOnly long_ton As Decimal = 2240D * pound

    ' angle in rad
    'Private ReadOnly degree As Decimal = PI / 180D

    ' time in second
    Private ReadOnly minute As Decimal = 60D
    Private ReadOnly hour As Decimal = 60D * minute
    'Private ReadOnly hour_inv As Decimal = 1D / hour
    Private ReadOnly day As Decimal = 24D * hour
    'Private ReadOnly week As Decimal = 7D * day
    Private ReadOnly year As Decimal = 365D * day
    'Private ReadOnly julian_year As Decimal = 365.25D * day

    ' length in meter
    Private ReadOnly cm As Decimal = 0.01D
    'Private ReadOnly mm As Decimal = 0.001D
    Private ReadOnly inch As Decimal = 0.0254D
    Private ReadOnly foot As Decimal = inch * 12D
    Private ReadOnly yard As Decimal = 3D * foot
    Private ReadOnly mile As Decimal = 1760D * yard

    ' area in square meter
    Private ReadOnly cm2 As Decimal = cm * cm
    'Private ReadOnly mm2 As Decimal = mm * mm
    Private ReadOnly inch2 As Decimal = inch * inch
    'Private ReadOnly foot2 As Decimal = foot * foot

    ' pressure in pascal
    Private ReadOnly bar As Decimal = 100000.0
    Private ReadOnly torr As Decimal = ATM / 760D
    Private ReadOnly inHg As Decimal = torr * inch * 1000D
    Private ReadOnly psi As Decimal = pound * G_constant / inch2

    ' volume in meter**3
    Private ReadOnly litre As Decimal = 0.001D
    Private ReadOnly liter As Decimal = litre
    'Private ReadOnly mm3 As Decimal = mm * mm * mm
    Private ReadOnly gallon As Decimal = 231D * inch * inch * inch ' gallon US
    'Private ReadOnly gallon_imp = 0.00454609D  ' gallon UK
    Private ReadOnly barrel As Decimal = 42D * gallon
    Private ReadOnly foot3 As Decimal = foot * foot * foot
    'Private ReadOnly pint As Decimal = 568.26125D * mm3

    ' temperature in kelvin
    Private ReadOnly zero_celsius As Decimal = 273.15D
    Private ReadOnly degree_f As Decimal = 5D / 9D   ' only for differences

    ' energy in joule
    Private ReadOnly calorie As Decimal = 4.184D
    Private ReadOnly calorie_IT = 4.1868D
    Private ReadOnly btu As Decimal = pound * degree_f * calorie_IT / gram

    ' power in watt
    Private ReadOnly horsepower As Decimal = 550D * foot * pound * G_constant
    'Private ReadOnly hp As Decimal = horsepower

    ' force in newton
    Private ReadOnly dyne As Decimal = 0.00001
    Private ReadOnly lbf As Decimal = pound * G_constant
    Private ReadOnly kgf As Decimal = G_constant '* 1 kg

    Private ReadOnly Exception As String() = {"kg/cm2", "kg/cm2g", "mmH2O", "inH2O"}

    Public PREFIXS As New Dictionary(Of String, UnitPrefix)
    Public UNITS As New Dictionary(Of String, Units)

    Sub New()
        '
        PREFIXS.Add("", New UnitPrefix("", "", 1D))
        PREFIXS.Add("y", New UnitPrefix("y", "yocto", 0.000000000000000000000001D))
        PREFIXS.Add("z", New UnitPrefix("z", "zepto", 0.000000000000000000001D))
        PREFIXS.Add("a", New UnitPrefix("a", "atto", 0.000000000000000001D))
        PREFIXS.Add("f", New UnitPrefix("f", "femto", 0.000000000000001D))
        PREFIXS.Add("p", New UnitPrefix("p", "pico", 0.000000000001D))
        PREFIXS.Add("n", New UnitPrefix("n", "nano", 0.000000001D))
        PREFIXS.Add("µ", New UnitPrefix("µ", "micro", 0.000001D))
        PREFIXS.Add("m", New UnitPrefix("m", "milli", 0.001D))
        PREFIXS.Add("c", New UnitPrefix("c", "centi", 0.01D))
        PREFIXS.Add("d", New UnitPrefix("d", "deci", 0.1D))
        PREFIXS.Add("da", New UnitPrefix("da", "deca", 10D))
        PREFIXS.Add("h", New UnitPrefix("h", "hecto", 100D))
        PREFIXS.Add("k", New UnitPrefix("k", "kilo", 1000D))
        PREFIXS.Add("MM", New UnitPrefix("MM", "million", 1000000D))
        PREFIXS.Add("M", New UnitPrefix("M", "mega", 1000000D))
        PREFIXS.Add("G", New UnitPrefix("G", "giga", 1000000000D))
        PREFIXS.Add("T", New UnitPrefix("T", "tera", 1000000000000D))
        PREFIXS.Add("P", New UnitPrefix("P", "peta", 1000000000000000D))
        PREFIXS.Add("E", New UnitPrefix("E", "exa", 1000000000000000000D))
        PREFIXS.Add("Z", New UnitPrefix("Z", "zetta", 1000000000000000000000D))
        PREFIXS.Add("Y", New UnitPrefix("Y", "yotta", 1000000000000000000000000D))

        ' Basic SI units
        UNITS.Add("m", New Units("m", "meter", _L:=1))
        UNITS.Add("g", New Units("g", "gram", _M:=1, _Coef:=gram))
        UNITS.Add("s", New Units("s", "second", _T:=1))
        UNITS.Add("A", New Units("A", "ampere", _I:=1))
        UNITS.Add("K", New Units("K", "kelvin", _THETA:=1))
        UNITS.Add("mol", New Units("mol", "mole", _N:=1))
        UNITS.Add("cd", New Units("cd", "candela", _J:=1))

        ' Derived SI units
        UNITS.Add("Hz", New Units("Hz", "hertz", _T:=-1))
        UNITS.Add("N", New Units("N", "newton", _M:=1, _L:=1, _T:=-2))
        UNITS.Add("kgf", New Units("kgf", "kilogram force", _M:=1, _L:=1, _T:=-2, _Coef:=G_constant))
        UNITS.Add("dyn", New Units("dyn", "dyne", _M:=1, _L:=1, _T:=-2, _Coef:=dyne))
        UNITS.Add("dyne", New Units("dyn", "dyne", _M:=1, _L:=1, _T:=-2, _Coef:=dyne))
        UNITS.Add("Pa", New Units("Pa", "pascal", _M:=1, _L:=-1, _T:=-2))
        UNITS.Add("J", New Units("J", "joule", _M:=1, _L:=2, _T:=-2))
        UNITS.Add("cal", New Units("cal", "calorie", _M:=1, _L:=2, _T:=-2, _Coef:=calorie))
        UNITS.Add("kcal", New Units("kcal", "kilo calorie", _M:=1, _L:=2, _T:=-2, _Coef:=calorie * 1000D))
        UNITS.Add("W", New Units("W", "watt", _M:=1, _L:=2, _T:=-3))
        UNITS.Add("L", New Units("L", "liter", _L:=3, _Coef:=liter))
        UNITS.Add("liter", New Units("L", "liter", _L:=3, _Coef:=liter))
        UNITS.Add("litre", New Units("L", "liter", _L:=3, _Coef:=liter))
        UNITS.Add("Poise", New Units("Poise", "poise", _L:=-1, _M:=1, _T:=-1, _Coef:=0.1D))
        UNITS.Add("P", New Units("P", "poise", _L:=-1, _M:=1, _T:=-1, _Coef:=0.1D))
        UNITS.Add("St", New Units("St", "stokes", _L:=2, _T:=-1, _Coef:=0.0001D))
        UNITS.Add("ton", New Units("ton", "tonne", _M:=1, _Coef:=ton))
        UNITS.Add("LT", New Units("long_ton", "long ton", _M:=1, _Coef:=2240D * pound))
        UNITS.Add("ST", New Units("short_ton", "short ton", _M:=1, _Coef:=2000D * pound))

        UNITS.Add("sec", New Units("sec", "second", _T:=1))
        UNITS.Add("min", New Units("min", "minute", _T:=1, _Coef:=minute))
        UNITS.Add("hr", New Units("hr", "hour", _T:=1, _Coef:=hour))
        UNITS.Add("h", New Units("h", "hour", _T:=1, _Coef:=hour))
        UNITS.Add("hour", New Units("hour", "hour", _T:=1, _Coef:=hour))
        UNITS.Add("day", New Units("day", "day", _T:=1, _Coef:=day))
        UNITS.Add("week", New Units("week", "week", _T:=1, _Coef:=hour))
        UNITS.Add("year", New Units("yr", "year", _T:=1, _Coef:=year))
        UNITS.Add("yr", New Units("yr", "year", _T:=1, _Coef:=year))

        UNITS.Add("C", New Units("C", "Celsius", _THETA:=1, _Coef:=1D, _Offset:=zero_celsius))

        ' Imperial system
        UNITS.Add("F", New Units("F", "Fahrenheit ", _THETA:=1, _Coef:=degree_f, _Offset:=zero_celsius - 32D * degree_f))
        UNITS.Add("R", New Units("R", "Rankin", _THETA:=1, _Coef:=degree_f))

        UNITS.Add("in", New Units("in", "inch", _L:=1, _Coef:=inch))
        UNITS.Add("ft", New Units("ft", "foot", _L:=1, _Coef:=foot))
        UNITS.Add("yard", New Units("yard", "yard", _L:=1, _Coef:=foot * 3D))
        UNITS.Add("mile", New Units("mile", "mile", _L:=1, _Coef:=mile))

        UNITS.Add("pound", New Units("lb", "pound (mass)", _M:=1, _Coef:=pound))
        UNITS.Add("lb", New Units("lb", "pound (mass)", _M:=1, _Coef:=pound))
        UNITS.Add("lbm", New Units("lb", "pound (mass)", _M:=1, _Coef:=pound))
        UNITS.Add("lbf", New Units("lbf", "pound (force)", _M:=1, _L:=1, _T:=-2, _Coef:=lbf))

        UNITS.Add("gal", New Units("gal", "US gallon", _L:=3, _Coef:=gallon))
        UNITS.Add("gallon", New Units("gallon", "US gallon", _L:=3, _Coef:=gallon))
        UNITS.Add("bbl", New Units("bbl", "US barrel", _Coef:=barrel))
        UNITS.Add("oz", New Units("oz", "ounce", _M:=1, _Coef:=pound))

        UNITS.Add("SCFM", New Units("SCFM", "standard cubic foot per minute", _L:=3, _T:=-1, _Coef:=foot3 / minute))
        UNITS.Add("SCFD", New Units("SCFD", "standard cubic foot per day", _L:=3, _T:=-1, _Coef:=foot3 / day))
        UNITS.Add("cfm", New Units("CFM", "standard cubic foot per minute", _L:=3, _T:=-1, _Coef:=foot3 / minute))
        UNITS.Add("cfd", New Units("CFD", "standard cubic foot per day", _L:=3, _T:=-1, _Coef:=foot3 / day))

        UNITS.Add("psi", New Units("psi", "pound per square inch", _M:=1, _L:=-1, _T:=-2, _Coef:=psi))
        UNITS.Add("psia", New Units("psi", "pound per square inch", _M:=1, _L:=-1, _T:=-2, _Coef:=psi))
        UNITS.Add("psig", New Units("psig", "pound per square inch (gauge)", _M:=1, _L:=-1, _T:=-2, _Coef:=psi, _Offset:=ATM))

        UNITS.Add("BTU", New Units("btu", "british thermal unit", _M:=1, _L:=2, _T:=-2, _Coef:=btu))
        UNITS.Add("btu", New Units("btu", "british thermal unit", _M:=1, _L:=2, _T:=-2, _Coef:=btu))

        UNITS.Add("horsepower", New Units("hp", "hores power", _M:=1, _L:=2, _T:=-3, _Coef:=horsepower))
        UNITS.Add("hp", New Units("hp", "hores power", _M:=1, _L:=2, _T:=-3, _Coef:=horsepower))

        ' velocity
        UNITS.Add("mph", New Units("mph", "mile per hour", _L:=1, _T:=-1, _Coef:=mile / hour))
        'UNITS.Add("", New Units("", "", ))

        ' pressure unit
        UNITS.Add("ATM", New Units("ATM", "atmosphere", _M:=1, _L:=-1, _T:=-2, _Coef:=ATM))
        UNITS.Add("bar", New Units("bar", "bar", _M:=1, _L:=-1, _T:=-2, _Coef:=bar))
        UNITS.Add("bara", New Units("bar", "bar (absolute)", _M:=1, _L:=-1, _T:=-2, _Coef:=bar))
        UNITS.Add("barg", New Units("barg", "bar (gauge)", _M:=1, _L:=-1, _T:=-2, _Coef:=bar, _Offset:=ATM))
        UNITS.Add("Pag", New Units("Pag", "pascal (gauge)", _M:=1, _L:=-1, _T:=-2, _Coef:=1, _Offset:=ATM))
        UNITS.Add("ksc", New Units("ksc", "kilogram per square centimeter", _M:=1, _L:=-1, _T:=-2, _Coef:=kgf / cm2))
        UNITS.Add("kscg", New Units("kscg", "kilogram per square centimeter (gauge)", _M:=1, _L:=-1, _T:=-2, _Coef:=kgf / cm2, _Offset:=ATM))
        UNITS.Add("kg/cm2", New Units("kg/cm2", "kilogram per square centimeter", _M:=1, _L:=-1, _T:=-2, _Coef:=kgf / cm2))
        UNITS.Add("kg/cm2g", New Units("kg/cm2g", "kilogram per square centimeter (gauge)", _M:=1, _L:=-1, _T:=-2, _Coef:=kgf / cm2, _Offset:=ATM))
        UNITS.Add("mmH2O", New Units("mmH2O", "millimeter of water", _M:=1, _L:=-1, _T:=-2, _Coef:=G_constant))
        UNITS.Add("inH2O", New Units("inH2O", "inch of water", _M:=1, _L:=-1, _T:=-2, _Coef:=G_constant * 1000D * inch))
        UNITS.Add("torr", New Units("torr", "torrr", _M:=1, _L:=-1, _T:=-2, _Coef:=torr))
        UNITS.Add("mmHg", New Units("mmHg", "millimeter of mercury", _M:=1, _L:=-1, _T:=-2, _Coef:=torr))
        UNITS.Add("inHg", New Units("inHg", "inch of mercury", _M:=1, _L:=-1, _T:=-2, _Coef:=inHg))

    End Sub

    Public Function Parse(s As String) As Units
        Dim i As Integer

        ' add exception of kg/cm2 and kg/cm2g
        If Array.IndexOf(Exception, s) >= 0 Then GoTo 1000

        i = s.IndexOf("/"c)
1000:
        If i > 0 Then
            Return ParseS(s.Substring(0, i)) / ParseS(s.Substring(i + 1))
        Else
            Return ParseS(s)
        End If

    End Function

    Private Function ParseS(s As String) As Units

        Dim matches As MatchCollection = Unit_re.Matches(s)
        Dim result As New Units
        Dim unit As Units
        Dim l_unit(matches.Count) As Units

        ' add exception of kg/cm2 and kg/cm2g
        If Array.IndexOf(Exception, s) >= 0 Then GoTo 1000

        ' Loop over matches
        For Each m As Match In matches
            unit = Me.ParseUnit(m.Groups(1).Value, m.Groups(2).Value)
            result *= unit
        Next

        Return result
        Exit Function
1000:
        Return Me.ParseSimpleUnit(s)
    End Function

    Public Function ParseUnit(s As String, pow As String) As Units
        Dim result As Units
        If pow = "" Or pow = "-" Or pow = "." Then
            result = Me.ParseSimpleUnit(s)
        Else
            result = Me.ParseSimpleUnit(s) ^ CDbl(pow)
        End If
        Return result
    End Function

    Private Function ParseSimpleUnit(s As String) As Units
        ' parse an unit without a power value
        Dim unit As New Units
        Dim prefix As New UnitPrefix
        Dim sUnit As String

        For Each pair As KeyValuePair(Of String, UnitPrefix) In PREFIXS
            ' pair.Key
            ' pair.Value
            If s.StartsWith(pair.Key) Then
                sUnit = s.Substring(Len(pair.Key)).Replace("(", "").Replace(")", "")

                Dim value As Units = Nothing

                If UNITS.TryGetValue(sUnit, value) Then
                    unit = New Units(value)
                    prefix = New UnitPrefix(pair.Value)
                    Exit For
                End If
            End If
        Next

        Return prefix * unit
    End Function
End Class