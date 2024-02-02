Imports System.Net.NetworkInformation
Imports System.Runtime.ExceptionServices
Imports System.Runtime.Remoting
Imports System.Threading
Imports ExcelDna.Integration

Public Module UOM

    ' constants for unit of measurement conversion
    Public Const PI As Double = 3.14159265358979

    ' physical constants
    Public Const G As Double = 9.80665 ' [m/s^2]
    Public Const Avogadro As Double = 6.02214076E+23 ' [1/mole]
    Public Const Boltzmann As Double = 1.380649E-23 ' [j/K]
    Public Const R As Double = Avogadro * Boltzmann ' [j/mole.K]
    Public Const C As Double = 299792458.0

    ' mass in kg
    Public Const gram As Double = 0.001
    Public Const metric_ton As Double = 1000.0
    Public Const ton As Double = 1000.0
    Public Const grain As Double = 0.00006479891
    Public Const pound As Double = grain * 7000.0
    Public Const lb As Double = pound
    Public Const lbm As Double = pound
    Public Const oz As Double = pound / 16.0
    Public Const stone As Double = 14.0 * pound

    ' angle in rad
    Public Const degree As Double = PI / 180.0

    ' time in second
    Public Const minute As Double = 60.0
    Public Const hour As Double = 60.0 * minute
    Public Const hour_inv As Double = 1 / hour
    Public Const day As Double = 24.0 * hour
    Public Const week As Double = 7.0 * day
    Public Const year As Double = 365.0 * day
    Public Const julian_year As Double = 365.25 * day

    ' length in meter
    Public Const cm As Double = 0.01
    Public Const mm As Double = 0.001
    Public Const inch As Double = 0.0254
    Public Const foot As Double = inch * 12.0
    Public Const yard As Double = 3.0 * foot
    Public Const mile As Double = 1760 * yard

    ' area in square meter
    Public Const cm2 As Double = cm * cm
    Public Const mm2 As Double = mm * mm
    Public Const inch2 As Double = inch * inch
    Public Const foot2 As Double = foot * foot

    ' pressure in pascal
    Public Const atm As Double = 101325.0
    Public Const bar As Double = 100000.0
    Public Const mmh2o As Double = G
    Public Const torr As Double = atm / 760
    Public Const mmHg As Double = torr
    Public Const inchHg As Double = torr * inch * 1000
    Public Const psi As Double = pound * G / inch2

    ' volume in meter**3
    Public Const litre As Double = 0.001
    Public Const liter As Double = litre
    Public Const l As Double = litre
    Public Const ml As Double = 0.001 * liter
    Public Const cm3 As Double = cm * cm * cm
    Public Const mm3 As Double = mm * mm * mm
    Public Const gallon As Double = 231.0 * inch3 ' gallon US
    Public Const gallon_imp = 0.00454609  ' gallon UK
    Public Const inch3 As Double = inch * inch * inch
    Public Const foot3 As Double = foot * foot * foot
    Public Const foot3_inv As Double = 1.0 / foot3
    Public Const pint As Double = 568.26125 * mm3

    ' temperature in kelvin
    Public Const zero_celsius As Double = 273.15
    Public Const degree_f As Double = 5.0 / 9.0   ' only for differences

    ' energy in joule
    Public Const calorie As Double = 4.184
    Public Const calorie_IT = 4.1868
    Public Const btu As Double = pound * degree_f * calorie_IT / gram

    ' power in watt
    Public Const horsepower As Double = 550.0 * foot * pound * G
    Public Const hp As Double = horsepower

    ' force in newton
    Public Const dyne As Double = 0.00001
    Public Const lbf As Double = pound * G
    Public Const kgf As Double = G '* 1 kg

    <ExcelFunction(Description:="convert UOM of area measurement", Category:="GCME E-PT | Unit Convertion")>
    Public Function AreaConvert(<ExcelArgument(Description:="measurement")> val As Double, <ExcelArgument(Description:="Original UOM")> fromUnit As String, <ExcelArgument(Description:="final UOM")> ToUnit As String) As Double
        Dim indexfrom As Integer
        Dim indexTo As Integer
        Dim Result As Double

        Dim Units = New String() {"m2", "cm2", "mm2", "ft2", "in2"}
        Dim factor = New Double() {1, cm2, mm2, foot2, inch2}

        indexfrom = Array.IndexOf(Units, LCase(fromUnit))
        indexTo = Array.IndexOf(Units, LCase(ToUnit))

        If IsNumeric(val) Then
            Result = val * factor(indexfrom) / factor(indexTo)
        End If

        AreaConvert = Result
    End Function

    <ExcelFunction(Description:="convert UOM of energy measurement", Category:="GCME E-PT | Unit Convertion")>
    Public Function EnergyConvert(<ExcelArgument(Description:="measurement")> val As Double, <ExcelArgument(Description:="Original UOM")> fromUnit As String, <ExcelArgument(Description:="final UOM")> ToUnit As String) As Double
        Dim indexfrom As Integer
        Dim indexTo As Integer
        Dim Result As Double

        Dim Units = New String() {"j", "kj", "mj", "Gj", "cal", "kcal", "mcal", "mmkcal", "btu", "kw-h", "kw-hr"}
        Dim factor = New Double() {1, 1000.0, 1000000.0, 1000000000.0, calorie, 1000.0 * calorie, 1000000.0 * calorie, 1000000000.0 * calorie, btu, 1000.0 * hour, 1000.0 * hour}

        indexfrom = Array.IndexOf(Units, LCase(fromUnit))
        indexTo = Array.IndexOf(Units, LCase(ToUnit))

        If IsNumeric(val) Then
            Result = val * factor(indexfrom) / factor(indexTo)
        End If

        EnergyConvert = Result
    End Function

    <ExcelFunction(Description:="convert UOM of heat intensity measurement", Category:="GCME E-PT | Unit Convertion")>
    Public Function HeatIntensityConvert(<ExcelArgument(Description:="measurement")> val As Double, <ExcelArgument(Description:="Original UOM")> fromUnit As String, <ExcelArgument(Description:="final UOM")> ToUnit As String) As Double
        Dim indexfrom As Integer
        Dim indexTo As Integer
        Dim Result As Double

        Dim Units = New String() {"kw/m2", "btu/h.ft2", "j/m2.s"}
        Dim factor = New Double() {1000.0, btu / (hour * foot2), 1}

        indexfrom = Array.IndexOf(Units, LCase(fromUnit))
        indexTo = Array.IndexOf(Units, LCase(ToUnit))

        If IsNumeric(val) Then
            Result = val * factor(indexfrom) / factor(indexTo)
        End If

        HeatIntensityConvert = Result
    End Function

    <ExcelFunction(Description:="convert UOM of heat transfer rate measurement", Category:="GCME E-PT | Unit Convertion")>
    Public Function HeatTransferRateConvert(<ExcelArgument(Description:="measurement")> val As Double, <ExcelArgument(Description:="Original UOM")> fromUnit As String, <ExcelArgument(Description:="final UOM")> ToUnit As String) As Double
        HeatTransferRateConvert = PowerConvert(val, fromUnit, ToUnit)
    End Function

    <ExcelFunction(Description:="convert UOM of latent heat measurement", Category:="GCME E-PT | Unit Convertion")>
    Public Function LatentHeatConvert(<ExcelArgument(Description:="measurement")> val As Double, <ExcelArgument(Description:="Original UOM")> fromUnit As String, <ExcelArgument(Description:="final UOM")> ToUnit As String) As Double
        Dim indexfrom As Integer
        Dim indexTo As Integer
        Dim Result As Double

        Dim Units = New String() {"j/kg", "kj/kg", "j/g", "btu/lb"}
        Dim factor = New Double() {1.0, 1000.0, 1 / gram, btu / pound}

        indexfrom = Array.IndexOf(Units, LCase(fromUnit))
        indexTo = Array.IndexOf(Units, LCase(ToUnit))

        If IsNumeric(val) Then
            Result = val * factor(indexfrom) / factor(indexTo)
        End If

        LatentHeatConvert = Result
    End Function

    <ExcelFunction(Description:="convert UOM of length measurement", Category:="GCME E-PT | Unit Convertion")>
    Public Function LengthConvert(<ExcelArgument(Description:="measurement")> val As Double, <ExcelArgument(Description:="Original UOM")> fromUnit As String, <ExcelArgument(Description:="final UOM")> ToUnit As String) As Double
        Dim indexfrom As Integer
        Dim indexTo As Integer
        Dim Result As Double

        Dim Units = New String() {"km", "m", "mm", "cm", "in", "ft", "yard", "mile", "inch"}
        Dim factor = New Double() {1000.0, 1, mm, cm, inch, foot, yard, mile, inch}

        indexfrom = Array.IndexOf(Units, LCase(fromUnit))
        indexTo = Array.IndexOf(Units, LCase(ToUnit))

        If IsNumeric(val) Then
            Result = val * factor(indexfrom) / factor(indexTo)
        End If

        LengthConvert = Result
    End Function

    <ExcelFunction(Description:="convert UOM of mass measurement", Category:="GCME E-PT | Unit Convertion")>
    Public Function MassConvert(<ExcelArgument(Description:="measurement")> val As Double, <ExcelArgument(Description:="Original UOM")> fromUnit As String, <ExcelArgument(Description:="final UOM")> ToUnit As String) As Double
        Dim indexfrom As Integer
        Dim indexTo As Integer
        Dim Result As Double

        Dim Units = New String() {"kg", "g", "mg", "ton", "lb", "ounce"}
        Dim factor = New Double() {1, gram, 0.000001, 1000.0, pound, oz}

        indexfrom = Array.IndexOf(Units, LCase(fromUnit))
        indexTo = Array.IndexOf(Units, LCase(ToUnit))

        If IsNumeric(val) Then
            Result = val * factor(indexfrom) / factor(indexTo)
        End If

        MassConvert = Result
    End Function

    <ExcelFunction(Description:="convert UOM of mass density measurement", Category:="GCME E-PT | Unit Convertion")>
    Public Function MassDensityConvert(<ExcelArgument(Description:="measurement")> val As Double, <ExcelArgument(Description:="Original UOM")> fromUnit As String, <ExcelArgument(Description:="final UOM")> ToUnit As String) As Double
        Dim indexfrom As Integer
        Dim indexTo As Integer
        Dim Result As Double

        Dim Units = New String() {"kg/m3", "lb/ft3", "g/l", "g/ml", "g/cm3"}
        Dim factor = New Double() {1, pound / foot3, gram / 0.001, gram / 0.000001, gram / 0.000001}

        indexfrom = Array.IndexOf(Units, LCase(fromUnit))
        indexTo = Array.IndexOf(Units, LCase(ToUnit))

        If IsNumeric(val) Then
            Result = val * factor(indexfrom) / factor(indexTo)
        End If

        MassDensityConvert = Result
    End Function

    <ExcelFunction(Description:="convert UOM of mass heat capacity measurement", Category:="GCME E-PT | Unit Convertion")>
    Public Function MassHeatcapacityConvert(<ExcelArgument(Description:="measurement")> val As Double, <ExcelArgument(Description:="Original UOM")> fromUnit As String, <ExcelArgument(Description:="final UOM")> ToUnit As String) As Double
        Dim indexfrom As Integer
        Dim indexTo As Integer
        Dim Result As Double

        Dim Units = New String() {"kj/kg.c", "kj/g.c", "j/g.c", "j/kg.c", "cal/kg.c", "cal/g.c", "kcal/kg.c", "kcal/g.c", "btu/lb.f"}
        Dim factor = New Double() {1000.0, 1000000.0, 1000.0, 1, calorie, calorie * 1000.0, 1000.0 * calorie, 1000000.0 * calorie, btu / (pound * degree_f)}

        indexfrom = Array.IndexOf(Units, LCase(fromUnit))
        indexTo = Array.IndexOf(Units, LCase(ToUnit))

        If IsNumeric(val) Then
            Result = val * factor(indexfrom) / factor(indexTo)
        End If

        MassHeatcapacityConvert = Result
    End Function

    <ExcelFunction(Description:="convert UOM of mass flow rate measurement", Category:="GCME E-PT | Unit Convertion")>
    Public Function MassflowRateConvert(<ExcelArgument(Description:="measurement")> val As Double, <ExcelArgument(Description:="Original UOM")> fromUnit As String, <ExcelArgument(Description:="final UOM")> ToUnit As String) As Double
        Dim indexfrom As Integer
        Dim indexTo As Integer
        Dim Result As Double

        Dim Units = New String() {"kg/day", "kg/h", "kg/s", "kg/min", "lb/h", "lb/s", "lb/min", "ton/h", "ton/day"}
        Dim factor = New Double() {1.0 / day, 1.0 / hour, 1.0, 1.0 / minute, pound / hour, pound, pound / minute, ton / hour, ton / day}

        indexfrom = Array.IndexOf(Units, LCase(fromUnit))
        indexTo = Array.IndexOf(Units, LCase(ToUnit))

        If IsNumeric(val) Then
            Result = val * factor(indexfrom) / factor(indexTo)
        End If

        MassflowRateConvert = Result
    End Function

    <ExcelFunction(Description:="convert UOM of power measurement", Category:="GCME E-PT | Unit Convertion")>
    Public Function PowerConvert(<ExcelArgument(Description:="measurement")> val As Double, <ExcelArgument(Description:="Original UOM")> fromUnit As String, <ExcelArgument(Description:="final UOM")> ToUnit As String) As Double
        Dim indexfrom As Integer
        Dim indexTo As Integer
        Dim Result As Double

        Dim Units = New String() {"j/s", "kj/s", "kj/min", "kj/h", "w", "kw", "mw", "kcal/h", "kcal/min", "kcal/s",
                        "cal/h", "cal/min", "cal/s", "btu/h", "btu/hr", "hp"}
        Dim factor = New Double() {1, 1000.0, 1000.0 / minute, 1000.0 / hour, 1, 1000.0, 1000000.0, 1000.0 * calorie / hour, 1000.0 * calorie / minute, 1000.0 * calorie,
                        calorie / hour, calorie / minute, calorie, btu / hour, btu / hour, horsepower}

        indexfrom = Array.IndexOf(Units, LCase(fromUnit))
        indexTo = Array.IndexOf(Units, LCase(ToUnit))

        If IsNumeric(val) Then
            Result = val * factor(indexfrom) / factor(indexTo)
        End If

        PowerConvert = Result
    End Function

    <ExcelFunction(Description:="convert UOM of pressure measurement", Category:="GCME E-PT | Unit Convertion")>
    Public Function PressureConvert(<ExcelArgument(Description:="measurement")> val As Double, <ExcelArgument(Description:="Original UOM")> fromUnit As String, <ExcelArgument(Description:="final UOM")> ToUnit As String) As Double
        Dim indexfrom As Integer
        Dim indexTo As Integer
        Dim Result As Double

        Dim Units = New String() {"atm", "bar", "barg", "inhg", "kg/cm2", "kg/cm2g", "kpa", "kpa(a}", "kpag", "mbar",
                      "mmh2o", "mmhg", "mpa", "mpag", "pa", "psi", "psig", "torr", "psia", "psi(a)", "psi(g)", "bara", "bar(a)", "inh2o"}
        Dim factor = New Double() {atm, bar, bar, inchHg, kgf / cm2, kgf / cm2, 1000.0, 1000.0, 1000.0, 100.0,
                    G * mm * 1000.0, torr, 1000000.0, 1000000.0, 1, psi, psi, torr, psi, psi, psi, bar, bar, G * inch * 1000.0}
        Dim Offset = New Double() {0, 0, atm, 0, 0, atm, 0, 0, atm, 0,
                        0, 0, 0, atm, 0, 0, atm, 0, 0, 0, atm, 0, 0, 0}

        indexfrom = Array.IndexOf(Units, LCase(fromUnit))
        indexTo = Array.IndexOf(Units, LCase(ToUnit))

        If IsNumeric(val) Then
            Result = ((val * factor(indexfrom) + Offset(indexfrom)) - Offset(indexTo)) / factor(indexTo)
        End If

        PressureConvert = Result

    End Function

    <ExcelFunction(Description:="convert UOM of pressure per length measurement", Category:="GCME E-PT | Unit Convertion")>
    Public Function PressperlenConvert(<ExcelArgument(Description:="measurement")> val As Double, <ExcelArgument(Description:="Original UOM")> fromUnit As String, <ExcelArgument(Description:="final UOM")> ToUnit As String) As Double
        Dim indexfrom As Integer
        Dim indexTo As Integer
        Dim Result As Double

        Dim Units = New String() {"kpa/100m", "bar/100m", "psi/100ft", "pa/m", "pa/100m"}
        Dim factor = New Double() {10.0, bar / 100.0, psi / (100.0 * foot), 1.0, 1 / 100.0}

        indexfrom = Array.IndexOf(Units, LCase(fromUnit))
        indexTo = Array.IndexOf(Units, LCase(ToUnit))

        If IsNumeric(val) Then
            Result = val * factor(indexfrom) / factor(indexTo)
        End If

        PressperlenConvert = Result
    End Function

    <ExcelFunction(Description:="convert UOM of temperature measurement", Category:="GCME E-PT | Unit Convertion")>
    Public Function TemperatureConvert(<ExcelArgument(Description:="measurement")> val As Double, <ExcelArgument(Description:="Original UOM")> fromUnit As String, <ExcelArgument(Description:="final UOM")> ToUnit As String) As Double
        Dim indexfrom As Integer
        Dim indexTo As Integer
        Dim Result As Double

        Dim Units = New String() {"c", "f", "k", "r"}
        Dim factor = New Double() {1.0, degree_f, 1.0, degree_f}
        Dim Offset = New Double() {zero_celsius, zero_celsius - 32.0 * degree_f, 0, 0}

        indexfrom = Array.IndexOf(Units, LCase(fromUnit))
        indexTo = Array.IndexOf(Units, LCase(ToUnit))

        If IsNumeric(val) Then
            Result = ((val * factor(indexfrom) + Offset(indexfrom)) - Offset(indexTo)) / factor(indexTo)
        Else
            Result = TemperatureConvert(0, "K", ToUnit)
        End If

        TemperatureConvert = Result
    End Function

    <ExcelFunction(Description:="convert UOM of temperature invert measurement", Category:="GCME E-PT | Unit Convertion")>
    Public Function TemperatureInvConvert(<ExcelArgument(Description:="measurement")> val As Double, <ExcelArgument(Description:="Original UOM")> fromUnit As String, <ExcelArgument(Description:="final UOM")> ToUnit As String) As Double
        Dim indexfrom As Integer
        Dim indexTo As Integer
        Dim Result As Double

        Dim Units = New String() {"1/c", "1/f", "1/k", "1/r"}
        Dim factor = New Double() {1.0, 1.0 / degree_f, 1.0, 1.0 / degree_f}

        indexfrom = Array.IndexOf(Units, LCase(fromUnit))
        indexTo = Array.IndexOf(Units, LCase(ToUnit))

        If IsNumeric(val) Then
            Result = val * factor(indexfrom) / factor(indexTo)
        End If

        TemperatureInvConvert = Result
    End Function

    <ExcelFunction(Description:="convert UOM of thermal conductivity measurement", Category:="GCME E-PT | Unit Convertion")>
    Public Function ThermalconductivityConvert(<ExcelArgument(Description:="measurement")> val As Double, <ExcelArgument(Description:="Original UOM")> fromUnit As String, <ExcelArgument(Description:="final UOM")> ToUnit As String) As Double
        Dim indexfrom As Integer
        Dim indexTo As Integer
        Dim Result As Double

        Dim Units = New String() {"w/m.K", "btu/hr.ft.f", "btu/h.ft.f", "Kcal/m.hr.c", "cal/s.cm.c"}
        Dim factor = New Double() {1.0, btu / hour / foot / degree_f, btu / hour / foot / degree_f, 1000.0 * calorie / hour, calorie / 0.01}

        indexfrom = Array.IndexOf(Units, LCase(fromUnit))
        indexTo = Array.IndexOf(Units, LCase(ToUnit))

        If IsNumeric(val) Then
            Result = val * factor(indexfrom) / factor(indexTo)
        End If

        ThermalconductivityConvert = Result
    End Function

    <ExcelFunction(Description:="convert UOM of UA measurement", Category:="GCME E-PT | Unit Convertion")>
    Public Function UAConvert(<ExcelArgument(Description:="measurement")> val As Double, <ExcelArgument(Description:="Original UOM")> fromUnit As String, <ExcelArgument(Description:="final UOM")> ToUnit As String) As Double
        Dim indexfrom As Integer
        Dim indexTo As Integer
        Dim Result As Double

        Dim Units = New String() {"kj/c.h", "kj/c.s", "w/c", "btu/f.h", "btu/f.hr", "kcal/c.h"}
        Dim factor = New Double() {1000.0 / hour, 1000.0, 1.0, btu / degree_f / hour, btu / degree_f / hour, 1000.0 * calorie / hour}

        indexfrom = Array.IndexOf(Units, LCase(fromUnit))
        indexTo = Array.IndexOf(Units, LCase(ToUnit))

        If IsNumeric(val) Then
            Result = val * factor(indexfrom) / factor(indexTo)
        End If

        UAConvert = Result
    End Function

    <ExcelFunction(Description:="convert UOM of velocity measurement", Category:="GCME E-PT | Unit Convertion")>
    Public Function VelocityConvert(<ExcelArgument(Description:="measurement")> val As Double, <ExcelArgument(Description:="Original UOM")> fromUnit As String, <ExcelArgument(Description:="final UOM")> ToUnit As String) As Double
        Dim indexfrom As Integer
        Dim indexTo As Integer
        Dim Result As Double

        Dim Units = New String() {"m/h", "m/s", "m/min", "ft/h", "ft/s", "ft/min", "in/s"}
        Dim factor = New Double() {1.0 / hour, 1.0, 1.0 / minute, foot / hour, foot, foot / minute, inch}

        indexfrom = Array.IndexOf(Units, LCase(fromUnit))
        indexTo = Array.IndexOf(Units, LCase(ToUnit))

        If IsNumeric(val) Then
            Result = val * factor(indexfrom) / factor(indexTo)
        End If

        VelocityConvert = Result
    End Function

    <ExcelFunction(Description:="convert UOM of dynamic viscosity measurement", Category:="GCME E-PT | Unit Convertion")>
    Public Function ViscosityConvert(<ExcelArgument(Description:="measurement")> val As Double, <ExcelArgument(Description:="Original UOM")> fromUnit As String, <ExcelArgument(Description:="final UOM")> ToUnit As String) As Double
        Dim indexfrom As Integer
        Dim indexTo As Integer
        Dim Result As Double

        Dim Units = New String() {"Pa-s", "Poise", "cP", "kg/m.s", "N-s/m2", "Pa.s", "dyne-s/cm2"}
        Dim factor = New Double() {1, 0.1, 0.001, G, 1, 1, 0.1}

        indexfrom = Array.IndexOf(Units, LCase(fromUnit))
        indexTo = Array.IndexOf(Units, LCase(ToUnit))

        If IsNumeric(val) Then
            Result = val * factor(indexfrom) / factor(indexTo)
        End If

        ViscosityConvert = Result
    End Function

    <ExcelFunction(Description:="convert UOM of kinematic viscosity measurement", Category:="GCME E-PT | Unit Convertion")>
    Public Function KinematicViscosityConvert(<ExcelArgument(Description:="measurement")> val As Double, <ExcelArgument(Description:="Original UOM")> fromUnit As String, <ExcelArgument(Description:="final UOM")> ToUnit As String) As Double
        Dim indexfrom As Integer
        Dim indexTo As Integer
        Dim Result As Double

        Dim Units = New String() {"St", "cSt", "cm2/s", "m2/s", "m2/h", "mm2/s", "ft2/s", "ft2/h", "in2/s"}
        Dim factor = New Double() {1000.0, 1000000.0, cm2, 1, 1 / hour, mm2, foot2, foot2 / hour, inch2}

        indexfrom = Array.IndexOf(Units, LCase(fromUnit))
        indexTo = Array.IndexOf(Units, LCase(ToUnit))

        If IsNumeric(val) Then
            Result = val * factor(indexfrom) / factor(indexTo)
        End If

        KinematicViscosityConvert = Result
    End Function

    <ExcelFunction(Description:="convert UOM of volume measurement", Category:="GCME E-PT | Unit Convertion")>
    Public Function VolumeConvert(<ExcelArgument(Description:="measurement")> val As Double, <ExcelArgument(Description:="Original UOM")> fromUnit As String, <ExcelArgument(Description:="final UOM")> ToUnit As String) As Double
        Dim indexfrom As Integer
        Dim indexTo As Integer
        Dim Result As Double

        Dim Units = New String() {"m3", "cm3", "mm3", "l", "liter", "litre", "ml", "ft3", "in3", "gal", "gallon", "pint"}
        Dim factor = New Double() {1.0, cm3, mm3, liter, liter, liter, mm * liter, foot3, inch3, gallon, gallon, pint}

        indexfrom = Array.IndexOf(Units, LCase(fromUnit))
        indexTo = Array.IndexOf(Units, LCase(ToUnit))

        If IsNumeric(val) Then
            Result = val * factor(indexfrom) / factor(indexTo)
        End If

        VolumeConvert = Result
    End Function

    <ExcelFunction(Description:="convert UOM of volumetric flow rate measurement", Category:="GCME E-PT | Unit Convertion")>
    Public Function VolumeflowRateConvert(<ExcelArgument(Description:="measurement")> val As Double, <ExcelArgument(Description:="Original UOM")> fromUnit As String, <ExcelArgument(Description:="final UOM")> ToUnit As String) As Double
        Dim indexfrom As Integer
        Dim indexTo As Integer
        Dim Result As Double

        Dim Units = New String() {"m3/h", "m3/s", "m3/min", "ft3/h", "ft3/s", "ft3/min", "gal/min", "l/min", "mmscfd", "scfd"}
        Dim factor = New Double() {1.0 / hour, 1.0, 1.0 / minute, foot3 / hour, foot3, foot3 / minute, gallon / minute, liter / minute, 1000000.0 * foot3 / day, foot3 / day}

        indexfrom = Array.IndexOf(Units, LCase(fromUnit))
        indexTo = Array.IndexOf(Units, LCase(ToUnit))

        If IsNumeric(val) Then
            Result = val * factor(indexfrom) / factor(indexTo)
        End If

        VolumeflowRateConvert = Result
    End Function

    <ExcelFunction(Description:="convert UOM of work (engergy) measurement", Category:="GCME E-PT | Unit Convertion")>
    Public Function WorkConvert(<ExcelArgument(Description:="measurement")> val As Double, <ExcelArgument(Description:="Original UOM")> fromUnit As String, <ExcelArgument(Description:="final UOM")> ToUnit As String) As Double
        Dim indexfrom As Integer
        Dim indexTo As Integer
        Dim Result As Double

        Dim Units = New String() {"kw", "hp"}
        Dim factor = New Double() {1000.0, horsepower}

        indexfrom = Array.IndexOf(Units, LCase(fromUnit))
        indexTo = Array.IndexOf(Units, LCase(ToUnit))

        If IsNumeric(val) Then
            Result = val * factor(indexfrom) / factor(indexTo)
        End If

        WorkConvert = Result
    End Function

    <ExcelFunction(Description:="convert gas mass flow rate to cubic foot per minute", Category:="GCME E-PT | Unit Convertion")>
    Public Function MassFlow2CFM(<ExcelArgument(Description:="Gas mass flow rate, [lb/min]")> m As Double, <ExcelArgument(Description:="Pressure, [psia]")> p As Double, <ExcelArgument(Description:="Temperature, [F]")> T As Double, <ExcelArgument(Description:="Gas molecular weight")> Optional MW As Double = 28.964, <ExcelArgument(Description:="Gas compressibility")> Optional Z As Double = 1.0) As Double
        MassFlow2CFM = (m * (1545.0 / MW) * TemperatureConvert(T, "F", "R") * Z) / (144.0 * p)
    End Function

    <ExcelFunction(Description:="convert gas cubic foot per minute to mass flow rate", Category:="GCME E-PT | Unit Convertion")>
    Public Function CFM2MassFlow(<ExcelArgument(Description:="Gas volumetric flow rate, [ft2/min]")> q As Double, <ExcelArgument(Description:="Pressure, [psia]")> p As Double, <ExcelArgument(Description:="Temperature, [F]")> T As Double, <ExcelArgument(Description:="Gas molecular weight")> Optional MW As Double = 28.964, <ExcelArgument(Description:="Gas compressibility")> Optional Z As Double = 1.0) As Double
        CFM2MassFlow = (144.0 * p * q) / ((1545.0 / MW) * TemperatureConvert(T, "F", "R") * Z)
    End Function

    <ExcelFunction(Description:="convert gas SCFM to pound per minute", Category:="GCME E-PT | Unit Convertion")>
    Public Function SCFM2MassFlow(<ExcelArgument(Description:="Gas volumetric flow rate, [SCFM]")> q As Double, <ExcelArgument(Description:="Gas molecular weight")> Optional MW As Double = 28.964, <ExcelArgument(Description:="Gas comperssibilty")> Optional Z As Double = 1.0) As Double
        SCFM2MassFlow = CFM2MassFlow(q, 14.696, 60, MW, Z) '144 * 14.696 * q / (1545.0 / MW * 60.0 * Z)
    End Function

    <ExcelFunction(Description:="convert gas SCFM to ACFM", Category:="GCME E-PT | Unit Convertion")>
    Public Function SCFM2ACFM(<ExcelArgument(Description:="Gas volumetric flow rate, [SCFM]")> q As Double, <ExcelArgument(Description:="Pressure, [psia]")> p As Double, <ExcelArgument(Description:="Temperature, [F]")> T As Double) As Double
        SCFM2ACFM = q * (14.696 / TemperatureConvert(60, "F", "R")) / (p / TemperatureConvert(T, "F", "R"))
    End Function

    <ExcelFunction(Description:="convert UOM -", Category:="GCME E-PT | Unit Convertion")>
    Public Function GCME_CONVERT(<ExcelArgument(Description:="measurement")> val As Double, <ExcelArgument(Description:="Original UOM")> fromUnit As String, <ExcelArgument(Description:="final UOM")> toUnit As String) As Double

CUSTOM:
        Dim UOM_ALL(,) = {
            {"s", "sec", "second", "", "min", "minute", "hr", "hrs", "hour", "hours", "day", "days", "yr", "year", "years", "week", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
            {"m", "km", "cm", "mm", "nm", "in", "ft", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
            {"kg", "g", "mg", "ton", "lb", "ounce", "pound", "stone", "oz", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
            {"A", "am", "amp", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
            {"K", "C", "F", "R", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
            {"mol", "mole", "kmol", "kmole", "lb-mol", "g-mol", "kg-mol", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
            {"m2", "cm2", "mm2", "in2", "ft2", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
            {"m3", "cm3", "mm3", "l", "liter", "litre", "ml", "in3", "ft3", "gal", "gallon", "pint", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
            {"Pa", "kPa", "kPag", "Pag", "ATM", "bar", "barg", "bara", "bar(a)", "kg/cm2", "kg/cm2g", "kcs", "kcsg", "psi", "psia", "psi(a)", "psig", "psi(g)", "torr", "mmHg", "mmH2O", "inchH2O", "inH2O", "", "", "", "", "", "", "", ""},
            {"J", "kJ", "MJ", "GJ", "cal", "kcal", "Mcal", "MMkcal", "BTU", "kW-h", "kW-hr", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
            {"kW/m2", " BTU/h.ft2", " J/m2.s", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
            {"J/kg", "kJ/kg", "J/g", "BTU/lb", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
            {"kg/m3", "lb/ft3", "g/L", "g/mL", "g/cm3", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
            {"kJ/kg.C", "kJ/g.C", "j/g.C", "j/kg.C", "cal/kg.C", "cal/g.C", "kcal/kg.C", "kcal/g.C", "BTU/lb.F", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
            {"kg/day", "kg/h", "kg/s", "kg/min", "lb/h", "lb/s", "lb/min", "ton/h", "ton/day", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
            {"J/s", "kJ/s", "kJ/min", "kJ/h", "W", "kW", "MW", "kcal/h", "kcal/min", "kcal/s", "cal/h", "cal/min", "cal/s", "BTU/h", "BTU/hr", "hp", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
            {"kPa/100m", "bar/100m", "psi/100ft", "Pa/m", "Pa/100m", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
            {"W/m.K", "BTU/hr.ft.F", "BTU/h.ft.F", "Kcal/m.hr.C", "cal/s.cm.C", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
            {"kJ/C.h", "kJ/C.s", "W/C", "BTU/F.h", "BTU/F.hr", "kcal/C.h", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
            {"m/h", "m/s", "m/min", "ft/h", "ft/s", "ft/min", "in/s", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
            {"Pa-s", "Poise", "cP", "kg/m.s", "N-s/m2", "Pa.s", "dyne-s/cm2", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
            {"St", "cSt", "cm2/s", "m2/s", "m2/h", "mm2/s", "ft2/s", "ft2/h", "in2/s", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
            {"m3/h", "m3/s", "m3/min", "ft3/h", "ft3/s", "ft3/min", "gal/min", "l/min", "mmscfd", "scfd", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""}}

        Dim UOM_type = {"time", "length", "mass", "electric_current", "temperature", "amount_of_substance", "area", "volume", "pressure", "energy",
            "heat_intensity", "latent_heat", "mass_density", "mass_heat_capacity", "mass_flow_rate", "power", "pressure_per_length", "thermal_conductivity", "ua", "velocity",
            "viscosity_dynamic", "viscosity_kinetic", "volume_flow_rate"}

        Dim idx As Integer
        Dim i, j

        For i = 0 To 22
            For j = 0 To 22
                If LCase(fromUnit) = LCase(UOM_ALL(i, j)) Then
                    idx = i
                    GoTo 1000
                End If
            Next
        Next

        If idx = 0 Then
            GCME_CONVERT = -9999
            Exit Function
        End If
1000:
        Select Case idx
            Case 1
                GCME_CONVERT = LengthConvert(val, fromUnit, toUnit)
            Case 2
                GCME_CONVERT = MassConvert(val, fromUnit, toUnit)
            Case 4
                GCME_CONVERT = TemperatureConvert(val, fromUnit, toUnit)
            Case 6
                GCME_CONVERT = AreaConvert(val, fromUnit, toUnit)
            Case 7
                GCME_CONVERT = VolumeConvert(val, fromUnit, toUnit)
            Case 8
                GCME_CONVERT = PressureConvert(val, fromUnit, toUnit)
            Case 9
                GCME_CONVERT = EnergyConvert(val, fromUnit, toUnit)
            Case 10
                GCME_CONVERT = HeatIntensityConvert(val, fromUnit, toUnit)
            Case 11
                GCME_CONVERT = LatentHeatConvert(val, fromUnit, toUnit)
            Case 12
                GCME_CONVERT = MassDensityConvert(val, fromUnit, toUnit)
            Case 13
                GCME_CONVERT = MassHeatcapacityConvert(val, fromUnit, toUnit)
            Case 14
                GCME_CONVERT = MassflowRateConvert(val, fromUnit, toUnit)
            Case 15
                GCME_CONVERT = PowerConvert(val, fromUnit, toUnit)
            Case 16
                GCME_CONVERT = PressperlenConvert(val, fromUnit, toUnit)
            Case 17
                GCME_CONVERT = ThermalconductivityConvert(val, fromUnit, toUnit)
            Case 18
                GCME_CONVERT = UAConvert(val, fromUnit, toUnit)
            Case 19
                GCME_CONVERT = VelocityConvert(val, fromUnit, toUnit)
            Case 20
                GCME_CONVERT = ViscosityConvert(val, fromUnit, toUnit)
            Case 21
                GCME_CONVERT = KinematicViscosityConvert(val, fromUnit, toUnit)
            Case 22
                GCME_CONVERT = VolumeflowRateConvert(val, fromUnit, toUnit)
            Case Else
                GCME_CONVERT = -999
        End Select

    End Function

End Module
