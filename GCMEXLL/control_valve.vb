Imports System.Security.Cryptography
Imports ExcelDna.Integration

Public Module control_valve
    Const MWair As Double = 28.96443

    Private Function Valve_fx_linear(x As Double) As Double
        Valve_fx_linear = x
    End Function

    Private Function Valve_fx_quickopening(x As Double) As Double
        Valve_fx_quickopening = Math.Sqrt(x)
    End Function

    Private Function Valve_fx_equalpercentage(x As Double, Optional R As Double = 0.0) As Double
        If R < 20.0 Or R > 50.0 Then R = 50.0
        Valve_fx_equalpercentage = Math.Pow(R, x - 1)
    End Function

    <ExcelFunction(Description:="determine vale characteristic Cv from percent openning (0.0-1.0)", Category:="GCME E-PT | Control Valve")>
    Public Function Valve_fx(<ExcelArgument(Description:="percent openning, 0.0-1.0")> x As Double,
                             <ExcelArgument(Description:="valve type [0-Linear, 1-Quick Opening, 2-Equal Percentage")> Optional type As Integer = 0,
                             <ExcelArgument(Description:="Equal percentage constant 20 < R < 50")> Optional R As Double = 0.0) As Double
        Select Case type
            Case 1
                Valve_fx = Valve_fx_quickopening(x)
            Case 2
                If R < 20 Or R > 50 Then R = 50.0
                Valve_fx = Valve_fx_equalpercentage(x, R)
            Case Else
                Valve_fx = Valve_fx_linear(x)
        End Select
    End Function

    <ExcelFunction(Description:="determine flow coefficient Cv for liquid flow", Category:="GCME E-PT | Control Valve")>
    Public Function LiquidCV(<ExcelArgument(Description:="volumetric flow rate, [gal/min]")> q As Double,
                              <ExcelArgument(Description:="upstream pressure, [psi]")> P1 As Double,
                              <ExcelArgument(Description:="downstream pressure, [psi]")> P2 As Double,
                              <ExcelArgument(Description:="specific gravity at flowing temperature (water = 1)")> Gf As Double,
                              <ExcelArgument(Description:="vapor pressure of liquid at flowing temperature, [psia]")> Optional Pv As Double = 0.0,
                              <ExcelArgument(Description:="critical pressure of liquid, [psia]")> Optional Pc As Double = 0.0,
                              <ExcelArgument(Description:="pressure at the vena contracta in the valve, [psia]")> Optional Pvc As Double = 0.0,
                              <ExcelArgument(Description:="valve end inside diameter, [inch]")> Optional d As Double = 0.0,
                              <ExcelArgument(Description:="inside diameter of upstream pipe, [inch]")> Optional D1 As Double = 0.0,
                              <ExcelArgument(Description:="inside diameter of downstream pipe, [inch]")> Optional D2 As Double = 0.0) As Double
        Dim F_F As Double = 1.0
        Dim F_L As Double = 1.0
        Dim F_P As Double
        Dim old As Double
        Dim CV As Double
        Dim Ktotal As Double
        Dim tol As Double = 0.0000000001

        If Math.Abs(P1 - P2) <= tol Then
            CV = 0.0
        Else
            If Pv > 0 And Pc > 0 Then
                F_F = 0.96 - 0.28 * Math.Sqrt(Pv / Pc)
            End If

            If Pvc > 0 Then
                F_L = Math.Sqrt((P1 - P2) / (P1 - Pvc))
            End If


            If d > 0 And D1 > 0 And D2 > 0 Then
                '
                F_P = 1.0
                Do
                    ' save old calculation
                    old = F_P
                    ' calc new cv
                    If Math.Abs(P1 - P2) >= F_L * F_L * (P1 - F_F * Pv) Then
                        CV = q / (F_L * F_P) * Math.Sqrt(Gf / (P1 - F_F * Pv))
                    Else
                        CV = q / F_P * Math.Sqrt(Gf / (P1 - P2))
                    End If
                    ' effect of pipe reducers = loss coefficients + Bernoulli coefficients
                    Ktotal = 0.5 * Math.Pow(1 - (d / D1) * (d / D1), 2) + Math.Pow(1 - (d / D2) * (d / D2), 2) + (1 - Math.Pow(d / D1, 4)) + (1 - Math.Pow(d / D2, 4))
                    ' calculate new F_P
                    F_P = 1 / Math.Sqrt((CV * CV * Ktotal) / (890.0 * Math.Pow(d, 4)) + 1)
                Loop While Math.Abs(F_P - old) > tol
            Else
                F_P = 1.0
            End If

            If Math.Abs(P1 - P2) >= F_L * F_L * (P1 - F_F * Pv) Then
                CV = q * Math.Sqrt(Gf / (P1 - F_L * Pv)) / (F_P * F_L)
            Else
                CV = q / F_P * Math.Sqrt(Gf / (P1 - P2))
            End If

        End If

        LiquidCV = CV
    End Function

    <ExcelFunction(Description:="determine flow coefficient Cv for gas and vapor flow", Category:="GCME E-PT | Control Valve")>
    Public Function VaporCV(<ExcelArgument(Description:="volumetric flow rate, [SCFH]")> q As Double,
                            <ExcelArgument(Description:="upstream pressure, [psia]")> P1 As Double,
                            <ExcelArgument(Description:="downstream pressure, [psia]")> P2 As Double,
                            <ExcelArgument(Description:="inlet temperature, [F]")> T1 As Double,
                            <ExcelArgument(Description:="gas molecular weight")> Optional M As Double = MWair,
                            <ExcelArgument(Description:="gas compressibilty factor")> Optional Z As Double = 1.0,
                            <ExcelArgument(Description:="gas specific heat ratio, Cp/Cv")> Optional k As Double = 1.4,
                            <ExcelArgument(Description:="pressure drop ratio factor (default = 1.0)")> Optional XT As Double = 0.0,
                            <ExcelArgument(Description:="valve end inside diameter, [inch]")> Optional d As Double = 0.0,
                            <ExcelArgument(Description:="inside diameter of upstream pipe, [inch]")> Optional D1 As Double = 0.0,
                            <ExcelArgument(Description:="inside diameter of downstream pipe, [inch]")> Optional D2 As Double = 0.0) As Double
        If M = 0.0 Then M = MWair
        If Z = 0.0 Then Z = 0.99
        If k = 0.0 Then k = 1.4
        If XT = 0.0 Then XT = 0.7

        Dim Gg As Double = M / MWair
        Dim Fk As Double = k / 1.4
        Dim x As Double = (P1 - P2) / P1
        Dim Y As Double = 1 - x / (3 * Fk * XT)
        Dim T As Double = UOM_CONVERT(T1, "F", "R")
        Dim F_P As Double
        Dim old As Double
        Dim CV As Double
        Dim Ktotal As Double
        Dim tol As Double = 0.0000000001
        Dim Xeff As Double

        If P1 <= 0 Or Math.Abs(P1 - P2) <= 0 Then
            CV = 0.0
        Else
            ' check for choke flow
            Xeff = Math.Min(x, Fk * XT)

            If d > 0 And D1 > 0 And D2 > 0 Then
                '
                F_P = 1.0
                Do
                    ' save old calculation
                    old = F_P
                    ' calc new cv
                    CV = q / (1360.0 * F_P * P1 * Y) * Math.Sqrt((Gg * T * Z) / Xeff)
                    ' effect of pipe reducers = loss coefficients + Bernoulli coefficients
                    Ktotal = 0.5 * Math.Pow(1 - (d / D1) * (d / D1), 2) + Math.Pow(1 - (d / D2) * (d / D2), 2) + (1 - Math.Pow(d / D1, 4)) + (1 - Math.Pow(d / D2, 4))
                    ' calculate new F_P
                    F_P = 1 / Math.Sqrt((CV * CV * Ktotal) / (890.0 * Math.Pow(d, 4)) + 1)
                Loop While Math.Abs(F_P - old) > tol
            Else
                F_P = 1.0
                CV = q / (1360.0 * F_P * P1 * Y) * Math.Sqrt((Gg * T * Z) / Xeff)
            End If

        End If

        VaporCV = CV
    End Function
End Module
