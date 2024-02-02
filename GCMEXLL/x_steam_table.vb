Imports ExcelDna.Integration

Public Module x_steam_table

    '***********************************************************************************************************
    '* Water and steam properties according to IAPWS IF-97                                                     *
    '* By Magnus Holmgren, www.x-eng.com                                                                       *
    '* The steam tables are free and provided as is.                                                           *
    '* We take no responsibilities for any errors in the code or damage thereby.                               *
    '* You are free to use, modify and distribute the code as long as authorship is properly acknowledged.     *
    '* Please notify me at magnus@x-eng.com if the code is used in commercial applications                     *
    '***********************************************************************************************************'
    '
    ' The code is also avalibale for matlab at www.x-eng.com
    '
    '*Contents.
    '*1 Calling functions
    '*1.1
    '*1.2 Temperature (T)
    '*1.3 Pressure (p)
    '*1.4 Enthalpy (h)
    '*1.5 Specific Volume (v)
    '*1.6 Density (rho)
    '*1.7 Specific entropy (s)
    '*1.8 Specific internal energy (u)
    '*1.9 Specific isobaric heat capacity (Cp)
    '*1.10 Specific isochoric heat capacity (Cv)
    '*1.11 Speed of sound
    '*1.12 Viscosity
    '*1.13 Prandtl
    '*1.14 Kappa
    '*1.15 Surface tension
    '*1.16 Heat conductivity
    '*1.17 Vapour fraction
    '*1.18 Vapour Volume Fraction
    '
    '*2 IAPWS IF 97 Calling functions
    '*2.1 Functions for region 1
    '*2.2 Functions for region 2
    '*2.3 Functions for region 3
    '*2.4 Functions for region 4
    '*2.5 Functions for region 5
    '
    '*3 Region Selection
    '*3.1 Regions as a function of pT
    '*3.2 Regions as a function of ph
    '*3.3 Regions as a function of ps
    '*3.4 Regions as a function of hs
    '*3.5 Regions as a function of p and rho
    '
    '4 Region Borders
    '4.1 Boundary between region 1 and 3.
    '4.2 Region 3. pSat_h and pSat_s
    '4.3 Region boundary 1to3 and 3to2 as a functions of s
    '
    '5 Transport properties
    '5.1 Viscosity (IAPWS formulation 1985)
    '5.2 Thermal Conductivity (IAPWS formulation 1985)
    '5.3 Surface Tension
    '
    '6 Units


    '***********************************************************************************************************
    '*1 Calling functions                                                                                      *
    '***********************************************************************************************************

    '***********************************************************************************************************
    '*1.1


    '***********************************************************************************************************
    '*1.2 Temperature

    <ExcelFunction(Description:="Saturated temperature as fuction of pressure, [C]", Category:="GCME E-PT | Steam Table")>
    Public Function Tsat_p(<ExcelArgument(Description:="Pressure, [bar]")> p As Double) As Double
        p = ToSIunit_p(p)
        If p >= 0.000611657 And p <= 22.06395 + 0.001 Then '0.001 Added to enable the tripple point.
            Tsat_p = FromSIunit_T(T4_p(p))
        Else
            Tsat_p = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Saturated temperature as fuction of entropy, [C]", Category:="GCME E-PT | Steam Table")>
    Public Function Tsat_s(<ExcelArgument(Description:="Entropy, [kJ/kg.K]")> s As Double) As Double
        s = ToSIunit_s(s)
        If s > -0.0001545495919 And s < 9.155759395 Then
            Tsat_s = FromSIunit_T(T4_p(P4_s(s)))
        Else
            Tsat_s = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Temperature as a function of pressure and enthalpy, [C]", Category:="GCME E-PT | Steam Table")>
    Public Function Temp_ph(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Enthalpy, [kJ/kg]")> h As Double) As Double
        p = ToSIunit_p(p)
        h = ToSIunit_h(h)
        Select Case Region_ph(p, h)
            Case 1
                Temp_ph = FromSIunit_T(T1_ph(p, h))
            Case 2
                Temp_ph = FromSIunit_T(T2_ph(p, h))
            Case 3
                Temp_ph = FromSIunit_T(T3_ph(p, h))
            Case 4
                Temp_ph = FromSIunit_T(T4_p(p))
            Case 5
                Temp_ph = FromSIunit_T(T5_ph(p, h))
            Case Else
                Temp_ph = -999.9
        End Select
    End Function

    <ExcelFunction(Description:="Temperature as a function of pressure and entropy, [C]", Category:="GCME E-PT | Steam Table")>
    Public Function Temp_ps(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Entropy, [kJ/kg.K]")> s As Double) As Double
        p = ToSIunit_p(p)
        s = ToSIunit_s(s)
        Select Case Region_ps(p, s)
            Case 1
                Temp_ps = FromSIunit_T(T1_ps(p, s))
            Case 2
                Temp_ps = FromSIunit_T(T2_ps(p, s))
            Case 3
                Temp_ps = FromSIunit_T(T3_ps(p, s))
            Case 4
                Temp_ps = FromSIunit_T(T4_p(p))
            Case 5
                Temp_ps = FromSIunit_T(T5_ps(p, s))
            Case Else
                Temp_ps = -999.9
        End Select
    End Function

    <ExcelFunction(Description:="Temperature as a function of enthalpy and entropy, [C]", Category:="GCME E-PT | Steam Table")>
    Public Function Temp_hs(<ExcelArgument(Description:="Enthalpy, [kJ/kg]")> h As Double, <ExcelArgument(Description:="Entropy, [kJ/kg.K]")> s As Double) As Double
        h = ToSIunit_h(h)
        s = ToSIunit_s(s)
        Select Case Region_hs(h, s)
            Case 1
                Temp_hs = FromSIunit_T(T1_ph(P1_hs(h, s), h))
            Case 2
                Temp_hs = FromSIunit_T(T2_ph(P2_hs(h, s), h))
            Case 3
                Temp_hs = FromSIunit_T(T3_ph(P3_hs(h, s), h))
            Case 4
                Temp_hs = FromSIunit_T(T4_hs(h, s))
            Case 5
                Temp_hs = -999.9 'Functions of hs is not implemented in region 5
            Case Else
                Temp_hs = -999.9
        End Select
    End Function
    '***********************************************************************************************************
    '*1.3 Pressure (p)

    <ExcelFunction(Description:="Saturated pressure as function oftemperature, [bar]", Category:="GCME E-PT | Steam Table")>
    Public Function Psat_T(<ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        T = ToSIunit_T(T)
        If T <= 647.096 And T > 273.15 Then
            Psat_T = FromSIunit_p(P4_T(T))
        Else
            Psat_T = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Saturated pressure as function of entropy, [bar]", Category:="GCME E-PT | Steam Table")>
    Public Function Psat_s(<ExcelArgument(Description:="Entropy, [kJ/kg.K]")> s As Double) As Double
        s = ToSIunit_s(s)
        If s > -0.0001545495919 And s < 9.155759395 Then
            Psat_s = FromSIunit_p(P4_s(s))
        Else
            Psat_s = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Pressure as a function of enthalpy and entropy, [bar]", Category:="GCME E-PT | Steam Table")>
    Public Function Pres_hs(<ExcelArgument(Description:="Enthalpy, [kJ/kg]")> h As Double, <ExcelArgument(Description:="Entropy, [kJ/kg.K]")> s As Double) As Double
        h = ToSIunit_h(h)
        s = ToSIunit_s(s)
        Select Case Region_hs(h, s)
            Case 1
                Pres_hs = FromSIunit_p(P1_hs(h, s))
            Case 2
                Pres_hs = FromSIunit_p(P2_hs(h, s))
            Case 3
                Pres_hs = FromSIunit_p(P3_hs(h, s))
            Case 4
                Pres_hs = FromSIunit_p(P4_T(T4_hs(h, s)))
            Case 5
                Pres_hs = -999.9 'Functions of hs is not implemented in region 5
            Case Else
                Pres_hs = -999.9
        End Select
    End Function

    <ExcelFunction(Description:="Pressure as a function of enthalpy and density, [bar]", Category:="GCME E-PT | Steam Table")>
    Public Function Pres_hrho(<ExcelArgument(Description:="Enthalpy, [kJ/kg]")> h As Double, <ExcelArgument(Description:="Density, [kg/m3]")> rho As Double) As Double
        '*Not valid for water or sumpercritical since water rho does not change very much with p.
        '*Uses iteration to find p.
        Dim High_Bound As Double
        Dim Low_Bound As Double
        Dim p As Double
        Dim rhos As Double
        High_Bound = FromSIunit_p(100)
        Low_Bound = FromSIunit_p(0.000611657)
        p = FromSIunit_p(10)
        rhos = 1 / Volume_ph(p, h)
        Do While Math.Abs(rho - rhos) > 0.0000001
            rhos = 1 / Volume_ph(p, h)
            If rhos >= rho Then
                High_Bound = p
            Else
                Low_Bound = p
            End If
            p = (Low_Bound + High_Bound) / 2
        Loop
        Pres_hrho = p
    End Function


    '***********************************************************************************************************
    '*1.4 Enthalpy (h)
    <ExcelFunction(Description:="Saturated vapor enthalpy as function ofpressure, [kJ/kg]", Category:="GCME E-PT | Steam Table")>
    Public Function EnthalpyV_p(<ExcelArgument(Description:="Pressure, [bar]")> p As Double) As Double
        p = ToSIunit_p(p)
        If p > 0.000611657 And p < 22.06395 Then
            EnthalpyV_p = FromSIunit_h(H4V_p(p))
        Else
            EnthalpyV_p = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Saturated liquid enthalpy as function ofpressure, [kJ/kg]", Category:="GCME E-PT | Steam Table")>
    Public Function EnthalpyL_p(<ExcelArgument(Description:="Pressure, [bar]")> p As Double) As Double
        p = ToSIunit_p(p)
        If p > 0.000611657 And p < 22.06395 Then
            EnthalpyL_p = FromSIunit_h(H4L_p(p))
        Else
            EnthalpyL_p = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Saturated vapor enthalpy as function oftemperature, [kJ/kg]", Category:="GCME E-PT | Steam Table")>
    Public Function EnthalpyV_T(<ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        T = ToSIunit_T(T)
        If T > 273.15 And T < 647.096 Then
            EnthalpyV_T = FromSIunit_h(H4V_p(P4_T(T)))
        Else
            EnthalpyV_T = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Saturated liquid enthalpy as function oftemperature, [kJ/kg]", Category:="GCME E-PT | Steam Table")>
    Public Function EnthalpyL_T(<ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        T = ToSIunit_T(T)
        If T > 273.15 And T < 647.096 Then
            EnthalpyL_T = FromSIunit_h(H4L_p(P4_T(T)))
        Else
            EnthalpyL_T = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Enthalpy as function of pressure and temperature, [kJ/kg]", Category:="GCME E-PT | Steam Table")>
    Public Function Enthalpy_pT(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        p = ToSIunit_p(p)
        T = ToSIunit_T(T)
        Select Case Region_pT(p, T)
            Case 1
                Enthalpy_pT = FromSIunit_h(H1_pT(p, T))
            Case 2
                Enthalpy_pT = FromSIunit_h(H2_pT(p, T))
            Case 3
                Enthalpy_pT = FromSIunit_h(H3_pT(p, T))
            Case 4
                Enthalpy_pT = -999.9
            Case 5
                Enthalpy_pT = FromSIunit_h(H5_pT(p, T))
            Case Else
                Enthalpy_pT = -999.9
        End Select
    End Function

    <ExcelFunction(Description:="Enthalpy as function of pressure and enthalpy, [kJ/kg]", Category:="GCME E-PT | Steam Table")>
    Public Function Enthalpy_ps(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Entropy, [kJ/kg.K]")> s As Double) As Double
        Dim xs As Double
        p = ToSIunit_p(p)
        s = ToSIunit_s(s)
        Select Case Region_ps(p, s)
            Case 1
                Enthalpy_ps = FromSIunit_h(H1_pT(p, T1_ps(p, s)))
            Case 2
                Enthalpy_ps = FromSIunit_h(H2_pT(p, T2_ps(p, s)))
            Case 3
                Enthalpy_ps = FromSIunit_h(H3_rhoT(1 / V3_ps(p, s), T3_ps(p, s)))
            Case 4
                xs = X4_ps(p, s)
                Enthalpy_ps = FromSIunit_h(xs * H4V_p(p) + (1 - xs) * H4L_p(p))
            Case 5
                Enthalpy_ps = FromSIunit_h(H5_pT(p, T5_ps(p, s)))
            Case Else
                Enthalpy_ps = -999.9
        End Select
    End Function

    <ExcelFunction(Description:="Enthalpy as function of pressure and vapor fraction, [kJ/kg]", Category:="GCME E-PT | Steam Table")>
    Public Function Enthalpy_px(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Vapor Fraction")> x As Double) As Double
        Dim hL As Double
        Dim hV As Double
        p = ToSIunit_p(p)
        x = ToSIunit_x(x)
        If x > 1 Or x < 0 Or p >= 22.064 Then
            Enthalpy_px = -999.9
            Exit Function
        End If
        hL = H4L_p(p)
        hV = H4V_p(p)
        Enthalpy_px = FromSIunit_h(hL + x * (hV - hL))
    End Function

    <ExcelFunction(Description:="Enthalpy as function of temperature and vapor, [kJ/kg]", Category:="GCME E-PT | Steam Table")>
    Public Function Enthalpy_Tx(<ExcelArgument(Description:="Temperature, [C]")> T As Double, <ExcelArgument(Description:="Vapor Fraction")> x As Double) As Double
        Dim hL As Double
        Dim hV As Double
        Dim p As Double
        T = ToSIunit_T(T)
        x = ToSIunit_x(x)
        If x > 1 Or x < 0 Or T >= 647.096 Then
            Enthalpy_Tx = -999.9
            Exit Function
        End If
        p = P4_T(T)
        hL = H4L_p(p)
        hV = H4V_p(p)
        Enthalpy_Tx = FromSIunit_h(hL + x * (hV - hL))
    End Function

    <ExcelFunction(Description:="Enthalpy as function of pressure and density, [kJ/kg]", Category:="GCME E-PT | Steam Table")>
    Public Function Enthalpy_prho(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Density, [kg/m3]")> rho As Double) As Double
        Dim hL, hV, vL, vV, x As Double
        p = ToSIunit_p(p)
        rho = 1 / ToSIunit_v(1 / rho)
        Select Case Region_prho(p, rho)
            Case 1
                Enthalpy_prho = FromSIunit_h(H1_pT(p, T1_prho(p, rho)))
            Case 2
                Enthalpy_prho = FromSIunit_h(H2_pT(p, T2_prho(p, rho)))
            Case 3
                Enthalpy_prho = FromSIunit_h(H3_rhoT(rho, T3_prho(p, rho)))
            Case 4
                If p < 16.529 Then
                    vV = V2_pT(p, T4_p(p))
                    vL = V1_pT(p, T4_p(p))
                Else
                    vV = V3_ph(p, H4V_p(p))
                    vL = V3_ph(p, H4L_p(p))
                End If
                hV = H4V_p(p)
                hL = H4L_p(p)
                x = (1 / rho - vL) / (vV - vL)
                Enthalpy_prho = FromSIunit_h((1 - x) * hL + x * hV)
            Case 5
                Enthalpy_prho = FromSIunit_h(H5_pT(p, T5_prho(p, rho)))
            Case Else
                Enthalpy_prho = -999.9
        End Select
    End Function


    '***********************************************************************************************************
    '*1.5 Specific Volume (v)

    <ExcelFunction(Description:="Saturated vapor volume as function of pressure, [m3/kg]", Category:="GCME E-PT | Steam Table")>
    Public Function VolumeV_p(<ExcelArgument(Description:="Pressure, [bar]")> p As Double) As Double
        p = ToSIunit_p(p)
        If p > 0.000611657 And p < 22.06395 Then
            If p < 16.529 Then
                VolumeV_p = FromSIunit_v(V2_pT(p, T4_p(p)))
            Else
                VolumeV_p = FromSIunit_v(V3_ph(p, H4V_p(p)))
            End If
        Else
            VolumeV_p = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Saturated vapor volume as function of pressure, [m3/kg]", Category:="GCME E-PT | Steam Table")>
    Public Function VolumeL_p(<ExcelArgument(Description:="Pressure, [bar]")> p As Double) As Double
        p = ToSIunit_p(p)
        If p > 0.000611657 And p < 22.06395 Then
            If p < 16.529 Then
                VolumeL_p = FromSIunit_v(V1_pT(p, T4_p(p)))
            Else
                VolumeL_p = FromSIunit_v(V3_ph(p, H4L_p(p)))
            End If
        Else
            VolumeL_p = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Saturated vapor volume as function of temperature, [m3/kg]", Category:="GCME E-PT | Steam Table")>
    Public Function VolumeV_T(<ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        T = ToSIunit_T(T)
        If T > 273.15 And T < 647.096 Then
            If T <= 623.15 Then
                VolumeV_T = FromSIunit_v(V2_pT(P4_T(T), T))
            Else
                VolumeV_T = FromSIunit_v(V3_ph(P4_T(T), H4V_p(P4_T(T))))
            End If
        Else
            VolumeV_T = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Saturated vapor volume as function of temperature, [m3/kg]", Category:="GCME E-PT | Steam Table")>
    Public Function VolumeL_T(<ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        T = ToSIunit_T(T)
        If T > 273.15 And T < 647.096 Then
            If T <= 623.15 Then
                VolumeL_T = FromSIunit_v(V1_pT(P4_T(T), T))
            Else
                VolumeL_T = FromSIunit_v(V3_ph(P4_T(T), H4L_p(P4_T(T))))
            End If
        Else
            VolumeL_T = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Specific vapor volume as function of pressure and temperature, [m3/kg]", Category:="GCME E-PT | Steam Table")>
    Public Function Volume_pT(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        p = ToSIunit_p(p)
        T = ToSIunit_T(T)
        Select Case Region_pT(p, T)
            Case 1
                Volume_pT = FromSIunit_v(V1_pT(p, T))
            Case 2
                Volume_pT = FromSIunit_v(V2_pT(p, T))
            Case 3
                Volume_pT = FromSIunit_v(V3_ph(p, H3_pT(p, T)))
            Case 4
                Volume_pT = -999.9
            Case 5
                Volume_pT = FromSIunit_v(V5_pT(p, T))
            Case Else
                Volume_pT = -999.9
        End Select
    End Function

    <ExcelFunction(Description:="Specific vapor volume as function of pressure and enthalpy, [m3/kg]", Category:="GCME E-PT | Steam Table")>
    Public Function Volume_ph(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Enthalpy, [kJ/kg]")> h As Double) As Double
        Dim xs As Double
        Dim v4V As Double
        Dim v4L As Double
        p = ToSIunit_p(p)
        h = ToSIunit_h(h)
        Select Case Region_ph(p, h)
            Case 1
                Volume_ph = FromSIunit_v(V1_pT(p, T1_ph(p, h)))
            Case 2
                Volume_ph = FromSIunit_v(V2_pT(p, T2_ph(p, h)))
            Case 3
                Volume_ph = FromSIunit_v(V3_ph(p, h))
            Case 4
                xs = X4_ph(p, h)
                If p < 16.529 Then
                    v4V = V2_pT(p, T4_p(p))
                    v4L = V1_pT(p, T4_p(p))
                Else
                    v4V = V3_ph(p, H4V_p(p))
                    v4L = V3_ph(p, H4L_p(p))
                End If
                Volume_ph = FromSIunit_v((xs * v4V + (1 - xs) * v4L))
            Case 5
                Volume_ph = FromSIunit_v(V5_pT(p, T5_ph(p, h)))
            Case Else
                Volume_ph = -999.9
        End Select
    End Function

    <ExcelFunction(Description:="Specific vapor volume as function of pressure and entrophy, [m3/kg]", Category:="GCME E-PT | Steam Table")>
    Public Function Volume_ps(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Entropy, [kJ/kg.K]")> s As Double) As Double
        Dim xs As Double
        Dim v4V As Double
        Dim v4L As Double
        p = ToSIunit_p(p)
        s = ToSIunit_s(s)
        Select Case Region_ps(p, s)
            Case 1
                Volume_ps = FromSIunit_v(V1_pT(p, T1_ps(p, s)))
            Case 2
                Volume_ps = FromSIunit_v(V2_pT(p, T2_ps(p, s)))
            Case 3
                Volume_ps = FromSIunit_v(V3_ps(p, s))
            Case 4
                xs = X4_ps(p, s)
                If p < 16.529 Then
                    v4V = V2_pT(p, T4_p(p))
                    v4L = V1_pT(p, T4_p(p))
                Else
                    v4V = V3_ph(p, H4V_p(p))
                    v4L = V3_ph(p, H4L_p(p))
                End If
                Volume_ps = FromSIunit_v((xs * v4V + (1 - xs) * v4L))
            Case 5
                Volume_ps = FromSIunit_v(V5_pT(p, T5_ps(p, s)))
            Case Else
                Volume_ps = -999.9
        End Select
    End Function

    '***********************************************************************************************************
    '*1.6 Density (rho)
    ' Density is calculated as 1/v

    <ExcelFunction(Description:="Saturated vapor density as function of pressure, [kg/m3]", Category:="GCME E-PT | Steam Table")>
    Public Function DensityV_p(<ExcelArgument(Description:="Pressure, [bar]")> p As Double) As Double
        DensityV_p = 1 / VolumeV_p(p)
    End Function

    <ExcelFunction(Description:="Saturated liquid density as function of pressure, [kg/m3]", Category:="GCME E-PT | Steam Table")>
    Public Function DensityL_p(<ExcelArgument(Description:="Pressure, [bar]")> p As Double) As Double
        DensityL_p = 1 / VolumeL_p(p)
    End Function

    <ExcelFunction(Description:="Saturated liquid density as function of temperature, [kg/m3]", Category:="GCME E-PT | Steam Table")>
    Public Function DensityL_T(<ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        DensityL_T = 1 / VolumeL_T(T)
    End Function

    <ExcelFunction(Description:="Saturated vapor density as function of temperature, [kg/m3]", Category:="GCME E-PT | Steam Table")>
    Public Function DensityV_T(<ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        DensityV_T = 1 / VolumeV_T(T)
    End Function

    <ExcelFunction(Description:="Density as function of pressure and temperature, [kg/m3]", Category:="GCME E-PT | Steam Table")>
    Public Function Density_pT(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        Density_pT = 1 / Volume_pT(p, T)
    End Function

    <ExcelFunction(Description:="Density as function of pressure and enthalpy, [kg/m3]", Category:="GCME E-PT | Steam Table")>
    Public Function Density_ph(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Enthalpy, [kJ/kg]")> h As Double) As Double
        Density_ph = 1 / Volume_ph(p, h)
    End Function

    <ExcelFunction(Description:="Density as function of pressure and entropy, [kg/m3]", Category:="GCME E-PT | Steam Table")>
    Public Function Density_ps(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Entropy, [kJ/kg.K]")> s As Double) As Double
        Density_ps = 1 / Volume_ps(p, s)
    End Function

    '***********************************************************************************************************
    '*1.7 Specific entropy (s)

    <ExcelFunction(Description:="Saturated vapor entropy as function of pressure, [kJ/kg.K]", Category:="GCME E-PT | Steam Table")>
    Public Function EntropyV_p(<ExcelArgument(Description:="Pressure, [bar]")> p As Double) As Double
        p = ToSIunit_p(p)
        If p > 0.000611657 And p < 22.06395 Then
            If p < 16.529 Then
                EntropyV_p = FromSIunit_s(S2_pT(p, T4_p(p)))
            Else
                EntropyV_p = FromSIunit_s(S3_rhoT(1 / (V3_ph(p, H4V_p(p))), T4_p(p)))
            End If
        Else
            EntropyV_p = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Saturated liquid entropy as function of pressure, [kJ/kg.K]", Category:="GCME E-PT | Steam Table")>
    Public Function EntropyL_p(<ExcelArgument(Description:="Pressure, [bar]")> p As Double) As Double
        p = ToSIunit_p(p)
        If p > 0.000611657 And p < 22.06395 Then
            If p < 16.529 Then
                EntropyL_p = FromSIunit_s(S1_pT(p, T4_p(p)))
            Else
                EntropyL_p = FromSIunit_s(S3_rhoT(1 / (V3_ph(p, H4L_p(p))), T4_p(p)))
            End If
        Else
            EntropyL_p = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Saturated vapor entropy as function of temperature, [kJ/kg.K]", Category:="GCME E-PT | Steam Table")>
    Public Function EntropyV_T(<ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        T = ToSIunit_T(T)
        If T > 273.15 And T < 647.096 Then
            If T <= 623.15 Then
                EntropyV_T = FromSIunit_s(S2_pT(P4_T(T), T))
            Else
                EntropyV_T = FromSIunit_s(S3_rhoT(1 / (V3_ph(P4_T(T), H4V_p(P4_T(T)))), T))
            End If
        Else
            EntropyV_T = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Saturated liquid entropy as function of temperature, [kJ/kg.K]", Category:="GCME E-PT | Steam Table")>
    Public Function EntropyL_T(<ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        T = ToSIunit_T(T)
        If T > 273.15 And T < 647.096 Then
            If T <= 623.15 Then
                EntropyL_T = FromSIunit_s(S1_pT(P4_T(T), T))
            Else
                EntropyL_T = FromSIunit_s(S3_rhoT(1 / (V3_ph(P4_T(T), H4L_p(P4_T(T)))), T))
            End If
        Else
            EntropyL_T = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Specific entropy as function of pressure and temperature, [kJ/kg.K]", Category:="GCME E-PT | Steam Table")>
    Public Function Entropy_pT(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        p = ToSIunit_p(p)
        T = ToSIunit_T(T)
        Select Case Region_pT(p, T)
            Case 1
                Entropy_pT = FromSIunit_s(S1_pT(p, T))
            Case 2
                Entropy_pT = FromSIunit_s(S2_pT(p, T))
            Case 3
                Entropy_pT = FromSIunit_s(S3_rhoT(1 / V3_ph(p, H3_pT(p, T)), T))
            Case 4
                Entropy_pT = -999.9
            Case 5
                Entropy_pT = FromSIunit_s(S5_pT(p, T))
            Case Else
                Entropy_pT = -999.9
        End Select
    End Function

    <ExcelFunction(Description:="Specific entropy as function of pressure and enthalpy, [kJ/kg.K]", Category:="GCME E-PT | Steam Table")>
    Public Function Entropy_ph(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Enthalpy, [kJ/kg]")> h As Double) As Double
        Dim Ts As Double
        Dim xs As Double
        Dim s4V As Double
        Dim s4L As Double
        Dim v4V As Double
        Dim v4L As Double
        p = ToSIunit_p(p)
        h = ToSIunit_h(h)
        Select Case Region_ph(p, h)
            Case 1
                Entropy_ph = FromSIunit_s(S1_pT(p, T1_ph(p, h)))
            Case 2
                Entropy_ph = FromSIunit_s(S2_pT(p, T2_ph(p, h)))
            Case 3
                Entropy_ph = FromSIunit_s(S3_rhoT(1 / V3_ph(p, h), T3_ph(p, h)))
            Case 4
                Ts = T4_p(p)
                xs = X4_ph(p, h)
                If p < 16.529 Then
                    s4V = S2_pT(p, Ts)
                    s4L = S1_pT(p, Ts)
                Else
                    v4V = V3_ph(p, H4V_p(p))
                    s4V = S3_rhoT(1 / v4V, Ts)
                    v4L = V3_ph(p, H4L_p(p))
                    s4L = S3_rhoT(1 / v4L, Ts)
                End If
                Entropy_ph = FromSIunit_s((xs * s4V + (1 - xs) * s4L))
            Case 5
                Entropy_ph = FromSIunit_s(S5_pT(p, T5_ph(p, h)))
            Case Else
                Entropy_ph = -999.9
        End Select
    End Function
    '***********************************************************************************************************
    '*1.8 Specific internal energy (u)

    <ExcelFunction(Description:="Saturated vapor internal energy as function of pressure, [kJ/kg]", Category:="GCME E-PT | Steam Table")>
    Public Function UV_p(<ExcelArgument(Description:="Pressure, [bar]")> p As Double) As Double
        p = ToSIunit_p(p)
        If p > 0.000611657 And p < 22.06395 Then
            If p < 16.529 Then
                UV_p = FromSIunit_u(U2_pT(p, T4_p(p)))
            Else
                UV_p = FromSIunit_u(U3_rhoT(1 / (V3_ph(p, H4V_p(p))), T4_p(p)))
            End If
        Else
            UV_p = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Saturated liquid internal energy as function of pressure, [kJ/kg]", Category:="GCME E-PT | Steam Table")>
    Public Function UL_p(<ExcelArgument(Description:="Pressure, [bar]")> p As Double) As Double
        p = ToSIunit_p(p)
        If p > 0.000611657 And p < 22.06395 Then
            If p < 16.529 Then
                UL_p = FromSIunit_u(U1_pT(p, T4_p(p)))
            Else
                UL_p = FromSIunit_u(U3_rhoT(1 / (V3_ph(p, H4L_p(p))), T4_p(p)))
            End If
        Else
            UL_p = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Saturated vapor internal energy as function of temperature, [kJ/kg]", Category:="GCME E-PT | Steam Table")>
    Public Function UV_T(<ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        T = ToSIunit_T(T)
        If T > 273.15 And T < 647.096 Then
            If T <= 623.15 Then
                UV_T = FromSIunit_u(U2_pT(P4_T(T), T))
            Else
                UV_T = FromSIunit_u(U3_rhoT(1 / (V3_ph(P4_T(T), H4V_p(P4_T(T)))), T))
            End If
        Else
            UV_T = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Saturated liquid internal energy as function of temperature, [kJ/kg]", Category:="GCME E-PT | Steam Table")>
    Public Function UL_T(<ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        T = ToSIunit_T(T)
        If T > 273.15 And T < 647.096 Then
            If T <= 623.15 Then
                UL_T = FromSIunit_u(U1_pT(P4_T(T), T))
            Else
                UL_T = FromSIunit_u(U3_rhoT(1 / (V3_ph(P4_T(T), H4L_p(P4_T(T)))), T))
            End If
        Else
            UL_T = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Specific internal energy as function of pressure and temperature, [kJ/kg]", Category:="GCME E-PT | Steam Table")>
    Public Function U_pT(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        p = ToSIunit_p(p)
        T = ToSIunit_T(T)
        Select Case Region_pT(p, T)
            Case 1
                U_pT = FromSIunit_u(U1_pT(p, T))
            Case 2
                U_pT = FromSIunit_u(U2_pT(p, T))
            Case 3
                U_pT = FromSIunit_u(U3_rhoT(1 / V3_ph(p, H3_pT(p, T)), T))
            Case 4
                U_pT = -999.9
            Case 5
                U_pT = FromSIunit_u(U5_pT(p, T))
            Case Else
                U_pT = -999.9
        End Select
    End Function

    <ExcelFunction(Description:="Specific internal energy as function of pressure and enthalpy, [kJ/kg]", Category:="GCME E-PT | Steam Table")>
    Public Function U_ph(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Enthalpy, [kJ/kg]")> h As Double) As Double
        Dim Ts As Double
        Dim xs As Double
        Dim u4v As Double
        Dim u4L As Double
        Dim v4V As Double
        Dim v4L As Double
        p = ToSIunit_p(p)
        h = ToSIunit_h(h)
        Select Case Region_ph(p, h)
            Case 1
                U_ph = FromSIunit_u(U1_pT(p, T1_ph(p, h)))
            Case 2
                U_ph = FromSIunit_u(U2_pT(p, T2_ph(p, h)))
            Case 3
                U_ph = FromSIunit_u(U3_rhoT(1 / V3_ph(p, h), T3_ph(p, h)))
            Case 4
                Ts = T4_p(p)
                xs = X4_ph(p, h)
                If p < 16.529 Then
                    u4v = U2_pT(p, Ts)
                    u4L = U1_pT(p, Ts)
                Else
                    v4V = V3_ph(p, H4V_p(p))
                    u4v = U3_rhoT(1 / v4V, Ts)
                    v4L = V3_ph(p, H4L_p(p))
                    u4L = U3_rhoT(1 / v4L, Ts)
                End If
                U_ph = FromSIunit_u((xs * u4v + (1 - xs) * u4L))
            Case 5
                Ts = T5_ph(p, h)
                U_ph = FromSIunit_u(U5_pT(p, Ts))
            Case Else
                U_ph = -999.9
        End Select
    End Function

    <ExcelFunction(Description:="Specific internal energy as function of pressure and entropy, [kJ/kg]", Category:="GCME E-PT | Steam Table")>
    Public Function U_ps(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Entropy, [kJ/kg.K]")> s As Double) As Double
        Dim x As Double
        Dim uLp, uVp As Double
        p = ToSIunit_p(p)
        s = ToSIunit_s(s)
        Select Case Region_ps(p, s)
            Case 1
                U_ps = FromSIunit_u(U1_pT(p, T1_ps(p, s)))
            Case 2
                U_ps = FromSIunit_u(U2_pT(p, T2_ps(p, s)))
            Case 3
                U_ps = FromSIunit_u(U3_rhoT(1 / V3_ps(p, s), T3_ps(p, s)))
            Case 4
                If p < 16.529 Then
                    uLp = U1_pT(p, T4_p(p))
                    uVp = U2_pT(p, T4_p(p))
                Else
                    uLp = U3_rhoT(1 / (V3_ph(p, H4L_p(p))), T4_p(p))
                    uVp = U3_rhoT(1 / (V3_ph(p, H4V_p(p))), T4_p(p))
                End If
                x = X4_ps(p, s)
                U_ps = FromSIunit_u((x * uVp + (1 - x) * uLp))
            Case 5
                U_ps = FromSIunit_u(U5_pT(p, T5_ps(p, s)))
            Case Else
                U_ps = -999.9
        End Select
    End Function
    '***********************************************************************************************************
    '*1.9 Specific isobaric heat capacity (Cp)

    <ExcelFunction(Description:="Saturated vapor isobaric heat capacity as function of pressure, [kJ/kg.K]", Category:="GCME E-PT | Steam Table")>
    Public Function CpV_p(<ExcelArgument(Description:="Pressure, [bar]")> p As Double) As Double
        p = ToSIunit_p(p)
        If p > 0.000611657 And p < 22.06395 Then
            If p < 16.529 Then
                CpV_p = FromSIunit_Cp(Cp2_pT(p, T4_p(p)))
            Else
                CpV_p = FromSIunit_Cp(Cp3_rhoT(1 / (V3_ph(p, H4V_p(p))), T4_p(p)))
            End If
        Else
            CpV_p = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Saturated liquid isobaric heat capacity as function of pressure, [kJ/kg.K]", Category:="GCME E-PT | Steam Table")>
    Public Function CpL_p(<ExcelArgument(Description:="Pressure, [bar]")> p As Double) As Double
        Dim T, h, v As Double
        p = ToSIunit_p(p)
        If p > 0.000611657 And p < 22.06395 Then
            If p < 16.529 Then
                CpL_p = FromSIunit_Cp(Cp1_pT(p, T4_p(p)))
            Else
                T = T4_p(p)
                h = H4L_p(p)
                v = V3_ph(p, H4L_p(p))

                CpL_p = FromSIunit_Cp(Cp3_rhoT(1 / (V3_ph(p, H4L_p(p))), T4_p(p)))
            End If
        Else
            CpL_p = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Saturated vapor isobaric heat capacity as function of temperature, [kJ/kg.K]", Category:="GCME E-PT | Steam Table")>
    Public Function CpV_T(<ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        T = ToSIunit_T(T)
        If T > 273.15 And T < 647.096 Then
            If T <= 623.15 Then
                CpV_T = FromSIunit_Cp(Cp2_pT(P4_T(T), T))
            Else
                CpV_T = FromSIunit_Cp(Cp3_rhoT(1 / (V3_ph(P4_T(T), H4V_p(P4_T(T)))), T))
            End If
        Else
            CpV_T = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Saturated liquid isobaric heat capacity as function of temperature, [kJ/kg.K]", Category:="GCME E-PT | Steam Table")>
    Public Function CpL_T(<ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        T = ToSIunit_T(T)
        If T > 273.15 And T < 647.096 Then
            If T <= 623.15 Then
                CpL_T = FromSIunit_Cp(Cp1_pT(P4_T(T), T))
            Else
                CpL_T = FromSIunit_Cp(Cp3_rhoT(1 / (V3_ph(P4_T(T), H4L_p(P4_T(T)))), T))
            End If
        Else
            CpL_T = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Specific isobaric heat capacity as function of pressure and temperature, [kJ/kg.K]", Category:="GCME E-PT | Steam Table")>
    Public Function Cp_pT(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        p = ToSIunit_p(p)
        T = ToSIunit_T(T)
        Select Case Region_pT(p, T)
            Case 1
                Cp_pT = FromSIunit_Cp(Cp1_pT(p, T))
            Case 2
                Cp_pT = FromSIunit_Cp(Cp2_pT(p, T))
            Case 3
                Cp_pT = FromSIunit_Cp(Cp3_rhoT(1 / V3_ph(p, H3_pT(p, T)), T))
            Case 4
                Cp_pT = -999.9
            Case 5
                Cp_pT = FromSIunit_Cp(Cp5_pT(p, T))
            Case Else
                Cp_pT = -999.9
        End Select
    End Function

    <ExcelFunction(Description:="Specific isobaric heat capacity as function of pressure and enthalpy, [kJ/kg.K]", Category:="GCME E-PT | Steam Table")>
    Public Function Cp_ph(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Enthalpy, [kJ/kg]")> h As Double) As Double
        p = ToSIunit_p(p)
        h = ToSIunit_h(h)
        Select Case Region_ph(p, h)
            Case 1
                Cp_ph = FromSIunit_Cp(Cp1_pT(p, T1_ph(p, h)))
            Case 2
                Cp_ph = FromSIunit_Cp(Cp2_pT(p, T2_ph(p, h)))
            Case 3
                Cp_ph = FromSIunit_Cp(Cp3_rhoT(1 / V3_ph(p, h), T3_ph(p, h)))
            Case 4
                Cp_ph = -999.9 '#Not def. for mixture"
            Case 5
                Cp_ph = FromSIunit_Cp(Cp5_pT(p, T5_ph(p, h)))
            Case Else
                Cp_ph = -999.9
        End Select
    End Function

    <ExcelFunction(Description:="Specific isobaric heat capacity as function of pressure and entropy, [kJ/kg.K]", Category:="GCME E-PT | Steam Table")>
    Public Function Cp_ps(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Entropy, [kJ/kg.K]")> s As Double) As Double
        p = ToSIunit_p(p)
        s = ToSIunit_s(s)
        Select Case Region_ps(p, s)
            Case 1
                Cp_ps = FromSIunit_Cp(Cp1_pT(p, T1_ps(p, s)))
            Case 2
                Cp_ps = FromSIunit_Cp(Cp2_pT(p, T2_ps(p, s)))
            Case 3
                Cp_ps = FromSIunit_Cp(Cp3_rhoT(1 / V3_ps(p, s), T3_ps(p, s)))
            Case 4
                Cp_ps = -999.9 '#Not def. for mixture"
            Case 5
                Cp_ps = FromSIunit_Cp(Cp5_pT(p, T5_ps(p, s)))
            Case Else
                Cp_ps = -999.9
        End Select
    End Function
    '***********************************************************************************************************
    '*1.10 Specific isochoric heat capacity (Cv)

    <ExcelFunction(Description:="Saturated vapor isochoric heat capacity as function of pressure, [kJ/kg.K]", Category:="GCME E-PT | Steam Table")>
    Public Function CvV_p(<ExcelArgument(Description:="Pressure, [bar]")> p As Double) As Double
        p = ToSIunit_p(p)
        If p > 0.000611657 And p < 22.06395 Then
            If p < 16.529 Then
                CvV_p = FromSIunit_Cv(Cv2_pT(p, T4_p(p)))
            Else
                CvV_p = FromSIunit_Cv(Cv3_rhoT(1 / (V3_ph(p, H4V_p(p))), T4_p(p)))
            End If
        Else
            CvV_p = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Saturated liquid isochoric heat capacity as function of pressure, [kJ/kg.K]", Category:="GCME E-PT | Steam Table")>
    Public Function CvL_p(<ExcelArgument(Description:="Pressure, [bar]")> p As Double) As Double
        p = ToSIunit_p(p)
        If p > 0.000611657 And p < 22.06395 Then
            If p < 16.529 Then
                CvL_p = FromSIunit_Cv(Cv1_pT(p, T4_p(p)))
            Else
                CvL_p = FromSIunit_Cv(Cv3_rhoT(1 / (V3_ph(p, H4L_p(p))), T4_p(p)))
            End If
        Else
            CvL_p = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Saturated vapor isochoric heat capacity as function of temperature, [kJ/kg.K]", Category:="GCME E-PT | Steam Table")>
    Public Function CvV_T(<ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        T = ToSIunit_T(T)
        If T > 273.15 And T < 647.096 Then
            If T <= 623.15 Then
                CvV_T = FromSIunit_Cv(Cv2_pT(P4_T(T), T))
            Else
                CvV_T = FromSIunit_Cv(Cv3_rhoT(1 / (V3_ph(P4_T(T), H4V_p(P4_T(T)))), T))
            End If
        Else
            CvV_T = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Saturated liquid isochoric heat capacity as function of temperature, [kJ/kg.K]", Category:="GCME E-PT | Steam Table")>
    Public Function CvL_T(<ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        T = ToSIunit_T(T)
        If T > 273.15 And T < 647.096 Then
            If T <= 623.15 Then
                CvL_T = FromSIunit_Cv(Cv1_pT(P4_T(T), T))
            Else
                CvL_T = FromSIunit_Cv(Cv3_rhoT(1 / (V3_ph(P4_T(T), H4L_p(P4_T(T)))), T))
            End If
        Else
            CvL_T = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Specific isochoric heat capacity as function of pressure and temperature, [kJ/kg.K]", Category:="GCME E-PT | Steam Table")>
    Public Function Cv_pT(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        p = ToSIunit_p(p)
        T = ToSIunit_T(T)
        Select Case Region_pT(p, T)
            Case 1
                Cv_pT = FromSIunit_Cv(Cv1_pT(p, T))
            Case 2
                Cv_pT = FromSIunit_Cv(Cv2_pT(p, T))
            Case 3
                Cv_pT = FromSIunit_Cv(Cv3_rhoT(1 / V3_ph(p, H3_pT(p, T)), T))
            Case 4
                Cv_pT = -999.9
            Case 5
                Cv_pT = FromSIunit_Cv(Cv5_pT(p, T))
            Case Else
                Cv_pT = -999.9
        End Select
    End Function

    <ExcelFunction(Description:="Specific isochoric heat capacity as function of pressure and enthalpy, [kJ/kg.K]", Category:="GCME E-PT | Steam Table")>
    Public Function Cv_ph(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Enthalpy, [kJ/kg]")> h As Double) As Double
        p = ToSIunit_p(p)
        h = ToSIunit_h(h)
        Select Case Region_ph(p, h)
            Case 1
                Cv_ph = FromSIunit_Cv(Cv1_pT(p, T1_ph(p, h)))
            Case 2
                Cv_ph = FromSIunit_Cv(Cv2_pT(p, T2_ph(p, h)))
            Case 3
                Cv_ph = FromSIunit_Cv(Cv3_rhoT(1 / V3_ph(p, h), T3_ph(p, h)))
            Case 4
                Cv_ph = -999.9 '#Not def. for mixture"
            Case 5
                Cv_ph = FromSIunit_Cv(Cv5_pT(p, T5_ph(p, h)))
            Case Else
                Cv_ph = -999.9
        End Select
    End Function

    <ExcelFunction(Description:="Specific isochoric heat capacity as function of pressure and enthropy, [kJ/kg.K]", Category:="GCME E-PT | Steam Table")>
    Public Function Cv_ps(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Entropy, [kJ/kg.K]")> s As Double) As Double
        p = ToSIunit_p(p)
        s = ToSIunit_s(s)
        Select Case Region_ps(p, s)
            Case 1
                Cv_ps = FromSIunit_Cv(Cv1_pT(p, T1_ps(p, s)))
            Case 2
                Cv_ps = FromSIunit_Cv(Cv2_pT(p, T2_ps(p, s)))
            Case 3
                Cv_ps = FromSIunit_Cv(Cv3_rhoT(1 / V3_ps(p, s), T3_ps(p, s)))
            Case 4
                Cv_ps = -999.9 '#Not def. for mixture
            Case 5
                Cv_ps = FromSIunit_Cv(Cv5_pT(p, T5_ps(p, s)))
            Case Else
                Cv_ps = -999.9
        End Select
    End Function


    '***********************************************************************************************************
    '*1.11 Speed of sound

    <ExcelFunction(Description:="Saturated vapor speed of sound as function of pressure, [m/s]", Category:="GCME E-PT | Steam Table")>
    Public Function WV_p(<ExcelArgument(Description:="Pressure, [bar]")> p As Double) As Double
        p = ToSIunit_p(p)
        If p > 0.000611657 And p < 22.06395 Then
            If p < 16.529 Then
                WV_p = FromSIunit_w(W2_pT(p, T4_p(p)))
            Else
                WV_p = FromSIunit_w(W3_rhoT(1 / (V3_ph(p, H4V_p(p))), T4_p(p)))
            End If
        Else
            WV_p = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Saturated liquid speed of sound as function of pressure, [m/s]", Category:="GCME E-PT | Steam Table")>
    Public Function WL_p(<ExcelArgument(Description:="Pressure, [bar]")> p As Double) As Double
        p = ToSIunit_p(p)
        If p > 0.000611657 And p < 22.06395 Then
            If p < 16.529 Then
                WL_p = FromSIunit_w(W1_pT(p, T4_p(p)))
            Else
                WL_p = FromSIunit_w(W3_rhoT(1 / (V3_ph(p, H4L_p(p))), T4_p(p)))
            End If
        Else
            WL_p = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Saturated vapor speed of sound as function of temperature, [m/s]", Category:="GCME E-PT | Steam Table")>
    Public Function WV_T(<ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        T = ToSIunit_T(T)
        If T > 273.15 And T < 647.096 Then
            If T <= 623.15 Then
                WV_T = FromSIunit_w(W2_pT(P4_T(T), T))
            Else
                WV_T = FromSIunit_w(W3_rhoT(1 / (V3_ph(P4_T(T), H4V_p(P4_T(T)))), T))
            End If
        Else
            WV_T = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Saturated liquid speed of sound as function of temperature, [m/s]", Category:="GCME E-PT | Steam Table")>
    Public Function WL_T(<ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        T = ToSIunit_T(T)
        If T > 273.15 And T < 647.096 Then
            If T <= 623.15 Then
                WL_T = FromSIunit_w(W1_pT(P4_T(T), T))
            Else
                WL_T = FromSIunit_w(W3_rhoT(1 / (V3_ph(P4_T(T), H4L_p(P4_T(T)))), T))
            End If
        Else
            WL_T = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Speed of sound as function of pressure and temperature, [m/s]", Category:="GCME E-PT | Steam Table")>
    Public Function W_pT(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        p = ToSIunit_p(p)
        T = ToSIunit_T(T)
        Select Case Region_pT(p, T)
            Case 1
                W_pT = FromSIunit_w(W1_pT(p, T))
            Case 2
                W_pT = FromSIunit_w(W2_pT(p, T))
            Case 3
                W_pT = FromSIunit_w(W3_rhoT(1 / V3_ph(p, H3_pT(p, T)), T))
            Case 4
                W_pT = -999.9
            Case 5
                W_pT = FromSIunit_w(W5_pT(p, T))
            Case Else
                W_pT = -999.9
        End Select
    End Function

    <ExcelFunction(Description:="Speed of sound as function of pressure and enthalpy, [m/s]", Category:="GCME E-PT | Steam Table")>
    Public Function W_ph(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Enthalpy, [kJ/kg]")> h As Double) As Double
        p = ToSIunit_p(p)
        h = ToSIunit_h(h)
        Select Case Region_ph(p, h)
            Case 1
                W_ph = FromSIunit_w(W1_pT(p, T1_ph(p, h)))
            Case 2
                W_ph = FromSIunit_w(W2_pT(p, T2_ph(p, h)))
            Case 3
                W_ph = FromSIunit_w(W3_rhoT(1 / V3_ph(p, h), T3_ph(p, h)))
            Case 4
                W_ph = -999.9 '#Not def. for mixture
            Case 5
                W_ph = FromSIunit_w(W5_pT(p, T5_ph(p, h)))
            Case Else
                W_ph = -999.9
        End Select
    End Function

    <ExcelFunction(Description:="Speed of sound as function of pressure and enthropy, [m/s]", Category:="GCME E-PT | Steam Table")>
    Public Function W_ps(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Entropy, [kJ/kg.K]")> s As Double) As Double
        p = ToSIunit_p(p)
        s = ToSIunit_s(s)
        Select Case Region_ps(p, s)
            Case 1
                W_ps = FromSIunit_w(W1_pT(p, T1_ps(p, s)))
            Case 2
                W_ps = FromSIunit_w(W2_pT(p, T2_ps(p, s)))
            Case 3
                W_ps = FromSIunit_w(W3_rhoT(1 / V3_ps(p, s), T3_ps(p, s)))
            Case 4
                W_ps = -999.9 '#Not def. for mixture
            Case 5
                W_ps = FromSIunit_w(W5_pT(p, T5_ps(p, s)))
            Case Else
                W_ps = -999.9
        End Select
    End Function
    '***********************************************************************************************************
    '*1.12 Viscosity

    <ExcelFunction(Description:="Viscosity function of pressure and temperature, [Pa-s]", Category:="GCME E-PT | Steam Table")>
    Public Function My_pT(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        p = ToSIunit_p(p)
        T = ToSIunit_T(T)
        Select Case Region_pT(p, T)
            Case 4
                My_pT = -999.9
            Case 1, 2, 3, 5
                My_pT = FromSIunit_my(My_AllRegions_pT(p, T))
            Case Else
                My_pT = -999.9
        End Select
    End Function

    <ExcelFunction(Description:="Viscosity function of pressure and enthalpy, [Pa-s]", Category:="GCME E-PT | Steam Table")>
    Public Function My_ph(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Enthalpy, [kJ/kg]")> h As Double) As Double
        p = ToSIunit_p(p)
        h = ToSIunit_h(h)
        Select Case Region_ph(p, h)
            Case 1, 2, 3, 5
                My_ph = FromSIunit_my(My_AllRegions_ph(p, h))
            Case 4
                My_ph = -999.9
            Case Else
                My_ph = -999.9
        End Select
    End Function

    <ExcelFunction(Description:="Viscosity function of pressure and enthropy, [Pa-s]", Category:="GCME E-PT | Steam Table")>
    Public Function My_ps(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Entropy, [kJ/kg.K]")> s As Double) As Double
        My_ps = My_ph(p, Enthalpy_ps(p, s))
    End Function

    '***********************************************************************************************************
    '*1.13 Prandtl

    <ExcelFunction(Description:="Prandtl number as function of pressure and temperature, [-]", Category:="GCME E-PT | Steam Table")>
    Public Function Pr_pT(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        Dim Cp As Double
        Dim my As Double
        Dim tc As Double
        Cp = ToSIunit_Cp(Cp_pT(p, T))
        my = ToSIunit_my(My_pT(p, T))
        tc = ToSIunit_tc(ThermalConductivity_pT(p, T))
        Pr_pT = Cp * 1000 * my / tc
    End Function

    <ExcelFunction(Description:="Prandtl number as function of pressure and enthalpy, [-]", Category:="GCME E-PT | Steam Table")>
    Public Function Pr_ph(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Enthalpy, [kJ/kg]")> h As Double) As Double
        Dim Cp As Double
        Dim my As Double
        Dim tc As Double
        Cp = ToSIunit_Cp(Cp_ph(p, h))
        my = ToSIunit_my(My_ph(p, h))
        tc = ToSIunit_tc(ThermalConductivity_ph(p, h))
        Pr_ph = Cp * 1000 * my / tc
    End Function
    '***********************************************************************************************************
    '*1.14 Kappa

    <ExcelFunction(Description:="Heat capacity ratio as function of pressure and enthalpy, [-]", Category:="GCME E-PT | Steam Table")>
    Public Function CPCV_pT(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        Dim Cp As Double
        Dim Cv As Double
        Cp = Cp_pT(p, T)
        Cv = Cv_pT(p, T)
        CPCV_pT = Cp / Cv
    End Function

    <ExcelFunction(Description:="Heat capacity ratio as function of pressure and enthalpy, [-]", Category:="GCME E-PT | Steam Table")>
    Public Function CPCV_ph(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Enthalpy, [kJ/kg]")> h As Double) As Double
        Dim Cp As Double
        Dim Cv As Double
        Cv = Cv_ph(p, h)
        Cp = Cp_ph(p, h)
        CPCV_ph = Cp / Cv
    End Function
    '***********************************************************************************************************
    '*1.15 Surface tension

    <ExcelFunction(Description:="Surface tension for phase water/steam as function of temperature, [N/m]", Category:="GCME E-PT | Steam Table")>
    Public Function SurfaceTension_t(<ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        T = ToSIunit_T(T)
        SurfaceTension_t = FromSIunit_st(Surface_Tension_T(T))
    End Function

    <ExcelFunction(Description:="Surface tension for phase water/steam as function of pressure, [N/m]", Category:="GCME E-PT | Steam Table")>
    Public Function SurfaceTension_p(<ExcelArgument(Description:="Pressure, [bar]")> p As Double) As Double
        Dim T As Double
        T = Tsat_p(p)
        T = ToSIunit_T(T)
        SurfaceTension_p = FromSIunit_st(Surface_Tension_T(T))
    End Function
    '***********************************************************************************************************
    '*1.16 Thermal conductivity

    <ExcelFunction(Description:="Saturated liquid thermal conductivity as function of pressure, [W/m.K]", Category:="GCME E-PT | Steam Table")>
    Public Function ThermalConductivityL_p(<ExcelArgument(Description:="Pressure, [bar]")> p As Double) As Double
        Dim T As Double
        Dim v As Double
        T = Tsat_p(p)
        v = VolumeL_p(p)
        p = ToSIunit_p(p)
        T = ToSIunit_T(T)
        v = ToSIunit_v(v)
        ThermalConductivityL_p = FromSIunit_tc(Tc_ptrho(p, T, 1 / v))
    End Function

    <ExcelFunction(Description:="Saturated vapor thermal conductivity as function of pressure, [W/m.K]", Category:="GCME E-PT | Steam Table")>
    Public Function ThermalConductivityV_p(<ExcelArgument(Description:="Pressure, [bar]")> p As Double) As Double
        Dim T As Double
        Dim v As Double
        T = Tsat_p(p)
        v = VolumeV_p(p)
        p = ToSIunit_p(p)
        T = ToSIunit_T(T)
        v = ToSIunit_v(v)
        ThermalConductivityV_p = FromSIunit_tc(Tc_ptrho(p, T, 1 / v))
    End Function

    <ExcelFunction(Description:="Saturated liquid thermal conductivity as function of temperature, [W/m.K]", Category:="GCME E-PT | Steam Table")>
    Public Function ThermalConductivityL_T(<ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        Dim p, v As Double
        p = Psat_T(T)
        v = VolumeL_T(T)
        p = ToSIunit_p(p)
        T = ToSIunit_T(T)
        v = ToSIunit_v(v)
        ThermalConductivityL_T = FromSIunit_tc(Tc_ptrho(p, T, 1 / v))
    End Function

    <ExcelFunction(Description:="Saturated vapor thermal conductivity as function of temperature, [W/m.K]", Category:="GCME E-PT | Steam Table")>
    Public Function ThermalConductivityV_T(<ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        Dim p, v As Double
        p = Psat_T(T)
        v = VolumeV_T(T)
        p = ToSIunit_p(p)
        T = ToSIunit_T(T)
        v = ToSIunit_v(v)
        ThermalConductivityV_T = FromSIunit_tc(Tc_ptrho(p, T, 1 / v))
    End Function

    <ExcelFunction(Description:="Thermal conductivity as function of pressure and temperature, [W/m.K]", Category:="GCME E-PT | Steam Table")>
    Public Function ThermalConductivity_pT(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Temperature, [C]")> T As Double) As Double
        Dim v As Double
        v = Volume_pT(p, T)
        p = ToSIunit_p(p)
        T = ToSIunit_T(T)
        v = ToSIunit_v(v)
        ThermalConductivity_pT = FromSIunit_tc(Tc_ptrho(p, T, 1 / v))
    End Function

    <ExcelFunction(Description:="Thermal conductivity as function of pressure and enthalpy, [W/m.K]", Category:="GCME E-PT | Steam Table")>
    Public Function ThermalConductivity_ph(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Enthalpy, [kJ/kg]")> h As Double) As Double
        Dim v As Double
        Dim T As Double
        v = Volume_ph(p, h)
        T = Temp_ph(p, h)
        p = ToSIunit_p(p)
        T = ToSIunit_T(T)
        v = ToSIunit_v(v)
        ThermalConductivity_ph = FromSIunit_tc(Tc_ptrho(p, T, 1 / v))
    End Function

    <ExcelFunction(Description:="Thermal conductivity as function of enthalpy and entropy, [W/m.K]", Category:="GCME E-PT | Steam Table")>
    Public Function ThermalConductivity_hs(<ExcelArgument(Description:="Enthalpy, [kJ/kg]")> h As Double, <ExcelArgument(Description:="Entropy, [kJ/kg.K]")> s As Double) As Double
        Dim p As Double
        Dim v As Double
        Dim T As Double
        p = Pres_hs(h, s)
        v = Volume_ph(p, h)
        T = Temp_ph(p, h)
        p = ToSIunit_p(p)
        T = ToSIunit_T(T)
        v = ToSIunit_v(v)
        ThermalConductivity_hs = FromSIunit_tc(Tc_ptrho(p, T, 1 / v))
    End Function
    '***********************************************************************************************************
    '*1.17 Vapour fraction

    <ExcelFunction(Description:="Vapor fraction as function of pressure and enthalpy", Category:="GCME E-PT | Steam Table")>
    Public Function Vfract_ph(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Enthalpy, [kJ/kg]")> h As Double) As Double
        p = ToSIunit_p(p)
        h = ToSIunit_h(h)
        If p > 0.000611657 And p < 22.06395 Then
            Vfract_ph = FromSIunit_x(X4_ph(p, h))
        Else
            Vfract_ph = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Vapor fraction as function of pressure and entropy", Category:="GCME E-PT | Steam Table")>
    Public Function Vfrac_ps(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Entropy, [kJ/kg.K]")> s As Double) As Double
        p = ToSIunit_p(p)
        s = ToSIunit_s(s)
        If p > 0.000611657 And p < 22.06395 Then
            Vfrac_ps = FromSIunit_x(X4_ps(p, s))
        Else
            Vfrac_ps = -999.9
        End If
    End Function
    '***********************************************************************************************************
    '*1.18 Vapour Volume Fraction

    <ExcelFunction(Description:="Vapor volumen fraction as function of pressure and enthalpy", Category:="GCME E-PT | Steam Table")>
    Public Function VVfrac_ph(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Enthalpy, [kJ/kg]")> h As Double) As Double
        Dim vL As Double
        Dim vV As Double
        Dim xs As Double
        p = ToSIunit_p(p)
        h = ToSIunit_h(h)
        If p > 0.000611657 And p < 22.06395 Then
            If p < 16.529 Then
                vL = V1_pT(p, T4_p(p))
                vV = V2_pT(p, T4_p(p))
            Else
                vL = V3_ph(p, H4L_p(p))
                vV = V3_ph(p, H4V_p(p))
            End If
            xs = X4_ph(p, h)
            VVfrac_ph = FromSIunit_vx((xs * vV / (xs * vV + (1 - xs) * vL)))
        Else
            VVfrac_ph = -999.9
        End If
    End Function

    <ExcelFunction(Description:="Vapor volume fraction as function of pressure and enthropy", Category:="GCME E-PT | Steam Table")>
    Public Function VVfrac_ps(<ExcelArgument(Description:="Pressure, [bar]")> p As Double, <ExcelArgument(Description:="Entropy, [kJ/kg.K]")> s As Double) As Double
        Dim vL As Double
        Dim vV As Double
        Dim xs As Double
        p = ToSIunit_p(p)
        s = ToSIunit_s(s)
        If p > 0.000611657 And p < 22.06395 Then
            If p < 16.529 Then
                vL = V1_pT(p, T4_p(p))
                vV = V2_pT(p, T4_p(p))
            Else
                vL = V3_ph(p, h4L_p(p))
                vV = V3_ph(p, h4V_p(p))
            End If
            xs = x4_ps(p, s)
            VVfrac_ps = FromSIunit_vx((xs * vV / (xs * vV + (1 - xs) * vL)))
        Else
            VVfrac_ps = -999.9
        End If
    End Function

    '***********************************************************************************************************
    '*2 IAPWS IF 97 Calling functions                                                                          *
    '***********************************************************************************************************
    '
    '***********************************************************************************************************
    '*2.1 Functions for region 1
    Private Function V1_pT(ByVal p As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        '5 Equations for Region 1, Section. 5.1 Basic Equation
        'Eqution 7, Table 3, Page 6
        Dim i As Integer
        Dim ps, tau, g_p As Double
        Const R As Double = 0.461526 'kJ/(kg K)
        Dim I1 = New Double() {0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 4, 4, 4, 5, 8, 8, 21, 23, 29, 30, 31, 32}
        Dim J1 = New Double() {-2, -1, 0, 1, 2, 3, 4, 5, -9, -7, -1, 0, 1, 3, -3, 0, 1, 3, 17, -4, 0, 6, -5, -2, 10, -8, -11, -6, -29, -31, -38, -39, -40, -41}
        Dim n1 = New Double() {0.14632971213167, -0.84548187169114, -3.756360367204, 3.3855169168385, -0.95791963387872, 0.15772038513228, -0.016616417199501, 0.00081214629983568, 0.00028319080123804, -0.00060706301565874, -0.018990068218419, -0.032529748770505, -0.021841717175414, -0.00005283835796993, -0.00047184321073267, -0.00030001780793026, 0.000047661393906987, -0.0000044141845330846, -0.00000000000000072694996297594, -0.000031679644845054, -0.0000028270797985312, -0.00000000085205128120103, -0.0000022425281908, -0.00000065171222895601, -0.00000000000014341729937924, -0.00000040516996860117, -0.0000000012734301741641, -0.00000000017424871230634, -6.8762131295531E-19, 1.4478307828521E-20, 2.6335781662795E-23, -1.1947622640071E-23, 1.8228094581404E-24, -9.3537087292458E-26}
        ps = p / 16.53
        tau = 1386 / T
        g_p = 0#
        For i = 0 To 33
            g_p -= n1(i) * I1(i) * (7.1 - ps) ^ (I1(i) - 1) * (tau - 1.222) ^ J1(i)
        Next i
        V1_pT = R * T / p * ps * g_p / 1000
    End Function
    Private Function H1_pT(ByVal p As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        '5 Equations for Region 1, Section. 5.1 Basic Equation
        'Eqution 7, Table 3, Page 6
        Dim i As Integer
        Dim tau, g_t As Double
        Const R As Double = 0.461526 'kJ/(kg K)
        Dim I1 = New Double() {0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 4, 4, 4, 5, 8, 8, 21, 23, 29, 30, 31, 32}
        Dim J1 = New Double() {-2, -1, 0, 1, 2, 3, 4, 5, -9, -7, -1, 0, 1, 3, -3, 0, 1, 3, 17, -4, 0, 6, -5, -2, 10, -8, -11, -6, -29, -31, -38, -39, -40, -41}
        Dim n1 = New Double() {0.14632971213167, -0.84548187169114, -3.756360367204, 3.3855169168385, -0.95791963387872, 0.15772038513228, -0.016616417199501, 0.00081214629983568, 0.00028319080123804, -0.00060706301565874, -0.018990068218419, -0.032529748770505, -0.021841717175414, -0.00005283835796993, -0.00047184321073267, -0.00030001780793026, 0.000047661393906987, -0.0000044141845330846, -0.00000000000000072694996297594, -0.000031679644845054, -0.0000028270797985312, -0.00000000085205128120103, -0.0000022425281908, -0.00000065171222895601, -0.00000000000014341729937924, -0.00000040516996860117, -0.0000000012734301741641, -0.00000000017424871230634, -6.8762131295531E-19, 1.4478307828521E-20, 2.6335781662795E-23, -1.1947622640071E-23, 1.8228094581404E-24, -9.3537087292458E-26}
        p /= 16.53
        tau = 1386 / T
        g_t = 0#
        For i = 0 To 33
            g_t += (n1(i) * (7.1 - p) ^ I1(i) * J1(i) * (tau - 1.222) ^ (J1(i) - 1))
        Next i
        H1_pT = R * T * tau * g_t
    End Function
    Private Function U1_pT(ByVal p As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        '5 Equations for Region 1, Section. 5.1 Basic Equation
        'Eqution 7, Table 3, Page 6
        Dim i As Integer
        Dim tau, g_t, g_p As Double
        Const R As Double = 0.461526 'kJ/(kg K)
        Dim I1 = New Double() {0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 4, 4, 4, 5, 8, 8, 21, 23, 29, 30, 31, 32}
        Dim J1 = New Double() {-2, -1, 0, 1, 2, 3, 4, 5, -9, -7, -1, 0, 1, 3, -3, 0, 1, 3, 17, -4, 0, 6, -5, -2, 10, -8, -11, -6, -29, -31, -38, -39, -40, -41}
        Dim n1 = New Double() {0.14632971213167, -0.84548187169114, -3.756360367204, 3.3855169168385, -0.95791963387872, 0.15772038513228, -0.016616417199501, 0.00081214629983568, 0.00028319080123804, -0.00060706301565874, -0.018990068218419, -0.032529748770505, -0.021841717175414, -0.00005283835796993, -0.00047184321073267, -0.00030001780793026, 0.000047661393906987, -0.0000044141845330846, -0.00000000000000072694996297594, -0.000031679644845054, -0.0000028270797985312, -0.00000000085205128120103, -0.0000022425281908, -0.00000065171222895601, -0.00000000000014341729937924, -0.00000040516996860117, -0.0000000012734301741641, -0.00000000017424871230634, -6.8762131295531E-19, 1.4478307828521E-20, 2.6335781662795E-23, -1.1947622640071E-23, 1.8228094581404E-24, -9.3537087292458E-26}
        p /= 16.53
        tau = 1386 / T
        g_t = 0#
        g_p = 0#
        For i = 0 To 33
            g_p -= n1(i) * I1(i) * (7.1 - p) ^ (I1(i) - 1) * (tau - 1.222) ^ J1(i)
            g_t += +(n1(i) * (7.1 - p) ^ I1(i) * J1(i) * (tau - 1.222) ^ (J1(i) - 1))
        Next i
        U1_pT = R * T * (tau * g_t - p * g_p)
    End Function
    Private Function S1_pT(ByVal p As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        '5 Equations for Region 1, Section. 5.1 Basic Equation
        'Eqution 7, Table 3, Page 6
        Dim i As Integer
        Dim g, g_t As Double
        Const R As Double = 0.461526 'kJ/(kg K)
        Dim I1 = New Double() {0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 4, 4, 4, 5, 8, 8, 21, 23, 29, 30, 31, 32}
        Dim J1 = New Double() {-2, -1, 0, 1, 2, 3, 4, 5, -9, -7, -1, 0, 1, 3, -3, 0, 1, 3, 17, -4, 0, 6, -5, -2, 10, -8, -11, -6, -29, -31, -38, -39, -40, -41}
        Dim n1 = New Double() {0.14632971213167, -0.84548187169114, -3.756360367204, 3.3855169168385, -0.95791963387872, 0.15772038513228, -0.016616417199501, 0.00081214629983568, 0.00028319080123804, -0.00060706301565874, -0.018990068218419, -0.032529748770505, -0.021841717175414, -0.00005283835796993, -0.00047184321073267, -0.00030001780793026, 0.000047661393906987, -0.0000044141845330846, -0.00000000000000072694996297594, -0.000031679644845054, -0.0000028270797985312, -0.00000000085205128120103, -0.0000022425281908, -0.00000065171222895601, -0.00000000000014341729937924, -0.00000040516996860117, -0.0000000012734301741641, -0.00000000017424871230634, -6.8762131295531E-19, 1.4478307828521E-20, 2.6335781662795E-23, -1.1947622640071E-23, 1.8228094581404E-24, -9.3537087292458E-26}
        p /= 16.53
        T = 1386 / T
        g = 0#
        g_t = 0#
        For i = 0 To 33
            g_t += (n1(i) * (7.1 - p) ^ I1(i) * J1(i) * (T - 1.222) ^ (J1(i) - 1))
            g += n1(i) * (7.1 - p) ^ I1(i) * (T - 1.222) ^ J1(i)
        Next i
        S1_pT = R * T * g_t - R * g
    End Function
    Private Function Cp1_pT(ByVal p As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        '5 Equations for Region 1, Section. 5.1 Basic Equation
        'Eqution 7, Table 3, Page 6
        Dim i As Integer
        Dim G_tt As Double
        Const R As Double = 0.461526 'kJ/(kg K)
        Dim I1 = New Double() {0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 4, 4, 4, 5, 8, 8, 21, 23, 29, 30, 31, 32}
        Dim J1 = New Double() {-2, -1, 0, 1, 2, 3, 4, 5, -9, -7, -1, 0, 1, 3, -3, 0, 1, 3, 17, -4, 0, 6, -5, -2, 10, -8, -11, -6, -29, -31, -38, -39, -40, -41}
        Dim n1 = New Double() {0.14632971213167, -0.84548187169114, -3.756360367204, 3.3855169168385, -0.95791963387872, 0.15772038513228, -0.016616417199501, 0.00081214629983568, 0.00028319080123804, -0.00060706301565874, -0.018990068218419, -0.032529748770505, -0.021841717175414, -0.00005283835796993, -0.00047184321073267, -0.00030001780793026, 0.000047661393906987, -0.0000044141845330846, -0.00000000000000072694996297594, -0.000031679644845054, -0.0000028270797985312, -0.00000000085205128120103, -0.0000022425281908, -0.00000065171222895601, -0.00000000000014341729937924, -0.00000040516996860117, -0.0000000012734301741641, -0.00000000017424871230634, -6.8762131295531E-19, 1.4478307828521E-20, 2.6335781662795E-23, -1.1947622640071E-23, 1.8228094581404E-24, -9.3537087292458E-26}
        p /= 16.53
        T = 1386 / T
        G_tt = 0#
        For i = 0 To 33
            G_tt += (n1(i) * (7.1 - p) ^ I1(i) * J1(i) * (J1(i) - 1) * (T - 1.222) ^ (J1(i) - 2))
        Next i
        Cp1_pT = -R * T ^ 2 * G_tt
    End Function
    Private Function Cv1_pT(ByVal p As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        '5 Equations for Region 1, Section. 5.1 Basic Equation
        'Eqution 7, Table 3, Page 6
        Dim i As Integer
        Dim g_p, g_pp, g_pt, G_tt As Double
        Const R As Double = 0.461526 'kJ/(kg K)
        Dim I1 = New Double() {0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 4, 4, 4, 5, 8, 8, 21, 23, 29, 30, 31, 32}
        Dim J1 = New Double() {-2, -1, 0, 1, 2, 3, 4, 5, -9, -7, -1, 0, 1, 3, -3, 0, 1, 3, 17, -4, 0, 6, -5, -2, 10, -8, -11, -6, -29, -31, -38, -39, -40, -41}
        Dim n1 = New Double() {0.14632971213167, -0.84548187169114, -3.756360367204, 3.3855169168385, -0.95791963387872, 0.15772038513228, -0.016616417199501, 0.00081214629983568, 0.00028319080123804, -0.00060706301565874, -0.018990068218419, -0.032529748770505, -0.021841717175414, -0.00005283835796993, -0.00047184321073267, -0.00030001780793026, 0.000047661393906987, -0.0000044141845330846, -0.00000000000000072694996297594, -0.000031679644845054, -0.0000028270797985312, -0.00000000085205128120103, -0.0000022425281908, -0.00000065171222895601, -0.00000000000014341729937924, -0.00000040516996860117, -0.0000000012734301741641, -0.00000000017424871230634, -6.8762131295531E-19, 1.4478307828521E-20, 2.6335781662795E-23, -1.1947622640071E-23, 1.8228094581404E-24, -9.3537087292458E-26}
        p /= 16.53
        T = 1386 / T
        g_p = 0#
        g_pp = 0#
        g_pt = 0#
        G_tt = 0#
        For i = 0 To 33
            g_p -= n1(i) * I1(i) * (7.1 - p) ^ (I1(i) - 1) * (T - 1.222) ^ J1(i)
            g_pp += n1(i) * I1(i) * (I1(i) - 1) * (7.1 - p) ^ (I1(i) - 2) * (T - 1.222) ^ J1(i)
            g_pt -= n1(i) * I1(i) * (7.1 - p) ^ (I1(i) - 1) * J1(i) * (T - 1.222) ^ (J1(i) - 1)
            G_tt += n1(i) * (7.1 - p) ^ I1(i) * J1(i) * (J1(i) - 1) * (T - 1.222) ^ (J1(i) - 2)
        Next i
        Cv1_pT = R * (-(T ^ 2 * G_tt) + (g_p - T * g_pt) ^ 2 / g_pp)
    End Function
    Private Function W1_pT(ByVal p As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        '5 Equations for Region 1, Section. 5.1 Basic Equation
        'Eqution 7, Table 3, Page 6
        Dim i As Integer
        Dim g_p, g_pp, g_pt, G_tt, tau As Double
        Const R As Double = 0.461526 'kJ/(kg K)
        Dim I1 = New Double() {0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 4, 4, 4, 5, 8, 8, 21, 23, 29, 30, 31, 32}
        Dim J1 = New Double() {-2, -1, 0, 1, 2, 3, 4, 5, -9, -7, -1, 0, 1, 3, -3, 0, 1, 3, 17, -4, 0, 6, -5, -2, 10, -8, -11, -6, -29, -31, -38, -39, -40, -41}
        Dim n1 = New Double() {0.14632971213167, -0.84548187169114, -3.756360367204, 3.3855169168385, -0.95791963387872, 0.15772038513228, -0.016616417199501, 0.00081214629983568, 0.00028319080123804, -0.00060706301565874, -0.018990068218419, -0.032529748770505, -0.021841717175414, -0.00005283835796993, -0.00047184321073267, -0.00030001780793026, 0.000047661393906987, -0.0000044141845330846, -0.00000000000000072694996297594, -0.000031679644845054, -0.0000028270797985312, -0.00000000085205128120103, -0.0000022425281908, -0.00000065171222895601, -0.00000000000014341729937924, -0.00000040516996860117, -0.0000000012734301741641, -0.00000000017424871230634, -6.8762131295531E-19, 1.4478307828521E-20, 2.6335781662795E-23, -1.1947622640071E-23, 1.8228094581404E-24, -9.3537087292458E-26}
        p /= 16.53
        tau = 1386 / T
        g_p = 0#
        g_pp = 0#
        g_pt = 0#
        G_tt = 0#
        For i = 0 To 33
            g_p -= n1(i) * I1(i) * (7.1 - p) ^ (I1(i) - 1) * (tau - 1.222) ^ J1(i)
            g_pp += n1(i) * I1(i) * (I1(i) - 1) * (7.1 - p) ^ (I1(i) - 2) * (tau - 1.222) ^ J1(i)
            g_pt -= n1(i) * I1(i) * (7.1 - p) ^ (I1(i) - 1) * J1(i) * (tau - 1.222) ^ (J1(i) - 1)
            G_tt += n1(i) * (7.1 - p) ^ I1(i) * J1(i) * (J1(i) - 1) * (tau - 1.222) ^ (J1(i) - 2)
        Next i
        W1_pT = (1000 * R * T * g_p ^ 2 / ((g_p - tau * g_pt) ^ 2 / (tau ^ 2 * G_tt) - g_pp)) ^ 0.5
    End Function
    Private Function T1_ph(ByVal p As Double, ByVal h As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        '5 Equations for Region 1, Section. 5.1 Basic Equation, 5.2.1 The Backward Equation T ( p,h )
        'Eqution 11, Table 6, Page 10
        Dim i As Integer
        Dim T As Double
        Dim I1 = New Double() {0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 2, 2, 3, 3, 4, 5, 6}
        Dim J1 = New Double() {0, 1, 2, 6, 22, 32, 0, 1, 2, 3, 4, 10, 32, 10, 32, 10, 32, 32, 32, 32}
        Dim n1 = New Double() {-238.72489924521, 404.21188637945, 113.49746881718, -5.8457616048039, -0.0001528548241314, -0.0000010866707695377, -13.391744872602, 43.211039183559, -54.010067170506, 30.535892203916, -6.5964749423638, 0.0093965400878363, 0.0000001157364750534, -0.000025858641282073, -0.0000000040644363084799, 0.000000066456186191635, 0.000000000080670734103027, -0.00000000000093477771213947, 0.0000000000000058265442020601, -1.5020185953503E-17}
        h /= 2500
        T = 0#
        For i = 0 To 19
            T += n1(i) * p ^ I1(i) * (h + 1) ^ J1(i)
        Next i
        T1_ph = T
    End Function
    Private Function T1_ps(ByVal p As Double, ByVal s As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        '5 Equations for Region 1, Section. 5.1 Basic Equation, 5.2.2 The Backward Equation T ( p, s )
        'Eqution 13, Table 8, Page 11
        Dim i As Integer
        Dim I1 = New Double() {0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 4}
        Dim J1 = New Double() {0, 1, 2, 3, 11, 31, 0, 1, 2, 3, 12, 31, 0, 1, 2, 9, 31, 10, 32, 32}
        Dim n1 = New Double() {174.78268058307, 34.806930892873, 6.5292584978455, 0.33039981775489, -0.00000019281382923196, -2.4909197244573E-23, -0.26107636489332, 0.22592965981586, -0.064256463395226, 0.0078876289270526, 0.00000000035672110607366, 1.7332496994895E-24, 0.00056608900654837, -0.00032635483139717, 0.000044778286690632, -0.00000000051322156908507, -4.2522657042207E-26, 0.00000000000026400441360689, 7.8124600459723E-29, -3.0732199903668E-31}
        T1_ps = 0#
        For i = 0 To 19
            T1_ps += n1(i) * p ^ I1(i) * (s + 2) ^ J1(i)
        Next i
    End Function
    Private Function P1_hs(ByVal h As Double, ByVal s As Double) As Double
        'Supplementary Release on Backward Equations for Pressure as a function of Enthalpy and Entropy p(h,s) to the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam
        '5 Backward Equation p(h,s) for Region 1
        'Eqution 1, Table 2, Page 5
        Dim i As Integer
        Dim p As Double
        Dim I1 = New Double() {0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 3, 4, 4, 5}
        Dim J1 = New Double() {0, 1, 2, 4, 5, 6, 8, 14, 0, 1, 4, 6, 0, 1, 10, 4, 1, 4, 0}
        Dim n1 = New Double() {-0.691997014660582, -18.361254878756, -9.28332409297335, 65.9639569909906, -16.2060388912024, 450.620017338667, 854.68067822417, 6075.23214001162, 32.6487682621856, -26.9408844582931, -319.9478483343, -928.35430704332, 30.3634537455249, -65.0540422444146, -4309.9131651613, -747.512324096068, 730.000345529245, 1142.84032569021, -436.407041874559}
        h /= 3400
        s /= 7.6
        p = 0#
        For i = 0 To 18
            p += n1(i) * (h + 0.05) ^ I1(i) * (s + 0.05) ^ J1(i)
        Next i
        P1_hs = p * 100
    End Function
    Private Function T1_prho(ByVal p As Double, ByVal rho As Double) As Double
        'Solve by iteration. Observe that fo low temperatures this equation has 2 solutions.
        'Solve with half interval method
        Dim Ts, Low_Bound, High_Bound, rhos As Double
        Low_Bound = 273.15
        High_Bound = T4_p(p)
        Do While Math.Abs(rho - rhos) > 0.00001
            Ts = (Low_Bound + High_Bound) / 2
            rhos = 1 / V1_pT(p, Ts)
            If rhos < rho Then
                High_Bound = Ts
            Else
                Low_Bound = Ts
            End If
        Loop
        T1_prho = Ts
    End Function
    '***********************************************************************************************************
    '*2.2 Functions for region 2

    Private Function V2_pT(ByVal p As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        '6 Equations for Region 2, Section. 6.1 Basic Equation
        'Table 11 and 12, Page 14 and 15
        Dim i As Integer
        Dim tau, g0_pi, gr_pi As Double
        Const R As Double = 0.461526 'kJ/(kg K)
        Dim J0 = New Double() {0, 1, -5, -4, -3, -2, -1, 2, 3}
        Dim n0 = New Double() {-9.6927686500217, 10.086655968018, -0.005608791128302, 0.071452738081455, -0.40710498223928, 1.4240819171444, -4.383951131945, -0.28408632460772, 0.021268463753307}
        Dim Ir = New Double() {1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 4, 4, 4, 5, 6, 6, 6, 7, 7, 7, 8, 8, 9, 10, 10, 10, 16, 16, 18, 20, 20, 20, 21, 22, 23, 24, 24, 24}
        Dim Jr = New Double() {0, 1, 2, 3, 6, 1, 2, 4, 7, 36, 0, 1, 3, 6, 35, 1, 2, 3, 7, 3, 16, 35, 0, 11, 25, 8, 36, 13, 4, 10, 14, 29, 50, 57, 20, 35, 48, 21, 53, 39, 26, 40, 58}
        Dim nr = New Double() {-0.0017731742473213, -0.017834862292358, -0.045996013696365, -0.057581259083432, -0.05032527872793, -0.000033032641670203, -0.00018948987516315, -0.0039392777243355, -0.043797295650573, -0.000026674547914087, 0.000000020481737692309, 0.00000043870667284435, -0.00003227767723857, -0.0015033924542148, -0.040668253562649, -0.00000000078847309559367, 0.000000012790717852285, 0.00000048225372718507, 0.0000022922076337661, -0.000000000016714766451061, -0.0021171472321355, -23.895741934104, -5.905956432427E-18, -0.0000012621808899101, -0.038946842435739, 0.000000000011256211360459, -8.2311340897998, 0.000000019809712802088, 1.0406965210174E-19, -0.00000000000010234747095929, -0.0000000010018179379511, -0.000000000080882908646985, 0.10693031879409, -0.33662250574171, 8.9185845355421E-25, 0.00000000000030629316876231997, -0.0000042002467698208, -5.9056029685639E-26, 0.0000037826947613457, -0.0000000000000012768608934681, 7.3087610595061E-29, 5.5414715350778E-17, -0.0000009436970724121}
        tau = 540 / T
        g0_pi = 1 / p
        gr_pi = 0#
        For i = 0 To 42
            gr_pi += nr(i) * Ir(i) * p ^ (Ir(i) - 1) * (tau - 0.5) ^ Jr(i)
        Next i
        V2_pT = R * T / p * p * (g0_pi + gr_pi) / 1000
    End Function
    Private Function H2_pT(ByVal p As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        '6 Equations for Region 2, Section. 6.1 Basic Equation
        'Table 11 and 12, Page 14 and 15
        Dim i As Integer
        Dim tau, g0_tau, gr_tau As Double
        Const R As Double = 0.461526 'kJ/(kg K)
        Dim J0 = New Double() {0, 1, -5, -4, -3, -2, -1, 2, 3}
        Dim n0 = New Double() {-9.6927686500217, 10.086655968018, -0.005608791128302, 0.071452738081455, -0.40710498223928, 1.4240819171444, -4.383951131945, -0.28408632460772, 0.021268463753307}
        Dim Ir = New Double() {1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 4, 4, 4, 5, 6, 6, 6, 7, 7, 7, 8, 8, 9, 10, 10, 10, 16, 16, 18, 20, 20, 20, 21, 22, 23, 24, 24, 24}
        Dim Jr = New Double() {0, 1, 2, 3, 6, 1, 2, 4, 7, 36, 0, 1, 3, 6, 35, 1, 2, 3, 7, 3, 16, 35, 0, 11, 25, 8, 36, 13, 4, 10, 14, 29, 50, 57, 20, 35, 48, 21, 53, 39, 26, 40, 58}
        Dim nr = New Double() {-0.0017731742473213, -0.017834862292358, -0.045996013696365, -0.057581259083432, -0.05032527872793, -0.000033032641670203, -0.00018948987516315, -0.0039392777243355, -0.043797295650573, -0.000026674547914087, 0.000000020481737692309, 0.00000043870667284435, -0.00003227767723857, -0.0015033924542148, -0.040668253562649, -0.00000000078847309559367, 0.000000012790717852285, 0.00000048225372718507, 0.0000022922076337661, -0.000000000016714766451061, -0.0021171472321355, -23.895741934104, -5.905956432427E-18, -0.0000012621808899101, -0.038946842435739, 0.000000000011256211360459, -8.2311340897998, 0.000000019809712802088, 1.0406965210174E-19, -0.00000000000010234747095929, -0.0000000010018179379511, -0.000000000080882908646985, 0.10693031879409, -0.33662250574171, 8.9185845355421E-25, 0.00000000000030629316876231997, -0.0000042002467698208, -5.9056029685639E-26, 0.0000037826947613457, -0.0000000000000012768608934681, 7.3087610595061E-29, 5.5414715350778E-17, -0.0000009436970724121}
        tau = 540 / T
        g0_tau = 0#
        For i = 0 To 8
            g0_tau += n0(i) * J0(i) * tau ^ (J0(i) - 1)
        Next i
        gr_tau = 0#
        For i = 0 To 42
            gr_tau += nr(i) * p ^ Ir(i) * Jr(i) * (tau - 0.5) ^ (Jr(i) - 1)
        Next i
        H2_pT = R * T * tau * (g0_tau + gr_tau)
    End Function
    Private Function U2_pT(ByVal p As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        '6 Equations for Region 2, Section. 6.1 Basic Equation
        'Table 11 and 12, Page 14 and 15
        Dim i As Integer
        Dim tau, g0_tau, g0_pi, gr_pi, gr_tau As Double
        Const R As Double = 0.461526 'kJ/(kg K)
        Dim J0 = New Double() {0, 1, -5, -4, -3, -2, -1, 2, 3}
        Dim n0 = New Double() {-9.6927686500217, 10.086655968018, -0.005608791128302, 0.071452738081455, -0.40710498223928, 1.4240819171444, -4.383951131945, -0.28408632460772, 0.021268463753307}
        Dim Ir = New Double() {1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 4, 4, 4, 5, 6, 6, 6, 7, 7, 7, 8, 8, 9, 10, 10, 10, 16, 16, 18, 20, 20, 20, 21, 22, 23, 24, 24, 24}
        Dim Jr = New Double() {0, 1, 2, 3, 6, 1, 2, 4, 7, 36, 0, 1, 3, 6, 35, 1, 2, 3, 7, 3, 16, 35, 0, 11, 25, 8, 36, 13, 4, 10, 14, 29, 50, 57, 20, 35, 48, 21, 53, 39, 26, 40, 58}
        Dim nr = New Double() {-0.0017731742473213, -0.017834862292358, -0.045996013696365, -0.057581259083432, -0.05032527872793, -0.000033032641670203, -0.00018948987516315, -0.0039392777243355, -0.043797295650573, -0.000026674547914087, 0.000000020481737692309, 0.00000043870667284435, -0.00003227767723857, -0.0015033924542148, -0.040668253562649, -0.00000000078847309559367, 0.000000012790717852285, 0.00000048225372718507, 0.0000022922076337661, -0.000000000016714766451061, -0.0021171472321355, -23.895741934104, -5.905956432427E-18, -0.0000012621808899101, -0.038946842435739, 0.000000000011256211360459, -8.2311340897998, 0.000000019809712802088, 1.0406965210174E-19, -0.00000000000010234747095929, -0.0000000010018179379511, -0.000000000080882908646985, 0.10693031879409, -0.33662250574171, 8.9185845355421E-25, 0.00000000000030629316876231997, -0.0000042002467698208, -5.9056029685639E-26, 0.0000037826947613457, -0.0000000000000012768608934681, 7.3087610595061E-29, 5.5414715350778E-17, -0.0000009436970724121}
        tau = 540 / T
        g0_pi = 1 / p
        g0_tau = 0#
        For i = 0 To 8
            g0_tau += n0(i) * J0(i) * tau ^ (J0(i) - 1)
        Next i
        gr_pi = 0#
        gr_tau = 0#
        For i = 0 To 42
            gr_pi += nr(i) * Ir(i) * p ^ (Ir(i) - 1) * (tau - 0.5) ^ Jr(i)
            gr_tau += nr(i) * p ^ Ir(i) * Jr(i) * (tau - 0.5) ^ (Jr(i) - 1)
        Next i
        U2_pT = R * T * (tau * (g0_tau + gr_tau) - p * (g0_pi + gr_pi))
    End Function

    Private Function S2_pT(ByVal p As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        '6 Equations for Region 2, Section. 6.1 Basic Equation
        'Table 11 and 12, Page 14 and 15
        Dim i As Integer
        Dim tau, g0, g0_tau, gr, gr_tau As Double
        Const R As Double = 0.461526 'kJ/(kg K)
        Dim J0 = New Double() {0, 1, -5, -4, -3, -2, -1, 2, 3}
        Dim n0 = New Double() {-9.6927686500217, 10.086655968018, -0.005608791128302, 0.071452738081455, -0.40710498223928, 1.4240819171444, -4.383951131945, -0.28408632460772, 0.021268463753307}
        Dim Ir = New Double() {1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 4, 4, 4, 5, 6, 6, 6, 7, 7, 7, 8, 8, 9, 10, 10, 10, 16, 16, 18, 20, 20, 20, 21, 22, 23, 24, 24, 24}
        Dim Jr = New Double() {0, 1, 2, 3, 6, 1, 2, 4, 7, 36, 0, 1, 3, 6, 35, 1, 2, 3, 7, 3, 16, 35, 0, 11, 25, 8, 36, 13, 4, 10, 14, 29, 50, 57, 20, 35, 48, 21, 53, 39, 26, 40, 58}
        Dim nr = New Double() {-0.0017731742473213, -0.017834862292358, -0.045996013696365, -0.057581259083432, -0.05032527872793, -0.000033032641670203, -0.00018948987516315, -0.0039392777243355, -0.043797295650573, -0.000026674547914087, 0.000000020481737692309, 0.00000043870667284435, -0.00003227767723857, -0.0015033924542148, -0.040668253562649, -0.00000000078847309559367, 0.000000012790717852285, 0.00000048225372718507, 0.0000022922076337661, -0.000000000016714766451061, -0.0021171472321355, -23.895741934104, -5.905956432427E-18, -0.0000012621808899101, -0.038946842435739, 0.000000000011256211360459, -8.2311340897998, 0.000000019809712802088, 1.0406965210174E-19, -0.00000000000010234747095929, -0.0000000010018179379511, -0.000000000080882908646985, 0.10693031879409, -0.33662250574171, 8.9185845355421E-25, 0.00000000000030629316876231997, -0.0000042002467698208, -5.9056029685639E-26, 0.0000037826947613457, -0.0000000000000012768608934681, 7.3087610595061E-29, 5.5414715350778E-17, -0.0000009436970724121}
        tau = 540 / T
        g0 = Math.Log(p)
        g0_tau = 0#
        For i = 0 To 8
            g0 += n0(i) * tau ^ J0(i)
            g0_tau += n0(i) * J0(i) * tau ^ (J0(i) - 1)
        Next i
        gr = 0#
        gr_tau = 0#
        For i = 0 To 42
            gr += nr(i) * p ^ Ir(i) * (tau - 0.5) ^ Jr(i)
            gr_tau += nr(i) * p ^ Ir(i) * Jr(i) * (tau - 0.5) ^ (Jr(i) - 1)
        Next i
        S2_pT = R * (tau * (g0_tau + gr_tau) - (g0 + gr))
    End Function
    Private Function Cp2_pT(ByVal p As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        '6 Equations for Region 2, Section. 6.1 Basic Equation
        'Table 11 and 12, Page 14 and 15
        Dim i As Integer
        Dim tau, g0_tautau, gr_tautau As Double
        Const R As Double = 0.461526 'kJ/(kg K)
        Dim J0 = New Double() {0, 1, -5, -4, -3, -2, -1, 2, 3}
        Dim n0 = New Double() {-9.6927686500217, 10.086655968018, -0.005608791128302, 0.071452738081455, -0.40710498223928, 1.4240819171444, -4.383951131945, -0.28408632460772, 0.021268463753307}
        Dim Ir = New Double() {1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 4, 4, 4, 5, 6, 6, 6, 7, 7, 7, 8, 8, 9, 10, 10, 10, 16, 16, 18, 20, 20, 20, 21, 22, 23, 24, 24, 24}
        Dim Jr = New Double() {0, 1, 2, 3, 6, 1, 2, 4, 7, 36, 0, 1, 3, 6, 35, 1, 2, 3, 7, 3, 16, 35, 0, 11, 25, 8, 36, 13, 4, 10, 14, 29, 50, 57, 20, 35, 48, 21, 53, 39, 26, 40, 58}
        Dim nr = New Double() {-0.0017731742473213, -0.017834862292358, -0.045996013696365, -0.057581259083432, -0.05032527872793, -0.000033032641670203, -0.00018948987516315, -0.0039392777243355, -0.043797295650573, -0.000026674547914087, 0.000000020481737692309, 0.00000043870667284435, -0.00003227767723857, -0.0015033924542148, -0.040668253562649, -0.00000000078847309559367, 0.000000012790717852285, 0.00000048225372718507, 0.0000022922076337661, -0.000000000016714766451061, -0.0021171472321355, -23.895741934104, -5.905956432427E-18, -0.0000012621808899101, -0.038946842435739, 0.000000000011256211360459, -8.2311340897998, 0.000000019809712802088, 1.0406965210174E-19, -0.00000000000010234747095929, -0.0000000010018179379511, -0.000000000080882908646985, 0.10693031879409, -0.33662250574171, 8.9185845355421E-25, 0.00000000000030629316876231997, -0.0000042002467698208, -5.9056029685639E-26, 0.0000037826947613457, -0.0000000000000012768608934681, 7.3087610595061E-29, 5.5414715350778E-17, -0.0000009436970724121}
        tau = 540 / T
        g0_tautau = 0#
        For i = 0 To 8
            g0_tautau += n0(i) * J0(i) * (J0(i) - 1) * tau ^ (J0(i) - 2)
        Next i
        gr_tautau = 0#
        For i = 0 To 42
            gr_tautau += nr(i) * p ^ Ir(i) * Jr(i) * (Jr(i) - 1) * (tau - 0.5) ^ (Jr(i) - 2)
        Next i
        Cp2_pT = -R * tau ^ 2 * (g0_tautau + gr_tautau)
    End Function
    Private Function Cv2_pT(ByVal p As Double, ByVal T As Double) As Double
        Dim i As Integer
        Dim tau, g0_tautau, gr_pi, gr_pitau, gr_pipi, gr_tautau As Double
        Const R As Double = 0.461526 'kJ/(kg K)
        Dim J0 = New Double() {0, 1, -5, -4, -3, -2, -1, 2, 3}
        Dim n0 = New Double() {-9.6927686500217, 10.086655968018, -0.005608791128302, 0.071452738081455, -0.40710498223928, 1.4240819171444, -4.383951131945, -0.28408632460772, 0.021268463753307}
        Dim Ir = New Double() {1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 4, 4, 4, 5, 6, 6, 6, 7, 7, 7, 8, 8, 9, 10, 10, 10, 16, 16, 18, 20, 20, 20, 21, 22, 23, 24, 24, 24}
        Dim Jr = New Double() {0, 1, 2, 3, 6, 1, 2, 4, 7, 36, 0, 1, 3, 6, 35, 1, 2, 3, 7, 3, 16, 35, 0, 11, 25, 8, 36, 13, 4, 10, 14, 29, 50, 57, 20, 35, 48, 21, 53, 39, 26, 40, 58}
        Dim nr = New Double() {-0.0017731742473213, -0.017834862292358, -0.045996013696365, -0.057581259083432, -0.05032527872793, -0.000033032641670203, -0.00018948987516315, -0.0039392777243355, -0.043797295650573, -0.000026674547914087, 0.000000020481737692309, 0.00000043870667284435, -0.00003227767723857, -0.0015033924542148, -0.040668253562649, -0.00000000078847309559367, 0.000000012790717852285, 0.00000048225372718507, 0.0000022922076337661, -0.000000000016714766451061, -0.0021171472321355, -23.895741934104, -5.905956432427E-18, -0.0000012621808899101, -0.038946842435739, 0.000000000011256211360459, -8.2311340897998, 0.000000019809712802088, 1.0406965210174E-19, -0.00000000000010234747095929, -0.0000000010018179379511, -0.000000000080882908646985, 0.10693031879409, -0.33662250574171, 8.9185845355421E-25, 0.00000000000030629316876231997, -0.0000042002467698208, -5.9056029685639E-26, 0.0000037826947613457, -0.0000000000000012768608934681, 7.3087610595061E-29, 5.5414715350778E-17, -0.0000009436970724121}
        tau = 540 / T
        g0_tautau = 0#
        For i = 0 To 8
            g0_tautau += n0(i) * J0(i) * (J0(i) - 1) * tau ^ (J0(i) - 2)
        Next i
        gr_pi = 0#
        gr_pitau = 0#
        gr_pipi = 0#
        gr_tautau = 0#
        For i = 0 To 42
            gr_pi += nr(i) * Ir(i) * p ^ (Ir(i) - 1) * (tau - 0.5) ^ Jr(i)
            gr_pipi += nr(i) * Ir(i) * (Ir(i) - 1) * p ^ (Ir(i) - 2) * (tau - 0.5) ^ Jr(i)
            gr_pitau += nr(i) * Ir(i) * p ^ (Ir(i) - 1) * Jr(i) * (tau - 0.5) ^ (Jr(i) - 1)
            gr_tautau += nr(i) * p ^ Ir(i) * Jr(i) * (Jr(i) - 1) * (tau - 0.5) ^ (Jr(i) - 2)
        Next i
        Cv2_pT = R * (-(tau ^ 2 * (g0_tautau + gr_tautau)) - ((1 + p * gr_pi - tau * p * gr_pitau) ^ 2) / (1 - p ^ 2 * gr_pipi))
    End Function
    Private Function W2_pT(ByVal p As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        '6 Equations for Region 2, Section. 6.1 Basic Equation
        'Table 11 and 12, Page 14 and 15
        Dim i As Integer
        Dim tau, g0_tautau, gr_pi, gr_pitau, gr_pipi, gr_tautau As Double
        Const R As Double = 0.461526 'kJ/(kg K)
        Dim J0 = New Double() {0, 1, -5, -4, -3, -2, -1, 2, 3}
        Dim n0 = New Double() {-9.6927686500217, 10.086655968018, -0.005608791128302, 0.071452738081455, -0.40710498223928, 1.4240819171444, -4.383951131945, -0.28408632460772, 0.021268463753307}
        Dim Ir = New Double() {1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 4, 4, 4, 5, 6, 6, 6, 7, 7, 7, 8, 8, 9, 10, 10, 10, 16, 16, 18, 20, 20, 20, 21, 22, 23, 24, 24, 24}
        Dim Jr = New Double() {0, 1, 2, 3, 6, 1, 2, 4, 7, 36, 0, 1, 3, 6, 35, 1, 2, 3, 7, 3, 16, 35, 0, 11, 25, 8, 36, 13, 4, 10, 14, 29, 50, 57, 20, 35, 48, 21, 53, 39, 26, 40, 58}
        Dim nr = New Double() {-0.0017731742473213, -0.017834862292358, -0.045996013696365, -0.057581259083432, -0.05032527872793, -0.000033032641670203, -0.00018948987516315, -0.0039392777243355, -0.043797295650573, -0.000026674547914087, 0.000000020481737692309, 0.00000043870667284435, -0.00003227767723857, -0.0015033924542148, -0.040668253562649, -0.00000000078847309559367, 0.000000012790717852285, 0.00000048225372718507, 0.0000022922076337661, -0.000000000016714766451061, -0.0021171472321355, -23.895741934104, -5.905956432427E-18, -0.0000012621808899101, -0.038946842435739, 0.000000000011256211360459, -8.2311340897998, 0.000000019809712802088, 1.0406965210174E-19, -0.00000000000010234747095929, -0.0000000010018179379511, -0.000000000080882908646985, 0.10693031879409, -0.33662250574171, 8.9185845355421E-25, 0.00000000000030629316876231997, -0.0000042002467698208, -5.9056029685639E-26, 0.0000037826947613457, -0.0000000000000012768608934681, 7.3087610595061E-29, 5.5414715350778E-17, -0.0000009436970724121}
        tau = 540 / T
        g0_tautau = 0#
        For i = 0 To 8
            g0_tautau += n0(i) * J0(i) * (J0(i) - 1) * tau ^ (J0(i) - 2)
        Next i
        gr_pi = 0#
        gr_pitau = 0#
        gr_pipi = 0#
        gr_tautau = 0#
        For i = 0 To 42
            gr_pi += nr(i) * Ir(i) * p ^ (Ir(i) - 1) * (tau - 0.5) ^ Jr(i)
            gr_pipi += nr(i) * Ir(i) * (Ir(i) - 1) * p ^ (Ir(i) - 2) * (tau - 0.5) ^ Jr(i)
            gr_pitau += nr(i) * Ir(i) * p ^ (Ir(i) - 1) * Jr(i) * (tau - 0.5) ^ (Jr(i) - 1)
            gr_tautau += nr(i) * p ^ Ir(i) * Jr(i) * (Jr(i) - 1) * (tau - 0.5) ^ (Jr(i) - 2)
        Next i
        W2_pT = (1000 * R * T * (1 + 2 * p * gr_pi + p ^ 2 * gr_pi ^ 2) / ((1 - p ^ 2 * gr_pipi) + (1 + p * gr_pi - tau * p * gr_pitau) ^ 2 / (tau ^ 2 * (g0_tautau + gr_tautau)))) ^ 0.5
    End Function
    Private Function T2_ph(ByVal p As Double, ByVal h As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        '6 Equations for Region 2,6.3.1 The Backward Equations T( p, h ) for Subregions 2a, 2b, and 2c
        Dim sub_reg As Integer
        Dim i As Integer
        Dim Ts, hs As Double

        If p < 4 Then
            sub_reg = 1
        Else
            If p < (905.84278514723 - 0.67955786399241 * h + 0.00012809002730136 * h ^ 2) Then
                sub_reg = 2
            Else
                sub_reg = 3
            End If
        End If

        Select Case sub_reg
            Case 1
                'Subregion A
                'Table 20, Eq 22, page 22
                Dim Ji = New Double() {0, 1, 2, 3, 7, 20, 0, 1, 2, 3, 7, 9, 11, 18, 44, 0, 2, 7, 36, 38, 40, 42, 44, 24, 44, 12, 32, 44, 32, 36, 42, 34, 44, 28}
                Dim Ii = New Double() {0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 2, 2, 3, 3, 4, 4, 4, 5, 5, 5, 6, 6, 7}
                Dim ni = New Double() {1089.8952318288, 849.51654495535, -107.81748091826, 33.153654801263, -7.4232016790248, 11.765048724356, 1.844574935579, -4.1792700549624, 6.2478196935812, -17.344563108114, -200.58176862096, 271.96065473796, -455.11318285818, 3091.9688604755, 252266.40357872, -0.0061707422868339, -0.31078046629583, 11.670873077107, 128127984.04046, -985549096.23276, 2822454697.3002, -3594897141.0703, 1722734991.3197, -13551.334240775, 12848734.66465, 1.3865724283226, 235988.32556514, -13105236.545054, 7399.9835474766, -551966.9703006, 3715408.5996233, 19127.72923966, -415351.64835634, -62.459855192507}
                Ts = 0
                hs = h / 2000
                For i = 0 To 33
                    Ts += ni(i) * p ^ (Ii(i)) * (hs - 2.1) ^ Ji(i)
                Next i
                T2_ph = Ts
            Case 2
                'Subregion B
                'Table 21, Eq 23, page 23
                Dim Ji = New Double() {0, 1, 2, 12, 18, 24, 28, 40, 0, 2, 6, 12, 18, 24, 28, 40, 2, 8, 18, 40, 1, 2, 12, 24, 2, 12, 18, 24, 28, 40, 18, 24, 40, 28, 2, 28, 1, 40}
                Dim Ii = New Double() {0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 4, 4, 5, 5, 5, 6, 7, 7, 9, 9}
                Dim ni = New Double() {1489.5041079516, 743.07798314034, -97.708318797837, 2.4742464705674, -0.63281320016026, 1.1385952129658, -0.47811863648625, 0.0085208123431544, 0.93747147377932, 3.3593118604916, 3.3809355601454, 0.16844539671904, 0.73875745236695, -0.47128737436186, 0.15020273139707, -0.002176411421975, -0.021810755324761, -0.10829784403677, -0.046333324635812, 0.000071280351959551, 0.00011032831789999, 0.00018955248387902, 0.0030891541160537, 0.0013555504554949, 0.00000028640237477456, -0.000010779857357512, -0.000076462712454814, 0.000014052392818316, -0.000031083814331434, -0.0000010302738212103, 0.0000002821728163504, 0.0000012704902271945, 0.000000073803353468292, -0.000000011030139238909, -0.000000000000081456365207833, -0.000000000025180545682962, -1.7565233969407E-18, 0.0000000000000086934156344163}
                Ts = 0
                hs = h / 2000
                For i = 0 To 37
                    Ts += ni(i) * (p - 2) ^ (Ii(i)) * (hs - 2.6) ^ Ji(i)
                Next i
                T2_ph = Ts
            Case Else
                'Subregion C
                'Table 22, Eq 24, page 24
                Dim Ji = New Double() {0, 4, 0, 2, 0, 2, 0, 1, 0, 2, 0, 1, 4, 8, 4, 0, 1, 4, 10, 12, 16, 20, 22}
                Dim Ii = New Double() {-7, -7, -6, -6, -5, -5, -2, -2, -1, -1, 0, 0, 1, 1, 2, 6, 6, 6, 6, 6, 6, 6, 6}
                Dim ni = New Double() {-3236839855524.2, 7326335090218.1, 358250899454.47, -583401318515.9, -10783068217.47, 20825544563.171, 610747.83564516, 859777.2253558, -25745.72360417, 31081.088422714, 1208.2315865936, 482.19755109255, 3.7966001272486, -10.842984880077, -0.04536417267666, 0.00000000000014559115658698, 0.000000000001126159740723, -0.000000000017804982240686, 0.00000012324579690832, -0.0000011606921130984, 0.000027846367088554, -0.00059270038474176, 0.0012918582991878}
                Ts = 0
                hs = h / 2000
                For i = 0 To 22
                    Ts += ni(i) * (p + 25) ^ (Ii(i)) * (hs - 1.8) ^ Ji(i)
                Next i
                T2_ph = Ts
        End Select
    End Function
    Private Function T2_ps(ByVal p As Double, ByVal s As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        '6 Equations for Region 2,6.3.2 The Backward Equations T( p, s ) for Subregions 2a, 2b, and 2c
        'Page 26
        Dim sub_reg As Integer
        Dim i As Integer
        Dim teta, sigma As Double

        If p < 4 Then
            sub_reg = 1
        Else
            If s < 5.85 Then
                sub_reg = 3
            Else
                sub_reg = 2
            End If
        End If
        Select Case sub_reg
            Case 1
                'Subregion A
                'Table 25, Eq 25, page 26
                Dim Ii = New Double() {-1.5, -1.5, -1.5, -1.5, -1.5, -1.5, -1.25, -1.25, -1.25, -1, -1, -1, -1, -1, -1, -0.75, -0.75, -0.5, -0.5, -0.5, -0.5, -0.25, -0.25, -0.25, -0.25, 0.25, 0.25, 0.25, 0.25, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.75, 0.75, 0.75, 0.75, 1, 1, 1.25, 1.25, 1.5, 1.5}
                Dim Ji = New Double() {-24, -23, -19, -13, -11, -10, -19, -15, -6, -26, -21, -17, -16, -9, -8, -15, -14, -26, -13, -9, -7, -27, -25, -11, -6, 1, 4, 8, 11, 0, 1, 5, 6, 10, 14, 16, 0, 4, 9, 17, 7, 18, 3, 15, 5, 18}
                Dim ni = New Double() {-392359.83861984, 515265.7382727, 40482.443161048, -321.93790923902, 96.961424218694, -22.867846371773, -449429.14124357, -5011.8336020166, 0.35684463560015, 44235.33584819, -13673.388811708, 421632.60207864, 22516.925837475, 474.42144865646, -149.31130797647, -197811.26320452, -23554.39947076, -19070.616302076, 55375.669883164, 3829.3691437363, -603.91860580567, 1936.3102620331, 4266.064369861, -5978.0638872718, -704.01463926862, 338.36784107553, 20.862786635187, 0.033834172656196, -0.000043124428414893, 166.53791356412, -139.86292055898, -0.78849547999872, 0.072132411753872, -0.0059754839398283, -0.000012141358953904, 0.00000023227096733871, -10.538463566194, 2.0718925496502, -0.072193155260427, 0.0000002074988708112, -0.018340657911379, 0.00000029036272348696, 0.21037527893619, 0.00025681239729999, -0.012799002933781, -0.0000082198102652018}
                sigma = s / 2
                teta = 0
                For i = 0 To 45
                    teta += ni(i) * p ^ Ii(i) * (sigma - 2) ^ Ji(i)
                Next i
                T2_ps = teta
            Case 2
                'Subregion B
                'Table 26, Eq 26, page 27
                Dim Ii = New Double() {-6, -6, -5, -5, -4, -4, -4, -3, -3, -3, -3, -2, -2, -2, -2, -1, -1, -1, -1, -1, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 2, 2, 2, 3, 3, 3, 4, 4, 5, 5, 5}
                Dim Ji = New Double() {0, 11, 0, 11, 0, 1, 11, 0, 1, 11, 12, 0, 1, 6, 10, 0, 1, 5, 8, 9, 0, 1, 2, 4, 5, 6, 9, 0, 1, 2, 3, 7, 8, 0, 1, 5, 0, 1, 3, 0, 1, 0, 1, 2}
                Dim ni = New Double() {316876.65083497, 20.864175881858, -398593.99803599, -21.816058518877, 223697.85194242, -2784.1703445817, 9.920743607148, -75197.512299157, 2970.8605951158, -3.4406878548526, 0.38815564249115, 17511.29508575, -1423.7112854449, 1.0943803364167, 0.89971619308495, -3375.9740098958, 471.62885818355, -1.9188241993679, 0.41078580492196, -0.33465378172097, 1387.0034777505, -406.63326195838, 41.72734715961, 2.1932549434532, -1.0320050009077, 0.35882943516703, 0.0052511453726066, 12.838916450705, -2.8642437219381, 0.56912683664855, -0.099962954584931, -0.0032632037778459, 0.00023320922576723, -0.1533480985745, 0.029072288239902, 0.00037534702741167, 0.0017296691702411, -0.00038556050844504, -0.000035017712292608, -0.000014566393631492, 0.0000056420857267269, 0.000000041286150074605, -0.000000020684671118824, 0.0000000016409393674725}
                sigma = s / 0.7853
                teta = 0
                For i = 0 To 43
                    teta += ni(i) * p ^ Ii(i) * (10 - sigma) ^ Ji(i)
                Next i
                T2_ps = teta
            Case Else
                'Subregion C
                'Table 27, Eq 27, page 28
                Dim Ii = New Double() {-2, -2, -1, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 3, 3, 3, 4, 4, 4, 5, 5, 5, 6, 6, 7, 7, 7, 7, 7}
                Dim Ji = New Double() {0, 1, 0, 0, 1, 2, 3, 0, 1, 3, 4, 0, 1, 2, 0, 1, 5, 0, 1, 4, 0, 1, 2, 0, 1, 0, 1, 3, 4, 5}
                Dim ni = New Double() {909.68501005365, 2404.566708842, -591.6232638713, 541.45404128074, -270.98308411192, 979.76525097926, -469.66772959435, 14.399274604723, -19.104204230429, 5.3299167111971, -21.252975375934, -0.3114733441376, 0.60334840894623, -0.042764839702509, 0.0058185597255259, -0.014597008284753, 0.0056631175631027, -0.000076155864584577, 0.00022440342919332, -0.000012561095013413, 0.00000063323132660934, -0.0000020541989675375, 0.000000036405370390082, -0.0000000029759897789215, 0.000000010136618529763, 0.0000000000059925719692351, -0.000000000020677870105164, -0.000000000020874278181886, 0.00000000010162166825089, -0.00000000016429828281347}
                sigma = s / 2.9251
                teta = 0
                For i = 0 To 29
                    teta += ni(i) * p ^ Ii(i) * (2 - sigma) ^ Ji(i)
                Next i
                T2_ps = teta
        End Select
    End Function
    Private Function P2_hs(ByVal h As Double, ByVal s As Double) As Double
        'Supplementary Release on Backward Equations for Pressure as a function of Enthalpy and Entropy p(h,s) to the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam
        'Chapter 6:Backward Equations p(h,s) for Region 2
        Dim sub_reg As Integer
        Dim i As Integer
        Dim eta, sigma, p As Double

        If h < (-3498.98083432139 + 2575.60716905876 * s - 421.073558227969 * s ^ 2 + 27.6349063799944 * s ^ 3) Then
            sub_reg = 1
        Else
            If s < 5.85 Then
                sub_reg = 3
            Else
                sub_reg = 2
            End If
        End If
        Select Case sub_reg
            Case 1
                'Subregion A
                'Table 6, Eq 3, page 8
                Dim Ii = New Double() {0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 3, 3, 3, 3, 3, 4, 5, 5, 6, 7}
                Dim Ji = New Double() {1, 3, 6, 16, 20, 22, 0, 1, 2, 3, 5, 6, 10, 16, 20, 22, 3, 16, 20, 0, 2, 3, 6, 16, 16, 3, 16, 3, 1}
                Dim ni = New Double() {-0.0182575361923032, -0.125229548799536, 0.592290437320145, 6.04769706185122, 238.624965444474, -298.639090222922, 0.051225081304075, -0.437266515606486, 0.413336902999504, -5.16468254574773, -5.57014838445711, 12.8555037824478, 11.414410895329, -119.504225652714, -2847.7798596156, 4317.57846408006, 1.1289404080265, 1974.09186206319, 1516.12444706087, 0.0141324451421235, 0.585501282219601, -2.97258075863012, 5.94567314847319, -6236.56565798905, 9659.86235133332, 6.81500934948134, -6332.07286824489, -5.5891922446576, 0.0400645798472063}
                eta = h / 4200
                sigma = s / 12
                p = 0
                For i = 0 To 28
                    p += ni(i) * (eta - 0.5) ^ Ii(i) * (sigma - 1.2) ^ Ji(i)
                Next i
                P2_hs = p ^ 4 * 4
            Case 2
                'Subregion B
                'Table 7, Eq 4, page 9
                Dim Ii = New Double() {0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 2, 2, 2, 3, 3, 3, 3, 4, 4, 5, 5, 6, 6, 6, 7, 7, 8, 8, 8, 8, 12, 14}
                Dim Ji = New Double() {0, 1, 2, 4, 8, 0, 1, 2, 3, 5, 12, 1, 6, 18, 0, 1, 7, 12, 1, 16, 1, 12, 1, 8, 18, 1, 16, 1, 3, 14, 18, 10, 16}
                Dim ni = New Double() {0.0801496989929495, -0.543862807146111, 0.337455597421283, 8.9055545115745, 313.840736431485, 0.797367065977789, -1.2161697355624, 8.72803386937477, -16.9769781757602, -186.552827328416, 95115.9274344237, -18.9168510120494, -4334.0703719484, 543212633.012715, 0.144793408386013, 128.024559637516, -67230.9534071268, 33697238.0095287, -586.63419676272, -22140322476.9889, 1716.06668708389, -570817595.806302, -3121.09693178482, -2078413.8463301, 3056059461577.86, 3221.57004314333, 326810259797.295, -1441.04158934487, 410.694867802691, 109077066873.024, -24796465425889.3, 1888019068.65134, -123651009018773.0#}
                eta = h / 4100
                sigma = s / 7.9
                For i = 0 To 32
                    p += ni(i) * (eta - 0.6) ^ Ii(i) * (sigma - 1.01) ^ Ji(i)
                Next i
                P2_hs = p ^ 4 * 100
            Case Else
                'Subregion C
                'Table 8, Eq 5, page 10
                Dim Ii = New Double() {0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 4, 5, 5, 5, 5, 6, 6, 10, 12, 16}
                Dim Ji = New Double() {0, 1, 2, 3, 4, 8, 0, 2, 5, 8, 14, 2, 3, 7, 10, 18, 0, 5, 8, 16, 18, 18, 1, 4, 6, 14, 8, 18, 7, 7, 10}
                Dim ni = New Double() {0.112225607199012, -3.39005953606712, -32.0503911730094, -197.5973051049, -407.693861553446, 13294.3775222331, 1.70846839774007, 37.3694198142245, 3581.44365815434, 423014.446424664, -751071025.760063, 52.3446127607898, -228.351290812417, -960652.417056937, -80705929.2526074, 1626980172256.69, 0.772465073604171, 46392.9973837746, -13731788.5134128, 1704703926305.12, -25110462818730.8, 31774883083552.0#, 53.8685623675312, -55308.9094625169, -1028615.22421405, 2042494187562.34, 273918446.626977, -2.63963146312685E+15, -1078908541.08088, -29649262098.0124, -1.11754907323424E+15}
                eta = h / 3500
                sigma = s / 5.9
                For i = 0 To 30
                    p += ni(i) * (eta - 0.7) ^ Ii(i) * (sigma - 1.1) ^ Ji(i)
                Next i
                P2_hs = p ^ 4 * 100
        End Select
    End Function
    Private Function T2_prho(ByVal p As Double, ByVal rho As Double) As Double
        'Solve by iteration. Observe that fo low temperatures this equation has 2 solutions.
        'Solve with half interval method
        Dim Low_Bound, High_Bound, rhos, Ts As Double

        If p < 16.5292 Then
            Low_Bound = T4_p(p)
        Else
            Low_Bound = B23T_p(p)
        End If
        High_Bound = 1073.15
        Do While Math.Abs(rho - rhos) > 0.000001
            Ts = (Low_Bound + High_Bound) / 2
            rhos = 1 / V2_pT(p, Ts)
            If rhos < rho Then
                High_Bound = Ts
            Else
                Low_Bound = Ts
            End If
        Loop
        T2_prho = Ts
    End Function
    '***********************************************************************************************************
    '*2.3 Functions for region 3

    Private Function P3_rhoT(ByVal rho As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        '7 Basic Equation for Region 3, Section. 6.1 Basic Equation
        'Table 30 and 31, Page 30 and 31
        Dim i As Integer, delta, tau, fidelta As Double
        Const R As Double = 0.461526, tc As Double = 647.096, rhoc As Double = 322
        Dim Ii = New Double() {0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 6, 6, 6, 7, 8, 9, 9, 10, 10, 11}
        Dim Ji = New Double() {0, 0, 1, 2, 7, 10, 12, 23, 2, 6, 15, 17, 0, 2, 6, 7, 22, 26, 0, 2, 4, 16, 26, 0, 2, 4, 26, 1, 3, 26, 0, 2, 26, 2, 26, 2, 26, 0, 1, 26}
        Dim ni = New Double() {1.0658070028513, -15.732845290239, 20.944396974307, -7.6867707878716, 2.6185947787954, -2.808078114862, 1.2053369696517, -0.0084566812812502, -1.2654315477714, -1.1524407806681, 0.88521043984318, -0.64207765181607, 0.38493460186671, -0.85214708824206, 4.8972281541877, -3.0502617256965, 0.039420536879154, 0.12558408424308, -0.2799932969871, 1.389979956946, -2.018991502357, -0.0082147637173963, -0.47596035734923, 0.0439840744735, -0.44476435428739, 0.90572070719733, 0.70522450087967, 0.10770512626332, -0.32913623258954, -0.50871062041158, -0.022175400873096, 0.094260751665092, 0.16436278447961, -0.013503372241348, -0.014834345352472, 0.00057922953628084, 0.0032308904703711, 0.000080964802996215, -0.00016557679795037, -0.000044923899061815}
        delta = rho / rhoc
        tau = tc / T
        fidelta = 0
        For i = 1 To 39
            fidelta += ni(i) * Ii(i) * delta ^ (Ii(i) - 1) * tau ^ Ji(i)
        Next i
        fidelta += ni(0) / delta
        P3_rhoT = rho * R * T * delta * fidelta / 1000
    End Function
    Private Function U3_rhoT(ByVal rho As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        '7 Basic Equation for Region 3, Section. 6.1 Basic Equation
        'Table 30 and 31, Page 30 and 31
        Dim i As Integer, delta, tau, fitau As Double
        Const R As Double = 0.461526, tc As Double = 647.096, rhoc As Double = 322
        Dim Ii = New Double() {0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 6, 6, 6, 7, 8, 9, 9, 10, 10, 11}
        Dim Ji = New Double() {0, 0, 1, 2, 7, 10, 12, 23, 2, 6, 15, 17, 0, 2, 6, 7, 22, 26, 0, 2, 4, 16, 26, 0, 2, 4, 26, 1, 3, 26, 0, 2, 26, 2, 26, 2, 26, 0, 1, 26}
        Dim ni = New Double() {1.0658070028513, -15.732845290239, 20.944396974307, -7.6867707878716, 2.6185947787954, -2.808078114862, 1.2053369696517, -0.0084566812812502, -1.2654315477714, -1.1524407806681, 0.88521043984318, -0.64207765181607, 0.38493460186671, -0.85214708824206, 4.8972281541877, -3.0502617256965, 0.039420536879154, 0.12558408424308, -0.2799932969871, 1.389979956946, -2.018991502357, -0.0082147637173963, -0.47596035734923, 0.0439840744735, -0.44476435428739, 0.90572070719733, 0.70522450087967, 0.10770512626332, -0.32913623258954, -0.50871062041158, -0.022175400873096, 0.094260751665092, 0.16436278447961, -0.013503372241348, -0.014834345352472, 0.00057922953628084, 0.0032308904703711, 0.000080964802996215, -0.00016557679795037, -0.000044923899061815}
        delta = rho / rhoc
        tau = tc / T
        fitau = 0
        For i = 1 To 39
            fitau += ni(i) * delta ^ Ii(i) * Ji(i) * tau ^ (Ji(i) - 1)
        Next i
        U3_rhoT = R * T * (tau * fitau)
    End Function

    Private Function H3_rhoT(ByVal rho As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        '7 Basic Equation for Region 3, Section. 6.1 Basic Equation
        'Table 30 and 31, Page 30 and 31
        Dim i As Integer, delta, tau, fidelta, fitau As Double
        Const R As Double = 0.461526, tc As Double = 647.096, rhoc As Double = 322
        Dim Ii = New Double() {0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 6, 6, 6, 7, 8, 9, 9, 10, 10, 11}
        Dim Ji = New Double() {0, 0, 1, 2, 7, 10, 12, 23, 2, 6, 15, 17, 0, 2, 6, 7, 22, 26, 0, 2, 4, 16, 26, 0, 2, 4, 26, 1, 3, 26, 0, 2, 26, 2, 26, 2, 26, 0, 1, 26}
        Dim ni = New Double() {1.0658070028513, -15.732845290239, 20.944396974307, -7.6867707878716, 2.6185947787954, -2.808078114862, 1.2053369696517, -0.0084566812812502, -1.2654315477714, -1.1524407806681, 0.88521043984318, -0.64207765181607, 0.38493460186671, -0.85214708824206, 4.8972281541877, -3.0502617256965, 0.039420536879154, 0.12558408424308, -0.2799932969871, 1.389979956946, -2.018991502357, -0.0082147637173963, -0.47596035734923, 0.0439840744735, -0.44476435428739, 0.90572070719733, 0.70522450087967, 0.10770512626332, -0.32913623258954, -0.50871062041158, -0.022175400873096, 0.094260751665092, 0.16436278447961, -0.013503372241348, -0.014834345352472, 0.00057922953628084, 0.0032308904703711, 0.000080964802996215, -0.00016557679795037, -0.000044923899061815}
        delta = rho / rhoc
        tau = tc / T
        fidelta = 0
        fitau = 0
        For i = 1 To 39
            fidelta += ni(i) * Ii(i) * delta ^ (Ii(i) - 1) * tau ^ Ji(i)
            fitau += ni(i) * delta ^ Ii(i) * Ji(i) * tau ^ (Ji(i) - 1)
        Next i
        fidelta += ni(0) / delta
        H3_rhoT = R * T * (tau * fitau + delta * fidelta)
    End Function
    Private Function S3_rhoT(ByVal rho As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        '7 Basic Equation for Region 3, Section. 6.1 Basic Equation
        'Table 30 and 31, Page 30 and 31
        Dim i As Integer, fi, delta, tau, fitau As Double
        Const R As Double = 0.461526, tc As Double = 647.096, rhoc As Double = 322
        Dim Ii = New Double() {0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 6, 6, 6, 7, 8, 9, 9, 10, 10, 11}
        Dim Ji = New Double() {0, 0, 1, 2, 7, 10, 12, 23, 2, 6, 15, 17, 0, 2, 6, 7, 22, 26, 0, 2, 4, 16, 26, 0, 2, 4, 26, 1, 3, 26, 0, 2, 26, 2, 26, 2, 26, 0, 1, 26}
        Dim ni = New Double() {1.0658070028513, -15.732845290239, 20.944396974307, -7.6867707878716, 2.6185947787954, -2.808078114862, 1.2053369696517, -0.0084566812812502, -1.2654315477714, -1.1524407806681, 0.88521043984318, -0.64207765181607, 0.38493460186671, -0.85214708824206, 4.8972281541877, -3.0502617256965, 0.039420536879154, 0.12558408424308, -0.2799932969871, 1.389979956946, -2.018991502357, -0.0082147637173963, -0.47596035734923, 0.0439840744735, -0.44476435428739, 0.90572070719733, 0.70522450087967, 0.10770512626332, -0.32913623258954, -0.50871062041158, -0.022175400873096, 0.094260751665092, 0.16436278447961, -0.013503372241348, -0.014834345352472, 0.00057922953628084, 0.0032308904703711, 0.000080964802996215, -0.00016557679795037, -0.000044923899061815}
        delta = rho / rhoc
        tau = tc / T
        fi = 0
        fitau = 0
        For i = 1 To 39
            fi += ni(i) * delta ^ Ii(i) * tau ^ Ji(i)
            fitau += ni(i) * delta ^ Ii(i) * Ji(i) * tau ^ (Ji(i) - 1)
        Next i
        fi += ni(0) * Math.Log(delta)
        S3_rhoT = R * (tau * fitau - fi)
    End Function
    Private Function Cp3_rhoT(ByVal rho As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        '7 Basic Equation for Region 3, Section. 6.1 Basic Equation
        'Table 30 and 31, Page 30 and 31
        Dim i As Integer, fideltatau, delta, tau, fitautau, fidelta, fideltadelta As Double
        Const R As Double = 0.461526, tc As Double = 647.096, rhoc As Double = 322
        Dim Ii = New Double() {0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 6, 6, 6, 7, 8, 9, 9, 10, 10, 11}
        Dim Ji = New Double() {0, 0, 1, 2, 7, 10, 12, 23, 2, 6, 15, 17, 0, 2, 6, 7, 22, 26, 0, 2, 4, 16, 26, 0, 2, 4, 26, 1, 3, 26, 0, 2, 26, 2, 26, 2, 26, 0, 1, 26}
        Dim ni = New Double() {1.0658070028513, -15.732845290239, 20.944396974307, -7.6867707878716, 2.6185947787954, -2.808078114862, 1.2053369696517, -0.0084566812812502, -1.2654315477714, -1.1524407806681, 0.88521043984318, -0.64207765181607, 0.38493460186671, -0.85214708824206, 4.8972281541877, -3.0502617256965, 0.039420536879154, 0.12558408424308, -0.2799932969871, 1.389979956946, -2.018991502357, -0.0082147637173963, -0.47596035734923, 0.0439840744735, -0.44476435428739, 0.90572070719733, 0.70522450087967, 0.10770512626332, -0.32913623258954, -0.50871062041158, -0.022175400873096, 0.094260751665092, 0.16436278447961, -0.013503372241348, -0.014834345352472, 0.00057922953628084, 0.0032308904703711, 0.000080964802996215, -0.00016557679795037, -0.000044923899061815}
        delta = rho / rhoc
        tau = tc / T
        fitautau = 0
        fidelta = 0
        fideltatau = 0
        fideltadelta = 0
        For i = 1 To 39
            fitautau += ni(i) * delta ^ Ii(i) * Ji(i) * (Ji(i) - 1) * tau ^ (Ji(i) - 2)
            fidelta += ni(i) * Ii(i) * delta ^ (Ii(i) - 1) * tau ^ Ji(i)
            fideltatau += ni(i) * Ii(i) * delta ^ (Ii(i) - 1) * Ji(i) * tau ^ (Ji(i) - 1)
            fideltadelta += ni(i) * Ii(i) * (Ii(i) - 1) * delta ^ (Ii(i) - 2) * tau ^ Ji(i)
        Next i
        fidelta += ni(0) / delta
        fideltadelta -= ni(0) / (delta ^ 2)
        Cp3_rhoT = R * (-(tau ^ 2 * fitautau) + (delta * fidelta - delta * tau * fideltatau) ^ 2 / (2 * delta * fidelta + delta ^ 2 * fideltadelta))
    End Function
    Private Function Cv3_rhoT(ByVal rho As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        '7 Basic Equation for Region 3, Section. 6.1 Basic Equation
        'Table 30 and 31, Page 30 and 31
        Dim i As Integer, delta, tau, fitautau As Double
        Const R As Double = 0.461526, tc As Double = 647.096, rhoc As Double = 322
        Dim Ii = New Double() {0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 6, 6, 6, 7, 8, 9, 9, 10, 10, 11}
        Dim Ji = New Double() {0, 0, 1, 2, 7, 10, 12, 23, 2, 6, 15, 17, 0, 2, 6, 7, 22, 26, 0, 2, 4, 16, 26, 0, 2, 4, 26, 1, 3, 26, 0, 2, 26, 2, 26, 2, 26, 0, 1, 26}
        Dim ni = New Double() {1.0658070028513, -15.732845290239, 20.944396974307, -7.6867707878716, 2.6185947787954, -2.808078114862, 1.2053369696517, -0.0084566812812502, -1.2654315477714, -1.1524407806681, 0.88521043984318, -0.64207765181607, 0.38493460186671, -0.85214708824206, 4.8972281541877, -3.0502617256965, 0.039420536879154, 0.12558408424308, -0.2799932969871, 1.389979956946, -2.018991502357, -0.0082147637173963, -0.47596035734923, 0.0439840744735, -0.44476435428739, 0.90572070719733, 0.70522450087967, 0.10770512626332, -0.32913623258954, -0.50871062041158, -0.022175400873096, 0.094260751665092, 0.16436278447961, -0.013503372241348, -0.014834345352472, 0.00057922953628084, 0.0032308904703711, 0.000080964802996215, -0.00016557679795037, -0.000044923899061815}
        delta = rho / rhoc
        tau = tc / T
        fitautau = 0
        For i = 1 To 39
            fitautau += ni(i) * delta ^ Ii(i) * Ji(i) * (Ji(i) - 1) * tau ^ (Ji(i) - 2)
        Next i
        Cv3_rhoT = R * -(tau * tau * fitautau)
    End Function
    Private Function W3_rhoT(ByVal rho As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        '7 Basic Equation for Region 3, Section. 6.1 Basic Equation
        'Table 30 and 31, Page 30 and 31
        Dim i As Integer, delta, tau, fitautau, fidelta, fideltatau, fideltadelta As Double
        Const R As Double = 0.461526, tc As Double = 647.096, rhoc As Double = 322
        Dim Ii = New Double() {0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 6, 6, 6, 7, 8, 9, 9, 10, 10, 11}
        Dim Ji = New Double() {0, 0, 1, 2, 7, 10, 12, 23, 2, 6, 15, 17, 0, 2, 6, 7, 22, 26, 0, 2, 4, 16, 26, 0, 2, 4, 26, 1, 3, 26, 0, 2, 26, 2, 26, 2, 26, 0, 1, 26}
        Dim ni = New Double() {1.0658070028513, -15.732845290239, 20.944396974307, -7.6867707878716, 2.6185947787954, -2.808078114862, 1.2053369696517, -0.0084566812812502, -1.2654315477714, -1.1524407806681, 0.88521043984318, -0.64207765181607, 0.38493460186671, -0.85214708824206, 4.8972281541877, -3.0502617256965, 0.039420536879154, 0.12558408424308, -0.2799932969871, 1.389979956946, -2.018991502357, -0.0082147637173963, -0.47596035734923, 0.0439840744735, -0.44476435428739, 0.90572070719733, 0.70522450087967, 0.10770512626332, -0.32913623258954, -0.50871062041158, -0.022175400873096, 0.094260751665092, 0.16436278447961, -0.013503372241348, -0.014834345352472, 0.00057922953628084, 0.0032308904703711, 0.000080964802996215, -0.00016557679795037, -0.000044923899061815}
        delta = rho / rhoc
        tau = tc / T
        fitautau = 0
        fidelta = 0
        fideltatau = 0
        fideltadelta = 0
        For i = 1 To 39
            fitautau += ni(i) * delta ^ Ii(i) * Ji(i) * (Ji(i) - 1) * tau ^ (Ji(i) - 2)
            fidelta += ni(i) * Ii(i) * delta ^ (Ii(i) - 1) * tau ^ Ji(i)
            fideltatau += ni(i) * Ii(i) * delta ^ (Ii(i) - 1) * Ji(i) * tau ^ (Ji(i) - 1)
            fideltadelta += ni(i) * Ii(i) * (Ii(i) - 1) * delta ^ (Ii(i) - 2) * tau ^ Ji(i)
        Next i
        fidelta += ni(0) / delta
        fideltadelta -= ni(0) / (delta ^ 2)
        W3_rhoT = (1000 * R * T * (2 * delta * fidelta + delta ^ 2 * fideltadelta - (delta * fidelta - delta * tau * fideltatau) ^ 2 / (tau ^ 2 * fitautau))) ^ 0.5
    End Function
    Private Function T3_ph(ByVal p As Double, ByVal h As Double) As Double
        'Revised Supplementary Release on Backward Equations for the Functions T(p,h), v(p,h) and T(p,s), v(p,s) for Region 3 of the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam
        '2004
        'Section 3.3 Backward Equations T(p,h) and v(p,h) for Subregions 3a and 3b
        'Boundary equation, Eq 1 Page 5
        Dim i As Integer, h3ab, ps, hs, Ts As Double
        'Const R As Double = 0.461526, tc As Double = 647.096, pc As Double = 22.064, rhoc As Double = 322
        h3ab = (2014.64004206875 + 3.74696550136983 * p - 0.0219921901054187 * p ^ 2 + 0.000087513168600995 * p ^ 3)
        If h < h3ab Then
            'Subregion 3a
            'Eq 2, Table 3, Page 7
            Dim Ii = New Double() {-12, -12, -12, -12, -12, -12, -12, -12, -10, -10, -10, -8, -8, -8, -8, -5, -3, -2, -2, -2, -1, -1, 0, 0, 1, 3, 3, 4, 4, 10, 12}
            Dim Ji = New Double() {0, 1, 2, 6, 14, 16, 20, 22, 1, 5, 12, 0, 2, 4, 10, 2, 0, 1, 3, 4, 0, 2, 0, 1, 1, 0, 1, 0, 3, 4, 5}
            Dim ni = New Double() {-0.000000133645667811215, 0.00000455912656802978, -0.0000146294640700979, 0.0063934131297008, 372.783927268847, -7186.54377460447, 573494.7521034, -2675693.29111439, -0.0000334066283302614, -0.0245479214069597, 47.8087847764996, 0.00000764664131818904, 0.00128350627676972, 0.0171219081377331, -8.51007304583213, -0.0136513461629781, -0.00000384460997596657, 0.00337423807911655, -0.551624873066791, 0.72920227710747, -0.00992522757376041, -0.119308831407288, 0.793929190615421, 0.454270731799386, 0.20999859125991, -0.00642109823904738, -0.023515586860454, 0.00252233108341612, -0.00764885133368119, 0.0136176427574291, -0.0133027883575669}
            ps = p / 100
            hs = h / 2300
            Ts = 0
            For i = 0 To 30
                Ts += ni(i) * (ps + 0.24) ^ Ii(i) * (hs - 0.615) ^ Ji(i)
            Next i
            T3_ph = Ts * 760
        Else
            'Subregion 3b
            'Eq 3, Table 4, Page 7,8
            Dim Ii = New Double() {-12, -12, -10, -10, -10, -10, -10, -8, -8, -8, -8, -8, -6, -6, -6, -4, -4, -3, -2, -2, -1, -1, -1, -1, -1, -1, 0, 0, 1, 3, 5, 6, 8}
            Dim Ji = New Double() {0, 1, 0, 1, 5, 10, 12, 0, 1, 2, 4, 10, 0, 1, 2, 0, 1, 5, 0, 4, 2, 4, 6, 10, 14, 16, 0, 2, 1, 1, 1, 1, 1}
            Dim ni = New Double() {0.000032325457364492, -0.000127575556587181, -0.000475851877356068, 0.00156183014181602, 0.105724860113781, -85.8514221132534, 724.140095480911, 0.00296475810273257, -0.00592721983365988, -0.0126305422818666, -0.115716196364853, 84.9000969739595, -0.0108602260086615, 0.0154304475328851, 0.0750455441524466, 0.0252520973612982, -0.0602507901232996, -3.07622221350501, -0.0574011959864879, 5.03471360939849, -0.925081888584834, 3.91733882917546, -77.314600713019, 9493.08762098587, -1410437.19679409, 8491662.30819026, 0.861095729446704, 0.32334644281172, 0.873281936020439, -0.436653048526683, 0.286596714529479, -0.131778331276228, 0.00676682064330275}
            hs = h / 2800
            ps = p / 100
            Ts = 0
            For i = 0 To 32
                Ts += ni(i) * (ps + 0.298) ^ Ii(i) * (hs - 0.72) ^ Ji(i)
            Next i
            T3_ph = Ts * 860
        End If
    End Function
    Private Function V3_ph(ByVal p As Double, ByVal h As Double) As Double
        'Revised Supplementary Release on Backward Equations for the Functions T(p,h), v(p,h) and T(p,s), v(p,s) for Region 3 of the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam
        '2004
        'Section 3.3 Backward Equations T(p,h) and v(p,h) for Subregions 3a and 3b
        'Boundary equation, Eq 1 Page 5
        Dim i As Integer, h3ab, ps, hs, vs As Double
        'Const R As Double = 0.461526, tc As Double = 647.096, pc As Double = 22.064, rhoc As Double = 322
        h3ab = (2014.64004206875 + 3.74696550136983 * p - 0.0219921901054187 * p ^ 2 + 0.000087513168600995 * p ^ 3)
        If h < h3ab Then
            'Subregion 3a
            'Eq 4, Table 6, Page 9
            Dim Ii = New Double() {-12, -12, -12, -12, -10, -10, -10, -8, -8, -6, -6, -6, -4, -4, -3, -2, -2, -1, -1, -1, -1, 0, 0, 1, 1, 1, 2, 2, 3, 4, 5, 8}
            Dim Ji = New Double() {6, 8, 12, 18, 4, 7, 10, 5, 12, 3, 4, 22, 2, 3, 7, 3, 16, 0, 1, 2, 3, 0, 1, 0, 1, 2, 0, 2, 0, 2, 2, 2}
            Dim ni = New Double() {0.00529944062966028, -0.170099690234461, 11.1323814312927, -2178.98123145125, -0.000506061827980875, 0.556495239685324, -9.43672726094016, -0.297856807561527, 93.9353943717186, 0.0192944939465981, 0.421740664704763, -3689141.2628233, -0.00737566847600639, -0.354753242424366, -1.99768169338727, 1.15456297059049, 5683.6687581596, 0.00808169540124668, 0.172416341519307, 1.04270175292927, -0.297691372792847, 0.560394465163593, 0.275234661176914, -0.148347894866012, -0.0651142513478515, -2.92468715386302, 0.0664876096952665, 3.52335014263844, -0.0146340792313332, -2.24503486668184, 1.10533464706142, -0.0408757344495612}
            ps = p / 100
            hs = h / 2100
            vs = 0
            For i = 0 To 31
                vs += ni(i) * (ps + 0.128) ^ Ii(i) * (hs - 0.727) ^ Ji(i)
            Next i
            V3_ph = vs * 0.0028
        Else
            'Subregion 3b
            'Eq 5, Table 7, Page 9
            Dim Ii = New Double() {-12, -12, -8, -8, -8, -8, -8, -8, -6, -6, -6, -6, -6, -6, -4, -4, -4, -3, -3, -2, -2, -1, -1, -1, -1, 0, 1, 1, 2, 2}
            Dim Ji = New Double() {0, 1, 0, 1, 3, 6, 7, 8, 0, 1, 2, 5, 6, 10, 3, 6, 10, 0, 2, 1, 2, 0, 1, 4, 5, 0, 0, 1, 2, 6}
            Dim ni = New Double() {-0.00000000225196934336318, 0.0000000140674363313486, 0.0000023378408528056, -0.0000331833715229001, 0.00107956778514318, -0.271382067378863, 1.07202262490333, -0.853821329075382, -0.0000215214194340526, 0.00076965608822273, -0.00431136580433864, 0.453342167309331, -0.507749535873652, -100.475154528389, -0.219201924648793, -3.21087965668917, 607.567815637771, 0.000557686450685932, 0.18749904002955, 0.00905368030448107, 0.285417173048685, 0.0329924030996098, 0.239897419685483, 4.82754995951394, -11.8035753702231, 0.169490044091791, -0.0179967222507787, 0.0371810116332674, -0.0536288335065096, 1.6069710109252}
            ps = p / 100
            hs = h / 2800
            vs = 0
            For i = 0 To 29
                vs += ni(i) * (ps + 0.0661) ^ Ii(i) * (hs - 0.72) ^ Ji(i)
            Next i
            V3_ph = vs * 0.0088
        End If
    End Function
    Private Function T3_ps(ByVal p As Double, ByVal s As Double) As Double
        'Revised Supplementary Release on Backward Equations for the Functions T(p,h), v(p,h) and T(p,s), v(p,s) for Region 3 of the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam
        '2004
        '3.4 Backward Equations T(p,s) and v(p,s) for Subregions 3a and 3b
        'Boundary equation, Eq 6 Page 11
        Dim i As Integer, ps, sigma, teta As Double
        'Const R As Double = 0.461526, tc As Double = 647.096, pc As Double = 22.064, rhoc As Double = 322

        If s <= 4.41202148223476 Then
            'Subregion 3a
            'Eq 6, Table 10, Page 11
            Dim Ii = New Double() {-12, -12, -10, -10, -10, -10, -8, -8, -8, -8, -6, -6, -6, -5, -5, -5, -4, -4, -4, -2, -2, -1, -1, 0, 0, 0, 1, 2, 2, 3, 8, 8, 10}
            Dim Ji = New Double() {28, 32, 4, 10, 12, 14, 5, 7, 8, 28, 2, 6, 32, 0, 14, 32, 6, 10, 36, 1, 4, 1, 6, 0, 1, 4, 0, 0, 3, 2, 0, 1, 2}
            Dim ni = New Double() {1500420082.63875, -159397258480.424, 0.000502181140217975, -67.2057767855466, 1450.58545404456, -8238.8953488889, -0.154852214233853, 11.2305046746695, -29.7000213482822, 43856513263.5495, 0.00137837838635464, -2.97478527157462, 9717779473494.13, -0.0000571527767052398, 28830.794977842, -74442828926270.3, 12.8017324848921, -368.275545889071, 6.64768904779177E+15, 0.044935925195888, -4.22897836099655, -0.240614376434179, -4.74341365254924, 0.72409399912611, 0.923874349695897, 3.99043655281015, 0.0384066651868009, -0.00359344365571848, -0.735196448821653, 0.188367048396131, 0.000141064266818704, -0.00257418501496337, 0.00123220024851555}
            sigma = s / 4.4
            ps = p / 100
            teta = 0
            For i = 0 To 32
                teta += ni(i) * (ps + 0.24) ^ Ii(i) * (sigma - 0.703) ^ Ji(i)
            Next i
            T3_ps = teta * 760
        Else
            'Subregion 3b
            'Eq 7, Table 11, Page 11
            Dim Ii = New Double() {-12, -12, -12, -12, -8, -8, -8, -6, -6, -6, -5, -5, -5, -5, -5, -4, -3, -3, -2, 0, 2, 3, 4, 5, 6, 8, 12, 14}
            Dim Ji = New Double() {1, 3, 4, 7, 0, 1, 3, 0, 2, 4, 0, 1, 2, 4, 6, 12, 1, 6, 2, 0, 1, 1, 0, 24, 0, 3, 1, 2}
            Dim ni = New Double() {0.52711170160166, -40.1317830052742, 153.020073134484, -2247.99398218827, -0.193993484669048, -1.40467557893768, 42.6799878114024, 0.752810643416743, 22.6657238616417, -622.873556909932, -0.660823667935396, 0.841267087271658, -25.3717501764397, 485.708963532948, 880.531517490555, 2650155.92794626, -0.359287150025783, -656.991567673753, 2.41768149185367, 0.856873461222588, 0.655143675313458, -0.213535213206406, 0.00562974957606348, -316955725450471.0#, -0.000699997000152457, 0.0119845803210767, 0.0000193848122022095, -0.0000215095749182309}
            sigma = s / 5.3
            ps = p / 100
            teta = 0
            For i = 0 To 27
                teta += ni(i) * (ps + 0.76) ^ Ii(i) * (sigma - 0.818) ^ Ji(i)
            Next i
            T3_ps = teta * 860
        End If
    End Function
    Private Function V3_ps(ByVal p As Double, ByVal s As Double) As Double
        'Revised Supplementary Release on Backward Equations for the Functions T(p,h), v(p,h) and T(p,s), v(p,s) for Region 3 of the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam
        '2004
        '3.4 Backward Equations T(p,s) and v(p,s) for Subregions 3a and 3b
        'Boundary equation, Eq 6 Page 11
        Dim i As Integer, ps, sigma, omega As Double
        'Const R As Double = 0.461526, tc As Double = 647.096, pc As Double = 22.064, rhoc As Double = 322
        If s <= 4.41202148223476 Then
            'Subregion 3a
            'Eq 8, Table 13, Page 14
            Dim Ii = New Double() {-12, -12, -12, -10, -10, -10, -10, -8, -8, -8, -8, -6, -5, -4, -3, -3, -2, -2, -1, -1, 0, 0, 0, 1, 2, 4, 5, 6}
            Dim Ji = New Double() {10, 12, 14, 4, 8, 10, 20, 5, 6, 14, 16, 28, 1, 5, 2, 4, 3, 8, 1, 2, 0, 1, 3, 0, 0, 2, 2, 0}
            Dim ni = New Double() {79.5544074093975, -2382.6124298459, 17681.3100617787, -0.00110524727080379, -15.3213833655326, 297.544599376982, -35031520.6871242, 0.277513761062119, -0.523964271036888, -148011.182995403, 1600148.99374266, 1708023226634.27, 0.000246866996006494, 1.6532608479798, -0.118008384666987, 2.537986423559, 0.965127704669424, -28.2172420532826, 0.203224612353823, 1.10648186063513, 0.52612794845128, 0.277000018736321, 1.08153340501132, -0.0744127885357893, 0.0164094443541384, -0.0680468275301065, 0.025798857610164, -0.000145749861944416}
            ps = p / 100
            sigma = s / 4.4
            omega = 0
            For i = 0 To 27
                omega += ni(i) * (ps + 0.187) ^ Ii(i) * (sigma - 0.755) ^ Ji(i)
            Next i
            V3_ps = omega * 0.0028
        Else
            'Subregion 3b
            'Eq 9, Table 14, Page 14
            Dim Ii = New Double() {-12, -12, -12, -12, -12, -12, -10, -10, -10, -10, -8, -5, -5, -5, -4, -4, -4, -4, -3, -2, -2, -2, -2, -2, -2, 0, 0, 0, 1, 1, 2}
            Dim Ji = New Double() {0, 1, 2, 3, 5, 6, 0, 1, 2, 4, 0, 1, 2, 3, 0, 1, 2, 3, 1, 0, 1, 2, 3, 4, 12, 0, 1, 2, 0, 2, 2}
            Dim ni = New Double() {0.0000591599780322238, -0.00185465997137856, 0.0104190510480013, 0.0059864730203859, -0.771391189901699, 1.72549765557036, -0.000467076079846526, 0.0134533823384439, -0.0808094336805495, 0.508139374365767, 0.00128584643361683, -1.63899353915435, 5.86938199318063, -2.92466667918613, -0.00614076301499537, 5.76199014049172, -12.1613320606788, 1.67637540957944, -7.44135838773463, 0.0378168091437659, 4.01432203027688, 16.0279837479185, 3.17848779347728, -3.58362310304853, -1159952.60446827, 0.199256573577909, -0.122270624794624, -19.1449143716586, -0.0150448002905284, 14.6407900162154, -3.2747778718823}
            ps = p / 100
            sigma = s / 5.3
            omega = 0
            For i = 0 To 30
                omega += ni(i) * (ps + 0.298) ^ Ii(i) * (sigma - 0.816) ^ Ji(i)
            Next i
            V3_ps = omega * 0.0088
        End If
    End Function

    Private Function P3_hs(ByVal h As Double, ByVal s As Double) As Double
        'Supplementary Release on Backward Equations ( ) , p h s for Region 3,
        'Equations as a function of h and s for the Region Boundaries, and an Equation
        '( ) sat , T hs for Region 4 of the IAPWS Industrial Formulation 1997 for the
        'Thermodynamic Properties of Water and Steam
        '2004
        'Section 3 Backward Functions p(h,s), T(h,s), and v(h,s) for Region 3
        Dim i As Integer, ps, sigma, eta As Double
        'Const R As Double = 0.461526, tc As Double = 647.096, pc As Double = 22.064, rhoc As Double = 322
        If s < 4.41202148223476 Then
            'Subregion 3a
            'Eq 1, Table 3, Page 8
            Dim Ii = New Double() {0, 0, 0, 1, 1, 1, 1, 1, 2, 2, 3, 3, 3, 4, 4, 4, 4, 5, 6, 7, 8, 10, 10, 14, 18, 20, 22, 22, 24, 28, 28, 32, 32}
            Dim Ji = New Double() {0, 1, 5, 0, 3, 4, 8, 14, 6, 16, 0, 2, 3, 0, 1, 4, 5, 28, 28, 24, 1, 32, 36, 22, 28, 36, 16, 28, 36, 16, 36, 10, 28}
            Dim ni = New Double() {7.70889828326934, -26.0835009128688, 267.416218930389, 17.2221089496844, -293.54233214597, 614.135601882478, -61056.2757725674, -65127225.1118219, 73591.9313521937, -11664650591.4191, 35.5267086434461, -596.144543825955, -475.842430145708, 69.6781965359503, 335.674250377312, 25052.6809130882, 146997.380630766, 5.38069315091534E+19, 1.43619827291346E+21, 3.64985866165994E+19, -2547.41561156775, 2.40120197096563E+27, -3.93847464679496E+29, 1.47073407024852E+24, -4.26391250432059E+31, 1.94509340621077E+38, 6.66212132114896E+23, 7.06777016552858E+33, 1.75563621975576E+41, 1.08408607429124E+28, 7.30872705175151E+43, 1.5914584739887E+24, 3.77121605943324E+40}
            sigma = s / 4.4
            eta = h / 2300
            ps = 0
            For i = 0 To 32
                ps += ni(i) * (eta - 1.01) ^ Ii(i) * (sigma - 0.75) ^ Ji(i)
            Next i
            P3_hs = ps * 99
        Else
            'Subregion 3b
            'Eq 2, Table 4, Page 8
            Dim Ii = New Double() {-12, -12, -12, -12, -12, -10, -10, -10, -10, -8, -8, -6, -6, -6, -6, -5, -4, -4, -4, -3, -3, -3, -3, -2, -2, -1, 0, 2, 2, 5, 6, 8, 10, 14, 14}
            Dim Ji = New Double() {2, 10, 12, 14, 20, 2, 10, 14, 18, 2, 8, 2, 6, 7, 8, 10, 4, 5, 8, 1, 3, 5, 6, 0, 1, 0, 3, 0, 1, 0, 1, 1, 1, 3, 7}
            Dim ni = New Double() {0.000000000000125244360717979, -0.0126599322553713, 5.06878030140626, 31.7847171154202, -391041.161399932, -0.0000000000975733406392044, -18.6312419488279, 510.973543414101, 373847.005822362, 0.0000000299804024666572, 20.0544393820342, -0.00000498030487662829, -10.230180636003, 55.2819126990325, -206.211367510878, -7940.12232324823, 7.82248472028153, -58.6544326902468, 3550.73647696481, -0.000115303107290162, -1.75092403171802, 257.98168774816, -727.048374179467, 0.000121644822609198, 0.0393137871762692, 0.00704181005909296, -82.910820069811, -0.26517881813125, 13.7531682453991, -52.2394090753046, 2405.56298941048, -22736.1631268929, 89074.6343932567, -23923456.5822486, 5687958081.29714}
            sigma = s / 5.3
            eta = h / 2800
            ps = 0
            For i = 0 To 34
                ps += ni(i) * (eta - 0.681) ^ Ii(i) * (sigma - 0.792) ^ Ji(i)
            Next i
            P3_hs = 16.6 / ps
        End If
    End Function
    Private Function H3_pT(ByVal p As Double, ByVal T As Double) As Double
        'Not avalible with IF 97
        'Solve function T3_ph-T=0 with half interval method.
        Dim Ts, Low_Bound, High_Bound, hs As Double
        'ver2.6 Start corrected bug
        If p < 22.06395 Then   'Bellow tripple point
            Ts = T4_p(p)    'Saturation temperature
            If T <= Ts Then   'Liquid side
                High_Bound = H4L_p(p) 'Max h är liauid h.
                Low_Bound = H1_pT(p, 623.15)
            Else
                Low_Bound = H4V_p(p)  'Min h är Vapour h.
                High_Bound = H2_pT(p, B23T_p(p))
            End If
        Else                  'Above tripple point. R3 from R2 till R3.
            Low_Bound = H1_pT(p, 623.15)
            High_Bound = H2_pT(p, B23T_p(p))
        End If
        'ver2.6 End corrected bug
        Ts = T + 1
        Do While Math.Abs(T - Ts) > 0.000001
            hs = (Low_Bound + High_Bound) / 2
            Ts = T3_ph(p, hs)
            If Ts > T Then
                High_Bound = hs
            Else
                Low_Bound = hs
            End If
        Loop
        H3_pT = hs
    End Function
    Private Function T3_prho(ByVal p As Double, ByVal rho As Double)
        'Solve by iteration. Observe that fo low temperatures this equation has 2 solutions.
        'Solve with half interval method
        Dim ps, Low_Bound, High_Bound, Ts As Double
        Low_Bound = 623.15
        High_Bound = 1073.15
        Do While Math.Abs(p - ps) > 0.00000001
            Ts = (Low_Bound + High_Bound) / 2
            ps = P3_rhoT(rho, Ts)
            If ps > p Then
                High_Bound = Ts
            Else
                Low_Bound = Ts
            End If
        Loop
        T3_prho = Ts
    End Function

    '***********************************************************************************************************
    '*2.4 Functions for region 4
    Private Function P4_T(ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        'Section 8.1 The Saturation-Pressure Equation
        'Eq 30, Page 33
        Dim teta, a, b, c As Double
        teta = T - 0.23855557567849 / (T - 650.17534844798)
        a = teta ^ 2 + 1167.0521452767 * teta - 724213.16703206
        b = -17.073846940092 * teta ^ 2 + 12020.82470247 * teta - 3232555.0322333
        c = 14.91510861353 * teta ^ 2 - 4823.2657361591 * teta + 405113.40542057
        P4_T = (2 * c / (-b + (b ^ 2 - 4 * a * c) ^ 0.5)) ^ 4
    End Function

    Private Function T4_p(ByVal p As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        'Section 8.2 The Saturation-Temperature Equation
        'Eq 31, Page 34
        Dim beta, e, f, g, d As Double
        beta = p ^ 0.25
        e = beta ^ 2 - 17.073846940092 * beta + 14.91510861353
        f = 1167.0521452767 * beta ^ 2 + 12020.82470247 * beta - 4823.2657361591
        g = -724213.16703206 * beta ^ 2 - 3232555.0322333 * beta + 405113.40542057
        d = 2 * g / (-f - (f ^ 2 - 4 * e * g) ^ 0.5)
        T4_p = (650.17534844798 + d - ((650.17534844798 + d) ^ 2 - 4 * (-0.23855557567849 + 650.17534844798 * d)) ^ 0.5) / 2
    End Function
    Private Function H4_s(ByVal s As Double) As Double
        'Supplementary Release on Backward Equations ( ) , p h s for Region 3,Equations as a function of h and s for the Region Boundaries, and an Equation( ) sat , T hs for Region 4 of the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam
        '4 Equations for Region Boundaries Given Enthalpy and Entropy
        ' Se picture page 14
        Dim eta, sigma, sigma1, sigma2 As Double, i As Integer
        If s > -0.0001545495919 And s <= 3.77828134 Then
            'hL1_s
            'Eq 3,Table 9,Page 16
            Dim Ii = New Double() {0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 4, 5, 5, 7, 8, 12, 12, 14, 14, 16, 20, 20, 22, 24, 28, 32, 32}
            Dim Ji = New Double() {14, 36, 3, 16, 0, 5, 4, 36, 4, 16, 24, 18, 24, 1, 4, 2, 4, 1, 22, 10, 12, 28, 8, 3, 0, 6, 8}
            Dim ni = New Double() {0.332171191705237, 0.000611217706323496, -8.82092478906822, -0.45562819254325, -0.0000263483840850452, -22.3949661148062, -4.28398660164013, -0.616679338856916, -14.682303110404, 284.523138727299, -113.398503195444, 1156.71380760859, 395.551267359325, -1.54891257229285, 19.4486637751291, -3.57915139457043, -3.35369414148819, -0.66442679633246, 32332.1885383934, 3317.66744667084, -22350.1257931087, 5739538.75852936, 173.226193407919, -0.0363968822121321, 0.000000834596332878346, 5.03611916682674, 65.5444787064505}
            sigma = s / 3.8
            eta = 0
            For i = 0 To 26
                eta += ni(i) * (sigma - 1.09) ^ Ii(i) * (sigma + 0.0000366) ^ Ji(i)
            Next i
            H4_s = eta * 1700
        ElseIf s > 3.77828134 And s <= 4.41202148223476 Then
            'hL3_s
            'Eq 4,Table 10,Page 16
            Dim Ii = New Double() {0, 0, 0, 0, 2, 3, 4, 4, 5, 5, 6, 7, 7, 7, 10, 10, 10, 32, 32}
            Dim Ji = New Double() {1, 4, 10, 16, 1, 36, 3, 16, 20, 36, 4, 2, 28, 32, 14, 32, 36, 0, 6}
            Dim ni = New Double() {0.822673364673336, 0.181977213534479, -0.011200026031362, -0.000746778287048033, -0.179046263257381, 0.0424220110836657, -0.341355823438768, -2.09881740853565, -8.22477343323596, -4.99684082076008, 0.191413958471069, 0.0581062241093136, -1655.05498701029, 1588.70443421201, -85.0623535172818, -31771.4386511207, -94589.0406632871, -0.0000013927384708869, 0.63105253224098}
            sigma = s / 3.8
            eta = 0
            For i = 0 To 18
                eta += ni(i) * (sigma - 1.09) ^ Ii(i) * (sigma + 0.0000366) ^ Ji(i)
            Next i
            H4_s = eta * 1700
        ElseIf s > 4.41202148223476 And s <= 5.85 Then
            'Section 4.4 Equations ( ) 2ab " h s and ( ) 2c3b "h s for the Saturated Vapor Line
            'Page 19, Eq 5
            'hV2c3b_s(s)
            Dim Ii = New Double() {0, 0, 0, 1, 1, 5, 6, 7, 8, 8, 12, 16, 22, 22, 24, 36}
            Dim Ji = New Double() {0, 3, 4, 0, 12, 36, 12, 16, 2, 20, 32, 36, 2, 32, 7, 20}
            Dim ni = New Double() {1.04351280732769, -2.27807912708513, 1.80535256723202, 0.420440834792042, -105721.24483466, 4.36911607493884E+24, -328032702839.753, -6.7868676080427E+15, 7439.57464645363, -3.56896445355761E+19, 1.67590585186801E+31, -3.55028625419105E+37, 396611982166.538, -4.14716268484468E+40, 3.59080103867382E+18, -1.16994334851995E+40}
            sigma = s / 5.9
            eta = 0
            For i = 0 To 15
                eta += ni(i) * (sigma - 1.02) ^ Ii(i) * (sigma - 0.726) ^ Ji(i)
            Next i
            H4_s = eta ^ 4 * 2800
        ElseIf s > 5.85 And s < 9.155759395 Then
            'Section 4.4 Equations ( ) 2ab " h s and ( ) 2c3b "h s for the Saturated Vapor Line
            'Page 20, Eq 6
            Dim Ii = New Double() {1, 1, 2, 2, 4, 4, 7, 8, 8, 10, 12, 12, 18, 20, 24, 28, 28, 28, 28, 28, 32, 32, 32, 32, 32, 36, 36, 36, 36, 36}
            Dim Ji = New Double() {8, 24, 4, 32, 1, 2, 7, 5, 12, 1, 0, 7, 10, 12, 32, 8, 12, 20, 22, 24, 2, 7, 12, 14, 24, 10, 12, 20, 22, 28}
            Dim ni = New Double() {-524.581170928788, -9269472.18142218, -237.385107491666, 21077015581.2776, -23.9494562010986, 221.802480294197, -5104725.33393438, 1249813.96109147, 2000084369.96201, -815.158509791035, -157.612685637523, -11420042233.2791, 6.62364680776872E+15, -2.27622818296144E+18, -1.71048081348406E+31, 6.60788766938091E+15, 1.66320055886021E+22, -2.18003784381501E+29, -7.87276140295618E+29, 1.51062329700346E+31, 7957321.70300541, 1.31957647355347E+15, -3.2509706829914E+23, -4.18600611419248E+25, 2.97478906557467E+34, -9.53588761745473E+19, 1.66957699620939E+24, -1.75407764869978E+32, 3.47581490626396E+34, -7.10971318427851E+38}
            sigma1 = s / 5.21
            sigma2 = s / 9.2
            eta = 0
            For i = 0 To 29
                eta += ni(i) * (1 / sigma1 - 0.513) ^ Ii(i) * (sigma2 - 0.524) ^ Ji(i)
            Next i
            H4_s = Math.Exp(eta) * 2800
        Else
            H4_s = -999.9
        End If
    End Function
    Private Function P4_s(ByVal s As Double) As Double
        'Uses h4_s and Pres_hs for the diffrent regions to determine p4_s
        Dim hsat As Double
        hsat = H4_s(s)
        If s > -0.0001545495919 And s <= 3.77828134 Then
            P4_s = P1_hs(hsat, s)
        ElseIf s > 3.77828134 And s <= 5.210887663 Then
            P4_s = P3_hs(hsat, s)
        ElseIf s > 5.210887663 And s < 9.155759395 Then
            P4_s = P2_hs(hsat, s)
        Else
            P4_s = -999.9
        End If
    End Function
    Private Function H4L_p(ByVal p As Double) As Double
        Dim Low_Bound, High_Bound, hs, ps, Ts As Double
        If p > 0.000611657 And p < 22.06395 Then
            Ts = T4_p(p)
            If p < 16.529 Then
                H4L_p = H1_pT(p, Ts)
            Else
                'Iterate to find the the backward solution of p3sat_h
                Low_Bound = 1670.858218
                High_Bound = 2087.23500164864
                Do While Math.Abs(p - ps) > 0.00001
                    hs = (Low_Bound + High_Bound) / 2
                    ps = P3sat_h(hs)
                    If ps > p Then
                        High_Bound = hs
                    Else
                        Low_Bound = hs
                    End If
                Loop

                H4L_p = hs
            End If
        Else
            H4L_p = -999.9
        End If
    End Function
    Private Function H4V_p(ByVal p As Double) As Double
        Dim Low_Bound, High_Bound, hs, ps, Ts As Double
        If p > 0.000611657 And p < 22.06395 Then
            Ts = T4_p(p)
            If p < 16.529 Then
                H4V_p = H2_pT(p, Ts)
            Else
                'Iterate to find the the backward solution of p3sat_h
                Low_Bound = 2087.23500164864
                High_Bound = 2563.592004 + 5 '5 added to extrapolate to ensure even the border ==350°C solved.
                Do While Math.Abs(p - ps) > 0.000001
                    hs = (Low_Bound + High_Bound) / 2
                    ps = P3sat_h(hs)
                    If ps < p Then
                        High_Bound = hs
                    Else
                        Low_Bound = hs
                    End If
                Loop
                H4V_p = hs
            End If
        Else
            H4V_p = -999.9
        End If
    End Function
    Private Function X4_ph(ByVal p As Double, ByVal h As Double) As Double
        'Calculate vapour fraction from hL and hV for given p
        Dim h4v, h4l As Double
        h4v = H4V_p(p)
        h4l = H4L_p(p)
        If h > h4v Then
            X4_ph = 1
        ElseIf h < h4l Then
            X4_ph = 0
        Else
            X4_ph = (h - h4l) / (h4v - h4l)
        End If
    End Function
    Private Function X4_ps(ByVal p As Double, ByVal s As Double) As Double
        Dim ssV, ssL As Double
        If p < 16.529 Then
            ssV = S2_pT(p, T4_p(p))
            ssL = S1_pT(p, T4_p(p))
        Else
            ssV = S3_rhoT(1 / (V3_ph(p, H4V_p(p))), T4_p(p))
            ssL = S3_rhoT(1 / (V3_ph(p, H4L_p(p))), T4_p(p))
        End If
        If s < ssL Then
            X4_ps = 0
        ElseIf s > ssV Then
            X4_ps = 1
        Else
            X4_ps = (s - ssL) / (ssV - ssL)
        End If
    End Function
    Private Function T4_hs(ByVal h As Double, ByVal s As Double) As Double
        'Supplementary Release on Backward Equations ( ) , p h s for Region 3,
        'Chapter 5.3 page 30.
        'The if 97 function is only valid for part of region4. Use iteration outsida.
        Dim hL, Ts, ss, p, sigma, eta, teta, High_Bound, Low_Bound, PL, s4V, v4V, s4L, v4L, xs As Double, i As Integer
        Dim Ii = New Double() {0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 3, 3, 3, 3, 4, 4, 5, 5, 5, 5, 6, 6, 6, 8, 10, 10, 12, 14, 14, 16, 16, 18, 18, 18, 20, 28}
        Dim Ji = New Double() {0, 3, 12, 0, 1, 2, 5, 0, 5, 8, 0, 2, 3, 4, 0, 1, 1, 2, 4, 16, 6, 8, 22, 1, 20, 36, 24, 1, 28, 12, 32, 14, 22, 36, 24, 36}
        Dim ni = New Double() {0.179882673606601, -0.267507455199603, 1.162767226126, 0.147545428713616, -0.512871635973248, 0.421333567697984, 0.56374952218987, 0.429274443819153, -3.3570455214214, 10.8890916499278, -0.248483390456012, 0.30415322190639, -0.494819763939905, 1.07551674933261, 0.0733888415457688, 0.0140170545411085, -0.106110975998808, 0.0168324361811875, 1.25028363714877, 1013.16840309509, -1.51791558000712, 52.4277865990866, 23049.5545563912, 0.0249459806365456, 2107964.67412137, 366836848.613065, -144814105.365163, -0.0017927637300359, 4899556021.00459, 471.262212070518, -82929439019.8652, -1715.45662263191, 3557776.82973575, 586062760258.436, -12988763.5078195, 31724744937.1057}
        If (s > 5.210887825 And s < 9.15546555571324) Then
            sigma = s / 9.2
            eta = h / 2800
            teta = 0
            For i = 0 To 35
                teta += ni(i) * (eta - 0.119) ^ Ii(i) * (sigma - 1.07) ^ Ji(i)
            Next i
            T4_hs = teta * 550
        Else
            'function psat_h
            If s > -0.0001545495919 And s <= 3.77828134 Then
                Low_Bound = 0.000611
                High_Bound = 165.291642526045
                Do While Math.Abs(hL - h) > 0.00001 And Math.Abs(High_Bound - Low_Bound) > 0.0001
                    PL = (Low_Bound + High_Bound) / 2
                    Ts = T4_p(PL)
                    hL = H1_pT(PL, Ts)
                    If hL > h Then
                        High_Bound = PL
                    Else
                        Low_Bound = PL
                    End If
                Loop
            ElseIf s > 3.77828134 And s <= 4.41202148223476 Then
                PL = P3sat_h(h)
            ElseIf s > 4.41202148223476 And s <= 5.210887663 Then
                PL = P3sat_h(h)
            End If
            Low_Bound = 0.000611
            High_Bound = PL
            Do While Math.Abs(s - ss) > 0.000001 And Math.Abs(High_Bound - Low_Bound) > 0.0000001
                p = (Low_Bound + High_Bound) / 2

                'Calculate s4_ph
                Ts = T4_p(p)
                xs = X4_ph(p, h)
                If p < 16.529 Then
                    s4V = S2_pT(p, Ts)
                    s4L = S1_pT(p, Ts)
                Else
                    v4V = V3_ph(p, H4V_p(p))
                    s4V = S3_rhoT(1 / v4V, Ts)
                    v4L = V3_ph(p, H4L_p(p))
                    s4L = S3_rhoT(1 / v4L, Ts)
                End If
                ss = (xs * s4V + (1 - xs) * s4L)

                If ss < s Then
                    High_Bound = p
                Else
                    Low_Bound = p
                End If
            Loop
            T4_hs = T4_p(p)
        End If
    End Function
    '***********************************************************************************************************
    '*2.5 Functions for region 5
    Private Function H5_pT(ByVal p As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        'Basic Equation for Region 5
        'Eq 32,33, Page 36, Tables 37-41
        Dim tau, gamma0_tau, gammar_tau As Double, i As Integer
        Const R As Double = 0.461526   'kJ/(kg K)
        Dim Ji0 = New Double() {0, 1, -3, -2, -1, 2}
        Dim ni0 = New Double() {-13.179983674201, 6.8540841634434, -0.024805148933466, 0.36901534980333, -3.1161318213925, -0.32961626538917}
        Dim Iir = New Double() {1, 1, 1, 2, 3}
        Dim Jir = New Double() {0, 1, 3, 9, 3}
        Dim nir = New Double() {-0.00012563183589592, 0.0021774678714571, -0.004594282089991, -0.0000039724828359569, 0.00000012919228289784}
        tau = 1000 / T
        gamma0_tau = 0
        For i = 0 To 5
            gamma0_tau += ni0(i) * Ji0(i) * tau ^ (Ji0(i) - 1)
        Next i
        gammar_tau = 0
        For i = 0 To 4
            gammar_tau += nir(i) * p ^ Iir(i) * Jir(i) * tau ^ (Jir(i) - 1)
        Next i
        H5_pT = R * T * tau * (gamma0_tau + gammar_tau)
    End Function


    Private Function V5_pT(ByVal p As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        'Basic Equation for Region 5
        'Eq 32,33, Page 36, Tables 37-41
        Dim tau, gamma0_pi, gammar_pi As Double, i As Integer
        Const R As Double = 0.461526   'kJ/(kg K)
        Dim Ji0 = New Double() {0, 1, -3, -2, -1, 2}
        Dim ni0 = New Double() {-13.179983674201, 6.8540841634434, -0.024805148933466, 0.36901534980333, -3.1161318213925, -0.32961626538917}
        Dim Iir = New Double() {1, 1, 1, 2, 3}
        Dim Jir = New Double() {0, 1, 3, 9, 3}
        Dim nir = New Double() {-0.00012563183589592, 0.0021774678714571, -0.004594282089991, -0.0000039724828359569, 0.00000012919228289784}
        tau = 1000 / T
        gamma0_pi = 1 / p
        gammar_pi = 0
        For i = 0 To 4
            gammar_pi += nir(i) * Iir(i) * p ^ (Iir(i) - 1) * tau ^ Jir(i)
        Next i
        V5_pT = R * T / p * p * (gamma0_pi + gammar_pi) / 1000
    End Function

    Private Function U5_pT(ByVal p As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        'Basic Equation for Region 5
        'Eq 32,33, Page 36, Tables 37-41
        Dim tau, gamma0_pi, gammar_pi, gamma0_tau, gammar_tau As Double, i As Integer
        Const R As Double = 0.461526   'kJ/(kg K)
        Dim Ji0 = New Double() {0, 1, -3, -2, -1, 2}
        Dim ni0 = New Double() {-13.179983674201, 6.8540841634434, -0.024805148933466, 0.36901534980333, -3.1161318213925, -0.32961626538917}
        Dim Iir = New Double() {1, 1, 1, 2, 3}
        Dim Jir = New Double() {0, 1, 3, 9, 3}
        Dim nir = New Double() {-0.00012563183589592, 0.0021774678714571, -0.004594282089991, -0.0000039724828359569, 0.00000012919228289784}
        tau = 1000 / T
        gamma0_pi = 1 / p
        gamma0_tau = 0
        For i = 0 To 5
            gamma0_tau += ni0(i) * Ji0(i) * tau ^ (Ji0(i) - 1)
        Next i
        gammar_pi = 0
        gammar_tau = 0
        For i = 0 To 4
            gammar_pi += nir(i) * Iir(i) * p ^ (Iir(i) - 1) * tau ^ Jir(i)
            gammar_tau += nir(i) * p ^ Iir(i) * Jir(i) * tau ^ (Jir(i) - 1)
        Next i
        U5_pT = R * T * (tau * (gamma0_tau + gammar_tau) - p * (gamma0_pi + gammar_pi))
    End Function
    Private Function Cp5_pT(ByVal p As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        'Basic Equation for Region 5
        'Eq 32,33, Page 36, Tables 37-41
        Dim tau, gamma0_tautau, gammar_tautau As Double, i As Integer
        Const R As Double = 0.461526   'kJ/(kg K)
        Dim Ji0 = New Double() {0, 1, -3, -2, -1, 2}
        Dim ni0 = New Double() {-13.179983674201, 6.8540841634434, -0.024805148933466, 0.36901534980333, -3.1161318213925, -0.32961626538917}
        Dim Iir = New Double() {1, 1, 1, 2, 3}
        Dim Jir = New Double() {0, 1, 3, 9, 3}
        Dim nir = New Double() {-0.00012563183589592, 0.0021774678714571, -0.004594282089991, -0.0000039724828359569, 0.00000012919228289784}
        tau = 1000 / T
        gamma0_tautau = 0
        For i = 0 To 5
            gamma0_tautau += ni0(i) * Ji0(i) * (Ji0(i) - 1) * tau ^ (Ji0(i) - 2)
        Next i
        gammar_tautau = 0
        For i = 0 To 4
            gammar_tautau += nir(i) * p ^ Iir(i) * Jir(i) * (Jir(i) - 1) * tau ^ (Jir(i) - 2)
        Next i
        Cp5_pT = -R * tau ^ 2 * (gamma0_tautau + gammar_tautau)
    End Function

    Private Function S5_pT(ByVal p As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        'Basic Equation for Region 5
        'Eq 32,33, Page 36, Tables 37-41
        Dim tau, gamma0, gamma0_tau, gammar, gammar_tau As Double, i As Integer
        Const R As Double = 0.461526   'kJ/(kg K)
        Dim Ji0 = New Double() {0, 1, -3, -2, -1, 2}
        Dim ni0 = New Double() {-13.179983674201, 6.8540841634434, -0.024805148933466, 0.36901534980333, -3.1161318213925, -0.32961626538917}
        Dim Iir = New Double() {1, 1, 1, 2, 3}
        Dim Jir = New Double() {0, 1, 3, 9, 3}
        Dim nir = New Double() {-0.00012563183589592, 0.0021774678714571, -0.004594282089991, -0.0000039724828359569, 0.00000012919228289784}
        tau = 1000 / T
        gamma0 = Math.Log(p)
        gamma0_tau = 0
        For i = 0 To 5
            gamma0_tau += ni0(i) * Ji0(i) * tau ^ (Ji0(i) - 1)
            gamma0 += ni0(i) * tau ^ Ji0(i)
        Next i
        gammar = 0
        gammar_tau = 0
        For i = 0 To 4
            gammar += nir(i) * p ^ Iir(i) * tau ^ Jir(i)
            gammar_tau += nir(i) * p ^ Iir(i) * Jir(i) * tau ^ (Jir(i) - 1)
        Next i
        S5_pT = R * (tau * (gamma0_tau + gammar_tau) - (gamma0 + gammar))
    End Function
    Private Function Cv5_pT(ByVal p As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        'Basic Equation for Region 5
        'Eq 32,33, Page 36, Tables 37-41
        Dim tau, gamma0_tautau, gammar_pi, gammar_pitau, gammar_pipi, gammar_tautau As Double, i As Integer
        Const R As Double = 0.461526   'kJ/(kg K)
        Dim Ji0 = New Double() {0, 1, -3, -2, -1, 2}
        Dim ni0 = New Double() {-13.179983674201, 6.8540841634434, -0.024805148933466, 0.36901534980333, -3.1161318213925, -0.32961626538917}
        Dim Iir = New Double() {1, 1, 1, 2, 3}
        Dim Jir = New Double() {0, 1, 3, 9, 3}
        Dim nir = New Double() {-0.00012563183589592, 0.0021774678714571, -0.004594282089991, -0.0000039724828359569, 0.00000012919228289784}
        tau = 1000 / T
        gamma0_tautau = 0
        For i = 0 To 5
            gamma0_tautau += ni0(i) * (Ji0(i) - 1) * Ji0(i) * tau ^ (Ji0(i) - 2)
        Next i
        gammar_pi = 0
        gammar_pitau = 0
        gammar_pipi = 0
        gammar_tautau = 0
        For i = 0 To 4
            gammar_pi += nir(i) * Iir(i) * p ^ (Iir(i) - 1) * tau ^ Jir(i)
            gammar_pitau += nir(i) * Iir(i) * p ^ (Iir(i) - 1) * Jir(i) * tau ^ (Jir(i) - 1)
            gammar_pipi += nir(i) * Iir(i) * (Iir(i) - 1) * p ^ (Iir(i) - 2) * tau ^ Jir(i)
            gammar_tautau += nir(i) * p ^ Iir(i) * Jir(i) * (Jir(i) - 1) * tau ^ (Jir(i) - 2)
        Next i
        Cv5_pT = R * (-(tau ^ 2 * (gamma0_tautau + gammar_tautau)) - (1 + p * gammar_pi - tau * p * gammar_pitau) ^ 2 / (1 - p ^ 2 * gammar_pipi))

    End Function
    Private Function W5_pT(ByVal p As Double, ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam, September 1997
        'Basic Equation for Region 5
        'Eq 32,33, Page 36, Tables 37-41
        Dim tau, gamma0_tautau, gammar_pi, gammar_pitau, gammar_pipi, gammar_tautau As Double, i As Integer
        Const R As Double = 0.461526   'kJ/(kg K)
        Dim Ji0 = New Double() {0, 1, -3, -2, -1, 2}
        Dim ni0 = New Double() {-13.179983674201, 6.8540841634434, -0.024805148933466, 0.36901534980333, -3.1161318213925, -0.32961626538917}
        Dim Iir = New Double() {1, 1, 1, 2, 3}
        Dim Jir = New Double() {0, 1, 3, 9, 3}
        Dim nir = New Double() {-0.00012563183589592, 0.0021774678714571, -0.004594282089991, -0.0000039724828359569, 0.00000012919228289784}
        tau = 1000 / T
        gamma0_tautau = 0
        For i = 0 To 5
            gamma0_tautau += ni0(i) * (Ji0(i) - 1) * Ji0(i) * tau ^ (Ji0(i) - 2)
        Next i
        gammar_pi = 0
        gammar_pitau = 0
        gammar_pipi = 0
        gammar_tautau = 0
        For i = 0 To 4
            gammar_pi += nir(i) * Iir(i) * p ^ (Iir(i) - 1) * tau ^ Jir(i)
            gammar_pitau += nir(i) * Iir(i) * p ^ (Iir(i) - 1) * Jir(i) * tau ^ (Jir(i) - 1)
            gammar_pipi += nir(i) * Iir(i) * (Iir(i) - 1) * p ^ (Iir(i) - 2) * tau ^ Jir(i)
            gammar_tautau += nir(i) * p ^ Iir(i) * Jir(i) * (Jir(i) - 1) * tau ^ (Jir(i) - 2)
        Next i
        W5_pT = (1000 * R * T * (1 + 2 * p * gammar_pi + p ^ 2 * gammar_pi ^ 2) / ((1 - p ^ 2 * gammar_pipi) + (1 + p * gammar_pi - tau * p * gammar_pitau) ^ 2 / (tau ^ 2 * (gamma0_tautau + gammar_tautau)))) ^ 0.5
    End Function

    Private Function T5_ph(ByVal p As Double, ByVal h As Double) As Double
        'Solve with half interval method
        Dim Low_Bound, High_Bound, Ts, hs As Double
        Low_Bound = 1073.15
        High_Bound = 2273.15
        Do While Math.Abs(h - hs) > 0.00001
            Ts = (Low_Bound + High_Bound) / 2
            hs = H5_pT(p, Ts)
            If hs > h Then
                High_Bound = Ts
            Else
                Low_Bound = Ts
            End If
        Loop
        T5_ph = Ts
    End Function

    Private Function T5_ps(ByVal p As Double, ByVal s As Double) As Double
        'Solve with half interval method
        Dim Low_Bound, High_Bound, Ts, ss As Double
        Low_Bound = 1073.15
        High_Bound = 2273.15
        Do While Math.Abs(s - ss) > 0.00001
            Ts = (Low_Bound + High_Bound) / 2
            ss = S5_pT(p, Ts)
            If ss > s Then
                High_Bound = Ts
            Else
                Low_Bound = Ts
            End If
        Loop
        T5_ps = Ts
    End Function
    Private Function T5_prho(ByVal p As Double, ByVal rho As Double) As Double
        'Solve by iteration. Observe that fo low temperatures this equation has 2 solutions.
        'Solve with half interval method
        Dim Low_Bound, High_Bound, Ts, rhos As Double
        Low_Bound = 1073.15
        High_Bound = 2073.15
        Do While Math.Abs(rho - rhos) > 0.000001
            Ts = (Low_Bound + High_Bound) / 2
            rhos = 1 / V2_pT(p, Ts)
            If rhos < rho Then
                High_Bound = Ts
            Else
                Low_Bound = Ts
            End If
        Loop
        T5_prho = Ts
    End Function
    '***********************************************************************************************************
    '*3 Region Selection
    '***********************************************************************************************************
    '*3.1 Regions as a function of pT
    Private Function Region_pT(ByVal p As Double, ByVal T As Double) As Integer
        Dim ps As Double
        If T > 1073.15 And p < 10 And T < 2273.15 And p > 0.000611 Then
            Region_pT = 5
        ElseIf T <= 1073.15 And T > 273.15 And p <= 100 And p > 0.000611 Then
            If T > 623.15 Then
                If p > B23p_T(T) Then
                    Region_pT = 3
                    If T < 647.096 Then
                        ps = P4_T(T)
                        If Math.Abs(p - ps) < 0.00001 Then
                            Region_pT = 4
                        End If
                    End If
                Else
                    Region_pT = 2
                End If
            Else
                ps = P4_T(T)
                If Math.Abs(p - ps) < 0.00001 Then
                    Region_pT = 4
                ElseIf p > ps Then
                    Region_pT = 1
                Else
                    Region_pT = 2
                End If
            End If
        Else
            Region_pT = 0 '**Error, Outside valid area
        End If
    End Function
    '***********************************************************************************************************
    '*3.2 Regions as a function of ph
    Private Function Region_ph(ByVal p, ByVal h) As Integer
        Dim hL, hV, h_45, h_5u, Ts As Double
        'Check if outside pressure limits
        If p < 0.000611657 Or p > 100 Then
            Region_ph = 0
            Exit Function
        End If

        'Check if outside low h.
        If h < 0.963 * p + 2.2 Then 'Linear adaption to h1_pt()+2 to speed up calcualations.
            If h < H1_pT(p, 273.15) Then
                Region_ph = 0
                Exit Function
            End If
        End If

        If p < 16.5292 Then 'Bellow region 3,Check  region 1,4,2,5
            'Check Region 1
            Ts = T4_p(p)
            hL = 109.6635 * Math.Log(p) + 40.3481 * p + 734.58 'Approximate function for EnthalpyL_p
            If Math.Abs(h - hL) < 100 Then 'If approximate is not god enough use real function
                hL = H1_pT(p, Ts)
            End If
            If h <= hL Then
                Region_ph = 1
                Exit Function
            End If
            'Check Region 4
            hV = 45.1768 * Math.Log(p) - 20.158 * p + 2804.4 'Approximate function for EnthalpyV_p
            If Math.Abs(h - hV) < 50 Then 'If approximate is not god enough use real function
                hV = H2_pT(p, Ts)
            End If
            If h < hV Then
                Region_ph = 4
                Exit Function
            End If
            'Check upper limit of region 2 Quick Test
            If h < 4000 Then
                Region_ph = 2
                Exit Function
            End If
            'Check region 2 (Real value)
            h_45 = H2_pT(p, 1073.15)
            If h <= h_45 Then
                Region_ph = 2
                Exit Function
            End If
            'Check region 5
            If p > 10 Then
                Region_ph = 0
                Exit Function
            End If
            h_5u = H5_pT(p, 2273.15)
            If h < h_5u Then
                Region_ph = 5
                Exit Function
            End If
            Region_ph = 0
            Exit Function
        Else 'For p>16.5292
            'Check if in region1
            If h < H1_pT(p, 623.15) Then
                Region_ph = 1
                Exit Function
            End If
            'Check if in region 3 or 4 (Bellow Reg 2)
            If h < H2_pT(p, B23T_p(p)) Then
                'Region 3 or 4
                If p > P3sat_h(h) Then
                    Region_ph = 3
                    Exit Function
                Else
                    Region_ph = 4
                    Exit Function
                End If
            End If
            'Check if region 2
            If h < H2_pT(p, 1073.15) Then
                Region_ph = 2
                Exit Function
            End If
        End If
        Region_ph = 0
    End Function
    '***********************************************************************************************************
    '*3.3 Regions as a function of ps
    Private Function Region_ps(ByVal p As Double, ByVal s As Double) As Integer
        Dim ss As Double
        If p < 0.000611657 Or p > 100 Or s < 0 Or s > S5_pT(p, 2273.15) Then
            Region_ps = 0
            Exit Function
        End If

        'Check region 5
        If s > S2_pT(p, 1073.15) Then
            If p <= 10 Then
                Region_ps = 5
                Exit Function
            Else
                Region_ps = 0
                Exit Function
            End If
        End If

        'Check region 2
        If p > 16.529 Then
            ss = S2_pT(p, B23T_p(p)) 'Between 5.047 and 5.261. Use to speed up!
        Else
            ss = S2_pT(p, T4_p(p))
        End If
        If s > ss Then
            Region_ps = 2
            Exit Function
        End If

        'Check region 3
        ss = S1_pT(p, 623.15)
        If p > 16.529 And s > ss Then
            If p > P3sat_s(s) Then
                Region_ps = 3
                Exit Function
            Else
                Region_ps = 4
                Exit Function
            End If
        End If

        'Check region 4 (Not inside region 3)
        If p < 16.529 And s > S1_pT(p, T4_p(p)) Then
            Region_ps = 4
            Exit Function
        End If

        'Check region 1
        If p > 0.000611657 And s > S1_pT(p, 273.15) Then
            Region_ps = 1
            Exit Function
        End If
        Region_ps = 1
    End Function
    '***********************************************************************************************************
    '*3.4 Regions as a function of hs
    Private Function Region_hs(ByVal h As Double, ByVal s As Double) As Integer
        Dim TMax, hMax, hB, hL, hV, vmax, Tmin, hMin As Double
        If s < -0.0001545495919 Then
            Region_hs = 0
            Exit Function
        End If
        'Check linear adaption to p=0.000611. If bellow region 4.
        hMin = (((-0.0415878 - 2500.89262) / (-0.00015455 - 9.155759)) * s)
        If s < 9.155759395 And h < hMin Then
            Region_hs = 0
            Exit Function
        End If

        '******Kolla 1 eller 4. (+liten bit över B13)
        If s >= -0.0001545495919 And s <= 3.77828134 Then
            If h < H4_s(s) Then
                Region_hs = 4
                Exit Function
            ElseIf s < 3.397782955 Then '100MPa line is limiting
                TMax = T1_ps(100, s)
                hMax = H1_pT(100, TMax)
                If h < hMax Then
                    Region_hs = 1
                    Exit Function
                Else
                    Region_hs = 0
                    Exit Function
                End If
            Else 'The point is either in region 4,1,3. Check B23
                hB = HB13_s(s)
                If h < hB Then
                    Region_hs = 1
                    Exit Function
                End If
                TMax = T3_ps(100, s)
                vmax = V3_ps(100, s)
                hMax = H3_rhoT(1 / vmax, TMax)
                If h < hMax Then
                    Region_hs = 3
                    Exit Function
                Else
                    Region_hs = 0
                    Exit Function
                End If
            End If
        End If

        '******Kolla region 2 eller 4. (Övre delen av område b23-> max)
        If s >= 5.260578707 And s <= 11.9212156897728 Then
            If s > 9.155759395 Then 'Above region 4
                Tmin = T2_ps(0.000611, s)
                hMin = H2_pT(0.000611, Tmin)
                'function adapted to h(1073.15,s)
                hMax = -0.07554022 * s ^ 4 + 3.341571 * s ^ 3 - 55.42151 * s ^ 2 + 408.515 * s + 3031.338
                If h > hMin And h < hMax Then
                    Region_hs = 2
                    Exit Function
                Else
                    Region_hs = 0
                    Exit Function
                End If
            End If


            hV = H4_s(s)

            If h < hV Then  'Region 4. Under region 3.
                Region_hs = 4
                Exit Function
            End If
            If s < 6.04048367171238 Then
                TMax = T2_ps(100, s)
                hMax = H2_pT(100, TMax)
            Else
                'function adapted to h(1073.15,s)
                hMax = -2.988734 * s ^ 4 + 121.4015 * s ^ 3 - 1805.15 * s ^ 2 + 11720.16 * s - 23998.33
            End If
            If h < hMax Then  'Region 2. Över region 4.
                Region_hs = 2
                Exit Function
            Else
                Region_hs = 0
                Exit Function
            End If
        End If

        'Kolla region 3 eller 4. Under kritiska punkten.
        If s >= 3.77828134 And s <= 4.41202148223476 Then
            hL = H4_s(s)
            If h < hL Then
                Region_hs = 4
                Exit Function
            End If
            TMax = T3_ps(100, s)
            vmax = V3_ps(100, s)
            hMax = H3_rhoT(1 / vmax, TMax)
            If h < hMax Then
                Region_hs = 3
                Exit Function
            Else
                Region_hs = 0
                Exit Function
            End If
        End If

        'Kolla region 3 eller 4 från kritiska punkten till övre delen av b23
        If s >= 4.41202148223476 And s <= 5.260578707 Then
            hV = H4_s(s)
            If h < hV Then
                Region_hs = 4
                Exit Function
            End If
            'Kolla om vi är under b23 giltighetsområde.
            If s <= 5.048096828 Then
                TMax = T3_ps(100, s)
                vmax = V3_ps(100, s)
                hMax = H3_rhoT(1 / vmax, TMax)
                If h < hMax Then
                    Region_hs = 3
                    Exit Function
                Else
                    Region_hs = 0
                    Exit Function
                End If
            Else 'Inom området för B23 i s led.
                If (h > 2812.942061) Then 'Ovanför B23 i h_led
                    If s > 5.09796573397125 Then
                        TMax = T2_ps(100, s)
                        hMax = H2_pT(100, TMax)
                        If h < hMax Then
                            Region_hs = 2
                            Exit Function
                        Else
                            Region_hs = 0
                            Exit Function
                        End If
                    Else
                        Region_hs = 0
                        Exit Function
                    End If
                End If
                If (h < 2563.592004) Then   'Nedanför B23 i h_led men vi har redan kollat ovanför hV2c3b
                    Region_hs = 3
                    Exit Function
                End If
                'Vi är inom b23 området i både s och h led.
                If P2_hs(h, s) > B23p_T(TB23_hs(h, s)) Then
                    Region_hs = 3
                    Exit Function
                Else
                    Region_hs = 2
                    Exit Function
                End If
            End If
        End If
        Region_hs = -999.9
    End Function
    '***********************************************************************************************************
    '*3.5 Regions as a function of p and rho
    Private Function Region_prho(ByVal p As Double, ByVal rho As Double) As Integer
        Dim v As Double
        v = 1 / rho
        If p < 0.000611657 Or p > 100 Then
            Region_prho = 0
            Exit Function
        End If
        If p < 16.5292 Then 'Bellow region 3, Check region 1,4,2
            If v < V1_pT(p, 273.15) Then 'Observe that this is not actually min of v. Not valid Water of 4°C is ligther.
                Region_prho = 0
                Exit Function
            End If
            If v <= V1_pT(p, T4_p(p)) Then
                Region_prho = 1
                Exit Function
            End If
            If v < V2_pT(p, T4_p(p)) Then
                Region_prho = 4
                Exit Function
            End If
            If v <= V2_pT(p, 1073.15) Then
                Region_prho = 2
                Exit Function
            End If
            If p > 10 Then 'Above region 5
                Region_prho = 0
                Exit Function
            End If
            If v <= V5_pT(p, 2073.15) Then
                Region_prho = 5
                Exit Function
            End If
        Else 'Check region 1,3,4,3,2 (Above the lowest point of region 3.)
            If v < V1_pT(p, 273.15) Then 'Observe that this is not actually min of v. Not valid Water of 4°C is ligther.
                Region_prho = 0
                Exit Function
            End If
            If v < V1_pT(p, 623.15) Then
                Region_prho = 1
                Exit Function
            End If
            'Check if in region 3 or 4 (Bellow Reg 2)
            If v < V2_pT(p, B23T_p(p)) Then
                'Region 3 or 4
                If p > 22.064 Then 'Above region 4
                    Region_prho = 3
                    Exit Function
                End If
                If v < V3_ph(p, H4L_p(p)) Or v > V3_ph(p, H4V_p(p)) Then 'Uses iteration!!
                    Region_prho = 3
                    Exit Function
                Else
                    Region_prho = 4
                    Exit Function
                End If
            End If
            'Check if region 2
            If v < V2_pT(p, 1073.15) Then
                Region_prho = 2
                Exit Function
            End If
        End If

        Region_prho = 0
    End Function


    '***********************************************************************************************************
    '*4 Region Borders
    '***********************************************************************************************************
    '***********************************************************************************************************
    '*4.1 Boundary between region 2 and 3.
    Private Function B23p_T(ByVal T As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam
        '1997
        'Section 4 Auxiliary Equation for the Boundary between Regions 2 and 3
        'Eq 5, Page 5
        B23p_T = 348.05185628969 - 1.1671859879975 * T + 0.0010192970039326 * T ^ 2
    End Function
    Private Function B23T_p(ByVal p As Double) As Double
        'Release on the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam
        '1997
        'Section 4 Auxiliary Equation for the Boundary between Regions 2 and 3
        'Eq 6, Page 6
        B23T_p = 572.54459862746 + ((p - 13.91883977887) / 0.0010192970039326) ^ 0.5
    End Function
    '***********************************************************************************************************
    '*4.2 Region 3. pSat_h and pSat_s
    Private Function P3sat_h(ByVal h As Double) As Double
        'Revised Supplementary Release on Backward Equations for the Functions T(p,h), v(p,h) and T(p,s), v(p,s) for Region 3 of the IAPWS Industrial Formulation 1997 for the Thermodynamic Properties of Water and Steam
        '2004
        'Section 4 Boundary Equations psat(h) and psat(s) for the Saturation Lines of Region 3
        'Se pictures Page 17, Eq 10, Table 17, Page 18
        Dim ps As Double, i As Integer
        Dim Ii = New Double() {0, 1, 1, 1, 1, 5, 7, 8, 14, 20, 22, 24, 28, 36}
        Dim Ji = New Double() {0, 1, 3, 4, 36, 3, 0, 24, 16, 16, 3, 18, 8, 24}
        Dim ni = New Double() {0.600073641753024, -9.36203654849857, 24.6590798594147, -107.014222858224, -91582131580576.8, -8623.32011700662, -23.5837344740032, 2.52304969384128E+17, -3.89718771997719E+18, -3.33775713645296E+22, 35649946963.6328, -1.48547544720641E+26, 3.30611514838798E+18, 8.13641294467829E+37}
        h /= 2600
        ps = 0
        For i = 0 To 13
            ps += ni(i) * (h - 1.02) ^ Ii(i) * (h - 0.608) ^ Ji(i)
        Next i
        P3sat_h = ps * 22
    End Function
    Private Function P3sat_s(ByVal s As Double) As Double
        Dim sigma, p As Double, i As Integer
        Dim Ii = New Double() {0, 1, 1, 4, 12, 12, 16, 24, 28, 32}
        Dim Ji = New Double() {0, 1, 32, 7, 4, 14, 36, 10, 0, 18}
        Dim ni = New Double() {0.639767553612785, -12.9727445396014, -2.24595125848403E+15, 1774667.41801846, 7170793495.71538, -3.78829107169011E+17, -9.55586736431328E+34, 1.87269814676188E+23, 119254746466.473, 1.10649277244882E+36}
        sigma = s / 5.2
        p = 0
        For i = 0 To 9
            p += ni(i) * (sigma - 1.03) ^ Ii(i) * (sigma - 0.699) ^ Ji(i)
        Next i
        P3sat_s = p * 22
    End Function
    '***********************************************************************************************************
    '4.3 Region boundary 1to3 and 3to2 as a functions of s
    Private Function HB13_s(ByVal s As Double) As Double
        'Supplementary Release on Backward Equations ( ) , p h s for Region 3,
        'Chapter 4.5 page 23.
        Dim sigma, eta As Double, i As Integer
        Dim Ii = New Double() {0, 1, 1, 3, 5, 6}
        Dim Ji = New Double() {0, -2, 2, -12, -4, -3}
        Dim ni = New Double() {0.913965547600543, -0.0000430944856041991, 60.3235694765419, 1.17518273082168E-18, 0.220000904781292, -69.0815545851641}
        sigma = s / 3.8
        eta = 0
        For i = 0 To 5
            eta += ni(i) * (sigma - 0.884) ^ Ii(i) * (sigma - 0.864) ^ Ji(i)
        Next i
        HB13_s = eta * 1700
    End Function
    Private Function TB23_hs(ByVal h As Double, ByVal s As Double) As Double
        'Supplementary Release on Backward Equations ( ) , p h s for Region 3,
        'Chapter 4.6 page 25.
        Dim sigma, eta, teta As Double, i As Integer
        Dim Ii = New Double() {-12, -10, -8, -4, -3, -2, -2, -2, -2, 0, 1, 1, 1, 3, 3, 5, 6, 6, 8, 8, 8, 12, 12, 14, 14}
        Dim Ji = New Double() {10, 8, 3, 4, 3, -6, 2, 3, 4, 0, -3, -2, 10, -2, -1, -5, -6, -3, -8, -2, -1, -12, -1, -12, 1}
        Dim ni = New Double() {0.00062909626082981, -0.000823453502583165, 0.0000000515446951519474, -1.17565945784945, 3.48519684726192, -0.00000000000507837382408313, -2.84637670005479, -2.36092263939673, 6.01492324973779, 1.48039650824546, 0.000360075182221907, -0.0126700045009952, -1221843.32521413, 0.149276502463272, 0.698733471798484, -0.0252207040114321, 0.0147151930985213, -1.08618917681849, -0.000936875039816322, 81.9877897570217, -182.041861521835, 0.00000261907376402688, -29162.6417025961, 0.0000140660774926165, 7832370.62349385}
        sigma = s / 5.3
        eta = h / 3000
        teta = 0
        For i = 0 To 24
            teta += ni(i) * (eta - 0.727) ^ Ii(i) * (sigma - 0.864) ^ Ji(i)
        Next i
        TB23_hs = teta * 900
    End Function

    '***********************************************************************************************************
    '*5 Transport properties
    '***********************************************************************************************************
    '*5.1 Viscosity (IAPWS formulation 1985, Revised 2003)
    '***********************************************************************************************************
    Private Function My_AllRegions_pT(ByVal p As Double, ByVal T As Double) As Double
        Dim rho, Ts, ps, my0, sum, my1, rhos As Double, i As Integer
        Dim h0 = New Double() {0.5132047, 0.3205656, 0, 0, -0.7782567, 0.1885447}
        Dim h1 = New Double() {0.2151778, 0.7317883, 1.241044, 1.476783, 0, 0}
        Dim h2 = New Double() {-0.2818107, -1.070786, -1.263184, 0, 0, 0}
        Dim h3 = New Double() {0.1778064, 0.460504, 0.2340379, -0.4924179, 0, 0}
        Dim h4 = New Double() {-0.0417661, 0, 0, 0.1600435, 0, 0}
        Dim h5 = New Double() {0, -0.01578386, 0, 0, 0, 0}
        Dim h6 = New Double() {0, 0, 0, -0.003629481, 0, 0}

        'Calcualte density.
        Select Case Region_pT(p, T)
            Case 1
                rho = 1 / V1_pT(p, T)
            Case 2
                rho = 1 / V2_pT(p, T)
            Case 3
                rho = 1 / V3_ph(p, H3_pT(p, T))
            Case 4
                rho = -999.9
            Case 5
                rho = 1 / v5_pT(p, T)
            Case Else
                My_AllRegions_pT = -999.9
                Exit Function
        End Select

        rhos = rho / 317.763
        Ts = T / 647.226
        ps = p / 22.115

        'Check valid area
        If T > 900 + 273.15 Or (T > 600 + 273.15 And p > 300) Or (T > 150 + 273.15 And p > 350) Or p > 500 Then
            My_AllRegions_pT = -999.9
            Exit Function
        End If
        my0 = Ts ^ 0.5 / (1 + 0.978197 / Ts + 0.579829 / (Ts ^ 2) - 0.202354 / (Ts ^ 3))
        sum = 0
        For i = 0 To 5
            sum = sum + h0(i) * (1 / Ts - 1) ^ i + h1(i) * (1 / Ts - 1) ^ i * (rhos - 1) ^ 1 + h2(i) * (1 / Ts - 1) ^ i * (rhos - 1) ^ 2 + h3(i) * (1 / Ts - 1) ^ i * (rhos - 1) ^ 3 + h4(i) * (1 / Ts - 1) ^ i * (rhos - 1) ^ 4 + h5(i) * (1 / Ts - 1) ^ i * (rhos - 1) ^ 5 + h6(i) * (1 / Ts - 1) ^ i * (rhos - 1) ^ 6
        Next i
        my1 = Math.Exp(rhos * sum)
        My_AllRegions_pT = my0 * my1 * 0.000055071
    End Function

    Private Function My_AllRegions_ph(ByVal p As Double, ByVal h As Double) As Double
        Dim rho, T, Ts, ps, my0, sum, my1, rhos, v4V, v4L, xs As Double, i As Integer
        Dim h0 = New Double() {0.5132047, 0.3205656, 0, 0, -0.7782567, 0.1885447}
        Dim h1 = New Double() {0.2151778, 0.7317883, 1.241044, 1.476783, 0, 0}
        Dim h2 = New Double() {-0.2818107, -1.070786, -1.263184, 0, 0, 0}
        Dim h3 = New Double() {0.1778064, 0.460504, 0.2340379, -0.4924179, 0, 0}
        Dim h4 = New Double() {-0.0417661, 0, 0, 0.1600435, 0, 0}
        Dim h5 = New Double() {0, -0.01578386, 0, 0, 0, 0}
        Dim h6 = New Double() {0, 0, 0, -0.003629481, 0, 0}

        'Calcualte density.
        Select Case Region_ph(p, h)
            Case 1
                Ts = T1_ph(p, h)
                T = Ts
                rho = 1 / V1_pT(p, Ts)
            Case 2
                Ts = T2_ph(p, h)
                T = Ts
                rho = 1 / V2_pT(p, Ts)
            Case 3
                rho = 1 / V3_ph(p, h)
                T = T3_ph(p, h)
            Case 4
                xs = x4_ph(p, h)
                If p < 16.529 Then
                    v4V = V2_pT(p, T4_p(p))
                    v4L = V1_pT(p, T4_p(p))
                Else
                    v4V = V3_ph(p, h4V_p(p))
                    v4L = V3_ph(p, h4L_p(p))
                End If
                rho = 1 / (xs * v4V + (1 - xs) * v4L)
                T = T4_p(p)
            Case 5
                Ts = T5_ph(p, h)
                T = Ts
                rho = 1 / v5_pT(p, Ts)
            Case Else
                My_AllRegions_ph = -999.9
                Exit Function
        End Select
        rhos = rho / 317.763
        Ts = T / 647.226
        ps = p / 22.115
        'Check valid area
        If T > 900 + 273.15 Or (T > 600 + 273.15 And p > 300) Or (T > 150 + 273.15 And p > 350) Or p > 500 Then
            My_AllRegions_ph = -999.9
            Exit Function
        End If
        my0 = Ts ^ 0.5 / (1 + 0.978197 / Ts + 0.579829 / (Ts ^ 2) - 0.202354 / (Ts ^ 3))

        sum = 0
        For i = 0 To 5
            sum = sum + h0(i) * (1 / Ts - 1) ^ i + h1(i) * (1 / Ts - 1) ^ i * (rhos - 1) ^ 1 + h2(i) * (1 / Ts - 1) ^ i * (rhos - 1) ^ 2 + h3(i) * (1 / Ts - 1) ^ i * (rhos - 1) ^ 3 + h4(i) * (1 / Ts - 1) ^ i * (rhos - 1) ^ 4 + h5(i) * (1 / Ts - 1) ^ i * (rhos - 1) ^ 5 + h6(i) * (1 / Ts - 1) ^ i * (rhos - 1) ^ 6
        Next i
        my1 = Math.Exp(rhos * sum)
        My_AllRegions_ph = my0 * my1 * 0.000055071
    End Function
    '***********************************************************************************************************
    '*5.2 Thermal Conductivity (IAPWS formulation 1985)
    Private Function Tc_ptrho(ByVal p As Double, ByVal T As Double, ByVal rho As Double) As Double
        'Revised release on the IAPS Formulation 1985 for the Thermal Conductivity of ordinary water
        'IAPWS September 1998
        'Page 8
        'ver2.6 Start corrected bug
        Dim tc0, tc1, dT, Q, s, tc2 As Double
        If T < 273.15 Then
            Tc_ptrho = -999.9 'Out of range of validity (para. B4)
            Exit Function
        ElseIf T < 500 + 273.15 Then
            If p > 100 Then
                Tc_ptrho = -999.9 'Out of range of validity (para. B4)
                Exit Function
            End If
        ElseIf T <= 650 + 273.15 Then
            If p > 70 Then
                Tc_ptrho = -999.9 'Out of range of validity (para. B4)
                Exit Function
            End If
        ElseIf T <= 800 + 273.15 Then
            If p > 40 Then
                Tc_ptrho = -999.9 'Out of range of validity (para. B4)
                Exit Function
            End If
        End If
        'ver2.6 End corrected bug

        T /= 647.26
        rho /= 317.7
        'rho = rho / 317.7
        tc0 = T ^ 0.5 * (0.0102811 + 0.0299621 * T + 0.0156146 * T ^ 2 - 0.00422464 * T ^ 3)
        tc1 = -0.39707 + 0.400302 * rho + 1.06 * Math.Exp(-0.171587 * (rho + 2.39219) ^ 2)
        dT = Math.Abs(T - 1) + 0.00308976
        Q = 2 + 0.0822994 / dT ^ (3 / 5)
        If T >= 1 Then
            s = 1 / dT
        Else
            s = 10.0932 / dT ^ (3 / 5)
        End If
        tc2 = (0.0701309 / T ^ 10 + 0.011852) * rho ^ (9 / 5) * Math.Exp(0.642857 * (1 - rho ^ (14 / 5))) + 0.00169937 * s * rho ^ Q * Math.Exp((Q / (1 + Q)) * (1 - rho ^ (1 + Q))) - 1.02 * Math.Exp(-4.11717 * T ^ (3 / 2) - 6.17937 / rho ^ 5)
        Tc_ptrho = tc0 + tc1 + tc2
    End Function
    '***********************************************************************************************************
    '5.3 Surface Tension
    Private Function Surface_Tension_T(ByVal T As Double)
        'IAPWS Release on Surface Tension of Ordinary Water Substance,
        'September 1994
        Dim tau As Double
        Const tc As Double = 647.096, b As Double = 0.2358, bb As Double = -0.625, my As Double = 1.256
        If T < 0.01 Or T > tc Then
            Surface_Tension_T = "Out of valid region"
            Exit Function
        End If
        tau = 1 - T / tc
        Surface_Tension_T = b * tau ^ my * (1 + bb * tau)
    End Function
    '***********************************************************************************************************
    '*6 Units                                                                                      *
    '***********************************************************************************************************

    Private Function ToSIunit_p(ByVal Ins As Double) As Double
        'Translate bar to MPa
        ToSIunit_p = Ins / 10
    End Function
    Private Function FromSIunit_p(ByVal Ins As Double) As Double
        'Translate bar to MPa
        FromSIunit_p = Ins * 10
    End Function
    Private Function ToSIunit_T(ByVal Ins As Double) As Double
        'Translate degC to Kelvon
        ToSIunit_T = Ins + 273.15
    End Function
    Private Function FromSIunit_T(ByVal Ins As Double) As Double
        'Translate Kelvin to degC
        FromSIunit_T = Ins - 273.15
    End Function
    Private Function ToSIunit_h(ByVal Ins As Double) As Double
        ToSIunit_h = Ins
    End Function
    Private Function FromSIunit_h(ByVal Ins As Double) As Double
        FromSIunit_h = Ins
    End Function
    Private Function ToSIunit_v(ByVal Ins As Double) As Double
        ToSIunit_v = Ins
    End Function
    Private Function FromSIunit_v(ByVal Ins As Double) As Double
        FromSIunit_v = Ins
    End Function
    Private Function ToSIunit_s(ByVal Ins As Double) As Double
        ToSIunit_s = Ins
    End Function
    Private Function FromSIunit_s(ByVal Ins As Double) As Double
        FromSIunit_s = Ins
    End Function
    Private Function ToSIunit_u(ByVal Ins As Double) As Double
        ToSIunit_u = Ins
    End Function
    Private Function FromSIunit_u(ByVal Ins As Double) As Double
        FromSIunit_u = Ins
    End Function
    Private Function ToSIunit_Cp(ByVal Ins As Double) As Double
        ToSIunit_Cp = Ins
    End Function
    Private Function FromSIunit_Cp(ByVal Ins As Double) As Double
        FromSIunit_Cp = Ins
    End Function
    Private Function ToSIunit_Cv(ByVal Ins As Double) As Double
        ToSIunit_Cv = Ins
    End Function
    Private Function FromSIunit_Cv(ByVal Ins As Double) As Double
        FromSIunit_Cv = Ins
    End Function
    Private Function ToSIunit_w(ByVal Ins As Double) As Double
        ToSIunit_w = Ins
    End Function
    Private Function FromSIunit_w(ByVal Ins As Double) As Double
        FromSIunit_w = Ins
    End Function
    Private Function ToSIunit_tc(ByVal Ins As Double) As Double
        ToSIunit_tc = Ins
    End Function
    Private Function FromSIunit_tc(ByVal Ins As Double) As Double
        FromSIunit_tc = Ins
    End Function
    Private Function ToSIunit_st(ByVal Ins As Double) As Double
        ToSIunit_st = Ins
    End Function
    Private Function FromSIunit_st(ByVal Ins As Double) As Double
        FromSIunit_st = Ins
    End Function
    Private Function ToSIunit_x(ByVal Ins As Double) As Double
        ToSIunit_x = Ins
    End Function
    Private Function FromSIunit_x(ByVal Ins As Double) As Double
        FromSIunit_x = Ins
    End Function
    Private Function ToSIunit_vx(ByVal Ins As Double) As Double
        ToSIunit_vx = Ins
    End Function
    Private Function FromSIunit_vx(ByVal Ins As Double) As Double
        FromSIunit_vx = Ins
    End Function
    Private Function ToSIunit_my(ByVal Ins As Double) As Double
        ToSIunit_my = Ins
    End Function
    Private Function FromSIunit_my(ByVal Ins As Double) As Double
        FromSIunit_my = Ins
    End Function



End Module
