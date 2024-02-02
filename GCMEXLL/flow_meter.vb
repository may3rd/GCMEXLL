Imports System.Net.Mime.MediaTypeNames
Imports ExcelDna.Integration

Public Module flow_meter
    'Chemical Engineering Design Library (ChEDL). Utilities for process modeling.
    'Copyright (C) 2018, 2019, 2020 Caleb Bell <Caleb.Andrew.Bell@gmail.com>
    'Modify to VBA by Maetee L. <may3rd@gmail.com>
    '
    'Permission is hereby granted, free of charge, to any person obtaining a copy
    'of this software and associated documentation files (the "Software"), to deal
    'in the Software without restriction, including without limitation the rights
    'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    'copies of the Software, and to permit persons to whom the Software is
    'furnished to do so, subject to the following conditions:
    '
    'The above copyright notice and this permission notice shall be included in all
    'copies or substantial portions of the Software.
    '
    'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    'SOFTWARE.
    '
    'This module contains correlations, standards, and solvers for orifice plates
    'and other flow metering devices. Both permanent and measured pressure drop
    'is included, and models work for both liquids and gases. A number of
    'non-standard devices are included, as well as limited two-phase functionality.

    'For reporting bugs, adding feature requests, or submitting pull requests,
    'please use the `GitHub issue tracker <https://github.com/CalebBell/fluids/>`_
    'or contact the author at Caleb.Andrew.Bell@gmail.com.

    '.. contents:: :local:

    'Flow Meter Solvers
    '------------------
    '.. autofunction:: differential_pressure_meter_solver

    'Flow Meter Interfaces
    '---------------------
    '.. autofunction:: differential_pressure_meter_dP
    '.. autofunction:: differential_pressure_meter_C_epsilon
    '.. autofunction:: differential_pressure_meter_beta
    '.. autofunction:: DP_orifice

    'Orifice Plate Correlations
    '--------------------------
    '.. autofunction:: C_Reader_Harris_Gallagher
    '.. autofunction:: C_eccentric_orifice_ISO_15377_1998
    '.. autofunction:: C_quarter_circle_orifice_ISO_15377_1998
    '.. autofunction:: C_Miller_1996
    '.. autofunction:: orifice_expansibility
    '.. autofunction:: orifice_expansibility_1989
    '.. autodata:: ISO_15377_CONICAL_ORIFICE_C

    'Nozzle Flow Meters
    '------------------
    '.. autofunction:: C_long_radius_nozzle
    '.. autofunction:: C_ISA_1932_nozzle
    '.. autofunction:: C_venturi_nozzle
    '.. autofunction:: nozzle_expansibility

    'Venturi Tube Meters
    '-------------------
    '.. autodata:: ROUGH_WELDED_CONVERGENT_VENTURI_TUBE_C
    '.. autodata:: MACHINED_CONVERGENT_VENTURI_TUBE_C
    '.. autodata:: AS_CAST_VENTURI_TUBE_C
    '.. autofunction:: DP_venturi_tube
    '.. autofunction:: C_Reader_Harris_Gallagher_wet_venturi_tube
    '.. autofunction:: DP_Reader_Harris_Gallagher_wet_venturi_tube

    'Cone Meters
    '-----------
    '.. autodata:: CONE_METER_C
    '.. autofunction:: diameter_ratio_cone_meter
    '.. autofunction:: cone_meter_expansibility_Stewart
    '.. autofunction:: DP_cone_meter

    'Wedge Meters
    '------------
    '.. autofunction:: C_wedge_meter_ISO_5167_6_2017
    '.. autofunction:: C_wedge_meter_Miller
    '.. autofunction:: diameter_ratio_wedge_meter
    '.. autofunction:: DP_wedge_meter

    'Flow Meter Utilities
    '--------------------
    '.. autofunction:: discharge_coefficient_to_K
    '.. autofunction:: K_to_discharge_coefficient
    '.. autofunction:: velocity_of_approach_factor
    '.. autofunction:: flow_coefficient
    '.. autofunction:: flow_meter_discharge
    '.. autodata:: all_meters
    Public Const PI As Double = 3.14159265358979

    Const CONCENTRIC_ORIFICE = "orifice" ' normal
    Const ECCENTRIC_ORIFICE = "eccentric orifice"
    Const CONICAL_ORIFICE = "conical orifice"
    Const SEGMENTAL_ORIFICE = "segmental orifice"
    Const QUARTER_CIRCLE_ORIFICE = "quarter circle orifice"
    Const CONDITIONING_4_HOLE_ORIFICE = "Rosemount 4 hole self conditioning"
    'ORIFICE_HOLE_TYPES = {CONCENTRIC_ORIFICE, ECCENTRIC_ORIFICE, CONICAL_ORIFICE, SEGMENTAL_ORIFICE, QUARTER_CIRCLE_ORIFICE)

    Const ORIFICE_CORNER_TAPS = "corner"
    Const ORIFICE_FLANGE_TAPS = "flange"
    Const ORIFICE_D_AND_D_2_TAPS = "D and D/2"
    Const ORIFICE_PIPE_TAPS = "pipe" ' Not in ISO 5167
    Const ORIFICE_VENA_CONTRACTA_TAPS = "vena contracta" ' Not in ISO 5167, normally segmental or eccentric orifices

    ' Used by miller; modifier on taps
    Const TAPS_OPPOSITE = "180 degree"
    Const TAPS_SIDE = "90 degree"

    Const ISO_5167_ORIFICE = "ISO 5167 orifice"
    Const ISO_15377_ECCENTRIC_ORIFICE = "ISO 15377 eccentric orifice"
    Const ISO_15377_QUARTER_CIRCLE_ORIFICE = "ISO 15377 quarter-circle orifice"
    Const ISO_15377_CONICAL_ORIFICE = "ISO 15377 conical orifice"

    Const MILLER_ORIFICE = "Miller orifice"
    Const MILLER_ECCENTRIC_ORIFICE = "Miller eccentric orifice"
    Const MILLER_SEGMENTAL_ORIFICE = "Miller segmental orifice"
    Const MILLER_CONICAL_ORIFICE = "Miller conical orifice"
    Const MILLER_QUARTER_CIRCLE_ORIFICE = "Miller quarter circle orifice"

    Const UNSPECIFIED_METER = "unspecified meter"

    Const LONG_RADIUS_NOZZLE = "long radius nozzle"
    Const ISA_1932_NOZZLE = "ISA 1932 nozzle"
    Const VENTURI_NOZZLE = "venturi nozzle"

    Const AS_CAST_VENTURI_TUBE = "as cast convergent venturi tube"
    Const MACHINED_CONVERGENT_VENTURI_TUBE = "machined convergent venturi tube"
    Const ROUGH_WELDED_CONVERGENT_VENTURI_TUBE = "rough welded convergent venturi tube"

    Const HOLLINGSHEAD_ORIFICE = "Hollingshead orifice"
    Const HOLLINGSHEAD_VENTURI_SMOOTH = "Hollingshead venturi smooth"
    Const HOLLINGSHEAD_VENTURI_SHARP = "Hollingshead venturi sharp"
    Const HOLLINGSHEAD_CONE = "Hollingshead v cone"
    Const HOLLINGSHEAD_WEDGE = "Hollingshead wedge"

    Const CONE_METER = "cone meter"
    Const WEDGE_METER = "wedge meter"

    Function IsInArray(stringToBeFound As String, arr As Array) As Boolean
        IsInArray = Array.IndexOf(arr, stringToBeFound) > 0
    End Function

    Function IsInBetaSimpleMeter(mtype As String) As Boolean
        Dim beta_simple_meters = New String() {ISO_5167_ORIFICE, ISO_15377_ECCENTRIC_ORIFICE,
                              ISO_15377_CONICAL_ORIFICE, ISO_15377_QUARTER_CIRCLE_ORIFICE,
                              MILLER_ORIFICE, MILLER_ECCENTRIC_ORIFICE,
                              MILLER_SEGMENTAL_ORIFICE, MILLER_CONICAL_ORIFICE,
                              MILLER_QUARTER_CIRCLE_ORIFICE,
                              CONCENTRIC_ORIFICE, ECCENTRIC_ORIFICE, CONICAL_ORIFICE,
                              SEGMENTAL_ORIFICE, QUARTER_CIRCLE_ORIFICE,
                              UNSPECIFIED_METER,
                              HOLLINGSHEAD_VENTURI_SHARP, HOLLINGSHEAD_VENTURI_SMOOTH, HOLLINGSHEAD_ORIFICE,
                              LONG_RADIUS_NOZZLE,
                              ISA_1932_NOZZLE, VENTURI_NOZZLE,
                              AS_CAST_VENTURI_TUBE,
                              MACHINED_CONVERGENT_VENTURI_TUBE,
                              ROUGH_WELDED_CONVERGENT_VENTURI_TUBE}
        IsInBetaSimpleMeter = IsInArray(mtype, beta_simple_meters)
    End Function

    <ExcelFunction(Description:="Calculate mass flow rate across orifice", Category:="GCME E-PT | Orifice")>
    Public Function Flow_meter_discharge(D, D0, P1, P2, rho, c, expansibility) As Double
        '    Calculates the flow rate of an orifice plate based on the geometry
        '    of the plate, measured pressures of the orifice, and the density of the
        '    fluid.
        '
        '    .. math::
        '        m = \left(\frac{\pi D_o^2}{4}\right) C \frac{\sqrt{2\Delta P \rho_1}}
        '        {\sqrt{1 - \beta^4}}\cdot \epsilon
        '
        '    Parameters
        '    ----------
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    Do : float
        '        Diameter of orifice at flow conditions, [m]
        '    P1 : float
        '        Static pressure of fluid upstream of orifice at the cross-section of
        '        the pressure tap, [Pa]
        '    P2 : float
        '        Static pressure of fluid downstream of orifice at the cross-section of
        '        the pressure tap, [Pa]
        '    rho : float
        '        Density of fluid at `P1`, [kg/m^3]
        '    C : float
        '        Coefficient of discharge of the orifice, [-]
        '    expansibility : float, optional
        '        Expansibility factor (1 for incompressible fluids, less than 1 for
        '        real fluids), [-]
        '
        '    Returns
        '    -------
        '    m : float
        '        Mass flow rate of fluid, [kg/s]
        '
        '    Notes
        '    -----
        '    This is formula 1-12 in [1]_ and also [2]_.
        '
        '    Examples
        '    --------
        '    >>> flow_meter_discharge(D=0.0739, Do=0.0222, P1=1E5, P2=9.9E4, rho=1.1646,
        '    ... C=0.5988, expansibility=0.9975)
        '    0.01120390943807026
        '
        '    References
        '    ----------
        '    .. [1] American Society of Mechanical Engineers. Mfc-3M-2004 Measurement
        '       Of Fluid Flow In Pipes Using Orifice, Nozzle, And Venturi. ASME, 2001.
        '    .. [2] ISO 5167-2:2003 - Measurement of Fluid Flow by Means of Pressure
        '       Differential Devices Inserted in Circular Cross-Section Conduits Running
        '       Full -- Part 2: Orifice Plates.
        '

        Dim beta, beta2
        beta = D0 / D
        beta2 = beta * beta

        Flow_meter_discharge = (0.25 * PI * D0 * D0) * c * expansibility * Math.Sqrt((2.0 * rho * (P1 - P2)) / (1.0 - beta2 * beta2))

    End Function

    Public Function Orifice_expansibility(D, D0, P1, P2, k) As Double
        '    Calculates the expansibility factor for orifice plate calculations
        '    based on the geometry of the plate, measured pressures of the orifice, and
        '    the isentropic exponent of the fluid.
        '
        '    .. math::
        '        \epsilon = 1 - (0.351 + 0.256\beta^4 + 0.93\beta^8)
        '        \left[1-\left(\frac{P_2}{P_1}\right)^{1/\kappa}\right]
        '
        '    Parameters
        '    ----------
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    Do : float
        '        Diameter of orifice at flow conditions, [m]
        '    P1 : float
        '        Static pressure of fluid upstream of orifice at the cross-section of
        '        the pressure tap, [Pa]
        '    P2 : float
        '        Static pressure of fluid downstream of orifice at the cross-section of
        '        the pressure tap, [Pa]
        '    k : float
        '        Isentropic exponent of fluid, [-]
        '
        '    Returns
        '    -------
        '    expansibility : float, optional
        '        Expansibility factor (1 for incompressible fluids, less than 1 for
        '        real fluids), [-]
        '
        '    Notes
        '    -----
        '    This formula was determined for the range of P2/P1 >= 0.80, and for fluids
        '    of air, steam, and natural gas. However, there is no objection to using
        '    it for other fluids.
        '
        '    It is said in [1]_ that for liquids this should not be used. The result
        '    can be forced by setting `k` to a really high number like 1E20.
        '
        '    Examples
        '    --------
        '    >>> orifice_expansibility(D=0.0739, Do=0.0222, P1=1E5, P2=9.9E4, k=1.4)
        '    0.9974739057343425
        '
        '    References
        '    ----------
        '    .. [1] American Society of Mechanical Engineers. Mfc-3M-2004 Measurement
        '       Of Fluid Flow In Pipes Using Orifice, Nozzle, And Venturi. ASME, 2001.
        '    .. [2] ISO 5167-2:2003 - Measurement of Fluid Flow by Means of Pressure
        '       Differential Devices Inserted in Circular Cross-Section Conduits Running
        '       Full -- Part 2: Orifice Plates.
        '
        Dim beta, beta2, beta4

        beta = D0 / D
        beta2 = beta * beta
        beta4 = beta2 * beta2
        Orifice_expansibility = (1.0 - (0.351 + beta4 * (0.93 * beta4 + 0.256)) * (1.0 - Math.Pow(P2 / P1, 1 / k)))

    End Function

    Function Orifice_expansibility_1989(D, D0, P1, P2, k) As Double
        '    Calculates the expansibility factor for orifice plate calculations
        '    based on the geometry of the plate, measured pressures of the orifice, and
        '    the isentropic exponent of the fluid.
        '
        '    .. math::
        '        \epsilon = 1- (0.41 + 0.35\beta^4)\Delta P/\kappa/P_1
        '
        '    Parameters
        '    ----------
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    Do : float
        '        Diameter of orifice at flow conditions, [m]
        '    P1 : float
        '        Static pressure of fluid upstream of orifice at the cross-section of
        '        the pressure tap, [Pa]
        '    P2 : float
        '        Static pressure of fluid downstream of orifice at the cross-section of
        '        the pressure tap, [Pa]
        '    k : float
        '        Isentropic exponent of fluid, [-]
        '
        '    Returns
        '    -------
        '    expansibility : float
        '        Expansibility factor (1 for incompressible fluids, less than 1 for
        '        real fluids), [-]
        '
        '    Notes
        '    -----
        '    This formula was determined for the range of P2/P1 >= 0.75, and for fluids
        '    of air, steam, and natural gas. However, there is no objection to using
        '    it for other fluids.
        '
        '    This is an older formula used to calculate expansibility factors for
        '    orifice plates.
        '
        '    In this standard, an expansibility factor formula transformation in terms
        '    of the pressure after the orifice is presented as well. This is the more
        '    standard formulation in terms of the upstream conditions. The other formula
        '    is below for reference only:
        '
        '    .. math::
        '        \epsilon_2 = \sqrt{1 + \frac{\Delta P}{P_2}} -  (0.41 + 0.35\beta^4)
        '        \frac{\Delta P}{\kappa P_2 \sqrt{1 + \frac{\Delta P}{P_2}}}
        '
        '    [2]_ recommends this formulation for wedge meters as well.
        '
        '    Examples
        '    --------
        '    >>> orifice_expansibility_1989(D=0.0739, Do=0.0222, P1=1E5, P2=9.9E4, k=1.4)
        '    0.9970510687411718
        '
        '    References
        '    ----------
        '    .. [1] American Society of Mechanical Engineers. MFC-3M-1989 Measurement
        '       Of Fluid Flow In Pipes Using Orifice, Nozzle, And Venturi. ASME, 2005.
        '    .. [2] Miller, Richard W. Flow Measurement Engineering Handbook. 3rd
        '       edition. New York: McGraw-Hill Education, 1996.
        '
        Dim beta, beta2, beta4

        beta = D0 / D
        beta2 = beta * beta
        beta4 = beta2 * beta2
        Orifice_expansibility_1989 = 1.0 - (0.41 + 0.35 * beta4) * (P1 - P2) / (k * P1)

    End Function

    Public Function C_Reader_Harris_Gallagher(D, D0, rho, mu, m, Optional taps = "D and D/2") As Double
        '    Calculates the coefficient of discharge of the orifice based on the
        '    geometry of the plate, measured pressures of the orifice, mass flow rate
        '    through the orifice, and the density and viscosity of the fluid.
        '
        '    .. math::
        '        C = 0.5961 + 0.0261\beta^2 - 0.216\beta^8 + 0.000521\left(\frac{
        '        10^6\beta}{Re_D}\right)^{0.7}\\
        '        + (0.0188 + 0.0063A)\beta^{3.5} \left(\frac{10^6}{Re_D}\right)^{0.3} \\
        '        +(0.043 + 0.080\exp(-10L_1) -0.123\exp(-7L_1))(1-0.11A)\frac{\beta^4}
        '        {1-\beta^4} \\
        '        -  0.031(M_2' - 0.8M_2'^{1.1})\beta^{1.3}
        '
        '    .. math::
        '        M_2' = \frac{2L_2'}{1-\beta}
        '
        '    .. math::
        '        A = \left(\frac{19000\beta}{Re_{D}}\right)^{0.8}
        '
        '    .. math::
        '        Re_D = \frac{\rho v D}{\mu}
        '
        '
        '    If D < 71.12 mm (2.8 in.) (Note this is a continuous addition; there is no
        '    discontinuity):
        '
        '    .. math::
        '        C += 0.11(0.75-\beta)\left(2.8-\frac{D}{0.0254}\right)
        '
        '    If the orifice has corner taps:
        '
        '    .. math::
        '        L_1 = L_2' = 0
        '
        '    If the orifice has D and D/2 taps:
        '
        '    .. math::
        '        L_1 = 1
        '
        '    .. math::
        '        L_2' = 0.47
        '
        '    If the orifice has Flange taps:
        '
        '    .. math::
        '        L_1 = L_2' = \frac{0.0254}{D}
        '
        '    Parameters
        '    ----------
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    Do : float
        '        Diameter of orifice at flow conditions, [m]
        '    rho : float
        '        Density of fluid at `P1`, [kg/m^3]
        '    mu : float
        '        Viscosity of fluid at `P1`, [Pa*s]
        '    m : float
        '        Mass flow rate of fluid through the orifice, [kg/s]
        '    taps : str
        '        The orientation of the taps; one of 1 = 'corner', 2 = 'flange', 3 = 'D', or 'D/2',
        '        [-]
        '
        '    Returns
        '    -------
        '    C : float
        '        Coefficient of discharge of the orifice, [-]
        '
        '    Notes
        '    -----
        '    The following limits apply to the orifice plate standard [1]_:
        '
        '    The measured pressure difference for the orifice plate should be under
        '    250 kPa.
        '
        '    There are roughness limits as well; the roughness should be under 6
        '    micrometers, although there are many more conditions to that given in [1]_.
        '
        '    For orifice plates with D and D/2 or corner pressure taps:
        '
        '    * Orifice bore diameter muse be larger than 12.5 mm (0.5 inches)
        '    * Pipe diameter between 50 mm and 1 m (2 to 40 inches)
        '    * Beta between 0.1 and 0.75 inclusive
        '    * Reynolds number larger than 5000 (for :math:`0.10 \le \beta \le 0.56`)
        '      or for :math:`\beta \ge 0.56, Re_D \ge 16000\beta^2`
        '
        '    For orifice plates with flange pressure taps:
        '
        '    * Orifice bore diameter muse be larger than 12.5 mm (0.5 inches)
        '    * Pipe diameter between 50 mm and 1 m (2 to 40 inches)
        '    * Beta between 0.1 and 0.75 inclusive
        '    * Reynolds number larger than 5000 and also larger than
        '      :math:`170000\beta^2 D`.
        '
        '    This is also presented in Crane's TP410 (2009) publication, whereas the
        '    1999 and 1982 editions showed only a graph for discharge coefficients.
        '
        '    Examples
        '    --------
        '    >>> C_Reader_Harris_Gallagher(D=0.07391, Do=0.0222, rho=1.165, mu=1.85E-5,
        '    ... m=0.12, taps="flange")
        '    0.5990326277163659
        '
        '    References
        '    ----------
        '    .. [1] American Society of Mechanical Engineers. Mfc-3M-2004 Measurement
        '       Of Fluid Flow In Pipes Using Orifice, Nozzle, And Venturi. ASME, 2001.
        '    .. [2] ISO 5167-2:2003 - Measurement of Fluid Flow by Means of Pressure
        '       Differential Devices Inserted in Circular Cross-Section Conduits Running
        '       Full -- Part 2: Orifice Plates.
        '    .. [3] Reader-Harris, M. J., "The Equation for the Expansibility Factor for
        '       Orifice Plates," Proceedings of FLOMEKO 1998, Lund, Sweden, 1998:
        '       209-214.
        '    .. [4] Reader-Harris, Michael. Orifice Plates and Venturi Tubes. Springer,
        '       2015.

        Dim A_pipe, V, Re_D, Re_D_inv, l1, L2_prime, beta, beta2, beta4, beta8, a, M2_prime
        Dim expnL1, expnL2, expnL3, delta_C_upstream, t1, delta_C_downstream, x1, x2, t2
        Dim C_inf_C_s, c, t3, delta_C_diameter

        A_pipe = 0.25 * PI * D * D
        V = m / (A_pipe * rho)
        Re_D = rho * V * D / mu
        Re_D_inv = 1.0 / Re_D

        beta = D0 / D
        If taps = "corner" Then 'Corner
            l1 = 0
            L2_prime = 0
        ElseIf taps = "flange" Then 'Flange
            l1 = 0.0254 / D
            L2_prime = l1
        Else ' D or D/2
            l1 = 1
            L2_prime = 0.47
        End If

        beta2 = beta * beta
        beta4 = beta2 * beta2
        beta8 = beta4 * beta4

        a = Math.Pow(19000 * beta * Re_D_inv, 0.8)
        M2_prime = 2.0 * L2_prime / (1.0 - beta)

        expnL1 = Math.Exp(-l1)
        expnL2 = expnL1 * expnL1
        expnL3 = expnL1 * expnL2
        delta_C_upstream = ((0.043 + expnL3 * expnL2 * expnL2 * (0.08 * expnL3 - 0.123)) * (1.0 - 0.11 * a) * beta4 / (1.0 - beta4))

        ' The max part is not in the ISO standard
        t1 = Math.Log(3700.0 * Re_D_inv)
        If t1 < 0 Then t1 = 0
        delta_C_downstream = (-0.031 * (M2_prime - 0.8 * Math.Pow(M2_prime, 1.1)) * Math.Pow(beta, 1.3) * (1.0 + 8.0 * t1))

        ' C_inf is discharge coefficient with corner taps for infinite Re
        x1 = 63.0957344480193 * Math.Pow(Re_D_inv, 0.3)
        x2 = 22.7 - 0.0047 * Re_D
        If x1 > x2 Then
            t2 = x1
        Else
            t2 = x2
        End If

        C_inf_C_s = (0.5961 + 0.0261 * beta2 - 0.216 * beta8 + 0.000521 * Math.Pow(1000000.0 * beta * Re_D_inv, 0.7) + (0.0188 + 0.0063 * a) * beta2 * beta * Math.Sqrt(beta) * (t2))
        c = (C_inf_C_s + delta_C_upstream + delta_C_downstream)
        If D < 0.07112 Then
            t3 = (2.8 - D / 0.0254)
            delta_C_diameter = 0.011 * (0.75 - beta) * t3
            c += delta_C_diameter
        End If

        C_Reader_Harris_Gallagher = c
    End Function

    Function C_Miller_1996(D As Double, D0 As Double, rho As Double, mu As Double, m As Double, subtype As String, taps As String, tap_position As String) As Double
        '    Calculates the coefficient of discharge of any of the orifice types
        '    supported by the Miller (1996) [1]_ correlation set. These correlations
        '    cover a wide range of industrial applications and sizes. Most of them are
        '    functions of `beta` ratio and Reynolds number. Unlike the ISO standards,
        '    these correlations do not come with well defined ranges of validity, so
        '    caution should be applied using there correlations.
        '
        '    The base equation is as follows, and each orifice type and range has
        '    different values or correlations for :math:`C_{\infty}`, `b`, and `n`.
        '
        '    .. math::
        '        C = C_{\infty} + \frac{b}{{Re}_D^n}
        '
        '    Parameters
        '    ----------
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    Do : float
        '        Diameter of orifice at flow conditions, [m]
        '    rho : float
        '        Density of fluid at `P1`, [kg/m^3]
        '    mu : float
        '        Viscosity of fluid at `P1`, [Pa*s]
        '    m : float
        '        Mass flow rate of fluid through the orifice, [kg/s]
        '    subtype : str, optional
        '        One of 'orifice', 'eccentric orifice', 'segmental orifice',
        '        'conical orifice', or 'quarter circle orifice', [-]
        '    taps : str, optional
        '        The orientation of the taps; one of 'corner', 'flange',
        '        'D and D/2', 'pipe', or 'vena contracta'; not all orifice subtypes
        '        support the all tap types [-]
        '    tap_position : str, optional
        '        The rotation of the taps, used **only for the eccentric orifice case**
        '        where the pressure profile is are not symmetric; '180 degree' for the
        '        normal case where the taps are opposite the orifice bore, and
        '        '90 degree' for the case where, normally for operational reasons, the
        '        taps are near the bore [-]
        '
        '    Returns
        '    -------
        '    C : float
        '        Coefficient of discharge of the orifice, [-]
        '
        '    Notes
        '    -----
        '    Many of the correlations transition at a pipe diameter of 100 mm to
        '    different equations, which will lead to discontinuous behavior.
        '
        '    It should also be noted the author of these correlations developed a
        '    commercial flow meter rating software package, at [2]_.
        '    He passed away in 2014, but contributed massively to the field of flow
        '    measurement.
        '
        '    The numerous equations for the different cases are as follows:
        '
        '    For all **regular (concentric) orifices**, the `b` equation is as follows
        '    and n = 0.75:
        '
        '    .. math::
        '        b = 91.706\beta^{2.5}
        '
        '    Regular (concentric) orifice, corner taps:
        '
        '    .. math::
        '         C_{\infty} = 0.5959 + 0.0312\beta^2.1 - 0.184\beta^8
        '
        '    Regular (concentric) orifice, flange taps, D > 58.4 mm:
        '
        '    .. math::
        '         C_{\infty} = 0.5959 + 0.0312\beta^{2.1} - 0.184\beta^8
        '         + \frac{2.286\beta^4}{(D_{mm}(1.0 - \beta^4))}
        '         - \frac{0.856\beta^3}{D_{mm}}
        '
        '    Regular (concentric) orifice, flange taps, D < 58.4 mm:
        '
        '    .. math::
        '         C_{\infty} = 0.5959 + 0.0312\beta^{2.1} - 0.184\beta^8
        '         + \frac{0.039\beta^4}{(1.0 - \beta^4)} - \frac{0.856\beta^3}{D_{mm}}
        '
        '    Regular (concentric) orifice, 'D and D/2' taps:
        '
        '    .. math::
        '         C_{\infty} = 0.5959 + 0.0312\beta^{2.1} - 0.184\beta^8
        '         + \frac{0.039\beta^4}{(1.0 - \beta^4)} - 0.01584
        '
        '    Regular (concentric) orifice, 'pipe' taps:
        '
        '    .. math::
        '         C_{\infty} = 0.5959 + 0.461\beta^{2.1} + 0.48\beta^8
        '         + \frac{0.039\beta^4}{(1.0 - \beta^4)}
        '
        '    For the case of a **conical orifice**, there is no tap dependence
        '    and one equation (`b` = 0, `n` = 0):
        '
        '    .. math::
        '         C_{\infty} = 0.734 \text{ if } 250\beta \le Re \le 500\beta \text{ else } 0.730
        '
        '    For the case of a **quarter circle orifice**, corner and flange taps have
        '    the same dependence (`b` = 0, `n` = 0):
        '
        '    .. math::
        '         C_{\infty} = (0.7746 - 0.1334\beta^{2.1} + 1.4098\beta^8
        '                        + \frac{0.0675\beta^4}{(1 - \beta^4)} + 0.3865\beta^3)
        '
        '    For all **segmental orifice** types, `b` = 0 and `n` = 0
        '
        '    Segmental orifice, 'flange' taps, D < 10 cm:
        '
        '    .. math::
        '         C_{\infty} = 0.6284 + 0.1462\beta^{2.1} - 0.8464\beta^8
        '         + \frac{0.2603\beta^4}{(1-\beta^4)} - 0.2886\beta^3
        '
        '    Segmental orifice, 'flange' taps, D > 10 cm:
        '
        '    .. math::
        '         C_{\infty} = 0.6276 + 0.0828\beta^{2.1} + 0.2739\beta^8
        '         - \frac{0.0934\beta^4}{(1-\beta^4)} - 0.1132\beta^3
        '
        '    Segmental orifice, 'vena contracta' taps, D < 10 cm:
        '
        '    .. math::
        '         C_{\infty} = 0.6261 + 0.1851\beta^{2.1} - 0.2879\beta^8
        '         + \frac{0.1170\beta^4}{(1-\beta^4)} - 0.2845\beta^3
        '
        '    Segmental orifice, 'vena contracta' taps, D > 10 cm:
        '
        '    .. math::
        '         C_{\infty} = 0.6276 + 0.0828\beta^{2.1} + 0.2739\beta^8
        '         - \frac{0.0934\beta^4}{(1-\beta^4)} - 0.1132\beta^3
        '
        '    For all **eccentric orifice** types,  `n` = 0.75 and `b` is fit to a
        '    polynomial of `beta`.
        '
        '    Eccentric orifice, 'flange' taps, 180 degree opposite taps, D < 10 cm:
        '
        '    .. math::
        '        C_{\infty} = 0.5917 + 0.3061\beta^{2.1} + .3406\beta^8 -\frac{.1019\beta^4}{(1-\beta^4)} - 0.2715\beta^3
        '
        '    .. math::
        '        b = 7.3 - 15.7\beta + 170.8\beta^2 - 399.7\beta^3 + 332.2\beta^4
        '
        '    Eccentric orifice, 'flange' taps, 180 degree opposite taps, D > 10 cm:
        '
        '    .. math::
        '        C_{\infty} = 0.6016 + 0.3312\beta^{2.1} -1.5581\beta^8 + \frac{0.6510\beta^4}{(1-\beta^4)} - 0.7308\beta^3
        '
        '    .. math::
        '        b = -139.7 + 1328.8\beta - 4228.2\beta^2 + 5691.9\beta^3 - 2710.4\beta^4
        '
        '    Eccentric orifice, 'flange' taps, 90 degree side taps, D < 10 cm:
        '
        '    .. math::
        '        C_{\infty} = 0.5866 + 0.3917\beta^{2.1} + .7586\beta^8 - \frac{.2273\beta^4}{(1-\beta^4)} - .3343\beta^3
        '
        '    .. math::
        '        b = 69.1 - 469.4\beta + 1245.6\beta^2 -1287.5\beta^3 + 486.2\beta^4
        '
        '    Eccentric orifice, 'flange' taps, 90 degree side taps, D > 10 cm:
        '
        '    .. math::
        '        C_{\infty} = 0.6037 + 0.1598\beta^{2.1} -.2918\beta^8 + \frac{0.0244\beta^4}{(1-\beta^4)} - 0.0790\beta^3
        '
        '    .. math::
        '        b = -103.2 + 898.3\beta - 2557.3\beta^2 + 2977.0\beta^3 - 1131.3\beta^4
        '
        '    Eccentric orifice, 'vena contracta' taps, 180 degree opposite taps, D < 10 cm:
        '
        '    .. math::
        '        C_{\infty} = 0.5925 + 0.3380\beta^{2.1} + 0.4016\beta^8 - \frac{.1046\beta^4}{(1-\beta^4)} - 0.3212\beta^3
        '
        '    .. math::
        '        b = 23.3 -207.0\beta + 821.5\beta^2 -1388.6\beta^3 + 900.3\beta^4
        '
        '    Eccentric orifice, 'vena contracta' taps, 180 degree opposite taps, D > 10 cm:
        '
        '    .. math::
        '        C_{\infty} = 0.5922 + 0.3932\beta^{2.1} + .3412\beta^8 - \frac{.0569\beta^4}{(1-\beta^4)} - 0.4628\beta^3
        '
        '    .. math::
        '        b = 55.7 - 471.4\beta + 1721.8\beta^2 - 2722.6\beta^3 + 1569.4\beta^4
        '
        '    Eccentric orifice, 'vena contracta' taps, 90 degree side taps, D < 10 cm:
        '
        '    .. math::
        '        C_{\infty} = 0.5875 + 0.3813\beta^{2.1} + 0.6898\beta^8 - \frac{0.1963\beta^4}{(1-\beta^4)} - 0.3366\beta^3
        '
        '    .. math::
        '        b = -69.3 + 556.9\beta - 1332.2\beta^2 + 1303.7\beta^3 - 394.8\beta^4
        '
        '    Eccentric orifice, 'vena contracta' taps, 90 degree side taps, D > 10 cm:
        '
        '    .. math::
        '        C_{\infty} = 0.5949 + 0.4078\beta^{2.1} + 0.0547\beta^8 + \frac{0.0955\beta^4}{(1-\beta^4)} - 0.5608\beta^3
        '
        '    .. math::
        '        b = 52.8 - 434.2\beta + 1571.2\beta^2 - 2460.9\beta^3 + 1420.2\beta^4
        '
        '
        '    Examples
        '    --------
        '    >>> C_Miller_1996(D=0.07391, Do=0.0222, rho=1.165, mu=1.85E-5, m=0.12, taps='flange', subtype='orifice')
        '    0.599065557156788
        '
        '    References
        '    ----------
        '    .. [1] Miller, Richard W. Flow Measurement Engineering Handbook.
        '       McGraw-Hill Education, 1996.
        '    .. [2] "RW Miller & Associates." Accessed April 13, 2020.
        '       http://rwmillerassociates.com/.
        '
        Dim A_pipe, V, Re, D_mm, beta, beta2, beta3, beta4, beta8, beta21
        Dim b, n, C_inf, c

        C_inf = 0.0
        b = 0

        A_pipe = 0.25 * PI * D * D
        V = m / (A_pipe * rho)
        Re = rho * V * D / mu
        D_mm = D * 1000.0

        beta = D0 / D
        beta2 = beta * beta
        beta3 = beta * beta * beta
        beta4 = beta * beta3
        beta8 = beta4 * beta4
        beta21 = Math.Pow(beta, 2.1)

        If IsInArray(subtype, {MILLER_ORIFICE, CONCENTRIC_ORIFICE}) Then
            b = 91.706 * Math.Pow(beta, 2.5)
            n = 0.75
            If taps = ORIFICE_CORNER_TAPS Then
                C_inf = 0.5959 + 0.0312 * beta21 - 0.184 * beta8
            ElseIf taps = ORIFICE_FLANGE_TAPS Then
                If D_mm >= 58.4 Then
                    C_inf = 0.5959 + 0.0312 * beta21 - 0.184 * beta8 + 2.286 * beta4 / (D_mm * (1.0 - beta4)) - 0.856 * beta3 / D_mm
                Else
                    C_inf = 0.5959 + 0.0312 * beta21 - 0.184 * beta8 + 0.039 * beta4 / (1.0 - beta4) - 0.856 * beta3 / D_mm
                End If
            ElseIf taps = ORIFICE_D_AND_D_2_TAPS Then
                C_inf = 0.5959 + 0.0312 * beta21 - 0.184 * beta8 + 0.039 * beta4 / (1.0 - beta4) - 0.01584
            ElseIf taps = ORIFICE_PIPE_TAPS Then
                C_inf = 0.5959 + 0.461 * beta21 + 0.48 * beta8 + 0.039 * beta4 / (1.0 - beta4)
            End If
        ElseIf IsInArray(subtype, {MILLER_ECCENTRIC_ORIFICE, ECCENTRIC_ORIFICE}) Then
            If tap_position <> TAPS_OPPOSITE And tap_position <> TAPS_SIDE Then
                'MsgBox "_Miller_1996_unsupported_tap_pos_eccentric"
                C_Miller_1996 = 0.0

                Exit Function
            End If
            n = 0.75
            If taps = ORIFICE_FLANGE_TAPS Then
                If tap_position = TAPS_OPPOSITE Then
                    If D < 0.1 Then
                        b = 7.3 - 15.7 * beta + 170.8 * beta * beta - 399.7 * beta3 + 332.2 * beta4
                        C_inf = 0.5917 + 0.3061 * beta21 + 0.3406 * beta8 - 0.1019 * beta4 / (1 - beta4) - 0.2715 * beta3
                    Else
                        b = -139.7 + 1328.8 * beta - 4228.2 * beta * beta + 5691.9 * beta3 - 2710.4 * beta4
                        C_inf = 0.6016 + 0.3312 * beta21 - 1.5581 * beta8 + 0.651 * beta4 / (1 - beta4) - 0.7308 * beta3
                    End If
                ElseIf tap_position = TAPS_SIDE Then
                    If D < 0.1 Then
                        b = 69.1 - 469.4 * beta + 1245.6 * beta2 - 1287.5 * beta3 + 486.2 * beta4
                        C_inf = 0.5866 + 0.3917 * beta21 + 0.7586 * beta8 - 0.2273 * beta4 / (1 - beta4) - 0.3343 * beta3
                    Else
                        b = -103.2 + 898.3 * beta - 2557.3 * beta2 + 2977.0 * beta3 - 1131.3 * beta4
                        C_inf = 0.6037 + 0.1598 * beta21 - 0.2918 * beta8 + 0.0244 * beta4 / (1 - beta4) - 0.079 * beta3
                    End If
                End If
            ElseIf taps = ORIFICE_VENA_CONTRACTA_TAPS Then
                If tap_position = TAPS_OPPOSITE Then
                    If D < 0.1 Then
                        b = 23.3 - 207.0 * beta + 821.5 * beta2 - 1388.6 * beta3 + 900.3 * beta4
                        C_inf = 0.5925 + 0.338 * beta21 + 0.4016 * beta8 - 0.1046 * beta4 / (1 - beta4) - 0.3212 * beta3
                    Else
                        b = 55.7 - 471.4 * beta + 1721.8 * beta2 - 2722.6 * beta3 + 1569.4 * beta4
                        C_inf = 0.5922 + 0.3932 * beta21 + 0.3412 * beta8 - 0.0569 * beta4 / (1 - beta4) - 0.4628 * beta3
                    End If
                ElseIf tap_position = TAPS_SIDE Then
                    If D < 0.1 Then
                        b = -69.3 + 556.9 * beta - 1332.2 * beta2 + 1303.7 * beta3 - 394.8 * beta4
                        C_inf = 0.5875 + 0.3813 * beta21 + 0.6898 * beta8 - 0.1963 * beta4 / (1 - beta4) - 0.3366 * beta3
                    Else
                        b = 52.8 - 434.2 * beta + 1571.2 * beta2 - 2460.9 * beta3 + 1420.2 * beta4
                        C_inf = 0.5949 + 0.4078 * beta21 + 0.0547 * beta8 + 0.0955 * beta4 / (1 - beta4) - 0.5608 * beta3
                    End If
                End If
            Else
                'MsgBox "Only support for FLANGE and VENA CONTRACTA taps"
                C_Miller_1996 = 0.0

                Exit Function
            End If
        ElseIf IsInArray(subtype, {MILLER_SEGMENTAL_ORIFICE, SEGMENTAL_ORIFICE}) Then
            n = 0
            b = 0
            If taps = ORIFICE_FLANGE_TAPS Then
                If D < 0.1 Then
                    C_inf = 0.6284 + 0.1462 * beta21 - 0.8464 * beta8 + 0.2603 * beta4 / (1 - beta4) - 0.2886 * beta3
                Else
                    C_inf = 0.6276 + 0.0828 * beta21 + 0.2739 * beta8 - 0.0934 * beta4 / (1 - beta4) - 0.1132 * beta3
                End If
            ElseIf taps = ORIFICE_VENA_CONTRACTA_TAPS Then
                If D < 0.1 Then
                    C_inf = 0.6261 + 0.1851 * beta21 - 0.2879 * beta8 + 0.117 * beta4 / (1 - beta4) - 0.2845 * beta3
                Else
                    ' Yes these are supposed to be the same as the flange, large set
                    C_inf = 0.6276 + 0.0828 * beta21 + 0.2739 * beta8 - 0.0934 * beta4 / (1 - beta4) - 0.1132 * beta3
                End If
            Else
                'MsgBox "Only support for FLANGE and VENA CONTRACTA taps"
                C_Miller_1996 = 0.0

                Exit Function
            End If
        ElseIf IsInArray(subtype, {MILLER_CONICAL_ORIFICE, CONICAL_ORIFICE}) Then
            n = 0
            b = 0
            If 250.0 * beta <= Re And Re <= 500.0 * beta Then
                C_inf = 0.734
            Else
                C_inf = 0.73
            End If
        ElseIf IsInArray(subtype, {MILLER_QUARTER_CIRCLE_ORIFICE, QUARTER_CIRCLE_ORIFICE}) Then
            n = 0
            b = 0
            C_inf = (0.7746 - 0.1334 * beta21 + 1.4098 * beta8 + 0.0675 * beta4 / (1.0 - beta4) + 0.3865 * beta3)
        Else
            'MsgBox "_Miller_1996_unsupported_type"
            C_Miller_1996 = 0.0

            Exit Function
        End If

        c = C_inf + b * Math.Pow(Re, -n)
        C_Miller_1996 = c

    End Function

    Function C_eccentric_orifice_ISO_15377_1998(D, D0) As Double
        '    Calculates the coefficient of discharge of an eccentric orifice based
        '    on the geometry of the plate according to ISO 15377, first introduced in
        '    1998 and also presented in the second 2007 edition. It also appears in BS
        '    1042-1.2: 1989.
        '
        '    .. math::
        '        C = 0.9355 - 1.6889\beta + 3.0428\beta^2 - 1.7989\beta^3
        '
        '    This type of plate is normally used to avoid obstructing entrained gas,
        '    liquid, or sediment.
        '
        '    Parameters
        '    ----------
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    Do : float
        '        Diameter of orifice at flow conditions, [m]
        '
        '    Returns
        '    -------
        '    C : float
        '        Coefficient of discharge of the eccentric orifice, [-]
        '
        '    Notes
        '    -----
        '    No correction for where the orifice bore is located is included.
        '
        '    The following limits apply to the orifice plate standard [1]_:
        '
        '    * Bore diameter above 50 mm.
        '    * Pipe diameter between 10 cm and 1 m.
        '    * Beta ratio between 0.46 and 0.84
        '    * :math:`2\times 10^5 \beta^2 \le Re_D \le 10^6 \beta`
        '
        '    The uncertainty of this equation for `C` is said to be 1% if `beta` is
        '    under 0.75, otherwise 2%.
        '
        '    The `orifice_expansibility` function should be used with this method as
        '    well.
        '
        '    Additional specifications are:
        '
        '    * The thickness of the orifice should be between 0.005`D` and 0.02`D`.
        '    * Corner tappings should be used, with hole diameter between 3 and 10 mm.
        '      The angular orientation of the tappings matters because the flow meter
        '      is not symmetrical. The angle should ideally be at the top or bottom of
        '      the plate, opposite which side the bore is on - but this can cause
        '      issues with deposition if the taps are on the bottom or gas bubbles if
        '      the taps are on the taps. The taps are often placed 30 degrees away from
        '      the ideal position to counteract this effect, with under an extra 2%
        '      error.
        '
        '    Some comparisons with CFD results can be found in [2]_.
        '
        '    Examples
        '    --------
        '    >>> C_eccentric_orifice_ISO_15377_1998(.2, .075)
        '    0.6351923828125
        '
        '    References
        '    ----------
        '    .. [1] TC 30/SC 2, ISO. ISO/TR 15377:1998, Measurement of Fluid Flow by
        '       Means of Pressure-Differential Devices - Guide for the Specification of
        '       Nozzles and Orifice Plates beyond the Scope of ISO 5167-1.
        '    .. [2] Yashvanth, S., Varadarajan Seshadri, and J. YogeshKumarK. "CFD
        '       Analysis of Flow through Single and Multi Stage Eccentric Orifice Plate
        '       Assemblies," 2017.

        Dim beta

        beta = D0 / D
        C_eccentric_orifice_ISO_15377_1998 = beta * (beta * (3.0428 - 1.7989 * beta) - 1.6889) + 0.9355

    End Function

    Function C_quarter_circle_orifice_ISO_15377_1998(D, D0) As Double
        '    Calculates the coefficient of discharge of a quarter circle orifice based
        '    on the geometry of the plate according to ISO 15377, first introduced in
        '    1998 and also presented in the second 2007 edition. It also appears in BS
        '    1042-1.2: 1989.
        '
        '    .. math::
        '        C = 0.73823 + 0.3309\beta - 1.1615\beta^2 + 1.5084\beta^3
        '
        '    Parameters
        '    ----------
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    Do : float
        '        Diameter of orifice at flow conditions, [m]
        '
        '    Returns
        '    -------
        '    C : float
        '        Coefficient of discharge of the quarter circle orifice, [-]
        '
        '    Notes
        '    -----
        '    The discharge coefficient of this type of orifice plate remains
        '    constant down to a lower than normal `Re`, as occurs in highly
        '    viscous applications.
        '
        '    The following limits apply to the orifice plate standard [1]_:
        '
        '    * Bore diameter >= 1.5 cm
        '    * Pipe diameter <= 50 cm
        '    * Beta ratio between 0.245 and 0.6
        '    * :math:`Re_d \le 10^5 \beta`
        '
        '    There is also a table in [1]_ which lists increased minimum
        '    upstream pipe diameters for pipes of different roughnesses; the
        '    higher the roughness, the larger the pipe diameter required,
        '    and the table goes up to 20 cm for rusty cast iron.
        '
        '    Corner taps should be used up to pipe diameters of 40 mm;
        '    for larger pipes, corner or flange taps can be used. No impact
        '    on the flow coefficient is included in the correlation.
        '
        '    The recommended expansibility method for this type of orifice is
        '    :obj:`orifice_expansibility`.
        '
        '    Examples
        '    --------
        '    >>> C_quarter_circle_orifice_ISO_15377_1998(.2, .075)
        '    0.77851484375000
        '
        '    References
        '    ----------
        '    .. [1] TC 30/SC 2, ISO. ISO/TR 15377:1998, Measurement of Fluid Flow by
        '       Means of Pressure-Differential Devices - Guide for the Specification of
        '       Nozzles and Orifice Plates beyond the Scope of ISO 5167-1.
        Dim beta

        beta = D0 / D
        C_quarter_circle_orifice_ISO_15377_1998 = beta * (beta * (1.5084 * beta - 1.16158) + 0.3309) + 0.73823

    End Function

    Function Discharge_coefficient_to_K(D, D0, c) As Double
        '    Converts a discharge coefficient to a standard loss coefficient,
        '    for use in computation of the actual pressure drop of an orifice or other
        '    device.
        '
        '    .. math::
        '        K = \left[\frac{\sqrt{1-\beta^4(1-C^2)}}{C\beta^2} - 1\right]^2
        '
        '    Parameters
        '    ----------
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    Do : float
        '        Diameter of orifice at flow conditions, [m]
        '    C : float
        '        Coefficient of discharge of the orifice, [-]
        '
        '    Returns
        '    -------
        '    K : float
        '        Loss coefficient with respect to the velocity and density of the fluid
        '        just upstream of the orifice, [-]
        '
        '    Notes
        '    -----
        '    If expansibility is used in the orifice calculation, the result will not
        '    match with the specified pressure drop formula in [1]_; it can almost
        '    be matched by dividing the calculated mass flow by the expansibility factor
        '    and using that mass flow with the loss coefficient.
        '
        '    Examples
        '    --------
        '    >>> discharge_coefficient_to_K(D=0.07366, Do=0.05, C=0.61512)
        '    5.2314291729754
        '
        '    References
        '    ----------
        '    .. [1] American Society of Mechanical Engineers. Mfc-3M-2004 Measurement
        '       Of Fluid Flow In Pipes Using Orifice, Nozzle, And Venturi. ASME, 2001.
        '    .. [2] ISO 5167-2:2003 - Measurement of Fluid Flow by Means of Pressure
        '       Differential Devices Inserted in Circular Cross-Section Conduits Running
        '       Full -- Part 2: Orifice Plates.
        Dim beta, beta2, beta4, root_K

        beta = D0 / D
        beta2 = beta * beta
        beta4 = beta2 * beta2
        root_K = (Math.Sqrt(1.0 - beta4 * (1.0 - c * c)) / (c * beta2) - 1.0)
        Discharge_coefficient_to_K = root_K * root_K

    End Function

    Function K_to_discharge_coefficient(D, D0, k) As Double
        '    Converts a standard loss coefficient to a discharge coefficient.
        '
        '    .. math::
        '        C = \sqrt{\frac{1}{2 \sqrt{K} \beta^{4} + K \beta^{4}}
        '        - \frac{\beta^{4}}{2 \sqrt{K} \beta^{4} + K \beta^{4}} }
        '
        '    Parameters
        '    ----------
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    Do : float
        '        Diameter of orifice at flow conditions, [m]
        '    K : float
        '        Loss coefficient with respect to the velocity and density of the fluid
        '        just upstream of the orifice, [-]
        '
        '    Returns
        '    -------
        '    C : float
        '        Coefficient of discharge of the orifice, [-]
        '
        '    Notes
        '    -----
        '    If expansibility is used in the orifice calculation, the result will not
        '    match with the specified pressure drop formula in [1]_; it can almost
        '    be matched by dividing the calculated mass flow by the expansibility factor
        '    and using that mass flow with the loss coefficient.
        '
        '    This expression was derived with SymPy, and checked numerically. There were
        '    three other, incorrect roots.
        '
        '    Examples
        '    --------
        '    >>> K_to_discharge_coefficient(D=0.07366, Do=0.05, K=5.2314291729754)
        '    0.6151200000000001
        '
        '    References
        '    ----------
        '    .. [1] American Society of Mechanical Engineers. Mfc-3M-2004 Measurement
        '       Of Fluid Flow In Pipes Using Orifice, Nozzle, And Venturi. ASME, 2001.
        '    .. [2] ISO 5167-2:2003 - Measurement of Fluid Flow by Means of Pressure
        '       Differential Devices Inserted in Circular Cross-Section Conduits Running
        '       Full -- Part 2: Orifice Plates.

        Dim beta, beta2, beta4, root_K

        beta = D0 / D
        beta2 = beta * beta
        beta4 = beta2 * beta2
        root_K = Math.Sqrt(k)
        K_to_discharge_coefficient = Math.Sqrt((1.0 - beta4) / ((2.0 * root_K + k) * beta4))

    End Function

    Function DP_orifice(D, D0, P1, P2, c) As Double
        '    Calculates the non-recoverable pressure drop of an orifice plate based
        '    on the pressure drop and the geometry of the plate and the discharge
        '    coefficient.
        '
        '    .. math::
        '        \Delta\bar w = \frac{\sqrt{1-\beta^4(1-C^2)}-C\beta^2}
        '        {\sqrt{1-\beta^4(1-C^2)}+C\beta^2} (P_1 - P_2)
        '
        '    Parameters
        '    ----------
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    Do : float
        '        Diameter of orifice at flow conditions, [m]
        '    P1 : float
        '        Static pressure of fluid upstream of orifice at the cross-section of
        '        the pressure tap, [Pa]
        '    P2 : float
        '        Static pressure of fluid downstream of orifice at the cross-section of
        '        the pressure tap, [Pa]
        '    C : float
        '        Coefficient of discharge of the orifice, [-]
        '
        '    Returns
        '    -------
        '    dP : float
        '        Non-recoverable pressure drop of the orifice plate, [Pa]
        '
        '    Notes
        '    -----
        '    This formula can be well approximated by:
        '
        '    .. math::
        '        \Delta\bar w = \left(1 - \beta^{1.9}\right)(P_1 - P_2)
        '
        '    The recoverable pressure drop should be recovered by 6 pipe diameters
        '    downstream of the orifice plate.
        '
        '    Examples
        '    --------
        '    >>> DP_orifice(0.07366, 0.05, 200000.0, 183000.0, 0.61512)
        '    9069.474705745388
        '
        '    References
        '    ----------
        '    .. [1] American Society of Mechanical Engineers. Mfc-3M-2004 Measurement
        '       Of Fluid Flow In Pipes Using Orifice, Nozzle, And Venturi. ASME, 2001.
        '    .. [2] ISO 5167-2:2003 - Measurement of Fluid Flow by Means of Pressure
        '       Differential Devices Inserted in Circular Cross-Section Conduits Running
        '       Full -- Part 2: Orifice Plates.
        Dim beta, beta2, beta4, dP, delta_w

        beta = D0 / D
        beta2 = beta * beta
        beta4 = beta2 * beta2
        dP = P1 - P2
        delta_w = (Math.Sqrt(1.0 - beta4 * (1.0 - c * c)) - c * beta2) / (Math.Sqrt(1.0 - beta4 * (1.0 - c * c)) + c * beta2) * dP
        DP_orifice = delta_w

    End Function

    Function Velocity_of_approach_factor(D, D0) As Double
        '    Calculates a factor for orifice plate design called the `velocity of
        '    approach`.
        '
        '    .. math::
        '        \text{Velocity of approach} = \frac{1}{\sqrt{1 - \beta^4}}
        '
        '    Parameters
        '    ----------
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    Do : float
        '        Diameter of orifice at flow conditions, [m]
        '
        '    Returns
        '    -------
        '    velocity_of_approach : float
        '        Coefficient of discharge of the orifice, [-]
        '
        '    Notes
        '    -----
        '
        '    Examples
        '    --------
        '    >>> velocity_of_approach_factor(0.0739, 0.0222)
        '    1.0040970074165514
        '
        '    References
        '    ----------
        '    .. [1] American Society of Mechanical Engineers. Mfc-3M-2004 Measurement
        '       Of Fluid Flow In Pipes Using Orifice, Nozzle, And Venturi. ASME, 2001.

        Velocity_of_approach_factor = 1.0 / Math.Sqrt(1.0 - Math.Pow(D0 / D, 4))

    End Function

    Function Flow_coefficient(D, D0, c) As Double
        '    Calculates a factor for differential pressure flow meter design called
        '    the `flow coefficient`. This should not be confused with the flow
        '    coefficient often used when discussing valves.
        '
        '    .. math::
        '        \text{Flow coefficient} = \frac{C}{\sqrt{1 - \beta^4}}
        '
        '    Parameters
        '    ----------
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    Do : float
        '        Diameter of flow meter characteristic dimension at flow conditions, [m]
        '    C : float
        '        Coefficient of discharge of the flow meter, [-]
        '
        '    Returns
        '    -------
        '    flow_coefficient : float
        '        Differential pressure flow meter flow coefficient, [-]
        '
        '    Notes
        '    -----
        '    This measure is used not just for orifices but for other differential
        '    pressure flow meters [2]_.
        '
        '    It is sometimes given the symbol K. It is also equal to the product of the
        '    diacharge coefficient and the velocity of approach factor [2]_.
        '
        '    Examples
        '    --------
        '    >>> flow_coefficient(0.0739, 0.0222, 0.6)
        '    0.6024582044499308
        '
        '    References
        '    ----------
        '    .. [1] American Society of Mechanical Engineers. Mfc-3M-2004 Measurement
        '       Of Fluid Flow In Pipes Using Orifice, Nozzle, And Venturi. ASME, 2001.
        '    .. [2] Miller, Richard W. Flow Measurement Engineering Handbook. 3rd
        '       edition. New York: McGraw-Hill Education, 1996.

        Flow_coefficient = c * 1.0 / Math.Sqrt(1.0 - Math.Pow(D0 / D, 4))

    End Function


    Function Nozzle_expansibility(D, D0, P1, P2, k, Optional beta = 0) As Double
        '    Calculates the expansibility factor for a nozzle or venturi nozzle,
        '    based on the geometry of the plate, measured pressures of the orifice, and
        '    the isentropic exponent of the fluid.
        '
        '    .. math::
        '        \epsilon = \left\{\left(\frac{\kappa \tau^{2/\kappa}}{\kappa-1}\right)
        '        \left(\frac{1 - \beta^4}{1 - \beta^4 \tau^{2/\kappa}}\right)
        '        \left[\frac{1 - \tau^{(\kappa-1)/\kappa}}{1 - \tau}
        '        \right] \right\}^{0.5}
        '
        '    Parameters
        '    ----------
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    Do : float
        '        Diameter of orifice of the venturi or nozzle, [m]
        '    P1 : float
        '        Static pressure of fluid upstream of orifice at the cross-section of
        '        the pressure tap, [Pa]
        '    P2 : float
        '        Static pressure of fluid downstream of orifice at the cross-section of
        '        the pressure tap, [Pa]
        '    k : float
        '        Isentropic exponent of fluid, [-]
        '    beta : float, optional
        '        Optional `beta` ratio, which is useful to specify for wedge meters or
        '        flow meters which have a different beta ratio calculation, [-]
        '
        '    Returns
        '    -------
        '    expansibility : float
        '        Expansibility factor (1 for incompressible fluids, less than 1 for
        '        real fluids), [-]
        '
        '    Notes
        '    -----
        '    This formula was determined for the range of P2/P1 >= 0.75.
        '
        '    Mathematically the equation cannot be evaluated at `k` = 1, but if the
        '    limit of the equation is taken the following equation is obtained and is
        '    implemented:
        '
        '
        '    .. math::
        '        \epsilon = \sqrt{\frac{- D^{4} P_{1} P_{2}^{2} \log{\left(\frac{P_{2}}
        '        {P_{1}} \right)} + Do^{4} P_{1} P_{2}^{2} \log{\left(\frac{P_{2}}{P_{1}}
        '        \right)}}{D^{4} P_{1}^{3} - D^{4} P_{1}^{2} P_{2} - Do^{4} P_{1}
        '        P_{2}^{2} + Do^{4} P_{2}^{3}}}
        '
        '    Note also there is a small amount of floating-point error around the range
        '    of `k` ~1+1e-5 to ~1-1e-5, starting with 1e-7 and increasing to the point
        '    of giving values larger than 1 or zero in the  `k` ~1+1e-12 to ~1-1e-12
        '    range.
        '
        '    Examples
        '    --------
        '    >>> nozzle_expansibility(0.0739, 0.0222, 1E5, 9.9E4, 1.4, 0)
        '    0.994570234456
        '
        '    References
        '    ----------
        '    .. [1] American Society of Mechanical Engineers. Mfc-3M-2004 Measurement
        '       Of Fluid Flow In Pipes Using Orifice, Nozzle, And Venturi. ASME, 2001.
        '    .. [2] ISO 5167-3:2003 - Measurement of Fluid Flow by Means of Pressure
        '       Differential Devices Inserted in Circular Cross-Section Conduits Running
        '       Full -- Part 3: Nozzles and Venturi Nozzles.
        Dim beta2, beta4, tau, limit_val, term1, term2, term3

        If beta = 0 Then beta = D0 / D
        beta2 = beta * beta
        beta4 = beta2 * beta2
        tau = P2 / P1
        If k = 1.0 Then
            'Avoid a zero division error:
            'from sympy import *
            'D, Do, P1, P2, k = symbols('D, Do, P1, P2, k')
            'beta = Do/D
            'tau = P2/P1
            'term1 = k*tau**(2/k )/(k - 1)
            'term2 = (1 - beta**4)/(1 - beta**4*tau**(2/k))
            'term3 = (1 - tau**((k - 1)/k))/(1 - tau)
            'val= Math.Sqrt(term1*term2*term3)
            'print(simplify(limit((term1*term2*term3), k, 1)))

            limit_val = (P1 * P2 * P2 * (-D * D * D * D + D0 * D0 * D0 * D0) * Math.Log(P2 / P1) / (D * D * D * D * P1 * P1 * P1 - D * D * D * D * P1 * P1 * P2 - D0 * D0 * D0 * D0 * P1 * P2 * P2 + D0 * D0 * D0 * D0 * P2 * P2 * P2))
            Nozzle_expansibility = Math.Sqrt(limit_val)
            Exit Function
        End If

        term1 = k * Math.Pow(tau, (2.0 / k)) / (k - 1.0)
        term2 = (1.0 - beta4) / (1.0 - beta4 * Math.Pow(tau, (2.0 / k)))
        If tau = 1.0 Then
            '"""Avoid a zero division error.
            'Obtained with:
            '    from sympy import *
            '    tau, k = symbols('tau, k')
            '    expr = (1 - tau**((k - 1)/k))/(1 - tau)
            '    limit(expr, tau, 1)
            '"""
            term3 = (k - 1.0) / k
        Else
            ' This form of the equation is mathematically equivalent but
            ' does not have issues where k = `.
            term3 = (P1 - P2 * Math.Pow(tau, (-1.0 / k))) / (P1 - P2)
            ' term3 = (1.0 - tau**((k - 1.0)/k))/(1.0 - tau)
        End If
        Nozzle_expansibility = Math.Sqrt(term1 * term2 * term3)

    End Function


    Function C_long_radius_nozzle(D, D0, rho, mu, m) As Double
        '    Calculates the coefficient of discharge of a long radius nozzle used
        '    for measuring flow rate of fluid, based on the geometry of the nozzle,
        '    mass flow rate through the nozzle, and the density and viscosity of the
        '    fluid.
        '
        '    .. math::
        '        C = 0.9965 - 0.00653\beta^{0.5} \left(\frac{10^6}{Re_D}\right)^{0.5}
        '
        '    Parameters
        '    ----------
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    D0 : float
        '        Diameter of long radius nozzle orifice at flow conditions, [m]
        '    rho : float
        '        Density of fluid at `P1`, [kg/m^3]
        '    mu : float
        '        Viscosity of fluid at `P1`, [Pa*s]
        '    m : float
        '        Mass flow rate of fluid through the nozzle, [kg/s]
        '
        '    Returns
        '    -------
        '    C : float
        '        Coefficient of discharge of the long radius nozzle orifice, [-]
        '
        '    Notes
        '    -----
        '
        '    Examples
        '    --------
        '    >>> C_long_radius_nozzle(0.07391, 0.0422, 1.2, 1.8E-5, 0.1)
        '    0.9805503704679863
        '
        '    References
        '    ----------
        '    .. [1] American Society of Mechanical Engineers. Mfc-3M-2004 Measurement
        '       Of Fluid Flow In Pipes Using Orifice, Nozzle, And Venturi. ASME, 2001.
        '    .. [2] ISO 5167-3:2003 - Measurement of Fluid Flow by Means of Pressure
        '       Differential Devices Inserted in Circular Cross-Section Conduits Running
        '       Full -- Part 3: Nozzles and Venturi Nozzles.
        Dim A_pipe, V, Re_D, beta

        A_pipe = PI / 4.0 * D * D
        V = m / (A_pipe * rho)
        Re_D = rho * V * D / mu
        beta = D0 / D
        C_long_radius_nozzle = 0.9965 - 0.00653 * Math.Sqrt(beta) * Math.Sqrt(1000000.0 / Re_D)

    End Function

    Function C_ISA_1932_nozzle(D, D0, rho, mu, m) As Double
        '    Calculates the coefficient of discharge of an ISA 1932 style nozzle
        '    used for measuring flow rate of fluid, based on the geometry of the nozzle,
        '    mass flow rate through the nozzle, and the density and viscosity of the
        '    fluid.
        '
        '    .. math::
        '        C = 0.9900 - 0.2262\beta^{4.1} - (0.00175\beta^2 - 0.0033\beta^{4.15})
        '        \left(\frac{10^6}{Re_D}\right)^{1.15}
        '
        '    Parameters
        '    ----------
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    D0 : float
        '        Diameter of nozzle orifice at flow conditions, [m]
        '    rho : float
        '        Density of fluid at `P1`, [kg/m^3]
        '    mu : float
        '        Viscosity of fluid at `P1`, [Pa*s]
        '    m : float
        '        Mass flow rate of fluid through the nozzle, [kg/s]
        '
        '    Returns
        '    -------
        '    C : float
        '        Coefficient of discharge of the nozzle orifice, [-]
        '
        '    Notes
        '    -----
        '
        '    Examples
        '    --------
        '    >>> C_ISA_1932_nozzle(0.07391, 0.0422, 1.2, 1.8E-5, 0.1)
        '    0.9635849973250495
        '
        '    References
        '    ----------
        '    .. [1] American Society of Mechanical Engineers. Mfc-3M-2004 Measurement
        '       Of Fluid Flow In Pipes Using Orifice, Nozzle, And Venturi. ASME, 2001.
        '    .. [2] ISO 5167-3:2003 - Measurement of Fluid Flow by Means of Pressure
        '       Differential Devices Inserted in Circular Cross-Section Conduits Running
        '       Full -- Part 3: Nozzles and Venturi Nozzles.
        Dim A_pipe, V, Re_D, beta

        A_pipe = PI / 4.0 * D * D
        V = m / (A_pipe * rho)
        Re_D = rho * V * D / mu
        beta = D0 / D
        C_ISA_1932_nozzle = 0.99 - 0.2262 * Math.Pow(beta, 4.1) - (0.00175 * beta * beta - 0.0033 * Math.Pow(beta, 4.15)) * Math.Pow(1000000.0 / Re_D, 1.15)

    End Function

    Function C_venturi_nozzle(D, D0) As Double
        '    Calculates the coefficient of discharge of an Venturi style nozzle
        '    used for measuring flow rate of fluid, based on the geometry of the nozzle.
        '
        '    .. math::
        '        C = 0.9858 - 0.196\beta^{4.5}
        '
        '    Parameters
        '    ----------
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    D0 : float
        '        Diameter of nozzle orifice at flow conditions, [m]
        '
        '    Returns
        '    -------
        '    C : float
        '        Coefficient of discharge of the nozzle orifice, [-]
        '
        '    Notes
        '    -----
        '
        '    Examples
        '    --------
        '    >>> C_venturi_nozzle(0.07391, 0.0422)
        '    0.9698996454169576
        '
        '    References
        '    ----------
        '    .. [1] American Society of Mechanical Engineers. Mfc-3M-2004 Measurement
        '       Of Fluid Flow In Pipes Using Orifice, Nozzle, And Venturi. ASME, 2001.
        '    .. [2] ISO 5167-3:2003 - Measurement of Fluid Flow by Means of Pressure
        '       Differential Devices Inserted in Circular Cross-Section Conduits Running
        '       Full -- Part 3: Nozzles and Venturi Nozzles.

        C_venturi_nozzle = 0.9858 - 0.198 * Math.Pow(D0 / D, 4.5)

    End Function

    Function DP_venturi_tube(D, D0, P1, P2) As Double
        '    Calculates the non-recoverable pressure drop of a venturi tube
        '    differential pressure meter based on the pressure drop and the geometry of
        '    the venturi meter.
        '
        '    .. math::
        '        \epsilon =  \frac{\Delta\bar w }{\Delta P}
        '
        '    The :math:`\epsilon` value is looked up in a table of values as a function
        '    of beta ratio and upstream pipe diameter (roughness impact).
        '
        '    Parameters
        '    ----------
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    D0 : float
        '        Diameter of venturi tube at flow conditions, [m]
        '    P1 : float
        '        Static pressure of fluid upstream of venturi tube at the cross-section
        '        of the pressure tap, [Pa]
        '    P2 : float
        '        Static pressure of fluid downstream of venturi tube at the
        '         cross-section of the pressure tap, [Pa]
        '
        '    Returns
        '    -------
        '    dP : float
        '        Non-recoverable pressure drop of the venturi tube, [Pa]
        '
        '    Notes
        '    -----
        '    The recoverable pressure drop should be recovered by 6 pipe diameters
        '    downstream of the venturi tube.
        '
        '    Note there is some information on the effect of Reynolds number as well
        '    in [1]_ and [2]_, with a curve showing an increased pressure drop
        '    from 1E5-6E5 to with a decreasing multiplier from 1.75 to 1; the multiplier
        '    is 1 for higher Reynolds numbers. This is not currently included in this
        '    implementation.
        '
        '    Examples
        '    --------
        '    >>> DP_venturi_tube(0.07366, 0.05, 200000.0, 183000.0)
        '    1788.5717754177406
        '
        '    References
        '    ----------
        '    .. [1] American Society of Mechanical Engineers. Mfc-3M-2004 Measurement
        '       Of Fluid Flow In Pipes Using Orifice, Nozzle, And Venturi. ASME, 2001.
        '    .. [2] ISO 5167-4:2003 - Measurement of Fluid Flow by Means of Pressure
        '       Differential Devices Inserted in Circular Cross-Section Conduits Running
        '       Full -- Part 4: Venturi Tubes.
        '
        ' Effect of Re is not currently included
        Dim beta, venturi_tube_betas, venturi_tube_dP_high, venturi_tube_dP_low
        Dim D_bound_venturi_tube, epsilon_D65, epsilon_D500, epsilon

        beta = D0 / D

        ' Relative pressure loss as a function of beta reatio for venturi nozzles
        ' Venturi nozzles should be between 65 mm and 500 mm; there are high and low
        ' loss ratios , with the high losses corresponding to small diameters,
        ' low high losses corresponding to large diameters
        ' Interpolation can be performed.

        venturi_tube_betas = New Double() {0.29916, 0.29947, 0.31239, 0.31901, 0.32658, 0.33729,
                  0.34202, 0.34706, 0.35903, 0.36596, 0.37258, 0.38487,
                  0.38581, 0.40125, 0.40535, 0.41574, 0.42425, 0.43401,
                  0.44788, 0.45259, 0.47181, 0.47309, 0.49354, 0.49924,
                  0.51653, 0.5238, 0.53763, 0.54806, 0.55684, 0.57389,
                  0.58235, 0.59782, 0.60156, 0.62265, 0.62649, 0.64948,
                  0.65099, 0.6687, 0.67587, 0.68855, 0.69318, 0.70618,
                  0.71333, 0.72351, 0.74954, 0.74965}

        venturi_tube_dP_high = New Double() {0.164534, 0.164504, 0.163591, 0.163508, 0.163439,
                0.162652, 0.162224, 0.161866, 0.161238, 0.160786,
                0.160295, 0.15928, 0.159193, 0.157776, 0.157467,
                0.156517, 0.155323, 0.153835, 0.151862, 0.151154,
                0.14784, 0.147613, 0.144052, 0.14305, 0.140107,
                0.138981, 0.136794, 0.134737, 0.132847, 0.129303,
                0.127637, 0.124758, 0.124006, 0.119269, 0.118449,
                0.113605, 0.113269, 0.108995, 0.107109, 0.103688,
                0.102529, 0.099567, 0.097791, 0.095055, 0.087681,
                0.087648}

        venturi_tube_dP_low = New Double() {0.089232, 0.089218, 0.088671, 0.088435, 0.088206,
           0.087853, 0.087655, 0.087404, 0.086693, 0.086241,
           0.085813, 0.085142, 0.085102, 0.084446, 0.084202,
           0.083301, 0.08247, 0.08165, 0.080582, 0.080213,
           0.078509, 0.078378, 0.075989, 0.075226, 0.0727,
           0.071598, 0.069562, 0.068128, 0.066986, 0.064658,
           0.063298, 0.060872, 0.060378, 0.057879, 0.057403,
           0.054091, 0.053879, 0.051726, 0.050931, 0.049362,
           0.048675, 0.046522, 0.045381, 0.04384, 0.039913,
           0.039896}

        'ratios_average = 0.5*(ratios_high + ratios_low)
        D_bound_venturi_tube = New Double() {0.065, 0.5}

        epsilon_D65 = Interp(beta, venturi_tube_betas, venturi_tube_dP_high, , , True)
        epsilon_D500 = Interp(beta, venturi_tube_betas, venturi_tube_dP_low, , , True)
        epsilon = Interp(D, D_bound_venturi_tube, {epsilon_D65, epsilon_D500}, , , True)

        DP_venturi_tube = epsilon * (P1 - P2)

    End Function

    Function Diameter_ratio_cone_meter(D, Dc) As Double
        '    Calculates the diameter ratio `beta` used to characterize a cone
        '    flow meter.
        '
        '    .. math::
        '        \beta = \sqrt{1 - \frac{d_c^2}{D^2}}
        '
        '    Parameters
        '    ----------
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    Dc : float
        '        Diameter of the largest end of the cone meter, [m]
        '
        '    Returns
        '    -------
        '    beta : float
        '        Cone meter diameter ratio, [-]
        '
        '    Notes
        '    -----
        '    A mathematically equivalent formula often written is:
        '
        '    .. math::
        '        \beta = \frac{\sqrt{D^2 - d_c^2}}{D}
        '
        '    Examples
        '    --------
        '    >>> diameter_ratio_cone_meter(0.2575, 0.184)
        '    0.6995709873957624
        '
        '    References
        '    ----------
        '    .. [1] Hollingshead, Colter. "Discharge Coefficient Performance of Venturi,
        '       Standard Concentric Orifice Plate, V-Cone, and Wedge Flow Meters at
        '       Small Reynolds Numbers." May 1, 2011.
        '       https://digitalcommons.usu.edu/etd/869.

        Diameter_ratio_cone_meter = Math.Sqrt(1.0 - (Dc / D) * (Dc / D))

    End Function

    Function Cone_meter_expansibility_Stewart(D, Dc, P1, P2, k) As Double
        '    Calculates the expansibility factor for a cone flow meter,
        '    based on the geometry of the cone meter, measured pressures of the orifice,
        '    and the isentropic exponent of the fluid. Developed in [1]_, also shown
        '    in [2]_.
        '
        '    .. math::
        '        \epsilon = 1 - (0.649 + 0.696\beta^4) \frac{\Delta P}{\kappa P_1}
        '
        '    Parameters
        '    ----------
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    Dc : float
        '        Diameter of the largest end of the cone meter, [m]
        '    P1 : float
        '        Static pressure of fluid upstream of cone meter at the cross-section of
        '        the pressure tap, [Pa]
        '    P2 : float
        '        Static pressure of fluid at the end of the center of the cone pressure
        '        tap, [Pa]
        '    k : float
        '        Isentropic exponent of fluid, [-]
        '
        '    Returns
        '    -------
        '    expansibility : float
        '        Expansibility factor (1 for incompressible fluids, less than 1 for
        '        real fluids), [-]
        '
        '    Notes
        '    -----
        '    This formula was determined for the range of P2/P1 >= 0.75; the only gas
        '    used to determine the formula is air.
        '
        '    Examples
        '    --------
        '    >>> cone_meter_expansibility_Stewart(1, 0.9, 1E6, 8.5E5, 1.2)
        '    0.9157343
        '
        '    References
        '    ----------
        '    .. [1] Stewart, D. G., M. Reader-Harris, and NEL Dr RJW Peters. "Derivation
        '       of an Expansibility Factor for the V-Cone Meter." In Flow Measurement
        '       International Conference, Peebles, Scotland, UK, 2001.
        '    .. [2] ISO 5167-5:2016 - Measurement of Fluid Flow by Means of Pressure
        '       Differential Devices Inserted in Circular Cross-Section Conduits Running
        '       Full -- Part 5: Cone meters.
        Dim dP, beta, beta4

        dP = P1 - P2
        beta = Diameter_ratio_cone_meter(D, Dc)
        beta4 = beta * beta * beta * beta
        Cone_meter_expansibility_Stewart = 1.0 - (0.649 + 0.696 * beta4) * dP / (k * P1)

    End Function

    Function DP_cone_meter(D, Dc, P1, P2) As Double
        '    Calculates the non-recoverable pressure drop of a cone meter
        '    based on the measured pressures before and at the cone end, and the
        '    geometry of the cone meter according to [1]_.
        '
        '    .. math::
        '        \Delta \bar \omega = (1.09 - 0.813\beta)\Delta P
        '
        '    Parameters
        '    ----------
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    Dc : float
        '        Diameter of the largest end of the cone meter, [m]
        '    P1 : float
        '        Static pressure of fluid upstream of cone meter at the cross-section of
        '        the pressure tap, [Pa]
        '    P2 : float
        '        Static pressure of fluid at the end of the center of the cone pressure
        '        tap, [Pa]
        '
        '    Returns
        '    -------
        '    dP : float
        '        Non-recoverable pressure drop of the orifice plate, [Pa]
        '
        '    Notes
        '    -----
        '    The recoverable pressure drop should be recovered by 6 pipe diameters
        '    downstream of the cone meter.
        '
        '    Examples
        '    --------
        '    >>> DP_cone_meter(1, .7, 1E6, 9.5E5)
        '    25470.093437973323
        '
        '    References
        '    ----------
        '    .. [1] ISO 5167-5:2016 - Measurement of Fluid Flow by Means of Pressure
        '       Differential Devices Inserted in Circular Cross-Section Conduits Running
        '       Full -- Part 5: Cone meters.
        Dim dP, beta

        dP = P1 - P2
        beta = Diameter_ratio_cone_meter(D, Dc)
        DP_cone_meter = (1.09 - 0.813 * beta) * dP

    End Function

    Function Diameter_ratio_wedge_meter(D, h) As Double
        '    Calculates the diameter ratio `beta` used to characterize a wedge
        '    flow meter as given in [1]_ and [2]_.
        '
        '    .. math::
        '        \beta = \left(\frac{1}{\pi}\left\{\arccos\left[1 - \frac{2H}{D}
        '        \right] - 2 \left[1 - \frac{2H}{D}
        '        \right]\left(\frac{H}{D} - \left[\frac{H}{D}\right]^2
        '        \right)^{0.5}\right\}\right)^{0.5}
        '
        '    Parameters
        '    ----------
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    H : float
        '        Portion of the diameter of the clear segment of the pipe up to the
        '        wedge blocking flow; the height of the pipe up to the wedge, [m]
        '
        '    Returns
        '    -------
        '    beta : float
        '        Wedge meter diameter ratio, [-]
        '
        '    Notes
        '    -----
        '
        '    Examples
        '    --------
        '    >>> diameter_ratio_wedge_meter(0.2027, 0.0608)
        '    0.5022531424646643
        '
        '    References
        '    ----------
        '    .. [1] Hollingshead, Colter. "Discharge Coefficient Performance of Venturi,
        '       Standard Concentric Orifice Plate, V-Cone, and Wedge Flow Meters at
        '       Small Reynolds Numbers." May 1, 2011.
        '       https://digitalcommons.usu.edu/etd/869.
        '    .. [2] IntraWedge WEDGE FLOW METER Type: IWM. January 2011.
        '       http://www.intra-automation.com/download.php?file=pdf/products/technical_information/en/ti_iwm_en.pdf

        Dim H_D, t0, t1, t2, t3, t4

        H_D = h / D
        t0 = 1.0 - 2.0 * H_D
        t1 = Math.Acos(t0)
        t2 = t0 + t0
        t3 = Math.Sqrt(H_D - H_D * H_D)
        t4 = t1 - t2 * t3
        Diameter_ratio_wedge_meter = Math.Sqrt(t4 / PI)

    End Function

    Function C_wedge_meter_Miller(D, h) As Double
        '    Calculates the coefficient of discharge of an wedge flow meter
        '    used for measuring flow rate of fluid, based on the geometry of the
        '    differential pressure flow meter.
        '
        '    For half-inch lines:
        '
        '    .. math::
        '        C = 0.7883 + 0.107(1 - \beta^2)
        '
        '    For 1 to 1.5 inch lines:
        '
        '    .. math::
        '        C = 0.6143 + 0.718(1 - \beta^2)
        '
        '    For 1.5 to 24 inch lines:
        '
        '    .. math::
        '        C = 0.5433 + 0.2453(1 - \beta^2)
        '
        '    Parameters
        '    ----------
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    H : float
        '        Portion of the diameter of the clear segment of the pipe up to the
        '        wedge blocking flow; the height of the pipe up to the wedge, [m]
        '
        '    Returns
        '    -------
        '    C : float
        '        Coefficient of discharge of the wedge flow meter, [-]
        '
        '    Notes
        '    -----
        '    There is an ISO standard being developed to cover wedge meters as of 2018.
        '
        '    Wedge meters can have varying angles; 60 and 90 degree wedge meters have
        '    been reported. Tap locations 1 or 2 diameters (upstream and downstream),
        '    and 2D upstream/1D downstream have been used. Some wedges are sharp;
        '    some are smooth. [2]_ gives some experimental values.
        '
        '    Examples
        '    --------
        '    >>> C_wedge_meter_Miller(0.1524, 0.3*0.1524)
        '    0.7267069372687651
        '
        '    References
        '    ----------
        '    .. [1] Miller, Richard W. Flow Measurement Engineering Handbook. 3rd
        '       edition. New York: McGraw-Hill Education, 1996.
        '    .. [2] Seshadri, V., S. N. Singh, and S. Bhargava. "Effect of Wedge Shape
        '       and Pressure Tap Locations on the Characteristics of a Wedge Flowmeter."
        '       IJEMS Vol.01(5), October 1994.
        Dim beta, beta2, c

        beta = Diameter_ratio_wedge_meter(D, h)
        beta2 = beta * beta

        If D <= 0.7 * 0.0254 Then
            ' suggested limit 0.5 inch for this equation
            c = 0.7883 + 0.107 * (1.0 - beta2)
        ElseIf D <= 1.4 * 0.0254 Then
            ' Suggested limit is under 1.5 inches
            c = 0.6143 + 0.718 * (1.0 - beta2)
        Else
            c = 0.5433 + 0.2453 * (1.0 - beta2)
        End If

        C_wedge_meter_Miller = c

    End Function

    Function C_wedge_meter_ISO_5167_6_2017(D, h) As Double
        '    Calculates the coefficient of discharge of an wedge flow meter
        '    used for measuring flow rate of fluid, based on the geometry of the
        '    differential pressure flow meter according to the ISO 5167-6 standard
        '    (draft 2017).
        '
        '    .. math::
        '        C = 0.77 - 0.09\beta
        '
        '    Parameters
        '    ----------
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    H : float
        '        Portion of the diameter of the clear segment of the pipe up to the
        '        wedge blocking flow; the height of the pipe up to the wedge, [m]
        '
        '    Returns
        '    -------
        '    C : float
        '        Coefficient of discharge of the wedge flow meter, [-]
        '
        '    Notes
        '    -----
        '    This standard applies for wedge meters in line sizes between 50 and 600 mm;
        '    and height ratios between 0.2 and 0.6. The range of allowable Reynolds
        '    numbers is large; between 1E4 and 9E6. The uncertainty of the flow
        '    coefficient is approximately 4%. Usually a 10:1 span of flow can be
        '    measured accurately. The discharge and entry length of the meters must be
        '    at least half a pipe diameter. The wedge angle must be 90 degrees, plus or
        '    minus two degrees.
        '
        '    The orientation of the wedge meter does not change the accuracy of this
        '    model.
        '
        '    There should be a straight run of 10 pipe diameters before the wedge meter
        '    inlet, and two of the same pipe diameters after it.
        '
        '    Examples
        '    --------
        '    >>> C_wedge_meter_ISO_5167_6_2017(0.1524, 0.3*0.1524)
        '    0.724792059539853
        '
        '    References
        '    ----------
        '    .. [1] ISO/DIS 5167-6 - Measurement of Fluid Flow by Means of Pressure
        '       Differential Devices Inserted in Circular Cross-Section Conduits Running
        '       Full -- Part 6: Wedge Meters.

        C_wedge_meter_ISO_5167_6_2017 = 0.77 - 0.09 * Diameter_ratio_wedge_meter(D, h)

    End Function

    Function DP_wedge_meter(D, h, P1, P2) As Double
        '    Calculates the non-recoverable pressure drop of a wedge meter
        '    based on the measured pressures before and at the wedge meter, and the
        '    geometry of the wedge meter according to [1]_.
        '
        '    .. math::
        '        \Delta \bar \omega = (1.09 - 0.79\beta)\Delta P
        '
        '    Parameters
        '    ----------
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    H : float
        '        Portion of the diameter of the clear segment of the pipe up to the
        '        wedge blocking flow; the height of the pipe up to the wedge, [m]
        '    P1 : float
        '        Static pressure of fluid upstream of wedge meter at the cross-section
        '        of the pressure tap, [Pa]
        '    P2 : float
        '        Static pressure of fluid at the end of the wedge meter pressure tap, [
        '        Pa]
        '
        '    Returns
        '    -------
        '    dP : float
        '        Non-recoverable pressure drop of the wedge meter, [Pa]
        '
        '    Notes
        '    -----
        '    The recoverable pressure drop should be recovered by 5 pipe diameters
        '    downstream of the wedge meter.
        '
        '    Examples
        '    --------
        '    >>> DP_wedge_meter(1, .7, 1E6, 9.5E5)
        '    20344.849697483587
        '
        '    References
        '    ----------
        '    .. [1] ISO/DIS 5167-6 - Measurement of Fluid Flow by Means of Pressure
        '       Differential Devices Inserted in Circular Cross-Section Conduits Running
        '       Full -- Part 6: Wedge Meters.

        DP_wedge_meter = (1.09 - 0.79 * Diameter_ratio_wedge_meter(D, h)) * (P1 - P2)

    End Function

    Function C_Reader_Harris_Gallagher_wet_venturi_tube(mg, ml, rhog, rhol, D, D0, h) As Double
        '    Calculates the coefficient of discharge of the wet gas venturi tube
        '    based on the  geometry of the tube, mass flow rates of liquid and vapor
        '    through the tube, the density of the liquid and gas phases, and an
        '    adjustable coefficient `H`.
        '
        '    .. math::
        '        C = 1 - 0.0463\exp(-0.05Fr_{gas, th}) \cdot \min\left(1,
        '        \sqrt{\frac{X}{0.016}}\right)
        '
        '    .. math::
        '        Fr_{gas, th} = \frac{Fr_{\text{gas, densionetric }}}{\beta^{2.5}}
        '
        '    .. math::
        '        \phi = \sqrt{1 + C_{Ch} X + X^2}
        '
        '    .. math::
        '        C_{Ch} = \left(\frac{\rho_l}{\rho_{1,g}}\right)^n +
        '        \left(\frac{\rho_{1, g}}{\rho_{l}}\right)^n
        '
        '    .. math::
        '        n = \max\left[0.583 - 0.18\beta^2 - 0.578\exp\left(\frac{-0.8
        '        Fr_{\text{gas, densiometric}}}{H}\right),0.392 - 0.18\beta^2 \right]
        '
        '    .. math::
        '        X = \left(\frac{m_l}{m_g}\right) \sqrt{\frac{\rho_{1,g}}{\rho_l}}
        '
        '    .. math::
        '        {Fr_{\text{gas, densiometric}}} = \frac{v_{gas}}{\sqrt{gD}}
        '        \sqrt{\frac{\rho_{1,g}}{\rho_l - \rho_{1,g}}}
        '        =  \frac{4m_g}{\rho_{1,g} \pi D^2 \sqrt{gD}}
        '        \sqrt{\frac{\rho_{1,g}}{\rho_l - \rho_{1,g}}}
        '
        '    Parameters
        '    ----------
        '    mg : float
        '        Mass flow rate of gas through the venturi tube, [kg/s]
        '    ml : float
        '        Mass flow rate of liquid through the venturi tube, [kg/s]
        '    rhog : float
        '        Density of gas at `P1`, [kg/m^3]
        '    rhol : float
        '        Density of liquid at `P1`, [kg/m^3]
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    D0 : float
        '        Diameter of venturi tube at flow conditions, [m]
        '    H : float, optional
        '        A surface-tension effect coefficient used to adjust for different
        '        fluids, (1 for a hydrocarbon liquid, 1.35 for water, 0.79 for water in
        '        steam) [-]
        '
        '    Returns
        '    -------
        '    C : float
        '        Coefficient of discharge of the wet gas venturi tube flow meter
        '        (includes flow rate of gas ONLY), [-]
        '
        '    Notes
        '    -----
        '    This model has more error than single phase differential pressure meters.
        '    The model was first published in [1]_, and became ISO 11583 later.
        '
        '    The limits of this correlation according to [2]_ are as follows:
        '
        '    .. math::
        '        0.4 \le \beta \le 0.75
        '
        '    .. math::
        '        0 < X \le 0.3
        '
        '    .. math::
        '        Fr_{gas, th} > 3
        '
        '    .. math::
        '        \frac{\rho_g}{\rho_l} > 0.02
        '
        '    .. math::
        '        D \ge 50 \text{ mm}
        '
        '    Examples
        '    --------
        '    >>> C_Reader_Harris_Gallagher_wet_venturi_tube(5.31926, 5.31926/2, 50.0, 800., 0.1, 0.06, 1)
        '    0.9754210845876333
        '
        '    References
        '    ----------
        '    .. [1] Reader-harris, Michael, and Tuv Nel. An Improved Model for
        '       Venturi-Tube Over-Reading in Wet Gas, 2009.
        '    .. [2] ISO/TR 11583:2012 Measurement of Wet Gas Flow by Means of Pressure
        '       Differential Devices Inserted in Circular Cross-Section Conduits.
        Dim V, Frg, beta, beta2, Fr_gas_th, n, t0, t1, C_Ch, x, c

        V = 4.0 * mg / (rhog * PI * D * D)
        Frg = Froude_densimetric(V, D, rhol, rhog, False)
        beta = D0 / D
        beta2 = beta * beta
        Fr_gas_th = Frg / (beta2 * Math.Sqrt(beta))

        n = Math.Max(0.583 - 0.18 * beta2 - 0.578 * Math.Exp(-0.8 * Frg / h), 0.392 - 0.18 * beta2)

        t0 = rhog / rhol
        t1 = Math.Pow(t0, n)
        C_Ch = t1 + 1.0 / t1
        x = ml / mg * Math.Sqrt(t0)
        ' OF = Math.Sqrt(1.0 + X*(C_Ch + X))

        c = 1.0 - 0.0463 * Math.Exp(-0.05 * Fr_gas_th) * Math.Min(1.0, Math.Sqrt(x / 0.016))
        C_Reader_Harris_Gallagher_wet_venturi_tube = c

    End Function

    Function DP_Reader_Harris_Gallagher_wet_venturi_tube(D, D0, P1, P2, ml, mg, rhol, rhog, h) As Double
        '    Calculates the non-recoverable pressure drop of a wet gas venturi
        '    nozzle based on the pressure drop and the geometry of the venturi nozzle,
        '    the mass flow rates of liquid and gas through it, the densities of the
        '    vapor and liquid phase, and an adjustable coefficient `H`.
        '
        '    .. math::
        '        Y = \frac{\Delta \bar \omega}{\Delta P} - 0.0896 - 0.48\beta^9
        '
        '    .. math::
        '        Y_{max} = 0.61\exp\left[-11\frac{\rho_{1,g}}{\rho_l}
        '        - 0.045 \frac{Fr_{gas}}{H}\right]
        '
        '    .. math::
        '        \frac{Y}{Y_{max}} = 1 - \exp\left[-35 X^{0.75} \exp
        '        \left( \frac{-0.28Fr_{gas}}{H}\right)\right]
        '
        '    .. math::
        '        X = \left(\frac{m_l}{m_g}\right) \sqrt{\frac{\rho_{1,g}}{\rho_l}}
        '
        '    .. math::
        '        {Fr_{\text{gas, densiometric}}} = \frac{v_{gas}}{\sqrt{gD}}
        '        \sqrt{\frac{\rho_{1,g}}{\rho_l - \rho_{1,g}}}
        '        =  \frac{4m_g}{\rho_{1,g} \pi D^2 \sqrt{gD}}
        '        \sqrt{\frac{\rho_{1,g}}{\rho_l - \rho_{1,g}}}
        '
        '    Parameters
        '    ----------
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    D0 : float
        '        Diameter of venturi tube at flow conditions, [m]
        '    P1 : float
        '        Static pressure of fluid upstream of venturi tube at the cross-section
        '        of the pressure tap, [Pa]
        '    P2 : float
        '        Static pressure of fluid downstream of venturi tube at the cross-
        '        section of the pressure tap, [Pa]
        '    ml : float
        '        Mass flow rate of liquid through the venturi tube, [kg/s]
        '    mg : float
        '        Mass flow rate of gas through the venturi tube, [kg/s]
        '    rhol : float
        '        Density of liquid at `P1`, [kg/m^3]
        '    rhog : float
        '        Density of gas at `P1`, [kg/m^3]
        '    H : float, optional
        '        A surface-tension effect coefficient used to adjust for different
        '        fluids, (1 for a hydrocarbon liquid, 1.35 for water, 0.79 for water in
        '        steam) [-]
        '
        '    Returns
        '    -------
        '    C : float
        '        Coefficient of discharge of the wet gas venturi tube flow meter
        '        (includes flow rate of gas ONLY), [-]
        '
        '    Notes
        '    -----
        '    The model was first published in [1]_, and became ISO 11583 later.
        '
        '    Examples
        '    --------
        '    >>> DP_Reader_Harris_Gallagher_wet_venturi_tube(0.1, 0.06, 6E6, 6E6-5E4, 5.31926/2, 5.31926, 800, 50.0, 1)
        '    16957.43843129572
        '
        '    References
        '    ----------
        '    .. [1] Reader-harris, Michael, and Tuv Nel. An Improved Model for
        '       Venturi-Tube Over-Reading in Wet Gas, 2009.
        '    .. [2] ISO/TR 11583:2012 Measurement of Wet Gas Flow by Means of Pressure
        '       Differential Devices Inserted in Circular Cross-Section Conduits.
        Dim dP, beta, x, V, Frg, Y_ratio, Y_max, y, rhs, dw

        dP = P1 - P2
        beta = D0 / D
        x = ml / mg * Math.Sqrt(rhog / rhol)

        V = 4 * mg / (rhog * PI * D * D)
        Frg = Froude_densimetric(V, D, rhol, rhog, False)

        Y_ratio = 1.0 - Math.Exp(-35.0 * Math.Pow(x, 0.75) * Math.Exp(-0.28 * Frg / h))
        Y_max = 0.61 * Math.Exp(-11.0 * rhog / rhol - 0.045 * Frg / h)
        y = Y_max * Y_ratio
        rhs = -0.0896 - 0.48 * Math.Pow(beta, 9)
        dw = dP * (y - rhs)

        DP_Reader_Harris_Gallagher_wet_venturi_tube = dw

    End Function

    Function Froude_densimetric(V, l, rho1, rho2, Optional heavy = True) As Double
        '    Calculates the densimetric Froude number :math:`Fr_{den}` for velocity
        '    `V` geometric length `L`, heavier fluid density `rho1`, and lighter fluid
        '    density `rho2`. If desired, gravity can be specified as well. Depending on
        '    the application, this dimensionless number may be defined with the heavy
        '    phase or the light phase density in the numerator of the square root.
        '    For some applications, both need to be calculated. The default is to
        '    calculate with the heavy liquid ensity on top; set `heavy` to False
        '    to reverse this.
        '
        '    .. math::
        '        Fr = \frac{V}{\sqrt{gL}} \sqrt{\frac{\rho_\text{(1 or 2)}}
        '        {\rho_1 - \rho_2}}
        '
        '    Parameters
        '    ----------
        'V:      float
        '        Velocity of the specified phase, [m/s]
        'L:      float
        '        Characteristic length, no typical definition [m]
        'rho1:      float
        '        Density of the heavier phase, [kg/m^3]
        'rho2:      float
        '        Density of the lighter phase, [kg/m^3]
        '    heavy : bool, optional
        '        Whether or not the density used in the numerator is the heavy phase or
        '        the light phase, [-]
        '    g : float, optional
        '        Acceleration due to gravity, [m/s^2]
        '
        '    Returns
        '    -------
        'Fr_den:      float
        '        Densimetric Froude number, [-]
        '
        '    Notes
        '    -----
        '    Many alternate definitions including density ratios have been used.
        '
        '    .. math::
        '        Fr = \frac{\text{Inertial Force}}{\text{Gravity Force}}'
        '
        '    Where the gravity force is reduced by the relative densities of one fluid
        '    in another.
        '
        '    Note that an Exception will be raised if rho1 > rho2, as the square root
        '    becomes negative.
        '
        '    Examples
        '    --------
        '    >>> Froude_densimetric(1.83, L=2., rho1=800, rho2=1.2, g=9.81)
        '    0.4134543386272418
        '    >>> Froude_densimetric(1.83, L=2., rho1=800, rho2=1.2, g=9.81, heavy=False)
        '    0.016013017679205096
        '
        '    References
        '    ----------
        '    .. [1] Hall, A, G Stobie, and R Steven. "Further Evaluation of the
        '       Performance of Horizontally Installed Orifice Plate and Cone
        '       Differential Pressure Meters with Wet Gas Flows." In International
        '       SouthEast Asia Hydrocarbon Flow Measurement Workshop, KualaLumpur,
        '       Malaysia , 2008

        Dim rho3

        If heavy Then
            rho3 = rho1
        Else
            rho3 = rho2
        End If

        Froude_densimetric = V / (Math.Sqrt(9.80665 * l)) * Math.Sqrt(rho3 / (rho1 - rho2))

    End Function

    Function Differential_pressure_meter_beta(D As Double, D2 As Double, meter_type As String)
        '    Calculates the beta ratio of a differential pressure meter.
        '
        '    Parameters
        '    ----------
        '    D : float
        '        Upstream internal pipe diameter, [m]
        '    D2 : float
        '        Diameter of orifice, or venturi meter orifice, or flow tube orifice,
        '        or cone meter end diameter, or wedge meter fluid flow height, [m]
        '    meter_type : str
        '        One of {'conical orifice', 'orifice', 'machined convergent venturi tube',
        '        'ISO 5167 orifice', 'Miller quarter circle orifice', 'Hollingshead venturi sharp',
        '        'segmental orifice', 'Miller conical orifice', 'Miller segmental orifice',
        '        'quarter circle orifice', 'Hollingshead v cone', 'wedge meter', 'eccentric orifice',
        '        'venuri nozzle', 'rough welded convergent venturi tube', 'ISA 1932 nozzle',
        '        'ISO 15377 quarter-circle orifice', 'Hollingshead venturi smooth',
        '        'Hollingshead orifice', 'cone meter', 'Hollingshead wedge', 'Miller orifice',
        '        'long radius nozzle', 'ISO 15377 conical orifice', 'unspecified meter',
        '        'as cast convergent venturi tube', 'Miller eccentric orifice',
        '        'ISO 15377 eccentric orifice'}, [-]
        '
        '    Returns
        '    -------
        '    beta : float
        '        Differential pressure meter diameter ratio, [-]
        '
        '    Notes
        '    -----
        '
        '    Examples
        '    --------
        '    >>> differential_pressure_meter_beta(0.2575, 0.184,'cone meter')
        '    0.6995709873957624
        Dim beta

        If IsInBetaSimpleMeter(meter_type) Then
            beta = D2 / D
        ElseIf IsInArray(meter_type, {CONE_METER, HOLLINGSHEAD_CONE}) Then
            beta = Diameter_ratio_cone_meter(D, D2)
        ElseIf IsInArray(meter_type, {WEDGE_METER, HOLLINGSHEAD_WEDGE}) Then
            beta = Diameter_ratio_wedge_meter(D, D2)
        Else
            beta = 0.0
        End If

        Differential_pressure_meter_beta = beta

    End Function

    Function Orifice_std_Hollingshead_tck(x, y)

        Dim tx, ty, c, kx, ky, Z

        tx = {0.5, 0.5, 0.5, 0.5, 0.7, 0.7, 0.7, 0.7}
        ty = {0, 0, 0, 0, 2.30258509299405, 2.99573227355399, 3.40119738166216, 3.68887945411394, 4.0943445622221, 4.38202663467388,
                4.60517018598809, 5.29831736654804, 5.7037824746562, 6.21460809842219, 6.90775527898214, 7.60090245954208, 8.00636756765025,
                8.51719319141624, 9.21034037197618, 11.5129254649702, 13.8155105579643, 17.7275335633924, 17.7275335633924, 17.7275335633924,
                17.7275335633924}
        c = {0.233, 0.304079384502282, 0.539769337938802, 0.650941432564864, 0.676141993726265, 0.690169740115681, 0.697224070790928,
                0.699675957250515, 0.704022336370595, 0.700874158771197, 0.692665226515394, 0.682638781867897, 0.672793064316652, 0.649054216185994,
                0.637878095969801, 0.630202750473631, 0.628490452361042, 0.616773266650063, 0.614410803002411, 0.613727077014918, 0.614,
                0.217222222222222, 0.26754856063815, 0.547178981607613, 0.682583584947149, 0.684825512088075, 0.712775784969247, 0.706684254500825,
                0.702034574426881, 0.693147673731604, 0.671088678547894, 0.650121869598914, 0.625716497557949, 0.58884635672329, 0.623750533639281,
                0.578149766754485, 0.576189016008046, 0.592230310398501, 0.565779097486493, 0.601337637367252, 0.569359355594998, 0.552888888888889,
                0.206777777777778, 0.264434235009685, 0.463098557203435, 0.63068495223115, 0.689926018874737, 0.70927038791343, 0.733141665407242,
                0.740386621990052, 0.753149363639563, 0.768501905339505, 0.771007019842085, 0.76495337729654, 0.77070200817463, 0.689783247209235,
                0.691061834137385, 0.680576352979605, 0.629188477215149, 0.647090424466067, 0.596287989949754, 0.635309679831602, 0.627777777777778,
                0.191, 0.237122768892702, 0.444828426613922, 0.63372254649304, 0.692646297813639, 0.731687488866313, 0.754205721153009,
                0.77172737538752, 0.787604977842911, 0.795143180926116, 0.797757098609426, 0.786144504322234, 0.777182818678971, 0.705734580065083,
                0.662669862852663, 0.660069043365499, 0.632339643107207, 0.621268403483029, 0.616281323630018, 0.603728515722033, 0.604}
        kx = 3
        ky = 3

        x = {x}
        y = {y}
        Z = Cy_bispev(tx, ty, c, kx, ky, x, y)

        Orifice_std_Hollingshead_tck = Z(0)

    End Function

    Function Cone_Hollingshead_tck(x, y)

        Dim tx, ty, c, kx, ky, Z

        tx = {0.6611, 0.6611, 0.6611, 0.8203, 0.8203, 0.8203}
        ty = {0, 0, 0, 0, 2.30258509299405, 2.99573227355399, 3.40119738166216, 3.68887945411394, 4.0943445622221, 4.38202663467388,
                4.60517018598809, 5.01063529409626, 5.29831736654804, 5.7037824746562, 6.21460809842219, 6.90775527898214, 7.60090245954208,
                8.00636756765025, 8.29404964010203, 8.51719319141624, 8.9226582995244, 9.21034037197618, 9.90348755253613, 10.3089526606443,
                11.5129254649702, 13.8155105579643, 17.7275335633924, 17.7275335633924, 17.7275335633924, 17.7275335633924}
        c = {0.066, 0.0918118088794429, 0.140634145301067, 0.273197698663, 0.341778399535323, 0.40258800767255, 0.456314932881035,
                0.503544530735729, 0.545847369335969, 0.583175639128474, 0.628052124545805, 0.664719813500578, 0.709152439678625, 0.725472982341933,
                0.748781696392684, 0.758814550281781, 0.762869253263183, 0.766048214721483, 0.764418831958338, 0.778264414400624, 0.772150813911649,
                0.799472879402824, 0.807674219471452, 0.79862214208228, 0.80862405328503, 0.802, 0.0701623206401766, 0.105916263570389,
                0.148968183859281, 0.288308157486292, 0.354052137069574, 0.403397955040637, 0.454457032305519, 0.503463771220107, 0.544819015669371,
                0.584016424503113, 0.621155959809806, 0.621864884498082, 0.662174576071073, 0.728237954629295, 0.734003073480127, 0.73963248657796,
                0.748973679895375, 0.748072641291472, 0.767156475116998, 0.756853660688892, 0.778702964227274, 0.774238113131269, 0.788758416244345,
                0.785761045021833, 0.769707664555196, 0.771830091059603, 0.057, 0.0761254485994355, 0.124017334157783, 0.240374522095959,
                0.296624635025932, 0.348595365868552, 0.394800857193225, 0.436616016224806, 0.480912591024548, 0.524069128618623, 0.559060928802062,
                0.61445560487167, 0.647171364056714, 0.690415880906118, 0.703259025205022, 0.712177974557301, 0.722184530368027, 0.721505707129694,
                0.724982237626455, 0.721889008528991, 0.722184847576871, 0.737175135451553, 0.725238506230463, 0.72789438039334, 0.749654660702909,
                0.734}
        kx = 2
        ky = 3

        x = {x}
        y = {y}
        Z = Cy_bispev(tx, ty, c, kx, ky, x, y)

        Cone_Hollingshead_tck = Z(0)

    End Function

    Function Wedge_Hollingshead_tck(x, y)

        Dim tx, ty, c, kx, ky, Z

        tx = {0.5023, 0.5023, 0.611, 0.611}
        ty = {0, 0, 0, 0, 2.30258509299405, 2.99573227355399, 3.40119738166216, 3.68887945411394, 4.0943445622221, 4.38202663467388,
                4.60517018598809, 5.29831736654804, 5.7037824746562, 5.99146454710798, 6.21460809842219, 8.51719319141624, 9.21034037197618,
                11.5129254649702, 17.7275335633924, 17.7275335633924, 17.7275335633924, 17.7275335633924}
        c = {0.145, 0.18231832425722, 0.333991713000692, 0.537946771022697, 0.60777006599409, 0.645954294392508, 0.672975700777023,
                0.689640500757623, 0.705486311458958, 0.715574060063264, 0.720544640761086, 0.723957681606897, 0.748362756816017, 0.723296335591993,
                0.736632532049095, 0.726422214356705, 0.733960539412601, 0.733, 0.127, 0.169398738651323, 0.282849493352567,
                0.488910700907784, 0.56231200435241, 0.613309237967695, 0.643709239468792, 0.662992336666202, 0.678293436601103, 0.687302374134782,
                0.692747005312891, 0.69939923642349, 0.722120448354685, 0.694757729328402, 0.706370130681081, 0.678161453435987, 0.718532681194841,
                0.705}
        kx = 1
        ky = 3

        x = {x}
        y = {y}
        Z = Cy_bispev(tx, ty, c, kx, ky, x, y)

        Wedge_Hollingshead_tck = Z(0)

    End Function

    Function Fpbspl(T, n, k, x, l, ByRef h, ByRef hh)
        Dim i, j, li, f

        h(0) = n
        h(0) = 1.0

        For j = 1 To k

            'hh[0:j] = h[0:j]

            For i = 0 To j - 1
                hh(i) = h(i)
            Next i

            h(0) = 0

            For i = 0 To j - 1
                li = l + i
                f = hh(i) / (T(li) - T(li - j))
                h(i) = h(i) + f * (T(li) - x)
                h(i + 1) = f * (x - T(li - j))
            Next i

        Next j

        Fpbspl = 0

    End Function

    Function Init_w(T, k, x, lx, ByRef w)
        Dim tb, n, m, h, hh, te, l1, l2, i, j, arg, arg_temp

        tb = T(k)
        n = UBound(T) - LBound(T) + 1
        m = UBound(x) - LBound(x) + 1

        ReDim h(0 To 5)
        ReDim hh(0 To 4)

        te = T(n - k - 1)
        l1 = k + 1
        l2 = l1 + 1

        For i = 0 To m - 1
            arg = x(i)

            If arg < tb Then arg = tb
            If arg > te Then arg = te

            While Not (arg < T(l1) Or l1 = (n - k - 1))
                l1 = l2
                l2 = l1 + 1
            End While

            arg_temp = Fpbspl(T, n, k, arg, l1, h, hh)

            lx(i) = l1 - k - 1

            For j = 0 To k
                w(i, j) = h(j)
            Next j
        Next i

        Init_w = 0

    End Function

    Function Cy_bispev(tx, ty, c, kx, ky, x, y)
        Dim nx, ny, mx, my, kx1, ky1, nkx1, nky1, wx, wy, lx, ly
        Dim size_z, Z, i, j, i1, j1, sp, err, l2, a, tmp

        nx = UBound(tx) - LBound(tx) + 1
        ny = UBound(ty) - LBound(ty) + 1
        mx = 1 ' hardcode to one point
        my = 1 ' hardcode to one point

        kx1 = kx + 1
        ky1 = ky + 1

        nkx1 = nx - kx1
        nky1 = ny - ky1

        ReDim wx(0 To mx - 1, 0 To kx1 - 1)
        ReDim wy(0 To my - 1, 0 To ky1 - 1)
        ReDim lx(0 To mx - 1)
        ReDim ly(0 To my - 1)

        size_z = mx * my

        ReDim Z(0 To size_z - 1)

        i = Init_w(tx, kx, x, lx, wx)
        j = Init_w(ty, ky, y, ly, wy)

        For j = 0 To my - 1
            For i = 0 To mx - 1
                sp = 0
                err = 0
                For i1 = 0 To kx1 - 1
                    For j1 = 0 To ky1 - 1
                        l2 = lx(i) * nky1 + ly(j) + i1 * nky1 + j1
                        a = c(l2) * wx(i, i1) * wy(j, j1) - err
                        tmp = sp + a
                        err = (tmp - sp) - a
                        sp = tmp
                    Next j1
                Next i1
                Z(j * mx + i) = Z(j * mx + i) + sp
            Next i
        Next j

        Cy_bispev = Z
    End Function

    Function Differential_pressure_meter_C_epsilon(D As Double, D2 As Double, m As Double, P1 As Double, P2 As Double, rho As Double, mu As Double, k As Double, meter_type As String, taps As String, tap_position As String, Optional C_specified As Double = 0.0, Optional epsilon_specified As Double = 0.0)
        '    Calculates the discharge coefficient and expansibility of a flow
        '    meter given the mass flow rate, the upstream pressure, the second
        '    pressure value, and the orifice diameter for a differential
        '    pressure flow meter based on the geometry of the meter, measured pressures
        '    of the meter, and the density, viscosity, and isentropic exponent of the
        '    fluid.
        '
        '    Parameters
        '    ----------
        '    D  then float
        '        Upstream internal pipe diameter, [m]
        '    D2  then float
        '        Diameter of orifice, or venturi meter orifice, or flow tube orifice,
        '        or cone meter end diameter, or wedge meter fluid flow height, [m]
        '    m  then float
        '        Mass flow rate of fluid through the flow meter, [kg/s]
        '    P1  then float
        '        Static pressure of fluid upstream of differential pressure meter at the
        '        cross-section of the pressure tap, [Pa]
        '    P2  then float
        '        Static pressure of fluid downstream of differential pressure meter or
        '        at the prescribed location (varies by type of meter) [Pa]
        '    rho  then float
        '        Density of fluid at `P1`, [kg/m^3]
        '    mu  then float
        '        Viscosity of fluid at `P1`, [Pa*s]
        '    k  then float
        '        Isentropic exponent of fluid, [-]
        '    meter_type  then str
        '        One of {'conical orifice', 'orifice', 'machined convergent venturi tube',
        '        'ISO 5167 orifice', 'Miller quarter circle orifice', 'Hollingshead venturi sharp',
        '        'segmental orifice', 'Miller conical orifice', 'Miller segmental orifice',
        '        'quarter circle orifice', 'Hollingshead v cone', 'wedge meter', 'eccentric orifice',
        '        'venuri nozzle', 'rough welded convergent venturi tube', 'ISA 1932 nozzle',
        '        'ISO 15377 quarter-circle orifice', 'Hollingshead venturi smooth',
        '        'Hollingshead orifice', 'cone meter', 'Hollingshead wedge', 'Miller orifice',
        '        'long radius nozzle', 'ISO 15377 conical orifice', 'unspecified meter',
        '        'as cast convergent venturi tube', 'Miller eccentric orifice',
        '        'ISO 15377 eccentric orifice'}, [-]
        '    taps  then int, optional 1, 2, or 3
        '        The orientation of the taps; one of 'corner', 'flange', 'D', or 'D/2';
        '        applies for orifice meters only, [-]
        '    tap_position  then str, optional
        '        The rotation of the taps, used **only for the eccentric orifice case**
        '        where the pressure profile is are not symmetric; '180 degree' for the
        '        normal case where the taps are opposite the orifice bore, and
        '        '90 degree' for the case where, normally for operational reasons, the
        '        taps are near the bore [-]
        '    C_specified  then float, optional
        '        If specified, the correlation for the meter type is not used - this
        '        value is returned for `C`
        '    epsilon_specified  then float, optional
        '        If specified, the correlation for the fluid expansibility is not used -
        '        this value is returned for  thenmath then`\epsilon`, [-]
        '
        '    Returns
        '    -------
        '    C  then float
        '        Coefficient of discharge of the specified flow meter type at the
        '        specified conditions, [-]
        '    expansibility  then float
        '        Expansibility factor (1 for incompressible fluids, less than 1 for
        '        real fluids), [-]
        '
        '    Notes
        '    -----
        '    This function should be called by an outer loop when solving for a
        '    variable.
        '
        '    The latest ISO formulations for `expansibility` are used with the Miller
        '    correlations.
        '
        '    Examples
        '    --------
        '    >>> differential_pressure_meter_C_epsilon(0.07366, 0.05, 7.702338035732168, P1=200000.0, 183000.0, 999.1, 0.0011, 1.33, "ISO 5167 orifice", 3, "180 degree", 0, 0)
        '    (0.6151252900244296, 0.9711026966676307)

        Dim as_cast_convergent_venturi_Res
        Dim as_cast_convergent_venturi_Cs
        Dim machined_convergent_venturi_Res
        Dim machined_convergent_venturi_Cs
        Dim rough_welded_convergent_venturi_Res
        Dim rough_welded_convergent_venturi_Cs
        Dim as_cast_convergent_entrance_machined_venturi_Res
        Dim as_cast_convergent_entrance_machined_venturi_Cs
        Dim venturi_Res_Hollingshead
        Dim venturi_logRes_Hollingshead
        Dim venturi_smooth_Cs_Hollingshead
        Dim venturi_sharp_Cs_Hollingshead
        Dim CONE_METER_C
        Dim ROUGH_WELDED_CONVERGENT_VENTURI_TUBE_C
        Dim MACHINED_CONVERGENT_VENTURI_TUBE_C
        Dim AS_CAST_VENTURI_TUBE_C
        Dim ISO_15377_CONICAL_ORIFICE_C
        Dim cone_Res_Hollingshead
        Dim cone_logRes_Hollingshead
        Dim cone_betas_Hollingshead
        Dim cone_beta_6611_Hollingshead_Cs
        Dim cone_beta_6995_Hollingshead_Cs
        Dim cone_beta_8203_Hollingshead_Cs
        Dim cone_Hollingshead_Cs
        'Dim cone_Hollingshead_tck
        Dim wedge_Res_Hollingshead
        Dim wedge_logRes_Hollingshead
        Dim wedge_beta_5023_Hollingshead
        Dim wedge_beta_611_Hollingshead
        Dim wedge_betas_Hollingshead
        Dim wedge_Hollingshead_Cs
        'Dim wedge_Hollingshead_tck

        ' Venturi tube loss coefficients as a function of Re
        as_cast_convergent_venturi_Res = {400000.0, 60000.0, 100000.0, 150000.0}
        as_cast_convergent_venturi_Cs = {0.957, 0.966, 0.976, 0.982}

        machined_convergent_venturi_Res = {50000.0, 100000.0, 200000.0, 300000.0,
                                           750000.0,
                                           1500000.0,
                                           5000000.0} ' 2E6 to 1E8
        machined_convergent_venturi_Cs = {0.97, 0.977, 0.992, 0.998, 0.995, 1.0, 1.01}

        rough_welded_convergent_venturi_Res = {40000.0, 60000.0, 100000.0}
        rough_welded_convergent_venturi_Cs = {0.96, 0.97, 0.98}

        as_cast_convergent_entrance_machined_venturi_Res = {10000.0, 60000.0, 100000.0, 150000.0,
                                                            350000.0,
                                                            3200000.0} ' 5E5 to 3.2E6
        as_cast_convergent_entrance_machined_venturi_Cs = {0.963, 0.978, 0.98, 0.987, 0.992, 0.995}

        venturi_Res_Hollingshead = {1.0, 5.0, 10.0, 20.0, 30.0, 40.0, 60.0, 80.0, 100.0, 200.0, 300.0, 500.0, 1000.0, 2000.0, 3000.0, 5000.0, 10000.0, 30000.0, 50000.0, 75000.0, 100000.0, 1000000.0, 10000000.0, 50000000.0}
        venturi_logRes_Hollingshead = {0, 1.6094379124341, 2.30258509299405, 2.99573227355399, 3.40119738166216,
        3.68887945411394, 4.0943445622221, 4.38202663467388, 4.60517018598809, 5.29831736654804, 5.7037824746562,
        6.21460809842219, 6.90775527898214, 7.60090245954208, 8.00636756765025, 8.51719319141624, 9.21034037197618,
        10.3089526606443, 10.8197782844103, 11.2252433925184, 11.5129254649702, 13.8155105579643, 16.1180956509583,
        17.7275335633924}
        venturi_smooth_Cs_Hollingshead = {0.163, 0.336, 0.432, 0.515, 0.586, 0.625, 0.679, 0.705, 0.727, 0.803, 0.841, 0.881, 0.921, 0.937, 0.944, 0.954, 0.961, 0.967, 0.967, 0.97, 0.971, 0.973, 0.974, 0.975}
        venturi_sharp_Cs_Hollingshead = {0.146, 0.3, 0.401, 0.498, 0.554, 0.596, 0.65, 0.688, 0.715, 0.801, 0.841, 0.884, 0.914, 0.94, 0.947, 0.944, 0.952, 0.959, 0.962, 0.963, 0.965, 0.967, 0.967, 0.967}


        CONE_METER_C = 0.82
        'Constant loss coefficient for flow cone meters

        ROUGH_WELDED_CONVERGENT_VENTURI_TUBE_C = 0.985
        'Constant loss coefficient for rough-welded convergent venturi tubes

        MACHINED_CONVERGENT_VENTURI_TUBE_C = 0.995
        'Constant loss coefficient for machined convergent venturi tubes

        AS_CAST_VENTURI_TUBE_C = 0.984
        'Constant loss coefficient for as-cast venturi tubes

        ISO_15377_CONICAL_ORIFICE_C = 0.734
        'Constant loss coefficient for conical orifice plates according to ISO 15377

        cone_Res_Hollingshead = {1.0, 5.0, 10.0, 20.0, 30.0, 40.0, 60.0, 80.0, 100.0, 150.0, 200.0, 300.0, 500.0, 1000.0, 2000.0, 3000.0, 4000.0, 5000.0, 7500.0, 10000.0, 20000.0, 30000.0, 100000.0, 1000000.0, 10000000.0, 50000000.0}
        cone_logRes_Hollingshead = {0, 1.6094379124341, 2.30258509299405, 2.99573227355399, 3.40119738166216, 3.68887945411394, 4.0943445622221,
            4.38202663467388, 4.60517018598809, 5.01063529409626, 5.29831736654804, 5.7037824746562, 6.21460809842219, 6.90775527898214, 7.60090245954208,
            8.00636756765025, 8.29404964010203, 8.51719319141624, 8.9226582995244, 9.21034037197618, 9.90348755253613, 10.3089526606443,
            11.5129254649702, 13.8155105579643, 16.1180956509583, 17.7275335633924}
        cone_betas_Hollingshead = {0.6611, 0.6995, 0.8203}

        cone_beta_6611_Hollingshead_Cs = {0.066, 0.147, 0.207, 0.289, 0.349, 0.396, 0.462, 0.506, 0.537, 0.588, 0.622, 0.661, 0.7, 0.727, 0.75, 0.759, 0.763, 0.765,
        0.767, 0.773, 0.778, 0.789, 0.804, 0.803, 0.805, 0.802}
        cone_beta_6995_Hollingshead_Cs = {0.067, 0.15, 0.21, 0.292, 0.35, 0.394, 0.458, 0.502, 0.533, 0.584, 0.615, 0.645, 0.682, 0.721, 0.742, 0.75, 0.755, 0.757,
            0.763, 0.766, 0.774, 0.781, 0.792, 0.792, 0.79, 0.787}
        cone_beta_8203_Hollingshead_Cs = {0.057, 0.128, 0.182, 0.253, 0.303, 0.343, 0.4, 0.44, 0.472, 0.526, 0.557, 0.605, 0.644, 0.685, 0.705, 0.714, 0.721, 0.722,
            0.724, 0.723, 0.725, 0.731, 0.73, 0.73, 0.741, 0.734}
        cone_Hollingshead_Cs = {cone_beta_6611_Hollingshead_Cs, cone_beta_6995_Hollingshead_Cs, cone_beta_8203_Hollingshead_Cs}


        wedge_Res_Hollingshead = {1.0, 5.0, 10.0, 20.0, 30.0, 40.0, 60.0, 80.0, 100.0, 200.0, 300.0, 400.0, 500.0, 5000.0, 10000.0, 100000.0, 1000000.0, 50000000.0}
        wedge_logRes_Hollingshead = {0, 1.6094379124341, 2.30258509299405, 2.99573227355399, 3.40119738166216, 3.68887945411394, 4.0943445622221,
            4.38202663467388, 4.60517018598809, 5.29831736654804, 5.7037824746562, 5.99146454710798, 6.21460809842219, 8.51719319141624, 9.21034037197618,
            11.5129254649702, 13.8155105579643, 17.7275335633924}

        wedge_beta_5023_Hollingshead = {0.145, 0.318, 0.432, 0.551, 0.61, 0.641, 0.674, 0.69, 0.699, 0.716, 0.721, 0.725, 0.73, 0.729, 0.732, 0.732, 0.731, 0.733}
        wedge_beta_611_Hollingshead = {0.127, 0.28, 0.384, 0.503, 0.567, 0.606, 0.645, 0.663, 0.672, 0.688, 0.694, 0.7, 0.705, 0.7, 0.702, 0.695, 0.699, 0.705}
        wedge_betas_Hollingshead = {0.5023, 0.611}
        wedge_Hollingshead_Cs = {0.145, 0.318, 0.432, 0.551, 0.61, 0.641, 0.674, 0.69, 0.699, 0.716, 0.721, 0.725, 0.73, 0.729, 0.732, 0.732, 0.731, 0.733,
            0.127, 0.28, 0.384, 0.503, 0.567, 0.606, 0.645, 0.663, 0.672, 0.688, 0.694, 0.7, 0.705, 0.7, 0.702, 0.695, 0.699, 0.705}

        Dim c, epsilon, V, Re_D ', 'beta
        Dim Results(0 To 1, 0 To 0)

        c = 0.0
        epsilon = 0.0

        ' Translate default meter type to implementation specific correlation
        If meter_type = CONCENTRIC_ORIFICE Then
            meter_type = ISO_5167_ORIFICE
        ElseIf meter_type = ECCENTRIC_ORIFICE Then
            meter_type = ISO_15377_ECCENTRIC_ORIFICE
        ElseIf meter_type = CONICAL_ORIFICE Then
            meter_type = ISO_15377_CONICAL_ORIFICE
        ElseIf meter_type = QUARTER_CIRCLE_ORIFICE Then
            meter_type = ISO_15377_QUARTER_CIRCLE_ORIFICE
        ElseIf meter_type = SEGMENTAL_ORIFICE Then
            meter_type = MILLER_SEGMENTAL_ORIFICE
        End If

        If meter_type = ISO_5167_ORIFICE Then
            c = C_Reader_Harris_Gallagher(D, D2, rho, mu, m, taps)
            epsilon = Orifice_expansibility(D, D2, P1, P2, k)
        ElseIf meter_type = ISO_15377_ECCENTRIC_ORIFICE Then
            c = C_eccentric_orifice_ISO_15377_1998(D, D2)
            epsilon = Orifice_expansibility(D, D2, P1, P2, k)
        ElseIf meter_type = ISO_15377_QUARTER_CIRCLE_ORIFICE Then
            c = C_quarter_circle_orifice_ISO_15377_1998(D, D2)
            epsilon = Orifice_expansibility(D, D2, P1, P2, k)
        ElseIf meter_type = ISO_15377_CONICAL_ORIFICE Then
            c = ISO_15377_CONICAL_ORIFICE_C
            ' Average of concentric square edge orifice and ISA 1932 nozzles
            epsilon = 0.5 * (Orifice_expansibility(D, D2, P1, P2, k) + Nozzle_expansibility(D, D2, P1, P2, k))
        ElseIf IsInArray(meter_type, {MILLER_ORIFICE, MILLER_ECCENTRIC_ORIFICE, MILLER_SEGMENTAL_ORIFICE, MILLER_QUARTER_CIRCLE_ORIFICE}) Then
            c = C_Miller_1996(D, D2, rho, mu, m, meter_type, taps, tap_position)
            epsilon = Orifice_expansibility(D, D2, P1, P2, k)
        ElseIf meter_type = MILLER_CONICAL_ORIFICE Then
            c = C_Miller_1996(D, D2, rho, mu, m, meter_type, taps, tap_position)
            epsilon = 0.5 * (Orifice_expansibility(D, D2, P1, P2, k) + Nozzle_expansibility(D, D2, P1, P2, k))
        ElseIf meter_type = LONG_RADIUS_NOZZLE Then
            epsilon = Nozzle_expansibility(D, D2, P1, P2, k)
            c = C_long_radius_nozzle(D, D2, rho, mu, m)
        ElseIf meter_type = ISA_1932_NOZZLE Then
            epsilon = Nozzle_expansibility(D, D2, P1, P2, k)
            c = C_ISA_1932_nozzle(D, D2, rho, mu, m)
        ElseIf meter_type = VENTURI_NOZZLE Then
            epsilon = Nozzle_expansibility(D, D2, P1, P2, k)
            c = C_venturi_nozzle(D, D2)
        ElseIf meter_type = AS_CAST_VENTURI_TUBE Then
            epsilon = Nozzle_expansibility(D, D2, P1, P2, k)
            c = AS_CAST_VENTURI_TUBE_C
        ElseIf meter_type = MACHINED_CONVERGENT_VENTURI_TUBE Then
            epsilon = Nozzle_expansibility(D, D2, P1, P2, k)
            c = MACHINED_CONVERGENT_VENTURI_TUBE_C
        ElseIf meter_type = ROUGH_WELDED_CONVERGENT_VENTURI_TUBE Then
            epsilon = Nozzle_expansibility(D, D2, P1, P2, k)
            c = ROUGH_WELDED_CONVERGENT_VENTURI_TUBE_C
        ElseIf meter_type = CONE_METER Then
            epsilon = Cone_meter_expansibility_Stewart(D, D2, P1, P2, k)
            c = CONE_METER_C
        ElseIf meter_type = WEDGE_METER Then
            'beta = Diameter_ratio_wedge_meter(D, D2)
            epsilon = Nozzle_expansibility(D, D2, P1, P1, k, Diameter_ratio_wedge_meter(D, D2))
            c = C_wedge_meter_ISO_5167_6_2017(D, D2)
        ElseIf meter_type = HOLLINGSHEAD_ORIFICE Then
            V = m / ((0.25 * PI * D * D) * rho)
            Re_D = rho * V * D / mu
            c = Orifice_std_Hollingshead_tck(D2 / D, Math.Log(Re_D))
            epsilon = Orifice_expansibility(D, D2, P1, P2, k)
        ElseIf meter_type = HOLLINGSHEAD_VENTURI_SMOOTH Then
            V = m / ((0.25 * PI * D * D) * rho)
            Re_D = rho * V * D / mu
            c = Interp(Math.Log(Re_D), venturi_logRes_Hollingshead, venturi_smooth_Cs_Hollingshead, , , True)
            epsilon = Nozzle_expansibility(D, D2, P1, P2, k)
        ElseIf meter_type = HOLLINGSHEAD_VENTURI_SHARP Then
            V = m / ((0.25 * PI * D * D) * rho)
            Re_D = rho * V * D / mu
            c = Interp(Math.Log(Re_D), venturi_logRes_Hollingshead, venturi_sharp_Cs_Hollingshead, , , True)
            epsilon = Nozzle_expansibility(D, D2, P1, P2, k)
        ElseIf meter_type = HOLLINGSHEAD_CONE Then
            V = m / ((0.25 * PI * D * D) * rho)
            Re_D = rho * V * D / mu
            'beta = Diameter_ratio_cone_meter(D, D2)
            c = Cone_Hollingshead_tck(D2 / D, Math.Log(Re_D))
            epsilon = Cone_meter_expansibility_Stewart(D, D2, P1, P2, k)
        ElseIf meter_type = HOLLINGSHEAD_WEDGE Then
            V = m / ((0.25 * PI * D * D) * rho)
            Re_D = rho * V * D / mu
            'beta = Diameter_ratio_wedge_meter(D, D2)
            c = Wedge_Hollingshead_tck(D2 / D, Math.Log(Re_D))
            epsilon = Nozzle_expansibility(D, D2, P1, P1, k, Diameter_ratio_wedge_meter(D, D2))
        ElseIf meter_type = UNSPECIFIED_METER Then
            epsilon = Orifice_expansibility(D, D2, P1, P2, k)
            If C_specified <= 0 Then
                MsgBox("For unspecified meter type, C_specified is required")
            End If
        End If

        If Not C_specified <= 0.0 Then c = C_specified
        If Not epsilon_specified <= 0.0 Then epsilon = epsilon_specified

        If c = 0 And epsilon = 0.0 Then
            Differential_pressure_meter_C_epsilon = 0.0
        Else
            Results(0, 0) = c
            Results(1, 0) = epsilon

            Differential_pressure_meter_C_epsilon = Results
        End If

    End Function

    Function Err_dp_meter_solver(D As Double, D2 As Double, m As Double, P1 As Double, P2 As Double, rho As Double, mu As Double, k As Double, meter_type As String, taps As String, tap_position As String, Optional C_specified As Double = 0.0, Optional epsilon_specified As Double = 0.0) As Double

        Dim c, epsilon, m_cal
        Dim Results

        Results = Differential_pressure_meter_C_epsilon(D, D2, m, P1, P2, rho, mu, k, meter_type, taps, tap_position, C_specified, epsilon_specified)

        c = Results(0, 0)
        epsilon = Results(1, 0)

        m_cal = Flow_meter_discharge(D, D2, P1, P2, rho, c, epsilon)
        Err_dp_meter_solver = m - m_cal

    End Function

    Public Function DP_meter_solver_m(D1 As Double, D2 As Double, P1 As Double, P2 As Double, rho As Double, mu As Double, k As Double, Optional meter_type As String = "orifice", Optional taps As String = "D and D/2", Optional tap_position As String = "180 degree", Optional C_specified As Double = 0.0, Optional epsilon_specified As Double = 0.0) As Double

        Dim m As Double
        Dim m_calc As Double
        Dim c As Double
        Dim epsilon
        Dim Results

        'Assume C = 0.7 for starting point
        c = 0.7
        epsilon = Orifice_expansibility(D1, D2, P1, P2, k)

        'First guess ofmass flow rate
        m = Flow_meter_discharge(D1, D2, P1, P2, rho, c, epsilon)

        Do
            m_calc = m
            Results = Differential_pressure_meter_C_epsilon(D1, D2, m, P1, P2, rho, mu, k, meter_type, taps, tap_position, C_specified, epsilon_specified)
            'c = C_Reader_Harris_Gallagher(D1, D2, rho, mu, m, taps)
            c = Results(0, 0)
            epsilon = Results(1, 0)
            m = Flow_meter_discharge(D1, D2, P1, P2, rho, c, epsilon)
        Loop While Math.Abs(m - m_calc) > 0.0000000001

        DP_meter_solver_m = m

    End Function

    Public Function DP_meter_solver_D2(D1 As Double, m As Double, P1 As Double, P2 As Double, rho As Double, mu As Double, k As Double, Optional meter_type As String = "orifice", Optional taps As String = "corner", Optional tap_position As String = "180 degree", Optional C_specified As Double = 0.0, Optional epsilon_specified As Double = 0.0) As Double

        Dim a, b, c, d, f_a, f_b, f_c, s, f_s, mflag, delta
        Dim flag1, flag2, flag3, flag4, flag5

        a = D1 * 0.00005
        b = D1 * (1 - 0.000000001)
        delta = 0.000000001

        ' Brent method for root finding
        f_a = Err_dp_meter_solver(D1, a, m, P1, P2, rho, mu, k, meter_type, taps, tap_position, C_specified, epsilon_specified)
        f_b = Err_dp_meter_solver(D1, b, m, P1, P2, rho, mu, k, meter_type, taps, tap_position, C_specified, epsilon_specified)

        If f_a * f_b >= 0 Then
            DP_meter_solver_D2 = -999
            Exit Function
        End If

        If Math.Abs(f_a) < Math.Abs(f_b) Then
            'swap a,b
            c = a
            a = b
            b = c
        End If

        c = a
        d = b
        mflag = True

        Do
            f_a = Err_dp_meter_solver(D1, a, m, P1, P2, rho, mu, k, meter_type, taps, tap_position, C_specified, epsilon_specified)
            f_b = Err_dp_meter_solver(D1, b, m, P1, P2, rho, mu, k, meter_type, taps, tap_position, C_specified, epsilon_specified)
            f_c = Err_dp_meter_solver(D1, c, m, P1, P2, rho, mu, k, meter_type, taps, tap_position, C_specified, epsilon_specified)

            If Not f_a = f_c And Not f_b = f_c Then
                s = a * f_b * f_c / ((f_a - f_b) * (f_a - f_c)) + b * f_a * f_c / ((f_b - f_a) * (f_b - f_c)) + c * f_a * f_b / ((f_c - f_a) * (f_c - f_b))
            Else
                s = b - f_b * (b - a) / (f_b - f_a)
            End If

            ' determine conditions
            flag1 = Not (s >= (3 * a + b) / 4 And s <= b)
            flag2 = mflag And Math.Abs(s - b) >= Math.Abs((b - c) / 2)
            flag3 = (Not mflag) And Math.Abs(s - b) >= Math.Abs((b - c) / 2)
            flag4 = mflag And Math.Abs(b - c) < delta
            flag5 = (Not mflag) And Math.Abs(c - d) < delta
            If flag1 Or flag2 Or flag3 Or flag4 Or flag5 Then
                s = (a + b) / 2
                mflag = True
            Else
                mflag = False
            End If

            f_s = Err_dp_meter_solver(D1, s, m, P1, P2, rho, mu, k, meter_type, taps, tap_position, C_specified, epsilon_specified)

            d = c
            c = b

            If f_a * f_s < 0 Then
                b = s
            Else
                a = s
            End If

            If Math.Abs(f_a) < Math.Abs(f_b) Then
                c = a
                a = b
                b = c
            End If
        Loop Until Math.Abs(f_b) < delta Or Math.Abs(b - a) < delta

        DP_meter_solver_D2 = b

    End Function

    Public Function DP_meter_solver_P1(D1 As Double, D2 As Double, m As Double, P2 As Double, rho As Double, mu As Double, k As Double, Optional meter_type As String = "orifice", Optional taps As String = "corner", Optional tap_position As String = "180 degree", Optional C_specified As Double = 0.0, Optional epsilon_specified As Double = 0.0) As Double

        Dim a, b, c, d, f_a, f_b, f_c, s, f_s, mflag, delta As Double
        Dim flag1, flag2, flag3, flag4, flag5

        a = P2 * (1 + 0.000000001)
        b = P2 * 1.5
        delta = 0.000000001

        ' Brent method for root finding
        f_a = Err_dp_meter_solver(D1, D2, m, a, P2, rho, mu, k, meter_type, taps, tap_position, C_specified, epsilon_specified)
        f_b = Err_dp_meter_solver(D1, D2, m, b, P2, rho, mu, k, meter_type, taps, tap_position, C_specified, epsilon_specified)

        If f_a * f_b >= 0 Then
            DP_meter_solver_P1 = -999
            Exit Function
        End If

        If Math.Abs(f_a) < Math.Abs(f_b) Then
            'swap a,b
            c = a
            a = b
            b = c
        End If

        c = a
        d = b
        mflag = True

        Do
            f_a = Err_dp_meter_solver(D1, D2, m, a, P2, rho, mu, k, meter_type, taps, tap_position, C_specified, epsilon_specified)
            f_b = Err_dp_meter_solver(D1, D2, m, b, P2, rho, mu, k, meter_type, taps, tap_position, C_specified, epsilon_specified)
            f_c = Err_dp_meter_solver(D1, D2, m, c, P2, rho, mu, k, meter_type, taps, tap_position, C_specified, epsilon_specified)

            If Not f_a = f_c And Not f_b = f_c Then
                s = a * f_b * f_c / ((f_a - f_b) * (f_a - f_c)) + b * f_a * f_c / ((f_b - f_a) * (f_b - f_c)) + c * f_a * f_b / ((f_c - f_a) * (f_c - f_b))
            Else
                s = b - f_b * (b - a) / (f_b - f_a)
            End If

            ' determine conditions
            flag1 = Not (s >= (3 * a + b) / 4 And s <= b)
            flag2 = mflag And Math.Abs(s - b) >= Math.Abs((b - c) / 2)
            flag3 = (Not mflag) And Math.Abs(s - b) >= Math.Abs((b - c) / 2)
            flag4 = mflag And Math.Abs(b - c) < delta
            flag5 = (Not mflag) And Math.Abs(c - d) < delta
            If flag1 Or flag2 Or flag3 Or flag4 Or flag5 Then
                s = (a + b) / 2
                mflag = True
            Else
                mflag = False
            End If

            f_s = Err_dp_meter_solver(D1, D2, m, s, P2, rho, mu, k, meter_type, taps, tap_position, C_specified, epsilon_specified)

            d = c
            c = b

            If f_a * f_s < 0 Then
                b = s
            Else
                a = s
            End If

            If Math.Abs(f_a) < Math.Abs(f_b) Then
                c = a
                a = b
                b = c
            End If
        Loop Until Math.Abs(f_b) < delta Or Math.Abs(b - a) < delta

        DP_meter_solver_P1 = b

    End Function

    Public Function DP_meter_solver_P2(D1 As Double, D2 As Double, m As Double, P1 As Double, rho As Double, mu As Double, k As Double, Optional meter_type As String = "orifice", Optional taps As String = "corner", Optional tap_position As String = "180 degree", Optional C_specified As Double = 0.0, Optional epsilon_specified As Double = 0.0) As Double

        Dim a, b, c, D, f_a, f_b, f_c, s, f_s, mflag, delta
        Dim flag1, flag2, flag3, flag4, flag5

        a = P1 * 0.01
        b = P1 * (1 - 0.000000001)
        delta = 0.000000001

        ' Brent method for root finding
        f_a = Err_dp_meter_solver(D1, D2, m, P1, a, rho, mu, k, meter_type, taps, tap_position, C_specified, epsilon_specified)
        f_b = Err_dp_meter_solver(D1, D2, m, P1, b, rho, mu, k, meter_type, taps, tap_position, C_specified, epsilon_specified)

        If f_a * f_b >= 0 Then
            DP_meter_solver_P2 = -999
            Exit Function
        End If

        If Math.Abs(f_a) < Math.Abs(f_b) Then
            'swap a,b
            c = a
            a = b
            b = c
        End If

        c = a
        D = b
        mflag = True

        Do
            f_a = Err_dp_meter_solver(D1, D2, m, P1, a, rho, mu, k, meter_type, taps, tap_position, C_specified, epsilon_specified)
            f_b = Err_dp_meter_solver(D1, D2, m, P1, b, rho, mu, k, meter_type, taps, tap_position, C_specified, epsilon_specified)
            f_c = Err_dp_meter_solver(D1, D2, m, P1, c, rho, mu, k, meter_type, taps, tap_position, C_specified, epsilon_specified)

            If Not f_a = f_c And Not f_b = f_c Then
                s = a * f_b * f_c / ((f_a - f_b) * (f_a - f_c)) + b * f_a * f_c / ((f_b - f_a) * (f_b - f_c)) + c * f_a * f_b / ((f_c - f_a) * (f_c - f_b))
            Else
                s = b - f_b * (b - a) / (f_b - f_a)
            End If

            ' determine conditions
            flag1 = Not (s >= (3 * a + b) / 4 And s <= b)
            flag2 = mflag And Math.Abs(s - b) >= Math.Abs((b - c) / 2)
            flag3 = (Not mflag) And Math.Abs(s - b) >= Math.Abs((b - c) / 2)
            flag4 = mflag And Math.Abs(b - c) < delta
            flag5 = (Not mflag) And Math.Abs(c - D) < delta
            If flag1 Or flag2 Or flag3 Or flag4 Or flag5 Then
                s = (a + b) / 2
                mflag = True
            Else
                mflag = False
            End If

            f_s = Err_dp_meter_solver(D1, D2, m, P1, s, rho, mu, k, meter_type, taps, tap_position, C_specified, epsilon_specified)

            D = c
            c = b

            If f_a * f_s < 0 Then
                b = s
            Else
                a = s
            End If

            If Math.Abs(f_a) < Math.Abs(f_b) Then
                c = a
                a = b
                b = c
            End If
        Loop Until Math.Abs(f_b) < delta Or Math.Abs(b - a) < delta

        DP_meter_solver_P2 = b

    End Function
End Module
