﻿Imports ExcelDna.Integration

Public Module nominal_pipe_schedule

    <ExcelFunction(Description:="Return pipe nominal ID", Category:="GCME E-PT | Pipe Sch")>
    Public Function PipeNorminalID(idx As Integer) As String
        Dim arr = New String() {"1/8", "1/4", "3/8", "1/2", "3/4", "1", "1 1/4", "1 1/2", "2", "2 1/2", "3", "3 1/2", "4", "5", "6", "8", "10", "12", "14", "16", "18", "20", "22", "24", "26", "28", "30", "32", "34", "36", "38", "40", "42", "44", "46", "48"}

        PipeNorminalID = arr(idx)
    End Function

    <ExcelFunction(Description:="Return pipe OD in inch", Category:="GCME E-PT | Pipe Sch")>
    Public Function PipeOD(<ExcelArgument(Description:="Norminal pipe ID as string")> NormID As String) As Double
        Dim strNormID = New String() {"1/8", "1/4", "3/8", "1/2", "3/4", "1", "1 1/4", "1 1/2", "2", "2 1/2", "3", "3 1/2", "4", "5", "6", "8", "10", "12", "14", "16", "18", "20", "22", "24", "26", "28", "30", "32", "34", "36", "38", "40", "42", "44", "46", "48"}
        Dim dblNormOD = New Double() {0.405, 0.54, 0.675, 0.84, 1.05, 1.315, 1.66, 1.9, 2.375, 2.875, 3.5, 4, 4.5, 5.563, 6.625, 8.625, 10.75, 12.75, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36, 38, 40, 42, 44, 46, 48}

        PipeOD = dblNormOD(Array.IndexOf(strNormID, NormID))
    End Function

    <ExcelFunction(Description:="Return pipe schedule", Category:="GCME E-PT | Pipe Sch")>
    Public Function PipeSchedule(idx As Integer) As String
        Dim arr = New String() {"5", "5S", "10", "10S", "20", "30", "40", "40S", "STD", "60", "80", "80S", "XS", "100", "120", "140", "160", "XXS"}

        PipeSchedule = arr(idx)
    End Function

    <ExcelFunction(Description:="Length of pipe nominal ID", Category:="GCME E-PT | Pipe Sch")>
    Public Function Len_PipeNorminalID() As Integer
        Dim arr = New String() {"1/8", "1/4", "3/8", "1/2", "3/4", "1", "1 1/4", "1 1/2", "2", "2 1/2", "3", "3 1/2", "4", "5", "6", "8", "10", "12", "14", "16", "18", "20", "22", "24", "26", "28", "30", "32", "34", "36", "38", "40", "42", "44", "46", "48"}
        Len_PipeNorminalID = arr.Length
    End Function

    <ExcelFunction(Description:="Length of pipe OD", Category:="GCME E-PT | Pipe Sch")>
    Public Function Len_PipeOD() As Integer
        Len_PipeOD = Len_PipeNorminalID()
    End Function

    <ExcelFunction(Description:="Length of pipe schedule", Category:="GCME E-PT | Pipe Sch")>
    Public Function Len_PipeSchedule() As Integer
        Dim arr = New String() {"5", "5S", "10", "10S", "20", "30", "40", "40S", "STD", "60", "80", "80S", "XS", "100", "120", "140", "160", "XXS"}

        Len_PipeSchedule = arr.Length
    End Function

    <ExcelFunction(Description:="Return pipe wall thickness", Category:="GCME E-PT | Pipe Sch")>
    Public Function PipeThickness(<ExcelArgument(Description:="Norminal Pipe ID")> NormID As String, <ExcelArgument(Description:="Pipe Schedule")> Sch As String) As Double
        Dim strNormID = New String() {"1/8", "1/4", "3/8", "1/2", "3/4", "1", "1 1/4", "1 1/2", "2", "2 1/2", "3", "3 1/2", "4", "5", "6", "8", "10", "12", "14", "16", "18", "20", "22", "24", "26", "28", "30", "32", "34", "36", "38", "40", "42", "44", "46", "48"}
        Dim strSchedule = New String() {"5", "5S", "10", "10S", "20", "30", "40", "40S", "STD", "60", "80", "80S", "XS", "100", "120", "140", "160", "XXS"}

        Dim arr = New Double(,) {{0, 0, 0.049, 0.049, 0, 0, 0.068, 0.068, 0.068, 0, 0.095, 0.095, 0.095, 0, 0, 0, 0, 0},
                                    {0, 0, 0.065, 0.065, 0, 0, 0.088, 0.088, 0.088, 0, 0.119, 0.119, 0.119, 0, 0, 0, 0, 0},
                                    {0, 0, 0.065, 0.065, 0, 0.073, 0.091, 0.091, 0.091, 0, 0.126, 0.126, 0.126, 0, 0, 0, 0, 0},
                                    {0.065, 0.065, 0.083, 0.083, 0, 0.095, 0.109, 0.109, 0.109, 0, 0.147, 0.147, 0.147, 0, 0, 0, 0.188, 0.294},
                                    {0.065, 0.065, 0.083, 0.083, 0, 0.095, 0.113, 0.113, 0.113, 0, 0.154, 0.154, 0.154, 0, 0, 0, 0.219, 0.308},
                                    {0.065, 0.065, 0.109, 0.109, 0, 0.114, 0.133, 0.133, 0.133, 0, 0.179, 0.179, 0.179, 0, 0, 0, 0.25, 0.358},
                                    {0.065, 0.065, 0.109, 0.109, 0, 0.117, 0.14, 0.14, 0.14, 0, 0.191, 0.191, 0.191, 0, 0, 0, 0.25, 0.382},
                                    {0.065, 0.065, 0.109, 0.109, 0, 0.125, 0.145, 0.145, 0.145, 0, 0.2, 0.2, 0.2, 0, 0, 0, 0.281, 0.4},
                                    {0.065, 0.065, 0.109, 0.109, 0, 0.125, 0.154, 0.154, 0.154, 0, 0.218, 0.218, 0.218, 0, 0, 0, 0.344, 0.436},
                                    {0.083, 0.083, 0.12, 0.12, 0, 0.188, 0.203, 0.203, 0.203, 0, 0.276, 0.276, 0.276, 0, 0, 0, 0.375, 0.552},
                                    {0.083, 0.083, 0.12, 0.12, 0, 0.188, 0.216, 0.216, 0.216, 0, 0.3, 0.3, 0.3, 0, 0, 0, 0.438, 0.6},
                                    {0.083, 0.083, 0.12, 0.12, 0, 0.188, 0.226, 0.226, 0.226, 0, 0.318, 0.318, 0.318, 0, 0, 0, 0, 0.636},
                                    {0.083, 0.083, 0.12, 0.12, 0, 0.188, 0.237, 0.237, 0.237, 0, 0.337, 0.337, 0.337, 0, 0.438, 0, 0.531, 0.674},
                                    {0.109, 0.109, 0.134, 0.134, 0, 0, 0.258, 0.258, 0.258, 0, 0.375, 0.375, 0.375, 0, 0.5, 0, 0.625, 0.75},
                                    {0.109, 0.109, 0.134, 0.134, 0, 0, 0.28, 0.28, 0.28, 0, 0.432, 0.432, 0.432, 0, 0.562, 0, 0.719, 0.864},
                                    {0.109, 0.109, 0.148, 0.148, 0.25, 0.277, 0.322, 0.322, 0.322, 0.406, 0.5, 0.5, 0.5, 0.594, 0.719, 0.812, 0.906, 0.875},
                                    {0.134, 0.134, 0.165, 0.165, 0.25, 0.307, 0.365, 0.365, 0.365, 0.5, 0.594, 0.5, 0.5, 0.719, 0.844, 1, 1.125, 1},
                                    {0.156, 0.156, 0.18, 0.18, 0.25, 0.33, 0.406, 0.375, 0.375, 0.562, 0.688, 0.5, 0.5, 0.844, 1, 1.125, 1.312, 1},
                                    {0.156, 0.156, 0.25, 0.188, 0.312, 0.375, 0.438, 0.375, 0.375, 0.594, 0.75, 0.5, 0.5, 0.938, 1.094, 1.25, 1.406, 0},
                                    {0.165, 0.165, 0.25, 0.188, 0.312, 0.375, 0.5, 0.375, 0.375, 0.656, 0.844, 0.5, 0.5, 1.031, 1.219, 1.438, 1.594, 0},
                                    {0.165, 0.165, 0.25, 0.188, 0.312, 0.438, 0.562, 0.375, 0.375, 0.75, 0.938, 0.5, 0.5, 1.156, 1.375, 1.562, 1.781, 0},
                                    {0.188, 0.188, 0.25, 0.218, 0.375, 0.5, 0.594, 0.375, 0.375, 0.812, 1.031, 0.5, 0.5, 1.281, 1.5, 1.75, 1.969, 0},
                                    {0.188, 0.188, 0.25, 0.218, 0.375, 0.5, 0, 0, 0.375, 0.875, 1.125, 0, 0.5, 1.375, 1.625, 1.875, 2.125, 0},
                                    {0.218, 0.218, 0.25, 0.25, 0.375, 0.562, 0.688, 0.375, 0.375, 0.969, 1.219, 0.5, 0.5, 1.531, 1.812, 2.062, 2.344, 0},
                                    {0, 0, 0.312, 0, 0.5, 0, 0, 0, 0.375, 0, 0, 0, 0.5, 0, 0, 0, 0, 0},
                                    {0, 0, 0.312, 0, 0.5, 0.625, 0, 0, 0.375, 0, 0, 0, 0.5, 0, 0, 0, 0, 0},
                                    {0.25, 0.25, 0.312, 0.312, 0.5, 0.625, 0, 0, 0.375, 0, 0, 0, 0.5, 0, 0, 0, 0, 0},
                                    {0, 0, 0.312, 0, 0.5, 0.625, 0.688, 0, 0.375, 0, 0, 0, 0.5, 0, 0, 0, 0, 0},
                                    {0, 0, 0.312, 0, 0.5, 0.625, 0.688, 0, 0.375, 0, 0, 0, 0.5, 0, 0, 0, 0, 0},
                                    {0, 0, 0.312, 0, 0.5, 0.625, 0.75, 0, 0.375, 0, 0, 0, 0.5, 0, 0, 0, 0, 0},
                                    {0, 0, 0, 0, 0, 0, 0, 0, 0.375, 0, 0, 0, 0.5, 0, 0, 0, 0, 0},
                                    {0, 0, 0, 0, 0, 0, 0, 0, 0.375, 0, 0, 0, 0.5, 0, 0, 0, 0, 0},
                                    {0, 0, 0, 0, 0, 0, 0, 0, 0.375, 0, 0, 0, 0.5, 0, 0, 0, 0, 0},
                                    {0, 0, 0, 0, 0, 0, 0, 0, 0.375, 0, 0, 0, 0.5, 0, 0, 0, 0, 0},
                                    {0, 0, 0, 0, 0, 0, 0, 0, 0.375, 0, 0, 0, 0.5, 0, 0, 0, 0, 0},
                                    {0, 0, 0, 0, 0, 0, 0, 0, 0.375, 0, 0, 0, 0.5, 0, 0, 0, 0, 0}}

        PipeThickness = arr(Array.IndexOf(strNormID, NormID), Array.IndexOf(strSchedule, UCase(Sch)))
    End Function

    <ExcelFunction(Description:="Return pipe inside diamtere in inch", Category:="GCME E-PT | Pipe Sch")>
    Public Function PipeID(<ExcelArgument(Description:="Norminal Pipe ID")> NormID As String, <ExcelArgument(Description:="Pipe Schedule")> Sch As String) As Double
        Dim od = PipeOD(NormID)
        Dim tck = PipeThickness(NormID, Sch)

        PipeID = od - 2 * tck
    End Function

    <ExcelFunction(Description:="Return available pipe schedule for input norminal ID", Category:="GCME E-PT | Pipe Sch")>
    Public Function AvailableSchedule(<ExcelArgument(Description:="Norminal Pipe ID")> NormID As String) As String
        Dim S1 As String
        Dim i As Integer

        S1 = ""

        For i = 0 To Len_PipeSchedule() - 1
            If PipeThickness(NormID, PipeSchedule(i)) > 0 Then
                If S1 = "" Then
                    S1 = PipeSchedule(i)
                Else
                    S1 += "," + PipeSchedule(i)
                End If
            End If
        Next

        AvailableSchedule = S1

    End Function

End Module
