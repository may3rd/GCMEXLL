Imports System.Net.Mime.MediaTypeNames

Module numerical
    Public Function Interp(x, dx, dy, Optional left = -999.99, Optional right = -999.99, Optional extrapolate = False) As Double

        Dim j As Integer

        If x < dx(LBound(dx)) Then
            If left <> -999.99 Then
                Interp = left
                Exit Function
            ElseIf extrapolate Then
                j = LBound(dx)
                Interp = (dy(j + 1) - dy(j)) / (dx(j + 1) - dx(j)) * (x - dx(j)) + dy(j)
                Exit Function
            Else
                Interp = dy(LBound(dy))
            End If
        ElseIf x > dx(UBound(dx)) Then
            If right <> -999.99 Then
                Interp = right
                Exit Function
            ElseIf extrapolate Then
                j = UBound(dx) - 1
                Interp = (dy(j + 1) - dy(j)) / (dx(j + 1) - dx(j)) * (x - dx(j)) + dy(j)
                Exit Function
            Else
                Interp = dy(UBound(dy))
            End If
        Else
            j = Array.IndexOf(dx, x)
            If j = UBound(dx) - LBound(dx) + 1 Then
                Interp = dy(UBound(dy))
                Exit Function
            Else
                Interp = (dy(j + 1) - dy(j)) / (dx(j + 1) - dx(j)) * (x - dx(j)) + dy(j)
                Exit Function
            End If
        End If

    End Function

    Public Function Interp2d_linear(x, y, xs, ys, vals) As Double
        'Dim result As Double
        Dim i, i0, i1 As Short
        Dim v_low, v_high, y_low, y_high As Double

        If y < ys(LBound(ys)) Then
            i0 = LBound(ys)
            i1 = i0 + 1
        ElseIf y > ys(UBound(ys)) Then
            i1 = UBound(ys)
            i0 = i1 - 1
        Else
            For i = LBound(ys) To UBound(ys)
                If ys(i) >= y Then
                    i1 = i
                    i0 = i - 1
                    Exit For
                End If
            Next
        End If

        y_low = ys(i0)
        y_high = ys(i1)

        v_low = Interp(x, xs, vals(i0),,, True)
        v_high = Interp(x, xs, vals(i1),,, True)

        Interp2d_linear = v_low + (v_high - v_low) * (y - y_low) / (y_high - y_low)
    End Function

    ' Math for Mere Mortals
    ' BicubicLagrangeInterpolation_v1

    ' Returns an interpolated point using local bicubic interpolation on table Table.
    ' The top row and left column in Table must be headers.
    Public Function BicubicInterpolation(Table As Array, TopArray As Array, LeftArray As Array, TopPoint As Double, LeftPoint As Double) As Double
        Dim LeftMinIndex As Long
        Dim TopMinIndex As Long
        Dim i As Long
        Dim j As Long
        Dim Numerator As Double
        Dim Denominator As Double
        Dim Weights(0 To 3) As Double
        Dim Subset(0 To 3) As Double
        ' Choose TopIndex and LeftIndex that yield the 4x4 subset that we will interpolate over...
        ' Which index is the lowest on the top side?
        TopMinIndex = FindIndexBelow(TopArray, TopPoint)
        ' The leftmost item should be invalid, so the return value should be higher than 1.
        If TopMinIndex <= 1 Then
            ' Slide the range over to the right if it is lower than the source data domain.
            TopMinIndex = 2
        End If
        ' Slide the range over to the left if it is higher than the source data domain.
        If TopMinIndex >= TopArray.Length - 2 Then
            TopMinIndex = TopArray.Length - 3
        End If
        ' Which index is the lowest on the left side?
        LeftMinIndex = FindIndexBelow(LeftArray, LeftPoint)
        ' The leftmost item should be invalid, so the return value should be higher than 1.
        If LeftMinIndex <= 1 Then
            ' Slide the range over to the right if it is lower than the source data domain.
            LeftMinIndex = 2
        End If
        ' Slide the range over to the left if it is higher than the source data domain.
        If LeftMinIndex >= LeftArray.Length - 2 Then
            LeftMinIndex = LeftArray.Length - 3
        End If
        ' Determine weights that will be used for all four rows...
        ' Loop once for each weight.
        For i = 0 To 3
            ' Initialize the numerator and denominator.
            Numerator = 1
            Denominator = 1
            ' Loop once for each potential Lagrange coefficient.
            For j = 0 To 3
                If i <> j Then
                    ' Calculate the numerator for this term.
                    Numerator = Numerator * (TopPoint - TopArray(TopMinIndex + j - 1))
                    ' Calculate the denominator for this term.
                    Denominator = Denominator * (TopArray(TopMinIndex + i - 1) - TopArray(TopMinIndex + j - 1))
                End If
            Next
            ' Populate the Weights array with this weight value.
            Weights(i) = Numerator / Denominator
        Next

        ' Generate the 4x1 data subset that will be interpolated over...
        ' Loop once for each interpolated value on the line.
        For i = 0 To 3
            ' Initialize this item in the data subset.
            Subset(i) = 0
            ' Loop once for each Lagrange polynomial term.
            For j = 0 To 3
                ' Include this Lagrange polynomial term in the data subset.
                Subset(i) = Subset(i) + Table(LeftMinIndex + i - 1, TopMinIndex - 1 + j) * Weights(j)
            Next
        Next
        ' Determine weights for the 4x1 data subset, which is the column interpolated within the dataset...
        ' Loop once for each weight.
        For i = 0 To 3
            ' Initialize the numerator and denominator.
            Numerator = 1
            Denominator = 1
            ' Loop once for each potential Lagrange coefficient.
            For j = 0 To 3
                If i <> j Then
                    ' Calculate the numerator for this term.
                    Numerator = Numerator * (LeftPoint - LeftArray(LeftMinIndex - 1 + j))
                    ' Calculate the denominator for this term.
                    Denominator = Denominator * (LeftArray(LeftMinIndex - 1 + i) - LeftArray(LeftMinIndex - 1 + j))
                End If
            Next
            ' Populate the Weights array with this weight value.
            Weights(i) = Numerator / Denominator
        Next
        ' Assume the result is zero.
        BicubicInterpolation = 0
        ' Finish the interpolation to find the interpolated value...
        ' Loop once for each interpolated value on the subset line.
        For i = 0 To 3
            ' The interpolated value is the sum of the product of each Lagrange coefficient and its corresponding function value.
            BicubicInterpolation = BicubicInterpolation + Subset(i) * Weights(i)
        Next
    End Function


    ' Find the index of a value that is less than or equal to Value.
    ' If the dataset appears to be in reverse, find the index ABOVE the value.
    ' If the value A is a Range type, the first cell is ignored.
    Function FindIndexBelow(A As Array, Value As Double) As Long
        Dim i As Long
        ' Assume there is no such value in the array.
        FindIndexBelow = -1

        ' Are the items in reverse order?
        If A(LBound(A)) > A(LBound(A) + 1) Then
            For i = LBound(A) To UBound(A)
                ' Is this array element less than or equal to Value?
                If A(i) >= Value Then
                    ' This is a valid value.
                    FindIndexBelow = i
                Else
                    ' Stop looking.
                    Exit For
                End If
            Next
        Else
            For i = LBound(A) To UBound(A)
                ' Is this array element less than or equal to Value?
                If A(i) <= Value Then
                    ' This is a valid value.
                    FindIndexBelow = i
                Else
                    ' Stop looking.
                    Exit For
                End If
            Next
        End If

    End Function


End Module
