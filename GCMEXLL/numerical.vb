Module numerical
    Public Function Interp(x, dx, dy, Optional left = -999.99, Optional right = -999.99, Optional extrapolate = False) As Double

        Dim j

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

End Module
