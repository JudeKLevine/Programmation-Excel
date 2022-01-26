Function nombreJours(Nbmois As Integer, years As Integer) As Integer

    Select Case Nbmois
        Case 1, 3, 5, 7, 8, 10, 12
            nombreJours = 31
        Case 2
            If bissextile(years) = True Then
                nombreJours = 28
            Else
                nombreJours = 29
            End If
        Case 4, 6, 9, 11
            nombreJours = 30
    End Select

End Function
