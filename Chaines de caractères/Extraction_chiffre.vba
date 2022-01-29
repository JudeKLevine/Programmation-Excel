Function Extraction(ByVal S As String) As String

    Dim a As Double
    
    S = Replace(S, ",", ".")
    
    Do While Len(S) > 0
    a = Val(S)
    Select Case a
        Case 0
            If StrComp(Mid(S, 1, 1), "0", 0) = 0 Then
                Debug.Print a
            End If
            S = Mid(S, 2, Len(S))
        Case Is <> 0
            Debug.Print a
            S = Mid(S, Len(CStr(a)) + 1, Len(S))
        End Select
    Loop
    Extraction = "LES VALEURS SONT AU DESSUS"
End Function
