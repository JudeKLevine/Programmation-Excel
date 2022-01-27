Function Nbracine(ByVal Delta As Double) As Integer

    'traduction informatique d'un calcul mathematique
    Select Case Delta
        Case 0
            Nbracine = 1
        Case Is > 0
            Nbracine = 2
        Case Is < 0
            Nbracine = 0
    End Select
            
End Function
