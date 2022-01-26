Function MINI_2(Nb1 As Integer, Nb2 As Integer) As Integer
    If Nb1 <= Nb2 Then
        MINI_2 = Nb1
    Else
        MINI_2 = Nb2
    End If
End Function

Function MAXI_2(Nb1 As Integer, Nb2 As Integer) As Integer
    If Nb1 <= Nb2 Then
        MAXI_2 = Nb2
    Else
        MAXI_2 = Nb1
    End If
End Function

Function SommeEntreEntierImpaire(ByVal N As Integer, ByVal M As Integer) As Integer

    Dim Sum As Integer
    Dim i As Integer
    
    For i = MINI_2(N, M) To MAXI_2(N, M) 'pas besoin de faire des distinctions
        If i Mod 2 = 1 Then            'de cas sur le plus grand des deux
            Sum = Sum + i
        End If
    Next
    
   SommeEntreEntierImpaire = Sum
End Function
