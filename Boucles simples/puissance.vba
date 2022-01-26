Function PUISSANCE(ByVal N As Integer, ByVal M As Integer) As Long
    
    Dim i As Integer
    Dim resultat As Long
    
    resultat = 1
    For i = 1 To M
        resultat = resultat * N
    Next
     PUISSANCE = resultat
     
End Function
