Function SommeEntreEntier(ByVal N As Integer, ByVal M As Integer) As Integer

    Dim Sum As Integer
    Dim i As Integer
    
    For i = mini2(N, M) To maxi2(N, M) 'pas besoin de faire des distinctions
        Sum = Sum + i                  'de cas sur le plus grand des deux
    Next
    
    SommeEntreEntier = Sum

End Function
