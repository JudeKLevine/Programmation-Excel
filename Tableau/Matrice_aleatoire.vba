Function getMatrice(ByVal n As Integer, _
                    ByVal m As Integer, _
                    ByVal min As Integer, _
                    ByVal max As Integer) As Integer()
                    
    Dim Matrice() As Integer
    Dim i As Integer
    Dim j As Integer
    ReDim Matrice(n, m)
    
    For i = 0 To n
        For j = 0 To m
            Matrice(i, j) = Int((max - min + 1) * Rnd + min)
        Next
    Next
    
    getMatrice = Matrice

End Function
