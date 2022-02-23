Function Matrice_Hilbert(ByVal taille As Integer) As Double()

    'Matrice de Hilbert de taille taille
    Dim Matrice() As Double
    Dim i As Integer
    Dim j As Integer
    
    ReDim Matrice(taille - 1, taille - 1)
    
    For i = 0 To taille - 1
        For j = 0 To taille - 1
            Matrice(i, j) = 1 / (i + j + 1)
        Next
    Next
    
    Matrice_Hilbert = Matrice
    
End Function
