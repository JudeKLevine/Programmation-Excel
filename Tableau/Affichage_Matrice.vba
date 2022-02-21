Sub affiche_Matrice(TABLEAU() As Integer) ' PERMET D'AFFICHER LES ELEMENTS
                                          ' D'UNE MATRICE
    Dim i As Integer
    Dim j As Integer
    Dim col As String
    Dim Mat As String
    
    For i = LBound(TABLEAU, 1) To UBound(TABLEAU, 1)
        col = ""
        For j = LBound(TABLEAU, 1) To UBound(TABLEAU, 1)
            col = col & TABLEAU(i, j) & " "
        Next
        Mat = Mat & col & vbNewLine
    Next

    MsgBox Mat
    
End Sub
