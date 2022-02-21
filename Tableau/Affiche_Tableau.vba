Sub affiche(TABLEAU() As Variant) ' PERMET D'AFFICHER LES ELEMENTS
                                  ' D'UN TABLEAU
    Dim i As Integer
    Dim chaine As String
    
    For i = LBound(TABLEAU) To UBound(TABLEAU)
        chaine = chaine & TABLEAU(i) & vbNewLine
    Next
    MsgBox chaine
    
End Sub
