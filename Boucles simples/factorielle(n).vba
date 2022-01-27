Function Fact(n As Integer) As Long 
'Renvoie la factorielle de n

    Dim produit As Long
    Dim i As Integer
    
    produit = 1
    For i = 1 To n
        produit = produit * i
        Next
     
    Fact = produit
End Function
