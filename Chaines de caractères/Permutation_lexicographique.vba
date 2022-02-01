Function fact(ByVal n As Integer) As Long

    Dim produit As Long
    Dim i As Integer
    
    produit = 1
    If n = 0 Then
        fact = produit
    Else
        For i = 1 To n
            produit = produit * i
        Next
    End If
    
    fact = produit
    
End Function

Function Lexico_permutation(ByVal nieme As Long) As String

     Dim chaine As String
     Dim a As Long
     Dim b As Long
     Dim c As Long
     Dim result As String
     
     c = nieme
     
     chaine = "0123456789"
     Longueur = Len(chaine)
     
     Do While Longueur > 1
     
        a = c \ fact(Longueur - 1)
        b = c Mod fact(Longueur - 1)
        c = b
        Longueur = Longueur - 1
        
        ' Teste sur le restede la division
        If b = 0 And a <> 0 Then
            result = result & Mid(chaine, a, 1)
            chaine = Replace(chaine, Mid(chaine, a, 1), "")
        ElseIf a = 0 And b = 0 Then
            result = result & Mid(chaine, a + 2, 1)
            chaine = Replace(chaine, Mid(chaine, a + 2, 1), "")
        Else
            result = result & Mid(chaine, a + 1, 1)
            chaine = Replace(chaine, Mid(chaine, a + 1, 1), "")
        End If

    Loop
    result = result & chaine
    
     Lexico_permutation = result
End Function

