Function sumEntier(ByVal n As Integer) As Integer ' renvoie la somme des n premiers entiers
' renvoie la somme de n premiers entiers

    Dim i As Integer
    Dim sum As Integer
    sum = 0
    
    'Boucle des entiers
    For i = 1 To n
        sum = sum + i
    Next
    sumEntier = sum
End Function
