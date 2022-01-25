Function sum_carreEntier(ByVal n As Integer) As Integer
'renvoie la somme des carres des n premiers entiers

    Dim i As Integer
    Dim sum As Integer
    sum = 0
    'Boucle des entiers
    For i = 1 To n
        sum = sum + i * i
    Next
    sum_carreEntier = sum
End Function
