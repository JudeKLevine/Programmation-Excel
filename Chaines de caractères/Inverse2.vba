Function Myreverse(ByVal phrase As String) As String

    'renvoie: renverse une chaine de caractere

    If Len(phrase) Then
        Myreverse = Right(phrase, 1) & Myreverse(Left(phrase, Len(phrase) - 1))
    Else
        Myreverse = phrase
    End If
    
    End Function
