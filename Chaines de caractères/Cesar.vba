Function Code_cesar(ByVal Phrase As String, ByVal entier As Integer) As String
    
        'les lettres doivent être en majiscule
        Dim pos As Integer
        Dim chaine As String
        Dim ord As Integer
        Dim dec As Integer
        
        chaine = ""
        
        For pos = 1 To Len(Phrase)
            ord = Asc(Mid(Phrase, pos, 1))
            
            If ord = 32 Or ord = 39 Or ord = 46 Or ord = 44 Or ord = 59 _
            Or ord = 45 Or ord = 63 Or ord = 33 Then
                dec = ord
            Else
                dec = 65 + (ord - 65 + entier) Mod 26
            End If
            chaine = chaine & Chr(dec)
        Next
        Code_cesar = chaine
        
    End Function
