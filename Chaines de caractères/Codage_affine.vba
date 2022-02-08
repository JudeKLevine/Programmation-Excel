Function Codage_affine(ByVal Phrase As String, ByVal int1 As Integer, ByVal int2 As String) As String
 
        'code suivant le decalage (int2 + int1*ord) Mod 26
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
                dec = 65 + ((ord - 65) * int1 + int2) Mod 26
            End If
            chaine = chaine & Chr(dec)
        Next
        Codage_affine = chaine
        
    End Function
