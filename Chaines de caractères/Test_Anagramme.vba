Function Anagramme(ByVal chaine1 As String, chaine2 As String) As Boolean
    
    Dim i As Integer
    Dim count As Integer
    
    If Len(chaine1) <> Len(chaine2) Then
        MsgBox "NE SONT PAS ANAGRAMMES"
    Else
        For i = 1 To Len(chaine1)
            If InStr(chaine1, Mid(chaine2, i, 1)) > 0 Then
                count = count + 1
                chaine2 = Replace(chaine2, Mid(chaine2, i, 1), " ")
            End If
        Next
    End If
    
    If count = Len(chaine1) Then
        Anagramme = True
    End If
    
End Function
