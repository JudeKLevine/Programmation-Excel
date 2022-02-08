Function Remplace(ByVal phrase As String, ByVal mot1 As String, ByVal mot2 As String) As String

    Dim i As Integer
    Dim chaine As String
    
    chaine = ""
    For i = 1 To Len(phrase)
        If StrComp(Mid(phrase, i, Len(mot1)), mot1, 0) = 0 Then
            chaine = chaine & mot2
            i = i + Len(mot1) - 1
        Else
            chaine = chaine & Mid(phrase, i, 1)
        End If
    Next
    Remplace = chaine
End Function
