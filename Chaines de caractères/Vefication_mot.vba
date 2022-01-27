Function VerifMots(ByVal S As String, ByVal mot1 As String, ByVal mot2 As String) As String
    
    Dim Len1 As Integer
    Dim Len2 As Integer
    Dim i As Integer
    Dim test, rang1 As Integer
    Dim test2, rang2 As Integer
   
    Len1 = Len(mot1)
    Len2 = Len(mot2)
    test = 0
    test2 = 0
    
    'test pour le premier mot
    For i = 1 To Len(S) - Len1 + 1
        If StrComp(Mid(S, i, Len1 + i - 1), mot1, 0) = 0 Then
            test = test + 1
            rang1 = i
        End If
    Next
    If test > 0 Then
        VerifMots = mot1 & " " & "est present au rang" & " " & rang1
    Else
        VerifMots = mot1 & " " & " est absent"
    End If
    'test pour le deuxième mot
    For i = 1 To Len(S) - Len2 + 1
        If StrComp(Mid(S, i, Len2 + i - 1), mot2, 0) = 0 Then
            test2 = test2 + 1
            rang2 = i
        End If
        Debug.Print Mid(S, i, Len1 + i - 1)
    Next
    If test2 > 0 Then
        VerifMots = VerifMots & " " & mot2 & " " & "est present au rang" & " " & rang2
    Else
        VerifMots = VerifMots & " " & mot2 & " " & "est absent"
    End If
    
End Function
