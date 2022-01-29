Function NbOcc(ByVal S As String, ByVal Mot As String) As Long

    Dim i As Integer
    For i = 1 To Len(S) - Len(Mot)
        If StrComp(Mid(S, i, Len(Mot)), Mot, 0) = 0 Then
            NbOcc = NbOcc + 1
        End If
    Next
        
End Function
