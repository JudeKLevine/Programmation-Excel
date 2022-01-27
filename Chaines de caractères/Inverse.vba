Function Myreserve(ByVal S As String) As String
    
    Dim i As Integer
    
    For i = 1 To Len(S)
        Myreserve = Myreserve & Mid(S, Len(S) - i + 1, 1)
    Next

End Function
